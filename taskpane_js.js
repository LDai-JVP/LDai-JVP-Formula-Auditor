/* ============================================================
   Formula Auditor — taskpane.js  v5
   New in v5:
   - [2] Cell value preview in each ref row
   - [4] Named range resolution
   - [1] Error cell flagging (red indicator)
   - [9] Formula syntax highlighting
   ============================================================ */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  // Must be called before any user interaction
  Office.actions.associate("ShowFormulaAuditor", openAuditor);

  document.getElementById("close-btn").addEventListener("click", () => {
    Office.context.ui.closeContainer();
  });

  Excel.run(async (ctx) => {
    ctx.workbook.onSelectionChanged.add(onSelectionChanged);
    await ctx.sync();
  });

  document.addEventListener("keydown", handleGlobalKey);

  // Initial load
  setTimeout(() => {
    refreshAuditor(true).then(() => focusFirstItem());
  }, 75);
});

/* ── State ─────────────────────────────────────────────── */
let refs        = [];   // [{addr, sheetName, value, isError, isNamed, resolvedFrom}]
let activeIdx   = -1;
let locked      = false;
let history     = [];
let currentAddr = null;
let homeCell    = null;

/* ── Entry points ──────────────────────────────────────── */

function openAuditor() {
  locked   = false;
  history  = [];
  homeCell = null;
  // Small delay ensures Excel has registered the active cell
  // before we attempt to read it, particularly when focus was in the pane
  setTimeout(() => {
    refreshAuditor(true).then(() => focusFirstItem());
  }, 75);
}

async function onSelectionChanged() {}

/* ── Core refresh ──────────────────────────────────────── */

async function refreshAuditor(captureHome) {
  try {
    await Excel.run(async (ctx) => {
      const cell  = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load(["address", "formulas"]);
      sheet.load("name");
      await ctx.sync();

      currentAddr       = cell.address;
      const formula     = cell.formulas[0][0];
      const sheetName   = sheet.name;
      const addr        = cell.address.replace(/^.*!/, "");

      if (captureHome || !homeCell) {
        homeCell = {
          addr, sheetName,
          formula: typeof formula === "string" && formula.startsWith("=")
            ? formula : String(formula)
        };
      }

      updateCellLabel(cell.address);
      updateBackButton();
      renderHomeCell();

      if (typeof formula !== "string" || !formula.startsWith("=")) {
        showNoFormula(formula);
        return;
      }

      showFormula(formula);
      refs = await resolveRefs(ctx, formula, sheetName);
      renderRefList();
      setActive(refs.length > 0 ? 0 : -1, false);
      focusFirstItem();
    });
  } catch (e) {
    console.error("Formula Auditor error:", e);
  }
}

async function refreshFromRef({ addr, sheetName }) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const cell  = sheet.getRange(addr);
      cell.load(["address", "formulas"]);
      await ctx.sync();

      currentAddr   = cell.address;
      const formula = cell.formulas[0][0];

      updateCellLabel(cell.address);
      updateBackButton();
      renderHomeCell();

      if (typeof formula !== "string" || !formula.startsWith("=")) {
        showNoFormula(formula);
        return;
      }

      showFormula(formula);
      refs = await resolveRefs(ctx, formula, sheetName);
      renderRefList();
      setActive(refs.length > 0 ? 0 : -1, false);
      focusFirstItem();
    });
  } catch (e) {
    console.error("Formula Auditor refreshFromRef error:", e);
  }
}

/* ── Ref resolution (refs + named ranges + values + errors) ─ */

async function resolveRefs(ctx, formula, activeSheet) {
  const rawRefs   = parseRefs(formula, activeSheet);
  const namedRefs = await resolveNamedRanges(ctx, formula, activeSheet);

  // Merge — named ranges come after direct refs, deduplicated
  const seen    = new Set(rawRefs.map(r => `${r.sheetName}!${r.addr}`));
  const allRefs = [...rawRefs];
  for (const nr of namedRefs) {
    const key = `${nr.sheetName}!${nr.addr}`;
    if (!seen.has(key)) { seen.add(key); allRefs.push(nr); }
  }

  // Load values for each ref
  for (const ref of allRefs) {
    try {
      const ws   = ctx.workbook.worksheets.getItem(ref.sheetName);
      const rng  = ws.getRange(ref.addr);
      rng.load(["values", "valueTypes"]);
      await ctx.sync();

      const val     = rng.values[0][0];
      const valType = rng.valueTypes[0][0];

      ref.isError = valType === "Error";
      ref.value   = formatValue(val, valType);
    } catch(e) {
      ref.value   = "—";
      ref.isError = false;
    }
  }

  return allRefs;
}

async function resolveNamedRanges(ctx, formula, activeSheet) {
  const results = [];
  try {
    const names = ctx.workbook.names;
    names.load("items");
    await ctx.sync();

    for (const namedItem of names.items) {
      namedItem.load(["name", "type"]);
    }
    await ctx.sync();

    // Check each named item to see if it appears in the formula
    for (const namedItem of names.items) {
      if (namedItem.type !== "Range") continue;
      const token = namedItem.name;
      // Check the formula contains this name as a whole word
      const regex = new RegExp(`(?<![A-Za-z0-9_])${escapeRegex(token)}(?![A-Za-z0-9_])`, "i");
      if (!regex.test(formula)) continue;

      try {
        const range = namedItem.getRange();
        range.load(["address", "worksheet"]);
        await ctx.sync();

        // address comes back as "SheetName!A1:B2"
        const fullAddr  = range.address;
        const bangIdx   = fullAddr.indexOf("!");
        const sheetName = bangIdx >= 0 ? fullAddr.substring(0, bangIdx) : activeSheet;
        const addr      = bangIdx >= 0 ? fullAddr.substring(bangIdx + 1) : fullAddr;

        results.push({
          addr,
          sheetName,
          isNamed:      true,
          resolvedFrom: token,
          value:        "",
          isError:      false,
        });
      } catch(e) {}
    }
  } catch(e) {}
  return results;
}

function formatValue(val, valType) {
  if (valType === "Error")   return String(val);
  if (valType === "Empty")   return "(empty)";
  if (valType === "Boolean") return val ? "TRUE" : "FALSE";
  if (typeof val === "number") {
    if (Math.abs(val) >= 1e6)  return val.toLocaleString();
    if (Number.isInteger(val)) return String(val);
    return parseFloat(val.toFixed(4)).toString();
  }
  if (typeof val === "string") {
    return val.length > 18 ? val.substring(0, 18) + "…" : val;
  }
  return String(val);
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/* ── Formula reference parser ───────────────────────────── */

function parseRefs(formula, activeSheet) {
  const results = [];
  const seen    = new Set();
  const pattern = /(?:'([^']+)'|([A-Za-z0-9_]+))?!?(\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?)/g;
  let m;
  while ((m = pattern.exec(formula)) !== null) {
    const sheetName = m[1] || m[2] || activeSheet;
    const addr      = m[3];
    const before    = formula[m.index - 1];
    if (before && /[A-Za-z]/.test(before)) continue;
    const key = `${sheetName}!${addr}`;
    if (!seen.has(key)) {
      seen.add(key);
      results.push({ addr, sheetName, isNamed: false, resolvedFrom: null, value: "", isError: false });
    }
  }
  return results;
}

/* ── Formula syntax highlighting ────────────────────────── */

function highlightFormula(formula) {
  // Tokenise and wrap in spans
  let result  = "";
  let i       = 0;
  const len   = formula.length;

  while (i < len) {
    // String literal
    if (formula[i] === '"') {
      let j = i + 1;
      while (j < len && formula[j] !== '"') j++;
      result += `<span class="f-str">${esc(formula.substring(i, j + 1))}</span>`;
      i = j + 1;
      continue;
    }
    // Sheet-qualified ref e.g. Sheet2!B4 or 'My Sheet'!B4
    const sheetRef = /^(?:'[^']+'|[A-Za-z0-9_]+)!(\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?)/;
    const srm = formula.substring(i).match(sheetRef);
    if (srm) {
      result += `<span class="f-ref">${esc(srm[0])}</span>`;
      i += srm[0].length;
      continue;
    }
    // Plain cell ref e.g. B4, $B$4, A1:C3
    const cellRef = /^(\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?)/;
    const crm = formula.substring(i).match(cellRef);
    if (crm && (i === 0 || !/[A-Za-z]/.test(formula[i - 1]))) {
      result += `<span class="f-ref">${esc(crm[0])}</span>`;
      i += crm[0].length;
      continue;
    }
    // Function name e.g. SUM(
    const funcRef = /^([A-Za-z_][A-Za-z0-9_.]*)\(/;
    const frm = formula.substring(i).match(funcRef);
    if (frm) {
      result += `<span class="f-func">${esc(frm[1])}</span>(`;
      i += frm[0].length;
      continue;
    }
    // Number
    const numRef = /^[0-9]+(\.[0-9]+)?/;
    const nrm = formula.substring(i).match(numRef);
    if (nrm) {
      result += `<span class="f-num">${esc(nrm[0])}</span>`;
      i += nrm[0].length;
      continue;
    }
    // Operator / punctuation
    if ("+-*/^&=<>(),;!".includes(formula[i])) {
      result += `<span class="f-op">${esc(formula[i])}</span>`;
      i++; continue;
    }
    result += esc(formula[i]);
    i++;
  }
  return result;
}

function esc(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

/* ── UI helpers ─────────────────────────────────────────── */

function updateCellLabel(addr) {
  document.getElementById("cell-addr").textContent = addr;
}

function updateBackButton() {
  const btn = document.getElementById("back-btn");
  if (!btn) return;
  btn.style.display = history.length > 0 ? "inline-flex" : "none";
}

function showNoFormula(val) {
  document.getElementById("formula-box").innerHTML =
    `<span style="color:#999;">${val === "" ? "(empty cell)" : esc(String(val))}</span>`;
  document.getElementById("no-formula").style.display  = "block";
  document.getElementById("refs-label").style.display  = "none";
  document.getElementById("refs-list").style.display   = "none";
  refs = []; activeIdx = -1; locked = false;
  renderHomeCell();
  document.getElementById("hint").style.display = "";
}

function showFormula(formula) {
  document.getElementById("formula-box").innerHTML = highlightFormula(formula);
  document.getElementById("no-formula").style.display = "none";
  document.getElementById("refs-label").style.display = "";
  document.getElementById("refs-list").style.display  = "";
  document.getElementById("hint").style.display       = "";
}

function focusFirstItem() {
  const first = document.querySelector(".ref-item");
  if (first) { first.focus(); }
  else {
    const home = document.getElementById("home-cell-box");
    if (home) home.focus();
  }
}

/* ── Home cell rendering ────────────────────────────────── */

function renderHomeCell() {
  const container = document.getElementById("home-cell-container");
  if (!container || !homeCell) return;

  const isAtHome     = history.length === 0;
  const shortFormula = homeCell.formula.length > 32
    ? homeCell.formula.substring(0, 32) + "…" : homeCell.formula;
  const label = isAtHome
    ? "Home cell — return to root"
    : "Home cell — return to initial formula";

  container.innerHTML = `
    <div id="home-cell-box" tabindex="0" style="
        margin: 0 12px 10px;
        background: ${isAtHome ? "#e8f5ee" : "#f0f0f0"};
        border: 1px solid ${isAtHome ? "#217346" : "#d0d0d0"};
        border-radius: 4px; padding: 6px 10px;
        display: flex; align-items: center; gap: 8px; cursor: pointer;">
      <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style="flex-shrink:0;">
        <path d="M7 1L1 6.5V13h4V9h4v4h4V6.5L7 1z" fill="${isAtHome ? "#217346" : "#888"}"/>
      </svg>
      <div style="display:flex;flex-direction:column;gap:1px;flex:1;overflow:hidden;">
        <span style="font-size:9px;color:${isAtHome ? "#217346" : "#888"};">${label}</span>
        <span style="font-family:'Consolas',monospace;font-size:10px;font-weight:600;color:${isAtHome ? "#217346" : "#444"};">
          ${homeCell.sheetName}!${homeCell.addr}
        </span>
        <span style="font-family:'Consolas',monospace;font-size:9px;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
          ${esc(shortFormula)}
        </span>
      </div>
    </div>`;

  const box = document.getElementById("home-cell-box");
  box.addEventListener("click", () => goHome());
  box.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); goHome(); }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      if (refs.length > 0) {
        setActive(refs.length - 1, false);
        document.querySelectorAll(".ref-item")[refs.length - 1]?.focus();
      }
    }
  });
}

/* ── Ref list rendering ─────────────────────────────────── */

function renderRefList() {
  const list = document.getElementById("refs-list");
  list.innerHTML = "";

  document.getElementById("ref-count").textContent =
    refs.length ? `(${refs.length})` : "(none found)";

  if (refs.length === 0) {
    list.innerHTML = `<div style="padding:12px;color:#aaa;text-align:center;font-size:12px;">No external cell references found.</div>`;
    return;
  }

  refs.forEach((ref, i) => {
    const item = document.createElement("div");
    item.className   = "ref-item";
    item.tabIndex    = 0;
    item.dataset.idx = i;

    // Dot colour: red for errors, amber for named, green for normal
    const dotColor = ref.isError ? "#c0392b" : ref.isNamed ? "#b7770d" : "#217346";

    // Value display
    const valueHtml = ref.isError
      ? `<span class="ref-value ref-error">${esc(ref.value)}</span>`
      : `<span class="ref-value">${esc(ref.value)}</span>`;

    // Named range badge
    const namedBadge = ref.isNamed
      ? `<span class="ref-named-badge">${esc(ref.resolvedFrom)}</span>`
      : "";

    item.innerHTML = `
      <div class="ref-icon" style="background:${dotColor};">
        <svg viewBox="0 0 10 10"><rect x="1" y="1" width="8" height="8" rx="1"/></svg>
      </div>
      <div class="ref-main">
        <div class="ref-top-row">
          <span class="ref-addr">${esc(ref.addr)}</span>
          ${namedBadge}
          <span class="ref-sheet">${esc(ref.sheetName)}</span>
        </div>
        ${valueHtml}
      </div>`;

    item.addEventListener("click", () => { locked = true; setActive(i, true); });
    item.addEventListener("keydown", (e) => handleItemKey(e, i));
    list.appendChild(item);
  });
}

/* ── Active row ─────────────────────────────────────────── */

function setActive(idx, navigate) {
  document.querySelectorAll(".ref-item").forEach(el => el.classList.remove("active"));
  const homeBox = document.getElementById("home-cell-box");
  if (homeBox) homeBox.style.outline = "none";

  activeIdx = idx;

  if (idx === refs.length) {
    if (homeBox) {
      homeBox.style.outline = "2px solid #217346";
      homeBox.scrollIntoView({ block: "nearest" });
      if (navigate) goHome();
    }
    return;
  }

  if (idx < 0 || idx >= refs.length) return;
  const el = document.querySelectorAll(".ref-item")[idx];
  if (el) { el.classList.add("active"); el.scrollIntoView({ block: "nearest" }); }
  if (navigate) navigateTo(refs[idx]);
}

/* ── Keyboard handlers ───────────────────────────────────── */

function handleGlobalKey(e) {
  if (e.key === "ArrowDown") {
    e.preventDefault();
    locked = true;
    const next = activeIdx + 1;
    if (next >= refs.length) {
      setActive(refs.length, history.length === 0);
    } else {
      setActive(next, true);
    }
  } else if (e.key === "ArrowUp") {
    e.preventDefault();
    locked = true;
    if (activeIdx === refs.length) {
      setActive(refs.length - 1, true);
    } else {
      setActive((activeIdx - 1 + refs.length) % refs.length, true);
    }
  } else if (e.key === "Enter") {
    e.preventDefault();
    if (activeIdx === refs.length) { goHome(); }
    else { drillInto(activeIdx); }
  } else if (e.key === "Backspace") {
    e.preventDefault();
    goBack();
  } else if (e.key === "Escape") {
    Office.context.ui.closeContainer();
  }
}

function handleItemKey(e, idx) {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); drillInto(idx); }
}

/* ── Drill in ───────────────────────────────────────────── */

async function drillInto(idx) {
  if (idx < 0 || idx >= refs.length) return;

  let fromAddr  = currentAddr ? currentAddr.replace(/^.*!/, "") : null;
  let fromSheet = null;
  try {
    await Excel.run(async (ctx) => {
      const cell  = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load("address"); sheet.load("name");
      await ctx.sync();
      fromAddr  = cell.address.replace(/^.*!/, "");
      fromSheet = sheet.name;
    });
  } catch(e) {}

  if (fromAddr && fromSheet) {
    history.push({ addr: fromAddr, sheetName: fromSheet, label: currentAddr });
  }

  locked = false;
  await navigateTo(refs[idx]);
  await refreshFromRef(refs[idx]);
  updateBackButton();
}

/* ── Go back ─────────────────────────────────────────────── */

async function goBack() {
  if (history.length === 0) return;
  const prev = history.pop();
  locked = false;
  await navigateTo(prev);
  await refreshFromRef(prev);
  updateBackButton();
}

/* ── Go home ─────────────────────────────────────────────── */

async function goHome() {
  if (!homeCell) return;
  history = [];
  locked  = false;
  await navigateTo(homeCell);
  await refreshFromRef(homeCell);
  updateBackButton();
  renderHomeCell();
}

/* ── Navigate Excel ─────────────────────────────────────── */

async function navigateTo({ addr, sheetName }) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(addr);
      sheet.activate();
      range.select();
      await ctx.sync();
    });
  } catch (e) {
    console.warn("Could not navigate to", sheetName, addr, e);
  }
}
