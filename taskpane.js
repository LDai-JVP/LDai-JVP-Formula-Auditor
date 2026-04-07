
/* ============================================================
   Formula Auditor — taskpane.js  v4
   Behaviour:
   - Ctrl+Shift+M → audits active cell, sets it as home cell
   - Arrow Down past last ref row → moves focus to home cell
   - Arrow Up from home cell → moves focus back to last ref row
   - Enter on a ref row → drills in, pushes history, re-audits
   - Enter on home cell → jumps back to home cell + re-audits
   - Backspace → pops history stack, navigates + re-audits
   - Click home cell → same as Enter on home cell
   - Esc → closes pane
   ============================================================ */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  Office.actions.associate("ShowFormulaAuditor", openAuditor);

  document.getElementById("close-btn").addEventListener("click", () => {
    Office.context.ui.closeContainer();
  });

  Excel.run(async (ctx) => {
    ctx.workbook.onSelectionChanged.add(onSelectionChanged);
    await ctx.sync();
  });

  refreshAuditor();
});

/* ── State ─────────────────────────────────────────────── */
let refs        = [];
let activeIdx   = -1;     // -1 = none, 0..n-1 = ref rows, n = home cell
let locked      = false;
let history     = [];
let currentAddr = null;
let homeCell    = null;   // { addr, sheetName, formula } — set on first audit

/* ── Entry points ──────────────────────────────────────── */

function openAuditor() {
  locked = false;
  history = [];
  homeCell = null;        // reset home on fresh Ctrl+Shift+M
  refreshAuditor(true);   // true = capture home
}

async function onSelectionChanged() {
  // ignored — auditor only updates via Ctrl+Shift+M or Enter
}

/* ── Core refresh ──────────────────────────────────────── */

async function refreshAuditor(captureHome) {
  try {
    await Excel.run(async (ctx) => {
      const cell  = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load(["address", "formulas"]);
      sheet.load("name");
      await ctx.sync();

      currentAddr = cell.address;
      const formula   = cell.formulas[0][0];
      const sheetName = sheet.name;
      const addr      = cell.address.replace(/^.*!/, "");

      // Capture home on first audit or Ctrl+Shift+M
      if (captureHome || !homeCell) {
        homeCell = {
          addr,
          sheetName,
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
      refs = parseRefs(formula, sheetName);
      renderRefList();
      setActive(refs.length > 0 ? 0 : -1, false);
    });
  } catch (e) {
    console.error("Formula Auditor error:", e);
  }
}

/* ── Refresh from a known ref (goBack / goHome) ─────────── */

async function refreshFromRef({ addr, sheetName }) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const cell  = sheet.getRange(addr);
      cell.load(["address", "formulas"]);
      await ctx.sync();

      currentAddr = cell.address;
      const formula = cell.formulas[0][0];

      updateCellLabel(cell.address);
      updateBackButton();
      renderHomeCell();

      if (typeof formula !== "string" || !formula.startsWith("=")) {
        showNoFormula(formula);
        return;
      }

      showFormula(formula);
      refs = parseRefs(formula, sheetName);
      renderRefList();
      setActive(refs.length > 0 ? 0 : -1, false);
    });
  } catch (e) {
    console.error("Formula Auditor refreshFromRef error:", e);
  }
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
    if (!seen.has(key)) { seen.add(key); results.push({ addr, sheetName }); }
  }
  return results;
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
  document.getElementById("formula-box").textContent =
    val === "" ? "(empty cell)" : String(val);
  document.getElementById("no-formula").style.display  = "block";
  document.getElementById("refs-label").style.display  = "none";
  document.getElementById("refs-list").style.display   = "none";
  refs = []; activeIdx = -1; locked = false;
  renderHomeCell();
  document.getElementById("hint").style.display = "";
}

function showFormula(formula) {
  document.getElementById("formula-box").textContent  = formula;
  document.getElementById("no-formula").style.display = "none";
  document.getElementById("refs-label").style.display = "";
  document.getElementById("refs-list").style.display  = "";
  document.getElementById("hint").style.display       = "";
}

/* ── Home cell rendering ────────────────────────────────── */

function renderHomeCell() {
  const container = document.getElementById("home-cell-container");
  if (!container || !homeCell) return;

  const isAtHome = history.length === 0;
  const shortFormula = homeCell.formula.length > 32
    ? homeCell.formula.substring(0, 32) + "…"
    : homeCell.formula;

  container.innerHTML = `
    <div id="home-cell-box"
      tabindex="0"
      style="
        margin: 0 12px 10px;
        background: ${isAtHome ? "#e8f5ee" : "#f0f0f0"};
        border: 1px solid ${isAtHome ? "#217346" : "#d0d0d0"};
        border-radius: 4px;
        padding: 6px 10px;
        display: flex;
        align-items: center;
        gap: 8px;
        cursor: ${isAtHome ? "default" : "pointer"};
      ">
      <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style="flex-shrink:0;">
        <path d="M7 1L1 6.5V13h4V9h4v4h4V6.5L7 1z" fill="${isAtHome ? "#217346" : "#888"}"/>
      </svg>
      <div style="display:flex;flex-direction:column;gap:1px;flex:1;overflow:hidden;">
        <span style="font-size:9px;color:${isAtHome ? "#217346" : "#888"};">
          ${isAtHome ? "Home cell — you are here" : "Home cell — click or ↓ to return"}
        </span>
        <span style="font-family:'Consolas',monospace;font-size:10px;font-weight:600;color:${isAtHome ? "#217346" : "#444"};">
          ${homeCell.sheetName}!${homeCell.addr}
        </span>
        <span style="font-family:'Consolas',monospace;font-size:9px;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
          ${shortFormula}
        </span>
      </div>
    </div>`;

  const box = document.getElementById("home-cell-box");
  if (!isAtHome) {
    box.addEventListener("click", () => goHome());
    box.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); goHome(); }
      if (e.key === "ArrowUp") { e.preventDefault(); setActive(refs.length - 1, false); document.querySelectorAll(".ref-item")[refs.length - 1]?.focus(); }
    });
  }
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
    item.innerHTML   = `
      <div class="ref-icon">
        <svg viewBox="0 0 10 10"><rect x="1" y="1" width="8" height="8" rx="1"/></svg>
      </div>
      <span class="ref-addr">${ref.addr}</span>
      <span class="ref-sheet">${ref.sheetName}</span>`;

    item.addEventListener("click", () => { locked = true; setActive(i, true); });
    item.addEventListener("keydown", (e) => handleItemKey(e, i));
    list.appendChild(item);
  });

  document.onkeydown = handleGlobalKey;
}

/* ── Active row ─────────────────────────────────────────── */

function setActive(idx, navigate) {
  document.querySelectorAll(".ref-item").forEach(el => el.classList.remove("active"));
  const homeBox = document.getElementById("home-cell-box");
  if (homeBox) homeBox.style.outline = "none";

  activeIdx = idx;

  // Home cell selected (idx === refs.length)
  if (idx === refs.length) {
    if (homeBox && history.length > 0) {
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
    // If at last ref or beyond, move to home cell
    const next = activeIdx + 1;
    if (next >= refs.length) {
      setActive(refs.length, history.length > 0); // navigate home only if drilled in
    } else {
      setActive(next, true);
    }
  } else if (e.key === "ArrowUp") {
    e.preventDefault();
    locked = true;
    if (activeIdx === refs.length) {
      // From home cell, go back to last ref
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

  // Capture current cell for back stack
  let fromAddr = currentAddr ? currentAddr.replace(/^.*!/, "") : null;
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
  history = [];           // clear entire stack
  locked  = false;
  await navigateTo(homeCell);
  await refreshFromRef(homeCell);
  updateBackButton();
  renderHomeCell();       // re-render as "you are here"
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
