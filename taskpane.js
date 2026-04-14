/* ============================================================
   Formula Auditor — taskpane.js  v4 final
   Behaviour:
   - Ctrl+Shift+M → audits active cell, sets it as home cell
   - Arrow Down past last ref → highlights home cell
       · At root: navigates immediately (no Enter needed)
       · In sub-formula: highlights only, Enter confirms
   - Arrow Up from home cell → moves back to last ref row
   - Enter on a ref row → drills in, pushes history, re-audits
   - Enter on home cell (sub-formula only) → goes home
   - Click home cell → always goes home immediately
   - Backspace → pops history stack, navigates + re-audits
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

  // Global keydown registered once at load — handles Esc from anywhere
  document.addEventListener("keydown", handleGlobalKey);

  refreshAuditor(true).then(() => focusFirstItem());
});

function focusFirstItem() {
  const first = document.querySelector(".ref-item");
  if (first) {
    first.focus();
  } else {
    const home = document.getElementById("home-cell-box");
    if (home) home.focus();
  }
}

/* ── State ─────────────────────────────────────────────── */
let refs        = [];
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
  refreshAuditor(true).then(() => focusFirstItem());
}

async function onSelectionChanged() {
  // Auditor only updates via Ctrl+Shift+M or Enter
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
      focusFirstItem();
    });
  } catch (e) {
    console.error("Formula Auditor error:", e);
  }
}

/* ── Refresh from a known ref ───────────────────────────── */

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
      focusFirstItem();
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

  const isAtHome    = history.length === 0;
  const shortFormula = homeCell.formula.length > 32
    ? homeCell.formula.substring(0, 32) + "…"
    : homeCell.formula;

  // Label changes based on depth
  const label = isAtHome
    ? "Home cell — return to root"
    : "Home cell — return to initial formula";

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
        cursor: pointer;
      ">
      <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style="flex-shrink:0;">
        <path d="M7 1L1 6.5V13h4V9h4v4h4V6.5L7 1z" fill="${isAtHome ? "#217346" : "#888"}"/>
      </svg>
      <div style="display:flex;flex-direction:column;gap:1px;flex:1;overflow:hidden;">
        <span style="font-size:9px;color:${isAtHome ? "#217346" : "#888"};">${label}</span>
        <span style="font-family:'Consolas',monospace;font-size:10px;font-weight:600;color:${isAtHome ? "#217346" : "#444"};">
          ${homeCell.sheetName}!${homeCell.addr}
        </span>
        <span style="font-family:'Consolas',monospace;font-size:9px;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
          ${shortFormula}
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
}

/* ── Active row ─────────────────────────────────────────── */

function setActive(idx, navigate) {
  document.querySelectorAll(".ref-item").forEach(el => el.classList.remove("active"));
  const homeBox = document.getElementById("home-cell-box");
  if (homeBox) homeBox.style.outline = "none";

  activeIdx = idx;

  // Home cell (idx === refs.length)
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
      // At root → navigate immediately. In sub-formula → highlight only.
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
    if (activeIdx === refs.length) {
      // Enter on home cell — always goes home
      goHome();
    } else {
      drillInto(activeIdx);
    }

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
