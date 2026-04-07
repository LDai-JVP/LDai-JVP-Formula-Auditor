
/* ============================================================
   Formula Auditor — taskpane.js  v3
   Behaviour:
   - Pane open + click new cell in Excel → auto-refreshes
   - Ctrl+Shift+M → always refreshes to active cell
   - Arrow Up/Down or mouse click in list → jumps Excel to
     that cell but keeps auditor locked on original cell
   - Enter on list item → drills into that cell, pushes
     current cell onto history stack, re-audits new cell
   - Backspace → pops history stack, navigates back and
     re-audits the previous cell
   - Esc → closes the pane
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
let activeIdx   = -1;
let locked      = false;   // true only while browsing ref list
let history     = [];      // stack of {addr, sheetName} for backspace nav
let currentAddr = null;    // full address of cell currently being audited

/* ── Entry points ──────────────────────────────────────── */

function openAuditor() {
  // Ctrl+Shift+M always unlocks and refreshes
  locked = false;
  refreshAuditor();
}

async function onSelectionChanged() {
  // Selection changes in the sheet are ignored —
  // the auditor only refreshes via Ctrl+Shift+M or Enter (drill in)
}

/* ── Core refresh ──────────────────────────────────────── */

async function refreshAuditor() {
  try {
    await Excel.run(async (ctx) => {
      const cell  = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load(["address", "formulas"]);
      sheet.load("name");
      await ctx.sync();

      currentAddr = cell.address;
      updateCellLabel(cell.address);
      updateBackButton();

      const formula = cell.formulas[0][0];

      if (typeof formula !== "string" || !formula.startsWith("=")) {
        showNoFormula(formula);
        return;
      }

      showFormula(formula);
      refs = parseRefs(formula, sheet.name);
      renderRefList();
      setActive(refs.length > 0 ? 0 : -1, false);
    });
  } catch (e) {
    console.error("Formula Auditor error:", e);
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
    if (!seen.has(key)) {
      seen.add(key);
      results.push({ addr, sheetName });
    }
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
  btn.style.display   = history.length > 0 ? "inline-flex" : "none";
  btn.title = history.length > 0
    ? `Back to ${history[history.length - 1].label}`
    : "";
}

function showNoFormula(val) {
  document.getElementById("formula-box").textContent =
    val === "" ? "(empty cell)" : String(val);
  document.getElementById("no-formula").style.display  = "block";
  document.getElementById("refs-label").style.display  = "none";
  document.getElementById("refs-list").style.display   = "none";
  document.getElementById("hint").style.display        = "none";
  refs = []; activeIdx = -1; locked = false;
}

function showFormula(formula) {
  document.getElementById("formula-box").textContent  = formula;
  document.getElementById("no-formula").style.display = "none";
  document.getElementById("refs-label").style.display = "";
  document.getElementById("refs-list").style.display  = "";
  document.getElementById("hint").style.display       = "";
}

function renderRefList() {
  const list = document.getElementById("refs-list");
  list.innerHTML = "";

  document.getElementById("ref-count").textContent =
    refs.length ? `(${refs.length})` : "(none found)";

  if (refs.length === 0) {
    list.innerHTML = `<div style="padding:12px;color:#aaa;text-align:center;font-size:12px;">
      No external cell references found.</div>`;
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

    // Mouse click → jump only, lock auditor
    item.addEventListener("click", () => {
      locked = true;
      setActive(i, true);
    });

    item.addEventListener("keydown", (e) => handleItemKey(e, i));
    list.appendChild(item);
  });

  document.onkeydown = handleGlobalKey;
}

/* ── Active row ─────────────────────────────────────────── */

function setActive(idx, navigate) {
  document.querySelectorAll(".ref-item").forEach(el => el.classList.remove("active"));
  activeIdx = idx;
  if (idx < 0 || idx >= refs.length) return;
  const el = document.querySelectorAll(".ref-item")[idx];
  if (el) { el.classList.add("active"); el.scrollIntoView({ block: "nearest" }); }
  if (navigate) navigateTo(refs[idx]);
}

/* ── Keyboard handlers ───────────────────────────────────── */

function handleGlobalKey(e) {
  if (!refs.length && e.key !== "Backspace" && e.key !== "Escape") return;

  if (e.key === "ArrowDown") {
    e.preventDefault();
    locked = true;
    setActive((activeIdx + 1) % refs.length, true);

  } else if (e.key === "ArrowUp") {
    e.preventDefault();
    locked = true;
    setActive((activeIdx - 1 + refs.length) % refs.length, true);

  } else if (e.key === "Enter") {
    e.preventDefault();
    drillInto(activeIdx);

  } else if (e.key === "Backspace") {
    e.preventDefault();
    goBack();

  } else if (e.key === "Escape") {
    Office.context.ui.closeContainer();
  }
}

function handleItemKey(e, idx) {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    drillInto(idx);
  }
}

/* ── Drill in: push history, navigate, re-audit ─────────── */

async function drillInto(idx) {
  if (idx < 0 || idx >= refs.length) return;

  // Push current cell onto history before navigating
  history.push({
    sheetName: refs[idx].sheetName, // placeholder — we capture actual below
    addr:      null,
    label:     currentAddr
  });

  // Capture the actual current address for back navigation
  try {
    await Excel.run(async (ctx) => {
      const cell  = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load("address");
      sheet.load("name");
      await ctx.sync();
      history[history.length - 1] = {
        addr:      cell.address.replace(/^.*!/, ""),
        sheetName: sheet.name,
        label:     cell.address
      };
    });
  } catch(e) {}

  locked = false;
  await navigateTo(refs[idx]);
  await refreshAuditor();
  updateBackButton();
}

/* ── Go back: pop history, navigate, re-audit ───────────── */

async function goBack() {
  if (history.length === 0) return;
  const prev = history.pop();
  locked = false;
  await navigateTo(prev);
  await refreshAuditor();
  updateBackButton();
}

/* ── Navigate Excel to a cell ───────────────────────────── */

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
