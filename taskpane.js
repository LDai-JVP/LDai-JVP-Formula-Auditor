/* ============================================================
   Formula Auditor — taskpane.js
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

let refs = [];
let activeIdx = -1;

function openAuditor() {
  refreshAuditor();
}

async function onSelectionChanged() {
  await refreshAuditor();
}

async function refreshAuditor() {
  try {
    await Excel.run(async (ctx) => {
      const cell  = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load(["address", "formulas"]);
      sheet.load("name");
      await ctx.sync();

      const addr    = cell.address;
      const formula = cell.formulas[0][0];

      updateCellLabel(addr);

      if (typeof formula !== "string" || !formula.startsWith("=")) {
        showNoFormula(formula);
        return;
      }

      showFormula(formula);
      refs = parseRefs(formula, sheet.name);
      renderRefList();
      setActive(refs.length > 0 ? 0 : -1);
    });
  } catch (e) {
    console.error("Formula Auditor error:", e);
  }
}

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
    const key = sheetName + "!" + addr;
    if (!seen.has(key)) {
      seen.add(key);
      results.push({ addr, sheetName });
    }
  }
  return results;
}

function updateCellLabel(addr) {
  document.getElementById("cell-addr").textContent = addr;
}

function showNoFormula(val) {
  document.getElementById("formula-box").textContent =
    val === "" ? "(empty cell)" : String(val);
  document.getElementById("no-formula").style.display   = "block";
  document.getElementById("refs-label").style.display   = "none";
  document.getElementById("refs-list").style.display    = "none";
  document.getElementById("hint").style.display         = "none";
  refs = []; activeIdx = -1;
}

function showFormula(formula) {
  document.getElementById("formula-box").textContent = formula;
  document.getElementById("no-formula").style.display  = "none";
  document.getElementById("refs-label").style.display  = "";
  document.getElementById("refs-list").style.display   = "";
  document.getElementById("hint").style.display        = "";
}

function renderRefList() {
  const list = document.getElementById("refs-list");
  list.innerHTML = "";

  document.getElementById("ref-count").textContent =
    refs.length ? "(" + refs.length + ")" : "(none found)";

  if (refs.length === 0) {
    list.innerHTML = '<div style="padding:12px;color:#aaa;text-align:center;font-size:12px;">No external cell references found.</div>';
    return;
  }

  refs.forEach((ref, i) => {
    const item = document.createElement("div");
    item.className   = "ref-item";
    item.tabIndex    = 0;
    item.dataset.idx = i;
    item.innerHTML   =
      '<div class="ref-icon"><svg viewBox="0 0 10 10"><rect x="1" y="1" width="8" height="8" rx="1"/></svg></div>' +
      '<span class="ref-addr">' + ref.addr + '</span>' +
      '<span class="ref-sheet">' + ref.sheetName + '</span>';
    item.addEventListener("click",   () => activateAndNavigate(i));
    item.addEventListener("keydown", (e) => handleItemKey(e, i));
    list.appendChild(item);
  });

  document.onkeydown = handleGlobalKey;
}

function setActive(idx) {
  document.querySelectorAll(".ref-item").forEach(el => el.classList.remove("active"));
  activeIdx = idx;
  if (idx < 0 || idx >= refs.length) return;
  const el = document.querySelectorAll(".ref-item")[idx];
  if (el) { el.classList.add("active"); el.scrollIntoView({ block: "nearest" }); }
}

function handleGlobalKey(e) {
  if (!refs.length) return;
  if (e.key === "ArrowDown")  { e.preventDefault(); activateAndNavigate((activeIdx + 1) % refs.length); }
  else if (e.key === "ArrowUp") { e.preventDefault(); activateAndNavigate((activeIdx - 1 + refs.length) % refs.length); }
  else if (e.key === "Escape") { Office.context.ui.closeContainer(); }
}

function handleItemKey(e, idx) {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); activateAndNavigate(idx); }
}

function activateAndNavigate(idx) {
  setActive(idx);
  navigateTo(refs[idx]);
}

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