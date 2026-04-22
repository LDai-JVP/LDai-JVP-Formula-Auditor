/* ============================================================
   Formula Auditor — taskpane.js  v6
   Complete clean build with all features:
   - Value preview, named ranges, error flags, syntax highlight
   - Home cell, back stack, drill in/out
   - Ctrl+Shift+M re-audit, Esc to close, auto-focus
   ============================================================ */

Office.onReady(function(info) {
  if (info.host !== Office.HostType.Excel) return;

  try { Office.actions.associate("ShowFormulaAuditor", openAuditor); } catch(e) {}

  document.getElementById("close-btn").addEventListener("click", function() {
    Office.context.ui.closeContainer();
  });

  Excel.run(function(ctx) {
    ctx.workbook.onSelectionChanged.add(onSelectionChanged);
    return ctx.sync();
  });

  document.addEventListener("keydown", function(e) {
    if (e.ctrlKey && e.shiftKey && (e.key === "M" || e.key === "m")) {
      e.preventDefault();
      openAuditor();
      return;
    }
    handleGlobalKey(e);
  });

  setTimeout(function() {
    refreshAuditor(true).then(function() { focusFirstItem(); });
  }, 75);
});

/* ── State ─────────────────────────────────────────────── */
var refs        = [];
var activeIdx   = -1;
var locked      = false;
var history     = [];
var currentAddr = null;
var homeCell    = null;

/* ── Entry points ──────────────────────────────────────── */

function openAuditor() {
  locked   = false;
  history  = [];
  homeCell = null;
  setTimeout(function() {
    refreshAuditor(true).then(function() { focusFirstItem(); });
  }, 75);
}

function onSelectionChanged() {}

/* ── Focus helper ──────────────────────────────────────── */

function focusFirstItem() {
  var first = document.querySelector(".ref-item");
  if (first) { first.focus(); }
  else {
    var home = document.getElementById("home-cell-box");
    if (home) home.focus();
  }
}

/* ── Core refresh ──────────────────────────────────────── */

function refreshAuditor(captureHome) {
  return Excel.run(function(ctx) {
    var cell  = ctx.workbook.getActiveCell();
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    cell.load(["address", "formulas"]);
    sheet.load("name");
    return ctx.sync().then(function() {
      currentAddr       = cell.address;
      var formula       = cell.formulas[0][0];
      var sheetName     = sheet.name;
      var addr          = cell.address.replace(/^.*!/, "");

      if (captureHome || !homeCell) {
        homeCell = {
          addr: addr,
          sheetName: sheetName,
          formula: (typeof formula === "string" && formula.startsWith("="))
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
      return resolveRefs(formula, sheetName).then(function(resolved) {
        refs = resolved;
        renderRefList();
        setActive(refs.length > 0 ? 0 : -1, false);
        focusFirstItem();
      });
    });
  }).catch(function(e) {
    console.error("Formula Auditor error:", e);
  });
}

function refreshFromRef(ref) {
  return Excel.run(function(ctx) {
    var sheet = ctx.workbook.worksheets.getItem(ref.sheetName);
    var cell  = sheet.getRange(ref.addr);
    cell.load(["address", "formulas"]);
    return ctx.sync().then(function() {
      currentAddr   = cell.address;
      var formula   = cell.formulas[0][0];

      updateCellLabel(cell.address);
      updateBackButton();
      renderHomeCell();

      if (typeof formula !== "string" || !formula.startsWith("=")) {
        showNoFormula(formula);
        return;
      }

      showFormula(formula);
      return resolveRefs(formula, ref.sheetName).then(function(resolved) {
        refs = resolved;
        renderRefList();
        setActive(refs.length > 0 ? 0 : -1, false);
        focusFirstItem();
      });
    });
  }).catch(function(e) {
    console.error("Formula Auditor refreshFromRef error:", e);
  });
}

/* ── Ref resolution (values + named ranges + errors) ───── */

function resolveRefs(formula, activeSheet) {
  var rawRefs = parseRefs(formula, activeSheet);

  return Excel.run(function(ctx) {
    // Resolve named ranges
    var names = ctx.workbook.names;
    names.load("items");
    return ctx.sync().then(function() {
      var namedPromises = [];
      for (var n = 0; n < names.items.length; n++) {
        names.items[n].load(["name", "type"]);
      }
      return ctx.sync();
    }).then(function() {
      var namedRefs = [];
      var rangeLoads = [];

      for (var n = 0; n < names.items.length; n++) {
        var item = names.items[n];
        if (item.type !== "Range") continue;
        var regex = new RegExp("(?<![A-Za-z0-9_])" + escapeRegex(item.name) + "(?![A-Za-z0-9_])", "i");
        if (!regex.test(formula)) continue;
        try {
          var range = item.getRange();
          range.load(["address"]);
          rangeLoads.push({ range: range, name: item.name });
        } catch(e) {}
      }

      return ctx.sync().then(function() {
        var seen = {};
        for (var i = 0; i < rawRefs.length; i++) {
          seen[rawRefs[i].sheetName + "!" + rawRefs[i].addr] = true;
        }

        for (var r = 0; r < rangeLoads.length; r++) {
          var fullAddr = rangeLoads[r].range.address;
          var bangIdx  = fullAddr.indexOf("!");
          var sn = bangIdx >= 0 ? fullAddr.substring(0, bangIdx) : activeSheet;
          var ad = bangIdx >= 0 ? fullAddr.substring(bangIdx + 1) : fullAddr;
          var key = sn + "!" + ad;
          if (!seen[key]) {
            seen[key] = true;
            rawRefs.push({
              addr: ad, sheetName: sn, isNamed: true,
              resolvedFrom: rangeLoads[r].name, value: "", isError: false
            });
          }
        }

        // Now load values for all refs
        var valueLoads = [];
        for (var v = 0; v < rawRefs.length; v++) {
          try {
            var ws  = ctx.workbook.worksheets.getItem(rawRefs[v].sheetName);
            var rng = ws.getRange(rawRefs[v].addr);
            rng.load(["values", "valueTypes"]);
            valueLoads.push({ ref: rawRefs[v], range: rng });
          } catch(e) {
            rawRefs[v].value = "—";
            rawRefs[v].isError = false;
          }
        }

        return ctx.sync().then(function() {
          for (var vl = 0; vl < valueLoads.length; vl++) {
            try {
              var val     = valueLoads[vl].range.values[0][0];
              var valType = valueLoads[vl].range.valueTypes[0][0];
              valueLoads[vl].ref.isError = (valType === "Error");
              valueLoads[vl].ref.value   = formatValue(val, valType);
            } catch(e) {
              valueLoads[vl].ref.value   = "—";
              valueLoads[vl].ref.isError = false;
            }
          }
          return rawRefs;
        });
      });
    });
  }).catch(function(e) {
    console.error("resolveRefs error:", e);
    return rawRefs;
  });
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
  var results = [];
  var seen    = {};
  var pattern = /(?:'([^']+)'|([A-Za-z0-9_]+))?!?(\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?)/g;
  var m;
  while ((m = pattern.exec(formula)) !== null) {
    var sheetName = m[1] || m[2] || activeSheet;
    var addr      = m[3];
    var before    = formula[m.index - 1];
    if (before && /[A-Za-z]/.test(before)) continue;
    var key = sheetName + "!" + addr;
    if (!seen[key]) {
      seen[key] = true;
      results.push({ addr: addr, sheetName: sheetName, isNamed: false, resolvedFrom: null, value: "", isError: false });
    }
  }
  return results;
}

/* ── Formula syntax highlighting ────────────────────────── */

function highlightFormula(formula) {
  var result = "";
  var i = 0;
  var len = formula.length;

  while (i < len) {
    if (formula[i] === '"') {
      var j = i + 1;
      while (j < len && formula[j] !== '"') j++;
      result += '<span class="f-str">' + esc(formula.substring(i, j + 1)) + '</span>';
      i = j + 1;
      continue;
    }
    var rest = formula.substring(i);
    var srm = rest.match(/^(?:'[^']+'|[A-Za-z0-9_]+)!(\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?)/);
    if (srm) {
      result += '<span class="f-ref">' + esc(srm[0]) + '</span>';
      i += srm[0].length;
      continue;
    }
    var crm = rest.match(/^(\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?)/);
    if (crm && (i === 0 || !/[A-Za-z]/.test(formula[i - 1]))) {
      result += '<span class="f-ref">' + esc(crm[0]) + '</span>';
      i += crm[0].length;
      continue;
    }
    var frm = rest.match(/^([A-Za-z_][A-Za-z0-9_.]*)\(/);
    if (frm) {
      result += '<span class="f-func">' + esc(frm[1]) + '</span>(';
      i += frm[0].length;
      continue;
    }
    var nrm = rest.match(/^[0-9]+(\.[0-9]+)?/);
    if (nrm) {
      result += '<span class="f-num">' + esc(nrm[0]) + '</span>';
      i += nrm[0].length;
      continue;
    }
    if ("+-*/^&=<>(),;!".indexOf(formula[i]) >= 0) {
      result += '<span class="f-op">' + esc(formula[i]) + '</span>';
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
  var btn = document.getElementById("back-btn");
  if (!btn) return;
  btn.style.display = history.length > 0 ? "inline-flex" : "none";
}

function showNoFormula(val) {
  document.getElementById("formula-box").innerHTML =
    '<span style="color:#999;">' + (val === "" ? "(empty cell)" : esc(String(val))) + '</span>';
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

/* ── Home cell rendering ────────────────────────────────── */

function renderHomeCell() {
  var container = document.getElementById("home-cell-container");
  if (!container || !homeCell) return;

  var isAtHome     = history.length === 0;
  var shortFormula = homeCell.formula.length > 32
    ? homeCell.formula.substring(0, 32) + "…" : homeCell.formula;
  var label = isAtHome
    ? "Home cell — return to root"
    : "Home cell — return to initial formula";

  container.innerHTML =
    '<div id="home-cell-box" tabindex="0" style="' +
      'margin: 0 12px 10px;' +
      'background:' + (isAtHome ? '#e8f5ee' : '#f0f0f0') + ';' +
      'border: 1px solid ' + (isAtHome ? '#217346' : '#d0d0d0') + ';' +
      'border-radius: 4px; padding: 6px 10px;' +
      'display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
      '<svg width="13" height="13" viewBox="0 0 14 14" fill="none" style="flex-shrink:0;">' +
        '<path d="M7 1L1 6.5V13h4V9h4v4h4V6.5L7 1z" fill="' + (isAtHome ? '#217346' : '#888') + '"/>' +
      '</svg>' +
      '<div style="display:flex;flex-direction:column;gap:1px;flex:1;overflow:hidden;">' +
        '<span style="font-size:9px;color:' + (isAtHome ? '#217346' : '#888') + ';">' + label + '</span>' +
        '<span style="font-family:Consolas,monospace;font-size:10px;font-weight:600;color:' + (isAtHome ? '#217346' : '#444') + ';">' +
          homeCell.sheetName + '!' + homeCell.addr +
        '</span>' +
        '<span style="font-family:Consolas,monospace;font-size:9px;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' +
          esc(shortFormula) +
        '</span>' +
      '</div>' +
    '</div>';

  var box = document.getElementById("home-cell-box");
  box.addEventListener("click", function() { goHome(); });
  box.addEventListener("keydown", function(e) {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); goHome(); }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      if (refs.length > 0) {
        setActive(refs.length - 1, false);
        var items = document.querySelectorAll(".ref-item");
        if (items[refs.length - 1]) items[refs.length - 1].focus();
      }
    }
  });
}

/* ── Ref list rendering ─────────────────────────────────── */

function renderRefList() {
  var list = document.getElementById("refs-list");
  list.innerHTML = "";

  document.getElementById("ref-count").textContent =
    refs.length ? "(" + refs.length + ")" : "(none found)";

  if (refs.length === 0) {
    list.innerHTML = '<div style="padding:12px;color:#aaa;text-align:center;font-size:12px;">No external cell references found.</div>';
    return;
  }

  for (var i = 0; i < refs.length; i++) {
    (function(idx) {
      var ref  = refs[idx];
      var item = document.createElement("div");
      item.className   = "ref-item";
      item.tabIndex    = 0;
      item.dataset.idx = idx;

      var dotColor  = ref.isError ? "#c0392b" : ref.isNamed ? "#b7770d" : "#217346";
      var valueHtml = ref.isError
        ? '<span class="ref-value ref-error">' + esc(ref.value) + '</span>'
        : '<span class="ref-value">' + esc(ref.value) + '</span>';
      var namedBadge = ref.isNamed
        ? '<span class="ref-named-badge">' + esc(ref.resolvedFrom) + '</span>'
        : "";

      item.innerHTML =
        '<div class="ref-icon" style="background:' + dotColor + ';">' +
          '<svg viewBox="0 0 10 10"><rect x="1" y="1" width="8" height="8" rx="1"/></svg>' +
        '</div>' +
        '<div class="ref-main">' +
          '<div class="ref-top-row">' +
            '<span class="ref-addr">' + esc(ref.addr) + '</span>' +
            namedBadge +
            '<span class="ref-sheet">' + esc(ref.sheetName) + '</span>' +
          '</div>' +
          valueHtml +
        '</div>';

      item.addEventListener("click", function() { locked = true; setActive(idx, true); });
      item.addEventListener("keydown", function(e) {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); drillInto(idx); }
      });
      list.appendChild(item);
    })(i);
  }
}

/* ── Active row ─────────────────────────────────────────── */

function setActive(idx, navigate) {
  var items = document.querySelectorAll(".ref-item");
  for (var i = 0; i < items.length; i++) items[i].classList.remove("active");
  var homeBox = document.getElementById("home-cell-box");
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
  var el = items[idx];
  if (el) { el.classList.add("active"); el.scrollIntoView({ block: "nearest" }); }
  if (navigate) navigateTo(refs[idx]);
}

/* ── Keyboard handler ────────────────────────────────────── */

function handleGlobalKey(e) {
  if (e.key === "Escape") {
    e.preventDefault();
    Office.context.ui.closeContainer();
    return;
  }

  if (e.key === "Backspace") {
    e.preventDefault();
    goBack();
    return;
  }

  if (e.key === "ArrowDown") {
    e.preventDefault();
    locked = true;
    var next = activeIdx + 1;
    if (next >= refs.length) {
      setActive(refs.length, history.length === 0);
    } else {
      setActive(next, true);
    }
    return;
  }

  if (e.key === "ArrowUp") {
    e.preventDefault();
    locked = true;
    if (activeIdx === refs.length) {
      setActive(refs.length - 1, true);
    } else if (refs.length > 0) {
      setActive((activeIdx - 1 + refs.length) % refs.length, true);
    }
    return;
  }

  if (e.key === "Enter") {
    e.preventDefault();
    if (activeIdx === refs.length) { goHome(); }
    else { drillInto(activeIdx); }
    return;
  }
}

/* ── Drill in ───────────────────────────────────────────── */

function drillInto(idx) {
  if (idx < 0 || idx >= refs.length) return;

  return Excel.run(function(ctx) {
    var cell  = ctx.workbook.getActiveCell();
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    cell.load("address"); sheet.load("name");
    return ctx.sync().then(function() {
      history.push({
        addr:      cell.address.replace(/^.*!/, ""),
        sheetName: sheet.name,
        label:     cell.address
      });
      locked = false;
      return navigateTo(refs[idx]).then(function() {
        return refreshFromRef(refs[idx]);
      }).then(function() {
        updateBackButton();
      });
    });
  }).catch(function(e) {
    console.error("drillInto error:", e);
  });
}

/* ── Go back ─────────────────────────────────────────────── */

function goBack() {
  if (history.length === 0) return;
  var prev = history.pop();
  locked = false;
  return navigateTo(prev).then(function() {
    return refreshFromRef(prev);
  }).then(function() {
    updateBackButton();
  });
}

/* ── Go home ─────────────────────────────────────────────── */

function goHome() {
  if (!homeCell) return;
  history = [];
  locked  = false;
  return navigateTo(homeCell).then(function() {
    return refreshFromRef(homeCell);
  }).then(function() {
    updateBackButton();
    renderHomeCell();
  });
}

/* ── Navigate Excel ─────────────────────────────────────── */

function navigateTo(ref) {
  return Excel.run(function(ctx) {
    var sheet = ctx.workbook.worksheets.getItem(ref.sheetName);
    var range = sheet.getRange(ref.addr);
    sheet.activate();
    range.select();
    return ctx.sync();
  }).catch(function(e) {
    console.warn("Could not navigate to", ref.sheetName, ref.addr, e);
  });
}
