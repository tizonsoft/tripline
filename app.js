(function () {
  "use strict";

  // ---------- small helpers ----------
  var $ = function (sel) { return document.querySelector(sel); };
  var isBlank = function (v) { return v === null || v === undefined || String(v).trim() === ""; };
  var clean = function (v) { return isBlank(v) ? "" : String(v).trim(); };
  var escapeHTML = function (s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };
  var escapeAttr = escapeHTML;

  var toastTimer = null;
  function showToast(msg, isError) {
    var el = $("#toast");
    el.textContent = msg;
    el.classList.toggle("error", !!isError);
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.hidden = true; }, 4200);
  }

  // ---------- workbook parsing ----------
  function findHeaderRow(rows) {
    var wanted = { date: /date/i, day: /^day$/i, acc: /accommodat/i, act: /activit/i, notes: /notes?/i };
    for (var r = 0; r < Math.min(rows.length, 12); r++) {
      var row = rows[r] || [];
      var cols = {};
      for (var c = 0; c < row.length; c++) {
        var val = clean(row[c]);
        if (!val) continue;
        if (wanted.date.test(val) && cols.date === undefined) cols.date = c;
        else if (wanted.day.test(val) && cols.day === undefined) cols.day = c;
        else if (wanted.acc.test(val) && cols.acc === undefined) cols.acc = c;
        else if (wanted.act.test(val) && cols.act === undefined) cols.act = c;
        else if (wanted.notes.test(val) && cols.notes === undefined) cols.notes = c;
      }
      if (cols.date !== undefined) {
        cols.label = cols.date > 0 ? cols.date - 1 : null;
        return { rowIdx: r, cols: cols };
      }
    }
    return null;
  }

  function parseItinerary(sheet) {
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    var found = findHeaderRow(rows);
    if (!found) return [];
    var cols = found.cols;
    var days = [];
    var current = null;

    for (var r = found.rowIdx + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var rowEmpty = row.every(function (c) { return isBlank(c); });
      if (rowEmpty) {
        if (current) { days.push(current); current = null; }
        continue;
      }

      var dateVal = cols.date !== undefined ? row[cols.date] : null;
      var dayVal = cols.day !== undefined ? row[cols.day] : null;
      var accVal = cols.acc !== undefined ? row[cols.acc] : null;
      var actVal = cols.act !== undefined ? row[cols.act] : null;
      var notesVal = cols.notes !== undefined ? row[cols.notes] : null;
      var labelVal = cols.label !== null ? row[cols.label] : null;

      var hasDate = dateVal instanceof Date || !isBlank(dateVal);
      if (hasDate || !current) {
        if (current) days.push(current);
        var dateObj = dateVal instanceof Date ? dateVal : (isBlank(dateVal) ? null : new Date(dateVal));
        if (dateObj && isNaN(dateObj.getTime())) dateObj = null;
        current = {
          label: clean(labelVal) || ("Day " + (days.length + 1)),
          date: dateObj,
          dow: clean(dayVal),
          entries: []
        };
      }
      if (!isBlank(accVal) || !isBlank(actVal) || !isBlank(notesVal)) {
        current.entries.push({ place: clean(accVal), activity: clean(actVal), notes: clean(notesVal) });
      }
    }
    if (current) days.push(current);
    return days;
  }

  function parseResources(sheet) {
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false });
    var urlRe = /^https?:\/\//i;
    var groups = [];
    var current = null;

    rows.forEach(function (row) {
      var val = null;
      for (var c = 0; c < row.length; c++) {
        if (!isBlank(row[c])) { val = clean(row[c]); break; }
      }
      if (val === null) return;
      if (urlRe.test(val)) {
        if (!current) { current = { title: "Links", links: [] }; groups.push(current); }
        current.links.push(makeLink(val));
      } else {
        current = { title: val, links: [] };
        groups.push(current);
      }
    });
    return groups.filter(function (g) { return g.links.length > 0; });
  }

  function makeLink(url) {
    var label = url, host = "";
    try {
      var u = new URL(url);
      host = u.hostname.replace(/^www\./, "");
      var ignore = { hotel: 1, es: 1, en: 1, city: 1, index: 1 };
      var segments = u.pathname.split("/").filter(Boolean).map(function (s) {
        return s.replace(/\.(html?|php)$/i, "");
      });
      var meaningful = segments.filter(function (s) {
        return s.length > 2 && !ignore[s.toLowerCase()] && !/^\d+$/.test(s);
      });
      var slug = meaningful.length ? meaningful[meaningful.length - 1] : host;
      var nice = slug.replace(/[-_]+/g, " ").replace(/\b\w/g, function (c) { return c.toUpperCase(); }).trim();
      label = nice && nice.toLowerCase() !== host.toLowerCase() ? nice : host;
    } catch (e) { /* leave label as raw url */ }
    return { url: url, label: label, host: host };
  }

  // ---------- rendering ----------
  function formatDateLong(d) {
    try { return d.toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric", year: "numeric" }); }
    catch (e) { return ""; }
  }
  function formatDateShort(d) {
    try { return d.toLocaleDateString(undefined, { month: "short", day: "numeric" }); }
    catch (e) { return ""; }
  }

  function entryHTML(entry, isFirst) {
    var html = "";
    if (!isFirst && entry.place) {
      html += '<p class="entry-location">📍 <strong>' + escapeHTML(entry.place) + '</strong></p>';
    }
    if (entry.activity) html += '<p class="entry-activity">' + escapeHTML(entry.activity) + "</p>";
    if (entry.notes) html += '<p class="entry-notes">📝 ' + escapeHTML(entry.notes) + "</p>";
    if (!html) return "";
    return '<div class="entry">' + html + "</div>";
  }

  function dateBadgeHTML(day) {
    if (day.date) {
      var dow = day.date.toLocaleDateString(undefined, { weekday: "short" }).toUpperCase();
      var mon = day.date.toLocaleDateString(undefined, { month: "short" }).toUpperCase();
      return (
        '<div class="date-badge">' +
          '<span class="date-badge-dow">' + escapeHTML(dow) + '</span>' +
          '<span class="date-badge-num">' + day.date.getDate() + '</span>' +
          '<span class="date-badge-mon">' + escapeHTML(mon) + '</span>' +
        '</div>'
      );
    }
    var fallback = day.dow ? day.dow.slice(0, 3).toUpperCase() : "\u2022";
    return '<div class="date-badge date-badge-fallback"><span class="date-badge-dow">' + escapeHTML(fallback) + '</span></div>';
  }

  function dayCardHTML(day, idx) {
    var dateStr = day.date ? formatDateLong(day.date) : day.dow;
    var stay = (day.entries[0] && day.entries[0].place) ? day.entries[0].place : "";
    var entriesHTML = day.entries.map(function (e, i) { return entryHTML(e, i === 0); }).join("");
    if (!entriesHTML) entriesHTML = '<p class="entry-empty">No details added for this day.</p>';
    return (
      '<li class="day">' +
        '<div class="day-rail"><div class="day-node" aria-hidden="true"></div><div class="day-rail-line" aria-hidden="true"></div></div>' +
        '<div class="day-card">' +
          '<div class="day-card-head">' +
            dateBadgeHTML(day) +
            '<div class="day-card-heading">' +
              '<p class="day-label">' + escapeHTML(day.label) + '</p>' +
              '<p class="day-date">' + escapeHTML(dateStr || "") + '</p>' +
              (stay ? '<p class="stay-line">🛏️ Staying at <strong>' + escapeHTML(stay) + '</strong></p>' : "") +
            '</div>' +
          '</div>' +
          '<div class="day-entries">' + entriesHTML + '</div>' +
        '</div>' +
      '</li>'
    );
  }

  function resourceGroupHTML(g) {
    var links = g.links.map(function (l) {
      return (
        '<a class="resource-link" href="' + escapeAttr(l.url) + '" target="_blank" rel="noopener noreferrer">' +
          '<span class="link-icon">🔗</span>' +
          '<span class="link-text"><strong>' + escapeHTML(l.label) + '</strong><small>' + escapeHTML(l.host) + '</small></span>' +
        '</a>'
      );
    }).join("");
    return '<div class="resource-group"><h3 class="resource-title">📌 ' + escapeHTML(g.title) + '</h3><div class="resource-links">' + links + '</div></div>';
  }

  function prettifyFilename(name) {
    return name.replace(/\.[^.]+$/, "").replace(/[-_]+/g, " ").replace(/\b\w/g, function (c) { return c.toUpperCase(); }).trim();
  }

  function render(days, groups, title) {
    $("#tripTitle").textContent = title;
    var validDates = days.map(function (d) { return d.date; }).filter(Boolean);
    var subtitle = days.length + (days.length === 1 ? " day" : " days");
    if (validDates.length) {
      subtitle = formatDateShort(validDates[0]) + " – " + formatDateShort(validDates[validDates.length - 1]) + " · " + subtitle;
    }
    $("#tripSubtitle").textContent = subtitle;

    $("#timeline").innerHTML = days.map(dayCardHTML).join("");

    var resBtn = $("#tabResourcesBtn");
    if (groups.length) {
      $("#resources").innerHTML = groups.map(resourceGroupHTML).join("");
      resBtn.hidden = false;
    } else {
      resBtn.hidden = true;
    }

    switchTab("itinerary");
    $("#view-upload").hidden = true;
    $("#view-trip").hidden = false;
    $("#topbarActions").hidden = false;
    $("#fabPrint").hidden = false;
  }

  function switchTab(name) {
    document.querySelectorAll(".tab").forEach(function (t) { t.classList.toggle("active", t.dataset.tab === name); });
    $("#panel-itinerary").classList.toggle("active", name === "itinerary");
    $("#panel-itinerary").hidden = name !== "itinerary";
    $("#panel-resources").classList.toggle("active", name === "resources");
    $("#panel-resources").hidden = name !== "resources";
  }

  function resetToUpload() {
    $("#view-trip").hidden = true;
    $("#view-upload").hidden = false;
    $("#topbarActions").hidden = true;
    $("#fabPrint").hidden = true;
    $("#fileInput").value = "";
  }

  function processWorkbook(wb, sourceName) {
    var itinName = wb.SheetNames.find(function (n) { return /itiner|agenda|trip|plan/i.test(n); }) || wb.SheetNames[0];
    var resName = wb.SheetNames.find(function (n) { return /resource|link/i.test(n); });

    var days = parseItinerary(wb.Sheets[itinName]);
    if (!days.length) {
      showToast("Couldn't find itinerary rows. Check that your headers match Date, Day, Accommodations, Activities, Notes.", true);
      return;
    }
    var groups = resName ? parseResources(wb.Sheets[resName]) : [];
    render(days, groups, prettifyFilename(sourceName));
  }

  function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      showToast("Please choose an .xlsx or .xls file.", true);
      return;
    }
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var wb = XLSX.read(new Uint8Array(reader.result), { type: "array", cellDates: true });
        processWorkbook(wb, file.name);
      } catch (e) {
        showToast("Could not read that file. Make sure it's a valid Excel export.", true);
      }
    };
    reader.onerror = function () { showToast("Could not read that file.", true); };
    reader.readAsArrayBuffer(file);
  }

  // ---------- wire up ----------
  document.addEventListener("DOMContentLoaded", function () {
    var fileInput = $("#fileInput");
    var dropzone = $("#dropzone");

    fileInput.addEventListener("change", function (e) { handleFile(e.target.files[0]); });

    ["dragenter", "dragover"].forEach(function (evt) {
      dropzone.addEventListener(evt, function (e) { e.preventDefault(); dropzone.classList.add("dragover"); });
    });
    ["dragleave", "drop"].forEach(function (evt) {
      dropzone.addEventListener(evt, function (e) { e.preventDefault(); dropzone.classList.remove("dragover"); });
    });
    dropzone.addEventListener("drop", function (e) {
      var file = e.dataTransfer.files && e.dataTransfer.files[0];
      handleFile(file);
    });

    $("#btnSample").addEventListener("click", function () {
      fetch("sample/Galicia.xlsx").then(function (r) {
        if (!r.ok) throw new Error("not found");
        return r.arrayBuffer();
      }).then(function (buf) {
        var wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
        processWorkbook(wb, "Galicia Sample Trip");
      }).catch(function () {
        showToast("Sample file isn't available here (needs to be served over http/https, not opened as a local file).", true);
      });
    });

    $("#btnNewFile").addEventListener("click", resetToUpload);
    $("#btnPrint").addEventListener("click", function () { window.print(); });
    $("#fabPrint").addEventListener("click", function () { window.print(); });

    document.querySelectorAll(".tab").forEach(function (t) {
      t.addEventListener("click", function () { switchTab(t.dataset.tab); });
    });
  });
})();
