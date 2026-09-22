(() => {
  "use strict";

  /* ---------------------------------------------------------------
   * Referenzzeit von Excel.Flo (Teil 2)
   * ------------------------------------------------------------- */
  const REFERENCE_TIME_SECONDS = 35.0;

  /* ---------------------------------------------------------------
   * Ranking (Supabase, dasselbe Projekt wie das Übungsportal) – komplett
   * anonym, aber pro Gerät dedupliziert: eine zufällige, in localStorage
   * gespeicherte ID (genau wie "excelflo_anon_id" im Übungsportal) sorgt
   * dafür, dass jeder Teilnehmer nur EINE Zeile mit seiner persönlichen
   * Bestzeit hat – beliebig viele Retries zählen nie als weitere
   * Teilnehmer. Die Tabelle selbst ist über die normale API weder lesbar
   * noch schreibbar, nur die RPC-Funktion unten (siehe
   * hindernisparkour2_runs.sql) darf sie ändern. Alles best effort:
   * schlägt die Verbindung fehl, bleibt einfach der lokale Platzhalter
   * "–" stehen bzw. die Rang-Anzeige leer.
   * ------------------------------------------------------------- */
  const SUPABASE_URL = "https://hbhagmmbowplzjzfvuao.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_Ee47HbgO5Ne8PP9Jh5ugag_iE_lZcp6";
  const ANON_ID_KEY = "hindernisparkour_anon_id";

  let supabaseClient = null;
  function getSupabaseClient() {
    if (!window.supabase || !window.supabase.createClient) return null; // SDK nicht geladen (z.B. offline/geblockt)
    if (!supabaseClient) supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    return supabaseClient;
  }

  function getAnonId() {
    try {
      let id = localStorage.getItem(ANON_ID_KEY);
      if (!id) {
        id = window.crypto && window.crypto.randomUUID
          ? window.crypto.randomUUID()
          : "anon-" + Date.now() + "-" + Math.random().toString(16).slice(2);
        localStorage.setItem(ANON_ID_KEY, id);
      }
      return id;
    } catch (e) {
      return null; // z.B. localStorage blockiert/privater Modus
    }
  }

  function fmtBestTime(seconds) {
    const m = Math.floor(seconds / 60);
    const s = seconds - m * 60;
    return m > 0 ? `${m}:${s.toFixed(0).padStart(2, "0")}` : `${seconds.toFixed(1)}s`;
  }

  function loadBestTime() {
    const client = getSupabaseClient();
    if (!client) return;
    client
      .rpc("hindernisparkour2_best_time")
      .then(({ data, error }) => {
        if (error || data == null) return;
        startBestTimeEl.textContent = fmtBestTime(Number(data));
      })
      .catch(() => {}); // best effort, Platzhalter "–" bleibt stehen
    // Teilnehmerzahl (siehe hindernisparkour_participant_count.sql)
    client
      .rpc("hindernisparkour2_participant_count")
      .then(({ data, error }) => {
        if (error || data == null) return;
        startParticipantsEl.textContent = Number(data).toLocaleString("de-DE");
      })
      .catch(() => {});
  }

  // Trägt den Lauf ein (überschreibt die eigene Zeile nur, wenn diese Runde
  // besser war) und liefert Rang, Teilnehmerzahl, Bestzeit und die
  // tatsächlich gespeicherte persönliche Bestzeit zurück – wird bei jedem
  // Abschluss der Challenge aufgerufen.
  function recordRun(seconds) {
    const client = getSupabaseClient();
    const anonId = getAnonId();
    if (!client || !anonId) return Promise.reject(new Error("Supabase/Anon-ID nicht verfügbar"));
    return client
      .rpc("hindernisparkour2_submit", { p_anon_id: anonId, p_seconds: seconds })
      .then(({ data, error }) => {
        if (error) throw error;
        loadBestTime(); // Startseiten-Bestzeit könnte sich gerade geändert haben
        return data && data[0] ? data[0] : null;
      });
  }

  /* ---------------------------------------------------------------
   * Plattform-Erkennung: Strg (Windows) vs. Cmd (Mac)
   * ------------------------------------------------------------- */
  const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.platform || "") ||
    (navigator.userAgentData && navigator.userAgentData.platform === "macOS");
  const MOD = isMac ? "Cmd" : "Strg";

  /* ---------------------------------------------------------------
   * Touch-only-Erkennung (kein Gerät mit physischer Tastatur/Maus)
   * ------------------------------------------------------------- */
  function isTouchOnlyDevice() {
    const noHover = window.matchMedia("(hover: none)").matches;
    const coarsePointer = window.matchMedia("(pointer: coarse)").matches;
    return noHover && coarsePointer;
  }

  function fmt(n) {
    return typeof n === "number" ? n.toLocaleString("de-DE") : n;
  }

  function fmtCurrencyDE(n) {
    return n.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
  }

  function colLetter(i) {
    return String.fromCharCode(65 + i);
  }

  function refStr(r, c) {
    return colLetter(c) + (r + 1);
  }

  /* ---------------------------------------------------------------
   * Grid-Rendering (Excel-Look-alike)
   * ------------------------------------------------------------- */
  // r1/c1 sind exklusiv (halboffenes Intervall), wie bei Array.slice – vereinfacht Breiten-/Überlappungs-Checks.
  function inRange(r, c, rng) {
    return !!rng && r >= rng.r0 && r < rng.r1 && c >= rng.c0 && c < rng.c1;
  }

  // Baut aus Anker- und aktiver Zelle (Umschalt+Pfeiltasten-Markierung) ein normalisiertes Rechteck (r1/c1 exklusiv).
  function normalizeRange(a, b) {
    return {
      r0: Math.min(a.r, b.r),
      r1: Math.max(a.r, b.r) + 1,
      c0: Math.min(a.c, b.c),
      c1: Math.max(a.c, b.c) + 1,
    };
  }

  const ARROW_DELTA = {
    ArrowLeft: { dr: 0, dc: -1 },
    ArrowRight: { dr: 0, dc: 1 },
    ArrowUp: { dr: -1, dc: 0 },
    ArrowDown: { dr: 1, dc: 0 },
  };

  // Nachbildung von Strg+Pfeil (bzw. der "bis"-Seite von Strg+Umschalt+Pfeil): springt von (r,c) in
  // Richtung (dr,dc) bis zum Rand des zusammenhängenden befüllten Bereichs, wie in echtem Excel.
  // Wichtig (echtes Excel-Verhalten): steht man auf einer befüllten Zelle, deren Nachbar in
  // Richtung (dr,dc) LEER ist, wird nicht einfach abgebrochen – man springt über die Lücke hinweg
  // bis zur nächsten befüllten "Dateninsel" (z. B. von Spalte B über die leere Spalte C direkt nach D).
  function findDataEdge(ctx, r, c, dr, dc) {
    const isFilled = (rr, cc) =>
      rr >= 0 && rr < ctx.rows && cc >= 0 && cc < ctx.cols &&
      ctx.data[rr] && ctx.data[rr][cc] !== "" && ctx.data[rr][cc] != null;
    const inBounds = (rr, cc) => rr >= 0 && rr < ctx.rows && cc >= 0 && cc < ctx.cols;
    let rr = r, cc = c;
    if (isFilled(rr, cc) && isFilled(rr + dr, cc + dc)) {
      while (isFilled(rr + dr, cc + dc)) { rr += dr; cc += dc; }
    } else {
      let tr = rr, tc = cc;
      while (inBounds(tr + dr, tc + dc) && !isFilled(tr + dr, tc + dc)) { tr += dr; tc += dc; }
      if (isFilled(tr + dr, tc + dc)) {
        rr = tr + dr; cc = tc + dc;
      } else if (!ctx.noEmptyEdgeJump) {
        // Kommt in dieser Richtung nichts Befülltes mehr, springt echtes Excel bis
        // ans Blattende. In unserem kleinen Raster wäre das die letzte sichtbare
        // Zeile/Spalte – Aufgaben, bei denen das zu einer irreführenden Markierung
        // führen würde, schalten das über ctx.noEmptyEdgeJump ab (Cursor bleibt stehen).
        rr = tr; cc = tc;
      }
    }
    return { r: rr, c: cc };
  }

  // Normalisiert ctx.colSelected/rowSelected (Zahl oder {from,to}) auf einen Bereich.
  function lineSelRange(sel) {
    if (sel == null) return null;
    if (typeof sel === "number") return { from: sel, to: sel };
    return { from: Math.min(sel.from, sel.to), to: Math.max(sel.from, sel.to) };
  }
  function inLineSel(range, i) {
    return !!range && i >= range.from && i <= range.to;
  }

  function buildGridHTML(ctx) {
    const { cols, rows, data, activeCell, colSelected, rowSelected, headerFilterRow, styledRange, columnColors, rowColors, zoneGrid, orangeGrid, cutRect, selAnchor, wideCols, cellClass, cellStyle, colWidth, minVisualCols, minVisualRows , hiddenRows } = ctx;
    // Bei Aufgaben mit Umschalt+Pfeiltasten-Markierung (selAnchor gesetzt) wird die aktuelle Markierung
    // aus Anker- und aktiver Zelle live berechnet, statt sie separat im State mitzuführen. Bei einer
    // reinen 1-Zellen-"Markierung" (kein Umschalt genutzt) wird nichts zusätzlich hervorgehoben –
    // die normale aktive-Zelle-Markierung reicht dafür aus.
    // colSelected/rowSelected dürfen eine einzelne Nummer oder ein Bereich
    // {from, to} sein – so lassen sich mehrere Spalten/Zeilen auf einmal markieren.
    const colSel = lineSelRange(colSelected);
    const rowSel = lineSelRange(rowSelected);
    const rawSelection = selAnchor && activeCell ? normalizeRange(selAnchor, activeCell) : null;
    const selection = rawSelection && (rawSelection.r1 - rawSelection.r0 > 1 || rawSelection.c1 - rawSelection.c0 > 1)
      ? rawSelection
      : null;
    // Optionale, aufgabenspezifische Mindestgröße: hält das Grid-Fenster stabil, auch wenn
    // "cols"/"rows" durch Löschen kleiner wird, und blendet bewusst noch eine leere Spalte/Zeile
    // über die eigentlichen Daten hinaus ein (wirkt sonst zu knapp abgeschnitten).
    const visualCols = Math.max(cols, minVisualCols || 0);
    const visualRows = Math.max(rows, minVisualRows || 0);
    let html = `<table class="sheet"><thead><tr><th class="corner"></th>`;
    for (let c = 0; c < visualCols; c++) {
      const headClasses = [];
      if (inLineSel(colSel, c)) headClasses.push("col-selected");
      if (activeCell && activeCell.c === c) headClasses.push("active-col-head");
      if (columnColors && columnColors[c]) headClasses.push("col-" + columnColors[c]);
      const w = typeof colWidth === "function" ? colWidth(c) : null;
      const headStyle = w ? ` style="width:${w}px;"` : "";
      html += `<th class="${headClasses.join(" ")}"${headStyle}>${colLetter(c)}</th>`;
    }
    html += "</tr></thead><tbody>";
    for (let r = 0; r < visualRows; r++) {
      // Ausgefilterte Zeilen werden wie in Excel komplett ausgeblendet - die
      // Zeilennummern springen dadurch sichtbar (5, 7, 8, ...), statt Lücken zu zeigen.
      if (hiddenRows && hiddenRows.has(r)) continue;
      const rowheadClasses = ["rowhead"];
      if (inLineSel(rowSel, r)) rowheadClasses.push("row-selected");
      if (activeCell && activeCell.r === r) rowheadClasses.push("active-row-head");
      if (rowColors && rowColors[r]) rowheadClasses.push("row-" + rowColors[r]);
      if (hiddenRows && hiddenRows.size) rowheadClasses.push("row-filtered");
      html += `<tr><td class="${rowheadClasses.join(" ")}">${r + 1}</td>`;
      for (let c = 0; c < visualCols; c++) {
        const classes = ["cell"];
        if (activeCell && activeCell.r === r && activeCell.c === c) classes.push("active");
        if (inLineSel(colSel, c)) classes.push("col-selected");
        if (columnColors && columnColors[c]) classes.push("col-" + columnColors[c]);
        if (inLineSel(rowSel, r)) classes.push("row-selected");
        if (rowColors && rowColors[r]) classes.push("row-" + rowColors[r]);
        if (typeof cellClass === "function") {
          const extra = cellClass(r, c);
          if (extra) classes.push(extra);
        }
        if (orangeGrid && orangeGrid[r] && orangeGrid[r][c]) {
          classes.push("piece-orange");
        } else if (zoneGrid && zoneGrid[r] && zoneGrid[r][c]) {
          classes.push("zone-green");
        }
        if (cutRect && inRange(r, c, cutRect)) classes.push("piece-cut");
        if (selection && inRange(r, c, selection)) classes.push("range-select");
        let content = data && data[r] && data[r][c] != null ? String(data[r][c]) : "";
        if (
          styledRange &&
          r >= styledRange.r0 && r <= styledRange.r1 &&
          c >= styledRange.c0 && c <= styledRange.c1
        ) {
          classes.push("range-styled", "range-border");
          if (r === styledRange.r0) classes.push("range-header");
          else if ((r - styledRange.r0) % 2 === 0) classes.push("alt-row");
        }
        if (headerFilterRow != null && r === headerFilterRow && content) {
          content += '<span class="filter-arrow">▾</span>';
        }
        const styleParts = [];
        if (typeof cellStyle === "function") {
          const style = cellStyle(r, c);
          if (style) styleParts.push(style);
        }
        const cw = typeof colWidth === "function" ? colWidth(c) : null;
        if (cw) styleParts.push(`min-width:${cw}px;max-width:${cw}px;`);
        const styleAttr = styleParts.length ? ` style="${styleParts.join(" ")}"` : "";
        html += `<td class="${classes.join(" ")}"${styleAttr}>${content}</td>`;
      }
      html += "</tr>";
    }
    html += "</tbody></table>";
    return html;
  }


  /* ---------------------------------------------------------------
   * Tastenkappen-Badges & Fortschritts-Pips (rein visuell, Aufgaben-Screen-Redesign)
   * ------------------------------------------------------------- */
  function escapeHTML(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // groups: Array von Key-Gruppen, z.B. [["Strg","Leertaste"],["Strg","Minus (-)"]] –
  // Tasten innerhalb einer Gruppe (gleichzeitig gedrückt) werden mit "+" verbunden,
  // zwischen Gruppen (nacheinander auszuführende Schritte) steht ein Pfeil "→".
  function buildBadgesHTML(groups) {
    return groups
      .map((group) =>
        group.map((key) => `<span class="keycap-badge">${escapeHTML(key)}</span>`).join('<span class="badge-plus">+</span>')
      )
      .join('<span class="badge-arrow">→</span>');
  }

  function buildPipsHTML(currentIndex, total) {
    let html = "";
    for (let i = 0; i < total; i++) {
      const cls = i < currentIndex ? "pip-done" : i === currentIndex ? "pip-current" : "";
      html += `<span class="pip ${cls}"></span>`;
    }
    return html;
  }

  /* ---------------------------------------------------------------
   * Gemeinsame Helfer für die Aufgaben von Teil 2
   *
   * Alle Raster bilden den Aufbau der Kursdatei nach: Spalte A und Zeile 1
   * bleiben leer, die Tabelle beginnt wie in Excel bei B4 (Kopfzeile) bzw.
   * B5 (erste Datenzeile) – so stimmen alle Zellbezüge (C10, D6 …) mit der
   * Excel-Vorlage überein.
   * ------------------------------------------------------------- */

  const HEAD_ROW = 3;   // Zeile 4
  const FIRST_ROW = 4;  // Zeile 5

  function sheetData(rows, cols) {
    const data = [];
    for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(""));
    return data;
  }

  function boolGrid(rows, cols) {
    const grid = [];
    for (let r = 0; r < rows; r++) grid.push(new Array(cols).fill(false));
    return grid;
  }

  // Baut die cellClass-Funktion für den immer gleichen Tabellen-Look: grüne
  // Kopfzeile, abwechselnd hell hinterlegte Datenzeilen, rechtsbündige
  // Zahlenspalten. "extra" ist der aufgabenspezifische Teil.
  function tableLook(spec) {
    const cols = spec.cols;
    return function (r, c) {
      const classes = [];
      const inTable = cols.indexOf(c) !== -1;
      if (inTable && r === (spec.headRow == null ? HEAD_ROW : spec.headRow)) classes.push("head-green");
      else if (inTable && r >= spec.from && r <= spec.to) {
        if ((r - spec.from) % 2 === 1) classes.push("zebra");
        if (spec.numCols && spec.numCols.indexOf(c) !== -1) classes.push("num");
      }
      const extra = spec.extra ? spec.extra(r, c) : null;
      if (extra) classes.push(extra);
      return classes.length ? classes.join(" ") : null;
    };
  }

  // Einheitliche Navigation (Pfeiltasten, Strg+Pfeil, Umschalt+Pfeil, Tab) für
  // alle Aufgaben. Gibt true zurück, wenn die Taste hier schon vollständig
  // behandelt wurde – die Aufgabe prüft danach nur noch ihren eigenen Shortcut.
  function navAndGuard(e, ctx, opts) {
    opts = opts || {};
    const key = e.key;
    const isArrow = Object.prototype.hasOwnProperty.call(ARROW_DELTA, key);
    const isTab = key === "Tab";
    const isAnyModCombo = e.ctrlKey || e.metaKey || e.altKey || e.shiftKey;

    if (isArrow || isTab || isAnyModCombo || key === " ") e.preventDefault();
    if (PURE_MODIFIER_KEYS.has(key)) return true;

    if (opts.noMove && (isArrow || isTab)) return true; // Aufgaben mit fester Markierung

    if (isTab) {
      const c = e.shiftKey
        ? Math.max(0, ctx.activeCell.c - 1)
        : Math.min(ctx.cols - 1, ctx.activeCell.c + 1);
      ctx.activeCell = { r: ctx.activeCell.r, c };
      ctx.selAnchor = { r: ctx.activeCell.r, c };
      if (opts.clearLineSelection) { ctx.colSelected = null; ctx.rowSelected = null; }
      renderCurrentTask();
      return true;
    }

    if (isArrow) {
      const delta = ARROW_DELTA[key];
      const ctrlOrMeta = e.ctrlKey || e.metaKey;

      // Ist bereits eine ganze Zeile/Spalte markiert, erweitert Umschalt+Pfeil
      // die Markierung um weitere ganze Zeilen bzw. Spalten – wie in Excel.
      const lineSel = ctx.rowSelected != null || ctx.colSelected != null;
      if (opts.clearLineSelection && lineSel && e.shiftKey && !ctrlOrMeta) {
        const anchor = ctx.selAnchor || ctx.activeCell;
        if (ctx.rowSelected != null && delta.dr !== 0) {
          const r = Math.max(0, Math.min(ctx.rows - 1, ctx.activeCell.r + delta.dr));
          ctx.activeCell = { r, c: ctx.activeCell.c };
          ctx.rowSelected = { from: anchor.r, to: r };
          renderCurrentTask();
          return true;
        }
        if (ctx.colSelected != null && delta.dc !== 0) {
          const c = Math.max(0, Math.min(ctx.cols - 1, ctx.activeCell.c + delta.dc));
          ctx.activeCell = { r: ctx.activeCell.r, c };
          ctx.colSelected = { from: anchor.c, to: c };
          renderCurrentTask();
          return true;
        }
        // Pfeil quer zur markierten Richtung ändert an der Markierung nichts.
        renderCurrentTask();
        return true;
      }

      if (ctrlOrMeta && e.shiftKey) {
        ctx.activeCell = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
      } else if (ctrlOrMeta) {
        const edge = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
        ctx.activeCell = edge;
        ctx.selAnchor = edge;
      } else {
        const r = Math.max(0, Math.min(ctx.rows - 1, ctx.activeCell.r + delta.dr));
        const c = Math.max(0, Math.min(ctx.cols - 1, ctx.activeCell.c + delta.dc));
        ctx.activeCell = { r, c };
        if (!e.shiftKey) ctx.selAnchor = { r, c };
      }
      // Nach dem Bewegen ist eine vorher markierte ganze Spalte/Zeile hinfällig.
      if (opts.clearLineSelection) { ctx.colSelected = null; ctx.rowSelected = null; }
      renderCurrentTask();
      return true;
    }

    if (e.repeat) return true;
    return false;
  }

  // Aktuelle Markierung als normalisiertes Rechteck (r1/c1 exklusiv).
  function currentSelection(ctx) {
    return normalizeRange(ctx.selAnchor || ctx.activeCell, ctx.activeCell);
  }
  function selectionIs(ctx, r0, r1, c0, c1) {
    const s = currentSelection(ctx);
    return s.r0 === r0 && s.r1 === r1 && s.c0 === c0 && s.c1 === c1;
  }
  function isCell(ctx, r, c) {
    return ctx.activeCell.r === r && ctx.activeCell.c === c;
  }

  function todayDE() {
    const d = new Date();
    return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
  }

  /* ---------------------------------------------------------------
   * Menüband-Mockup für die Spaltenbreiten-Aufgabe: Alt, R, F, F, I
   * (Start > Format > Spaltenbreite automatisch anpassen).
   * Die Register-Buchstaben sind dieselben wie in Teil 1.
   * ------------------------------------------------------------- */
  const RIBBON_TABS = [
    { letter: "R", label: "Start" },
    { letter: "I", label: "Einfügen" },
    { letter: "S", label: "Seitenlayout" },
    { letter: "O", label: "Formeln" },
    { letter: "V", label: "Daten" },
    { letter: "P", label: "Überprüfung" },
    { letter: "F", label: "Ansicht" },
  ];

  function buildFormatRibbonHTML(ctx) {
    // 0 = aus, 1 = Alt gedrückt, 2 = Start-Register, 3 = erstes F, 4 = Format-Menü
    const stage = ctx.keytipStage;
    let html = '<div class="ribbon-tabs">';
    RIBBON_TABS.forEach((t) => {
      const isActive = t.label === "Start" && stage >= 2;
      const badge = stage === 1 ? `<span class="ribbon-keytip">${t.letter}</span>` : "";
      html += `<div class="ribbon-tab${isActive ? " active" : ""}">${t.label}${badge}</div>`;
    });
    html += "</div>";

    if (stage < 2) {
      html += `<div class="ribbon-content"><div class="ribbon-group">
        <div class="ribbon-btn">Einfügen</div>
        <div class="ribbon-btn">Ausschneiden</div>
        <div class="ribbon-btn">Format übertragen</div>
      </div></div>`;
      return html;
    }

    // Nach dem ersten F ist der zweite Buchstabe der Folge "FF" noch offen.
    const formatBadge = stage === 2
      ? '<span class="ribbon-keytip">FF</span>'
      : stage === 3
        ? '<span class="ribbon-keytip"><span class="keytip-used">F</span>F</span>'
        : "";
    html += `<div class="ribbon-content"><div class="ribbon-group">
        <div class="ribbon-btn">Einfügen</div>
        <div class="ribbon-btn">Löschen</div>
        <div class="ribbon-btn ribbon-btn-target">Format${formatBadge}</div>
      </div>`;

    if (stage === 4) {
      html += `<div class="ribbon-dropdown ribbon-dropdown-format">
          <div class="ribbon-dropdown-item">Zeilenhöhe … <span class="ribbon-keytip">H</span></div>
          <div class="ribbon-dropdown-item">Zeilenhöhe automatisch anpassen <span class="ribbon-keytip">A</span></div>
          <div class="ribbon-dropdown-item">Spaltenbreite … <span class="ribbon-keytip">B</span></div>
          <div class="ribbon-dropdown-item is-target">Spaltenbreite automatisch anpassen <span class="ribbon-keytip">I</span></div>
          <div class="ribbon-dropdown-item">Standardbreite … <span class="ribbon-keytip">S</span></div>
        </div>`;
    }
    html += "</div>";
    return html;
  }

  /* ---------------------------------------------------------------
   * Die 8 Aufgaben von Teil 2 – in der Reihenfolge der Tabellenblätter
   * 0 bis 7 aus "Shortcut-Hindernisparcours - Teil 2.xlsx".
   * ------------------------------------------------------------- */
  const TASKS = [

    /* -----------------------------------------------------------
     * Blatt 0 – Zellen entfernen: ganze Spalte (Strg+Leertaste) bzw.
     * ganze Zeile (Umschalt+Leertaste) markieren und mit Strg+Minus
     * löschen. Wer vorher mit Umschalt+Pfeiltasten einen Bereich
     * markiert, erwischt damit gleich mehrere Spalten/Zeilen auf
     * einmal – wie in Excel. Der grüne Block in der Mitte muss stehen
     * bleiben, deshalb lässt sich die mittlere Spalte nicht einfach
     * als Ganzes löschen.
     * --------------------------------------------------------- */
    {
      instruction: "Entferne die orangen Zellen.",
      keys: () => [["Umschalt", "Pfeiltasten"], [MOD, "Leertaste"], ["Umschalt", "Leertaste"], [MOD, "Minus (-)"]],
      hint(ctx) {
        return `Mit ${MOD}+Leertaste ganze Spalten, mit Umschalt+Leertaste ganze Zeilen markieren – die grünen Zellen müssen bleiben. (noch ${countOrange(ctx)} orange Zellen)`;
      },
      init() {
        const cols = 5, rows = 15;
        const orange = boolGrid(rows, cols);
        const green = boolGrid(rows, cols);
        // B4:D12 orange, in der Mitte C7:C9 grün – der grüne Block ist tabu.
        for (let r = 3; r <= 11; r++) {
          for (let c = 1; c <= 3; c++) {
            if (c === 2 && r >= 6 && r <= 8) green[r][c] = true;
            else orange[r][c] = true;
          }
        }
        return {
          cols, rows,
          data: sheetData(rows, cols),
          orangeGrid: orange,
          zoneGrid: green,
          activeCell: { r: 3, c: 1 },
          selAnchor: { r: 3, c: 1 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          minVisualCols: cols, minVisualRows: rows,
          colWidth: (c) => (c === 0 ? 48 : null),
        };
      },
      customHandleKey(e, ctx) {
        if (navAndGuard(e, ctx, { clearLineSelection: true })) return;
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const isSpace = e.code === "Space" || e.key === " " || e.key === "Spacebar";

        // Strg+Leertaste markiert die ganzen Spalten der aktuellen Markierung,
        // Umschalt+Leertaste die ganzen Zeilen – genau wie in Excel.
        if (isSpace && ctrlOrMeta && !e.shiftKey) {
          const sel = currentSelection(ctx);
          ctx.colSelected = { from: sel.c0, to: sel.c1 - 1 };
          ctx.rowSelected = null;
          renderCurrentTask();
          return;
        }
        if (isSpace && e.shiftKey && !ctrlOrMeta) {
          const sel = currentSelection(ctx);
          ctx.rowSelected = { from: sel.r0, to: sel.r1 - 1 };
          ctx.colSelected = null;
          renderCurrentTask();
          return;
        }

        const isDelete = ctrlOrMeta && !e.altKey &&
          (e.code === "Minus" || e.code === "NumpadSubtract" || e.key === "-");

        if (isDelete) {
          const colSel = lineSelRange(ctx.colSelected);
          const rowSel = lineSelRange(ctx.rowSelected);
          if (!colSel && !rowSel) {
            showWrongToast(); // ohne markierte Spalte/Zeile gibt es nichts zu löschen
            return;
          }

          let orange = 0, green = 0;
          for (let r = 0; r < ctx.rows; r++) {
            for (let c = 0; c < ctx.cols; c++) {
              const hit = colSel ? inLineSel(colSel, c) : inLineSel(rowSel, r);
              if (!hit) continue;
              if (ctx.orangeGrid[r][c]) orange++;
              if (ctx.zoneGrid[r][c]) green++;
            }
          }
          // Der grüne Block darf nicht mitgelöscht werden, und irgendetwas
          // Oranges muss die Auswahl auch treffen.
          if (green > 0 || orange === 0) {
            showWrongToast();
            return;
          }

          if (colSel) {
            // Spalten löschen: alles rechts davon rückt nach links.
            const count = colSel.to - colSel.from + 1;
            [ctx.orangeGrid, ctx.zoneGrid].forEach((grid) => {
              grid.forEach((row) => {
                row.splice(colSel.from, count);
                for (let i = 0; i < count; i++) row.push(false);
              });
            });
            ctx.activeCell = { r: ctx.activeCell.r, c: Math.min(colSel.from, ctx.cols - 1) };
          } else {
            // Zeilen löschen: alles darunter rückt nach oben.
            const count = rowSel.to - rowSel.from + 1;
            [ctx.orangeGrid, ctx.zoneGrid].forEach((grid) => {
              grid.splice(rowSel.from, count);
              for (let i = 0; i < count; i++) grid.push(new Array(ctx.cols).fill(false));
            });
            ctx.activeCell = { r: Math.min(rowSel.from, ctx.rows - 1), c: ctx.activeCell.c };
          }
          ctx.selAnchor = { r: ctx.activeCell.r, c: ctx.activeCell.c };
          ctx.colSelected = null;
          ctx.rowSelected = null;

          if (countOrange(ctx) === 0) {
            handleSuccess();
            return;
          }
          renderCurrentTask();
          return;
        }
        if (e.ctrlKey || e.metaKey || e.altKey) showWrongToast();
      },
    },

    /* -----------------------------------------------------------
     * Blatt 1 – Blitzvorschau (Strg+E): E-Mail-Adressen nach dem
     * Muster der ersten Zeile automatisch ausfüllen lassen.
     * --------------------------------------------------------- */
    {
      instruction: "Fülle die E-Mail-Adressen für alle anderen Agenten aus – nach dem Muster von Agent Brezel.",
      keys: () => [[MOD, "E"]],
      hint: "Die Geheimagentur K.Ä.S.E. stellt auf E-Mail um. Stell dich in die Spalte E-Mail und lass Excel das Muster erkennen.",
      init() {
        const AGENTS = [
          ["Bruno", "Brezel"], ["Gerda", "Gurke"], ["Otto", "Oktopus"], ["Susi", "Sonnenschein"],
          ["Hugo", "Hering"], ["Paula", "Pudding"], ["Max", "Muffin"], ["Lena", "Lakritz"],
          ["Timo", "Toast"], ["Frida", "Flamingo"], ["Kurt", "Keks"], ["Nina", "Nudel"],
        ];
        const cols = 5, rows = FIRST_ROW + AGENTS.length + 1;
        const lastRow = FIRST_ROW + AGENTS.length - 1; // Zeile 16
        const data = sheetData(rows, cols);
        data[HEAD_ROW][1] = "Vorname";
        data[HEAD_ROW][2] = "Nachname";
        data[HEAD_ROW][3] = "E-Mail";
        AGENTS.forEach((agent, i) => {
          data[FIRST_ROW + i][1] = agent[0];
          data[FIRST_ROW + i][2] = agent[1];
        });
        const mail = (first, last) => `${first}.${last}@kaese-agentur.de`.toLowerCase();
        data[FIRST_ROW][3] = mail(AGENTS[0][0], AGENTS[0][1]);

        const state = {
          cols, rows, data, AGENTS, mail, lastRow,
          activeCell: { r: FIRST_ROW + 1, c: 3 },
          selAnchor: { r: FIRST_ROW + 1, c: 3 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          done: false,
          minVisualCols: cols, minVisualRows: rows,
          colWidth: (c) => (c === 0 ? 48 : c === 3 ? 290 : null),
        };
        state.cellClass = tableLook({
          cols: [1, 2, 3], from: FIRST_ROW, to: lastRow,
          extra(r, c) {
            if (c !== 3) return null;
            if (r === FIRST_ROW) return "cell-source";
            if (r > FIRST_ROW && r <= lastRow) return state.done ? "cell-done" : "cell-pending";
            return null;
          },
        });
        return state;
      },
      customHandleKey(e, ctx) {
        if (navAndGuard(e, ctx)) return;
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const isFlashFill = ctrlOrMeta && !e.altKey && !e.shiftKey &&
          (e.code === "KeyE" || e.key === "e" || e.key === "E");

        if (isFlashFill) {
          if (ctx.done) return;
          // Wie in Excel: der Cursor muss in der Spalte stehen, die gefüllt werden soll.
          const inMailColumn = ctx.activeCell.c === 3 &&
            ctx.activeCell.r >= FIRST_ROW && ctx.activeCell.r <= ctx.lastRow;
          if (!inMailColumn) {
            showWrongToast();
            return;
          }
          ctx.AGENTS.forEach((agent, i) => {
            ctx.data[FIRST_ROW + i][3] = ctx.mail(agent[0], agent[1]);
          });
          ctx.done = true;
          handleSuccess();
          return;
        }
        if (e.ctrlKey || e.metaKey || e.altKey) showWrongToast();
      },
    },

    /* -----------------------------------------------------------
     * Blatt 2 – Navigation: mit Strg+Pfeiltasten von Start (D4) im
     * Zickzack über alle x-Zellen zum Ziel (D13). Der ganze Pfad ist
     * hellgelb vorgezeichnet und wird grün, sobald man ihn abläuft.
     * Abkürzungen (z. B. direkt D4 -> D7) zählen nicht.
     * --------------------------------------------------------- */
    {
      instruction: "Navigiere von Start zu Ziel.",
      keys: () => [[MOD, "Pfeiltasten"]],
      hint(ctx) {
        const left = ctx.path.length - 1 - ctx.currentIndex;
        return `Folge dem gelben Pfad – nur mit Strg+Pfeiltasten und in der vorgegebenen Reihenfolge. (noch ${left} Stationen)`;
      },
      init() {
        const cols = 5, rows = 15;
        const data = sheetData(rows, cols);
        // Zickzack-Route aus der Kursdatei: Start D4, dann über alle x nach D13.
        const path = [
          { r: 3, c: 3 },  // D4  – Start
          { r: 3, c: 1 },  // B4
          { r: 6, c: 1 },  // B7
          { r: 6, c: 3 },  // D7
          { r: 9, c: 3 },  // D10
          { r: 9, c: 1 },  // B10
          { r: 12, c: 1 }, // B13
          { r: 12, c: 3 }, // D13 – Ziel
        ];
        path.forEach((p) => { data[p.r][p.c] = "x"; });
        data[path[0].r][path[0].c] = "Start";
        data[path[path.length - 1].r][path[path.length - 1].c] = "Ziel";

        // Zellen zwischen zwei Stationen – sie zeichnen den Pfad sichtbar vor
        // und werden grün, sobald das jeweilige Teilstück gelaufen ist.
        const segments = [];
        for (let i = 0; i < path.length - 1; i++) {
          const from = path[i], to = path[i + 1];
          const dr = Math.sign(to.r - from.r), dc = Math.sign(to.c - from.c);
          const cells = [];
          let r = from.r + dr, c = from.c + dc;
          while (r !== to.r || c !== to.c) {
            cells.push({ r, c });
            r += dr; c += dc;
          }
          segments.push(cells);
        }

        const state = {
          cols, rows, data, path, segments,
          currentIndex: 0,
          visited: new Set([`${path[0].r},${path[0].c}`]),
          activeCell: { r: path[0].r, c: path[0].c },
          selAnchor: { r: path[0].r, c: path[0].c },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          minVisualCols: cols, minVisualRows: rows,
          colWidth: (c) => (c === 0 ? 48 : null),
        };
        state.cellClass = (r, c) => {
          const onPath = state.path.some((p) => p.r === r && p.c === c) ||
            state.segments.some((seg) => seg.some((p) => p.r === r && p.c === c));
          if (!onPath) return null;
          const idx = state.path.findIndex((p) => p.r === r && p.c === c);
          const classes = [];
          if (idx === 0 || idx === state.path.length - 1) classes.push("cell-endpoint");
          else if (idx > 0) classes.push("cell-waypoint");
          classes.push(state.visited.has(`${r},${c}`) ? "cell-done" : "cell-path");
          return classes.join(" ");
        };
        return state;
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        const isArrow = Object.prototype.hasOwnProperty.call(ARROW_DELTA, key);
        if (!isArrow) {
          if (e.ctrlKey || e.metaKey || e.altKey) {
            e.preventDefault();
            if (!PURE_MODIFIER_KEYS.has(key) && !e.repeat) showWrongToast();
          }
          return;
        }
        e.preventDefault();
        if (e.repeat) return;

        const delta = ARROW_DELTA[key];
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const next = ctrlOrMeta
          ? findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc)
          : {
              r: Math.max(0, Math.min(ctx.rows - 1, ctx.activeCell.r + delta.dr)),
              c: Math.max(0, Math.min(ctx.cols - 1, ctx.activeCell.c + delta.dc)),
            };
        ctx.activeCell = next;
        ctx.selAnchor = { r: next.r, c: next.c };

        const expected = ctx.path[ctx.currentIndex + 1];
        const ziel = ctx.path[ctx.path.length - 1];
        if (expected && next.r === expected.r && next.c === expected.c) {
          ctx.segments[ctx.currentIndex].forEach((p) => ctx.visited.add(`${p.r},${p.c}`));
          ctx.visited.add(`${expected.r},${expected.c}`);
          ctx.currentIndex += 1;
          if (ctx.currentIndex === ctx.path.length - 1) {
            handleSuccess();
            return;
          }
        } else if (next.r === ziel.r && next.c === ziel.c) {
          // Ziel per Abkürzung erreicht, ohne die Reihenfolge einzuhalten – zählt nicht.
          showWrongToast();
        }
        renderCurrentTask();
      },
    },

    /* -----------------------------------------------------------
     * Blatt 3 – Heutiges Datum (Strg+.) in die orange Zelle C10.
     * --------------------------------------------------------- */
    {
      instruction: "Trage das heutige Datum in die orange Zelle ein.",
      keys: () => [[MOD, "."]],
      hint: "Das Kaffeemaschinen-Logbuch: In die orange Zelle gehört das heutige Datum – ohne es zu tippen.",
      init() {
        const LOG = [
          ["Maschine gekauft", "01.01.2019"],
          ["Erster Kaffee (zu dünn)", "02.01.2019"],
          ["Entkalkt (angeblich)", "04.01.2021"],
          ["Brühgruppe verklemmt", "15.03.2023"],
          ["Jemand hat Tee reingemacht", "29.01.2024"],
        ];
        const cols = 4, rows = FIRST_ROW + LOG.length + 2;
        const targetRow = FIRST_ROW + LOG.length; // Zeile 10
        const data = sheetData(rows, cols);
        data[HEAD_ROW][1] = "Ereignis";
        data[HEAD_ROW][2] = "Datum";
        LOG.forEach((entry, i) => {
          data[FIRST_ROW + i][1] = entry[0];
          data[FIRST_ROW + i][2] = entry[1];
        });
        data[targetRow][1] = "Endlich richtig entkalkt!";

        const state = {
          cols, rows, data, targetRow,
          activeCell: { r: targetRow, c: 2 },
          selAnchor: { r: targetRow, c: 2 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          done: false,
          minVisualCols: cols, minVisualRows: rows,
          colWidth: (c) => (c === 0 ? 48 : c === 1 ? 300 : null),
        };
        state.cellClass = tableLook({
          cols: [1, 2], from: FIRST_ROW, to: targetRow - 1,
          extra(r, c) {
            if (r !== state.targetRow) return null;
            if (c === 1) return "cell-total";
            if (c === 2) return state.done ? "cell-done" : "cell-pending";
            return null;
          },
        });
        return state;
      },
      customHandleKey(e, ctx) {
        if (navAndGuard(e, ctx)) return;
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const isToday = ctrlOrMeta && !e.altKey && !e.shiftKey &&
          (e.code === "Period" || e.code === "NumpadDecimal" || e.key === ".");

        if (isToday) {
          if (ctx.done) return;
          if (!isCell(ctx, ctx.targetRow, 2)) {
            showWrongToast();
            return;
          }
          ctx.data[ctx.targetRow][2] = todayDE();
          ctx.done = true;
          handleSuccess();
          return;
        }
        if (e.ctrlKey || e.metaKey || e.altKey) showWrongToast();
      },
    },

    /* -----------------------------------------------------------
     * Blatt 4 – Daten verschieben: B5:D8 markieren, ausschneiden
     * (Strg+X) und in die zweite Tabelle ab F9 einfügen (Strg+V).
     * --------------------------------------------------------- */
    {
      instruction: "Verschiebe die Daten in die zweite Tabelle.",
      keys: () => [["Umschalt", "Pfeiltasten"], [MOD, "X"], [MOD, "V"]],
      hint(ctx) {
        return ctx.cut
          ? "Jetzt in die erste orange Zelle der zweiten Tabelle stellen und einfügen."
          : "Markiere die vier Datenzeilen (ohne Überschrift), schneide sie aus und füg sie unter den neuen Überschriften ein.";
      },
      init() {
        const ROWS = [
          ["Anna Müller", "Vertrieb", 12500],
          ["Ben Schmidt", "Marketing", 9800],
          ["Clara Weber", "Einkauf", 11350],
          ["David Fischer", "Service", 8750],
        ];
        const HEADS = ["Name", "Abteilung", "Umsatz"];
        // Spalte I bleibt als leere Spalte rechts neben der zweiten Tabelle stehen.
        const cols = 9, rows = 14;
        const srcHeadRow = HEAD_ROW;      // Zeile 4
        const srcFirstRow = FIRST_ROW;    // Zeile 5
        const srcLastRow = srcFirstRow + ROWS.length - 1; // Zeile 8
        const dstHeadRow = 7;             // Zeile 8
        const dstFirstRow = 8;            // Zeile 9
        const dstLastRow = dstFirstRow + ROWS.length - 1; // Zeile 12
        const data = sheetData(rows, cols);
        HEADS.forEach((h, i) => {
          data[srcHeadRow][1 + i] = h;
          data[dstHeadRow][5 + i] = h;
        });
        ROWS.forEach((row, i) => {
          data[srcFirstRow + i][1] = row[0];
          data[srcFirstRow + i][2] = row[1];
          data[srcFirstRow + i][3] = fmtCurrencyDE(row[2]);
        });

        const state = {
          cols, rows, data, ROWS,
          srcHeadRow, srcFirstRow, srcLastRow, dstHeadRow, dstFirstRow, dstLastRow,
          activeCell: { r: srcFirstRow, c: 1 },
          selAnchor: { r: srcFirstRow, c: 1 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          cutRect: null, cut: false, done: false,
          minVisualCols: cols, minVisualRows: rows,
          colWidth: (c) => (c === 0 ? 48 : c === 4 ? 48 : c === 8 ? 64 : null),
        };
        state.cellClass = (r, c) => {
          const classes = [];
          const inSrc = c >= 1 && c <= 3;
          const inDst = c >= 5 && c <= 7;
          if ((inSrc && r === state.srcHeadRow) || (inDst && r === state.dstHeadRow)) classes.push("head-green");
          if (inSrc && r >= state.srcFirstRow && r <= state.srcLastRow) {
            if ((r - state.srcFirstRow) % 2 === 1) classes.push("zebra");
            if (c === 3) classes.push("num");
          }
          if (inDst && r >= state.dstFirstRow && r <= state.dstLastRow) {
            classes.push(state.done ? "cell-done" : "cell-pending");
            if (c === 7) classes.push("num");
          }
          return classes.length ? classes.join(" ") : null;
        };
        return state;
      },
      customHandleKey(e, ctx) {
        if (navAndGuard(e, ctx)) return;
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        if (!ctrlOrMeta || e.altKey) {
          if (e.altKey) showWrongToast();
          return;
        }
        const isCut = e.code === "KeyX" || e.key === "x" || e.key === "X";
        const isPaste = e.code === "KeyV" || e.key === "v" || e.key === "V";

        if (isCut) {
          if (ctx.done) return;
          // Genau die vier Datenzeilen, ohne die Überschriften.
          if (!selectionIs(ctx, ctx.srcFirstRow, ctx.srcLastRow + 1, 1, 4)) {
            showWrongToast();
            return;
          }
          ctx.cut = true;
          ctx.cutRect = { r0: ctx.srcFirstRow, r1: ctx.srcLastRow + 1, c0: 1, c1: 4 };
          renderCurrentTask();
          return;
        }

        if (isPaste) {
          if (ctx.done) return;
          if (!ctx.cut) return; // nichts ausgeschnitten – in Excel passiert dann auch nichts
          // Eingefügt wird ab der linken oberen Zelle des Zielbereichs (F9).
          const okSingle = isCell(ctx, ctx.dstFirstRow, 5) &&
            selectionIs(ctx, ctx.dstFirstRow, ctx.dstFirstRow + 1, 5, 6);
          const okRange = selectionIs(ctx, ctx.dstFirstRow, ctx.dstLastRow + 1, 5, 8);
          if (!okSingle && !okRange) {
            showWrongToast();
            return;
          }
          ctx.ROWS.forEach((row, i) => {
            ctx.data[ctx.dstFirstRow + i][5] = row[0];
            ctx.data[ctx.dstFirstRow + i][6] = row[1];
            ctx.data[ctx.dstFirstRow + i][7] = fmtCurrencyDE(row[2]);
            for (let c = 1; c <= 3; c++) ctx.data[ctx.srcFirstRow + i][c] = "";
          });
          ctx.cutRect = null;
          ctx.done = true;
          handleSuccess();
          return;
        }
        showWrongToast();
      },
    },

    /* -----------------------------------------------------------
     * Blatt 5 – Spaltenbreite automatisch anpassen, ohne Maus:
     * Alt, R, F, F, I (Start > Format > Spaltenbreite automatisch
     * anpassen). Die Tasten werden nacheinander gedrückt.
     * --------------------------------------------------------- */
    {
      instruction: "Mach Spalte C so breit, dass man den Gewinn lesen kann – ohne Maus.",
      keys: () => [["Alt"], ["R"], ["F"], ["F"], ["I"]],
      hint: "Der Chef behauptet, er hat im Lotto gewonnen – man sieht nur Rauten. Tasten nacheinander drücken: Alt, R, F, F, I.",
      renderRibbon: (ctx) => buildFormatRibbonHTML(ctx),
      init() {
        const WINNERS = [
          ["Chef", 8765432.1],
          ["Praktikant Torben", 2.5],
          ["Frau Yilmaz", 12450],
        ];
        const cols = 4, rows = FIRST_ROW + WINNERS.length + 2;
        const lastRow = FIRST_ROW + WINNERS.length - 1; // Zeile 7
        const data = sheetData(rows, cols);
        data[HEAD_ROW][1] = "Wer?";
        data[HEAD_ROW][2] = "Gewinn";
        WINNERS.forEach((w, i) => {
          data[FIRST_ROW + i][1] = w[0];
          data[FIRST_ROW + i][2] = "########";
        });

        const state = {
          cols, rows, data, WINNERS, lastRow,
          activeCell: { r: FIRST_ROW, c: 2 },
          selAnchor: { r: FIRST_ROW, c: 2 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          keytipStage: 0,
          done: false,
          minVisualCols: cols, minVisualRows: rows,
          // Die Zielspalte ist bewusst zu schmal und wird erst am Ende breit.
          colWidth: (c) => (c === 0 ? 48 : c === 1 ? 240 : c === 2 ? (state.done ? 240 : 96) : null),
        };
        state.cellClass = tableLook({
          cols: [1, 2], from: FIRST_ROW, to: lastRow, numCols: [2],
          extra(r, c) {
            if (c !== 2 || r < FIRST_ROW || r > state.lastRow) return null;
            return state.done ? "cell-done num" : "cell-pending num";
          },
        });
        return state;
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        // Alt muss hier selbst ausgewertet werden (Start der Key-Tip-Folge),
        // deshalb nicht der übliche navAndGuard-Weg mit seiner Modifier-Abkürzung.
        if (key === "Alt") {
          e.preventDefault();
          if (e.repeat) return;
          ctx.keytipStage = 1;
          renderCurrentTask();
          return;
        }
        if (Object.prototype.hasOwnProperty.call(ARROW_DELTA, key) || key === "Tab") {
          if (navAndGuard(e, ctx)) return;
        }
        if (PURE_MODIFIER_KEYS.has(key)) return;
        if (e.repeat) return;
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault();
          showWrongToast();
          return;
        }
        if (key === "Escape") {
          ctx.keytipStage = 0;
          renderCurrentTask();
          return;
        }
        if (ctx.keytipStage === 0) return; // ohne Alt passiert nichts

        e.preventDefault();
        const code = e.code;
        const wrong = () => {
          ctx.keytipStage = 0;
          showWrongToast();
          renderCurrentTask();
        };

        // Stufen: 1 = Alt, 2 = Start-Register (wartet auf F), 3 = erstes F
        // (wartet auf das zweite F), 4 = Format-Menü offen (wartet auf I).
        if (ctx.keytipStage === 1) {
          if (code === "KeyR") { ctx.keytipStage = 2; renderCurrentTask(); } else wrong();
          return;
        }
        if (ctx.keytipStage === 2) {
          if (code === "KeyF") { ctx.keytipStage = 3; renderCurrentTask(); } else wrong();
          return;
        }
        if (ctx.keytipStage === 3) {
          if (code === "KeyF") { ctx.keytipStage = 4; renderCurrentTask(); } else wrong();
          return;
        }
        if (ctx.keytipStage === 4) {
          if (code !== "KeyI") { wrong(); return; }
          // "Spaltenbreite automatisch anpassen" wirkt auf die markierte Spalte.
          if (ctx.activeCell.c !== 2) { wrong(); return; }
          ctx.WINNERS.forEach((w, i) => {
            ctx.data[FIRST_ROW + i][2] = fmtCurrencyDE(w[1]);
          });
          ctx.keytipStage = 0;
          ctx.done = true;
          handleSuccess();
        }
      },
    },

    /* -----------------------------------------------------------
     * Blatt 6 – Kopieren (Strg+C), Zielbereich markieren
     * (Umschalt+Pfeiltasten) und einfügen (Strg+V). Die Ursprungs-
     * zelle darf dabei mit überschrieben werden.
     * --------------------------------------------------------- */
    {
      instruction: "Kopiere Annas Bestellung in alle orangen Zellen darunter.",
      keys: () => [[MOD, "C"], ["Umschalt", "Pfeiltasten"], [MOD, "V"]],
      hint(ctx) {
        return ctx.copied
          ? "Jetzt die orangen Zellen markieren (Umschalt+Pfeil nach unten) und einfügen."
          : "Donut-Freitag: Erst Annas Bestellung kopieren, dann die orangen Zellen markieren und einfügen.";
      },
      init() {
        const NAMES = ["Anna", "Ben", "Carla", "David", "Elena", "Felix", "Greta", "Hans", "Iris", "Jonas"];
        const ORDER = "Donut mit Streuseln";
        const cols = 4, rows = FIRST_ROW + NAMES.length + 1;
        const lastRow = FIRST_ROW + NAMES.length - 1; // Zeile 14
        const data = sheetData(rows, cols);
        data[HEAD_ROW][1] = "Name";
        data[HEAD_ROW][2] = "Bestellung";
        NAMES.forEach((n, i) => { data[FIRST_ROW + i][1] = n; });
        data[FIRST_ROW][2] = ORDER;

        const state = {
          cols, rows, data, ORDER, lastRow,
          activeCell: { r: FIRST_ROW, c: 2 },
          selAnchor: { r: FIRST_ROW, c: 2 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          cutRect: null, copied: false, done: false,
          // Unter C5 ist die Spalte leer: Strg+Umschalt+Pfeil nach unten würde
          // sonst bis zur letzten Rasterzeile durchlaufen und eine Markierung
          // erzeugen, die gar nicht dem Ziel entspricht. Hier bleibt der Cursor
          // deshalb stehen – markiert wird mit Umschalt+Pfeil oder von unten
          // (über Spalte B nach C14 und von dort mit Strg+Umschalt+Pfeil hoch).
          noEmptyEdgeJump: true,
          minVisualCols: cols, minVisualRows: rows,
          colWidth: (c) => (c === 0 ? 48 : c === 2 ? 280 : null),
        };
        state.cellClass = tableLook({
          cols: [1, 2], from: FIRST_ROW, to: lastRow,
          extra(r, c) {
            if (c !== 2) return null;
            if (r === FIRST_ROW) return "cell-source";
            if (r > FIRST_ROW && r <= state.lastRow) return state.done ? "cell-done" : "cell-pending";
            return null;
          },
        });
        return state;
      },
      customHandleKey(e, ctx) {
        if (navAndGuard(e, ctx)) return;
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        if (!ctrlOrMeta) {
          if (e.altKey) showWrongToast();
          return;
        }
        if (e.altKey) { showWrongToast(); return; }
        const isCopy = e.code === "KeyC" || e.key === "c" || e.key === "C";
        const isPaste = e.code === "KeyV" || e.key === "v" || e.key === "V";

        if (isCopy) {
          if (!selectionIs(ctx, FIRST_ROW, FIRST_ROW + 1, 2, 3)) {
            showWrongToast();
            return;
          }
          ctx.copied = true;
          ctx.cutRect = { r0: FIRST_ROW, r1: FIRST_ROW + 1, c0: 2, c1: 3 };
          renderCurrentTask();
          return;
        }

        if (isPaste) {
          if (ctx.done) return;
          if (!ctx.copied) return; // nichts kopiert – in Excel passiert dann auch nichts
          // Die orangen Zellen C6:C14 – die Ursprungszelle C5 darf mitmarkiert
          // sein, dann wird sie einfach mit demselben Text überschrieben.
          const okWithoutSource = selectionIs(ctx, FIRST_ROW + 1, ctx.lastRow + 1, 2, 3);
          const okWithSource = selectionIs(ctx, FIRST_ROW, ctx.lastRow + 1, 2, 3);
          if (!okWithoutSource && !okWithSource) {
            showWrongToast();
            return;
          }
          for (let r = FIRST_ROW + 1; r <= ctx.lastRow; r++) ctx.data[r][2] = ctx.ORDER;
          ctx.cutRect = null;
          ctx.done = true;
          handleSuccess();
          return;
        }
        showWrongToast();
      },
    },

    /* -----------------------------------------------------------
     * Blatt 7 (Finale) – Filter setzen (Strg+Umschalt+L), Filtermenü
     * der Status-Spalte mit Alt+Pfeil nach unten öffnen und wie in
     * echtem Excel über die Kontrollkästchen nur "ausgebüxt" stehen
     * lassen (Leertaste setzt/entfernt Haken), Enter bestätigt.
     * --------------------------------------------------------- */
    {
      instruction: "Filtere die Tabelle nach ausgebüxten Tieren.",
      keys: () => [[MOD, "Umschalt", "L"], ["Alt", "Pfeil ↓ / ↑"], ["Leertaste"], ["Enter"]],
      hint(ctx) {
        if (!ctx.filterOn) return "Alarm im Zoo! Schalte zuerst die Filter ein.";
        if (!ctx.dropdownOpen) return "Stell dich in die Überschrift Status und öffne das Filtermenü mit Alt+Pfeil nach unten.";
        return "Haken mit der Leertaste setzen oder entfernen – am Ende darf nur „ausgebüxt“ angehakt sein. Enter bestätigt.";
      },
      init() {
        const ANIMALS = [
          ["Löwe", "Leo", "im Gehege"],
          ["Giraffe", "Gisela", "ausgebüxt"],
          ["Pinguin", "Paul", "im Gehege"],
          ["Erdmännchen", "Egon", "ausgebüxt"],
          ["Waschbär", "Wastl", "ausgebüxt"],
          ["Elefant", "Elsa", "im Gehege"],
          ["Lama", "Lotte", "ausgebüxt"],
          ["Faultier", "Fritz", "schläft"],
          ["Otter", "Olli", "ausgebüxt"],
        ];
        const cols = 5, rows = FIRST_ROW + ANIMALS.length + 1;
        const lastRow = FIRST_ROW + ANIMALS.length - 1; // Zeile 13
        const data = sheetData(rows, cols);
        data[HEAD_ROW][1] = "Tier";
        data[HEAD_ROW][2] = "Name";
        data[HEAD_ROW][3] = "Status";
        ANIMALS.forEach((row, i) => {
          data[FIRST_ROW + i][1] = row[0];
          data[FIRST_ROW + i][2] = row[1];
          data[FIRST_ROW + i][3] = row[2];
        });

        const state = {
          cols, rows, data, ANIMALS, lastRow,
          filterCol: 3,
          // Reihenfolge wie im echten Excel-Filtermenü: (Alles auswählen) zuerst,
          // danach die Werte alphabetisch. Anfangs ist alles angehakt.
          values: ["ausgebüxt", "im Gehege", "schläft"],
          checked: [true, true, true],
          optionIndex: 0, // 0 = (Alles auswählen), danach 1..n = Werte
          filterOn: false, dropdownOpen: false, done: false,
          hiddenRows: null,
          // Start direkt in der Kopfzeile – von dort geht es mit den Pfeiltasten
          // zur Spalte Status, ohne vorher nach oben zu müssen.
          activeCell: { r: HEAD_ROW, c: 1 },
          selAnchor: { r: HEAD_ROW, c: 1 },
          colSelected: null, rowSelected: null, headerFilterRow: null,
          minVisualCols: cols,
          // Ein paar Zeilen mehr als die Tabelle braucht, damit das aufgeklappte
          // Filtermenü vollständig ins Rasterfenster passt und nicht abgeschnitten wird.
          minVisualRows: rows + 4,
          colWidth: (c) => (c === 0 ? 48 : null),
        };
        state.cellClass = tableLook({ cols: [1, 2, 3], from: FIRST_ROW, to: lastRow });
        return state;
      },
      renderOverlay(ctx) {
        if (!ctx.dropdownOpen) return "";
        // Linksbündig unter dem Filterknopf – rutscht aber nach links, wenn das
        // Menü sonst rechts aus dem Rasterfenster laufen würde (wie in Excel).
        const left = `min(calc(86px + ${ctx.filterCol} * (100% - 86px) / ${ctx.cols}), calc(100% - 276px))`;
        const top = `${30 + (HEAD_ROW + 1) * 30}px`;
        const allChecked = ctx.checked.every(Boolean);
        const box = (on, active) =>
          `<span class="xl-check${on ? " is-on" : ""}">${on ? "&#10003;" : ""}</span>`;
        const row = (label, on, idx) =>
          `<div class="xl-value${idx === ctx.optionIndex ? " is-active" : ""}">${box(on)}<span>${escapeHTML(label)}</span></div>`;

        let list = row("(Alles auswählen)", allChecked, 0);
        ctx.values.forEach((v, i) => { list += row(v, ctx.checked[i], i + 1); });

        return `<div class="xl-filter" style="left:${left};top:${top};">
            <div class="xl-menu">
              <div class="xl-item"><span class="xl-icon">A&#8595;</span>Von A bis Z sortieren</div>
              <div class="xl-item"><span class="xl-icon">Z&#8593;</span>Von Z bis A sortieren</div>
              <div class="xl-item is-muted"><span class="xl-icon"></span>Nach Farbe sortieren<span class="xl-more">&#9656;</span></div>
            </div>
            <div class="xl-menu">
              <div class="xl-item is-muted"><span class="xl-icon"></span>Filter aus &bdquo;Status&ldquo; l&ouml;schen</div>
              <div class="xl-item is-muted"><span class="xl-icon"></span>Nach Farbe filtern<span class="xl-more">&#9656;</span></div>
              <div class="xl-item is-muted"><span class="xl-icon"></span>Textfilter<span class="xl-more">&#9656;</span></div>
            </div>
            <div class="xl-search">Suchen</div>
            <div class="xl-list">${list}</div>
            <div class="xl-actions"><span class="xl-btn is-primary">OK</span><span class="xl-btn">Abbrechen</span></div>
          </div>`;
      },
      customHandleKey(e, ctx) {
        const key = e.key;

        // Solange das Filtermenü offen ist, gehören Pfeiltasten, Leertaste und
        // Enter dem Menü – wie in Excel.
        if (ctx.dropdownOpen) {
          e.preventDefault();
          const lastIndex = ctx.values.length;
          if (key === "ArrowDown" || key === "ArrowUp") {
            const dir = key === "ArrowDown" ? 1 : -1;
            ctx.optionIndex = Math.max(0, Math.min(lastIndex, ctx.optionIndex + dir));
            renderCurrentTask();
            return;
          }
          if (key === " " || e.code === "Space" || key === "Spacebar") {
            if (ctx.optionIndex === 0) {
              // "(Alles auswählen)" setzt oder entfernt alle Haken auf einmal.
              const turnOn = !ctx.checked.every(Boolean);
              ctx.checked = ctx.checked.map(() => turnOn);
            } else {
              ctx.checked[ctx.optionIndex - 1] = !ctx.checked[ctx.optionIndex - 1];
            }
            renderCurrentTask();
            return;
          }
          if (key === "Escape") {
            ctx.dropdownOpen = false;
            renderCurrentTask();
            return;
          }
          if (key === "Enter") {
            const wanted = ctx.values.map((v) => v === "ausgebüxt");
            const ok = ctx.checked.every((v, i) => v === wanted[i]);
            if (!ok) {
              showWrongToast();
              return;
            }
            const hidden = new Set();
            ctx.ANIMALS.forEach((row, i) => {
              if (row[2] !== "ausgebüxt") hidden.add(FIRST_ROW + i);
            });
            ctx.hiddenRows = hidden;
            ctx.dropdownOpen = false;
            ctx.done = true;
            handleSuccess();
            return;
          }
          return;
        }

        // Alt+Pfeil nach unten (oder nach oben) öffnet das Filtermenü der Spalte
        // unter dem Cursor – beides funktioniert auch in echtem Excel.
        if (e.altKey && (key === "ArrowDown" || key === "ArrowUp") && !e.ctrlKey && !e.metaKey) {
          e.preventDefault();
          if (e.repeat) return;
          if (!ctx.filterOn || ctx.activeCell.r !== HEAD_ROW || ctx.activeCell.c !== ctx.filterCol) {
            showWrongToast();
            return;
          }
          ctx.dropdownOpen = true;
          ctx.optionIndex = 0;
          renderCurrentTask();
          return;
        }

        if (navAndGuard(e, ctx)) return;

        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const isFilterToggle = ctrlOrMeta && e.shiftKey && !e.altKey &&
          (e.code === "KeyL" || e.key === "l" || e.key === "L");

        if (isFilterToggle) {
          ctx.filterOn = !ctx.filterOn;
          ctx.headerFilterRow = ctx.filterOn ? HEAD_ROW : null;
          if (!ctx.filterOn) ctx.hiddenRows = null;
          renderCurrentTask();
          return;
        }
        if (e.ctrlKey || e.metaKey || e.altKey) showWrongToast();
      },
    },
  ];

  // Restliche orange Zellen der ersten Aufgabe.
  function countOrange(ctx) {
    let n = 0;
    for (let r = 0; r < ctx.rows; r++) {
      for (let c = 0; c < ctx.cols; c++) if (ctx.orangeGrid[r][c]) n++;
    }
    return n;
  }

  /* ---------------------------------------------------------------
   * DOM-Referenzen
   * ------------------------------------------------------------- */
  const screens = {
    touch: document.getElementById("screen-touch"),
    start: document.getElementById("screen-start"),
    challenge: document.getElementById("screen-challenge"),
    result: document.getElementById("screen-result"),
  };
  const progressLabelEl = document.getElementById("progress-label");
  const taskPipsEl = document.getElementById("task-pips");
  const timerDisplayEl = document.getElementById("timer-display");
  const taskInstructionEl = document.getElementById("task-instruction");
  const taskBadgesEl = document.getElementById("task-badges");
  const taskHintEl = document.getElementById("task-hint");
  const nameboxEl = document.getElementById("namebox");
  const formulabarTextEl = document.getElementById("formulabar-text");
  const gridContainerEl = document.getElementById("grid-container");
  const ribbonContainerEl = document.getElementById("ribbon-container");
  const successCheckEl = document.getElementById("success-check");
  const wrongToastEl = document.getElementById("wrong-toast");
  const resultHeadlineEl = document.getElementById("result-headline");
  const resultTimeEl = document.getElementById("result-time");
  const resultCompareEl = document.getElementById("result-compare");
  const confettiLayerEl = document.getElementById("confetti-layer");
  const btnStart = document.getElementById("btn-start");
  const btnRetry = document.getElementById("btn-retry");
  const btnReset = document.getElementById("btn-reset");
  const topBarEl = document.querySelector(".top-bar");
  const startBestTimeEl = document.getElementById("start-best-time");
  const startParticipantsEl = document.getElementById("start-participants");
  const rankResultEl = document.getElementById("rank-result");
  const resultBadgeEl = document.getElementById("result-badge");
  const resultRankEl = document.getElementById("result-rank");
  const resultBestEl = document.getElementById("result-best");
  const resultRefEl = document.getElementById("result-ref");

  function showScreen(name) {
    Object.entries(screens).forEach(([key, el]) => {
      el.classList.toggle("active", key === name);
    });
    // Start- und Aufgaben-Screen haben ihr eigenes, full-bleed Logo — die globale Top-Bar würde es doppeln.
    topBarEl.classList.toggle("is-hidden", name === "start" || name === "challenge" || name === "result");
  }

  /* ---------------------------------------------------------------
   * State
   * ------------------------------------------------------------- */
  let currentTaskIndex = 0;
  let taskCtx = null;
  let challengeActive = false;
  let locked = false;
  let timerStartMs = 0;
  let timerInterval = null;
  let wrongToastTimer = null;

  function renderCurrentTask() {
    const task = TASKS[currentTaskIndex];
    progressLabelEl.textContent = `Aufgabe ${currentTaskIndex + 1}/${TASKS.length}`;
    taskPipsEl.innerHTML = buildPipsHTML(currentTaskIndex, TASKS.length);
    taskInstructionEl.textContent = task.instruction;
    taskBadgesEl.innerHTML = typeof task.keys === "function" ? buildBadgesHTML(task.keys()) : "";
    const hint = typeof task.hint === "function" ? task.hint(taskCtx) : task.hint;
    taskHintEl.textContent = hint || "";

    // renderOverlay ist der Haken für aufgabeneigene Einblendungen über dem Raster
    // (Filtermenü in Aufgabe 7, Dialog "Tabelle erstellen" in Aufgabe 8).
    const overlay = typeof task.renderOverlay === "function" ? task.renderOverlay(taskCtx) : "";
    gridContainerEl.innerHTML = buildGridHTML(taskCtx) + overlay;
    ribbonContainerEl.innerHTML = typeof task.renderRibbon === "function" ? task.renderRibbon(taskCtx) : "";

    const colSel = lineSelRange(taskCtx.colSelected);
    const rowSel = lineSelRange(taskCtx.rowSelected);
    if (colSel) {
      formulabarTextEl.textContent = colSel.from === colSel.to
        ? `Spalte ${colLetter(colSel.from)} markiert`
        : `Spalten ${colLetter(colSel.from)}:${colLetter(colSel.to)} markiert`;
    } else if (rowSel) {
      formulabarTextEl.textContent = rowSel.from === rowSel.to
        ? `Zeile ${rowSel.from + 1} markiert`
        : `Zeilen ${rowSel.from + 1}:${rowSel.to + 1} markiert`;
    } else if (taskCtx.formulaOverride != null) {
      formulabarTextEl.innerHTML = taskCtx.formulaOverride;
    } else {
      const val = taskCtx.activeCell
        ? (taskCtx.data[taskCtx.activeCell.r] && taskCtx.data[taskCtx.activeCell.r][taskCtx.activeCell.c]) || ""
        : "";
      formulabarTextEl.textContent = String(val).replace(/<[^>]*>/g, "");
    }

    nameboxEl.textContent = taskCtx.activeCell ? refStr(taskCtx.activeCell.r, taskCtx.activeCell.c) : "";
  }

  /* ---------------------------------------------------------------
   * Timer
   * ------------------------------------------------------------- */
  function startTimer() {
    timerStartMs = performance.now();
    updateTimerDisplay();
    timerInterval = setInterval(updateTimerDisplay, 50);
  }

  function updateTimerDisplay() {
    const elapsed = (performance.now() - timerStartMs) / 1000;
    timerDisplayEl.textContent = elapsed.toFixed(2);
  }

  function stopTimer() {
    clearInterval(timerInterval);
    timerInterval = null;
  }

  /* ---------------------------------------------------------------
   * Erfolg / Fehlversuch
   * ------------------------------------------------------------- */
  function showCheckmark() {
    successCheckEl.classList.add("show");
  }
  function hideCheckmark() {
    successCheckEl.classList.remove("show");
  }

  function showWrongToast() {
    clearTimeout(wrongToastTimer);
    wrongToastEl.classList.add("show");
    wrongToastTimer = setTimeout(() => wrongToastEl.classList.remove("show"), 1300);
  }

  function handleSuccess() {
    locked = true;
    const task = TASKS[currentTaskIndex];
    if (task.applySuccess) task.applySuccess(taskCtx);
    renderCurrentTask();
    showCheckmark();
    setTimeout(() => {
      hideCheckmark();
      advanceTask();
    }, 450);
  }

  function advanceTask() {
    currentTaskIndex++;
    if (currentTaskIndex >= TASKS.length) {
      finishChallenge();
    } else {
      taskCtx = TASKS[currentTaskIndex].init();
      renderCurrentTask();
      locked = false;
    }
  }

  function finishChallenge() {
    challengeActive = false;
    stopTimer();
    const finalSeconds = (performance.now() - timerStartMs) / 1000;
    progressLabelEl.textContent = `Aufgabe ${TASKS.length}/${TASKS.length}`;
    showResultScreen(finalSeconds);
  }

  function showResultScreen(finalSeconds) {
    rankResultEl.hidden = true;
    rankResultEl.textContent = "";
    resultRankEl.textContent = "–";
    resultBestEl.textContent = startBestTimeEl.textContent;
    animateResultTime(finalSeconds);
    const diff = Math.abs(finalSeconds - REFERENCE_TIME_SECONDS).toFixed(1);
    resultRefEl.textContent = REFERENCE_TIME_SECONDS.toFixed(1) + "s";
    const beatReference = finalSeconds < REFERENCE_TIME_SECONDS;
    if (beatReference) {
      resultHeadlineEl.textContent = "Stark gemacht!";
      resultBadgeEl.textContent = "\u{1F3C6}";
      resultCompareEl.textContent = `⚡ ${diff}s schneller als Flo`;
      resultCompareEl.className = "result-compare faster";
    } else {
      resultHeadlineEl.textContent = "Geschafft!";
      resultBadgeEl.textContent = "\u{1F3C1}";
      resultCompareEl.textContent = `${diff}s langsamer als Flo`;
      resultCompareEl.className = "result-compare slower";
    }
    showScreen("result");
    // Konfetti ist die besondere Belohnung dafür, Flos Referenzzeit geschlagen zu
    // haben (sofort bekannt) oder unter den Top 10 der Teilnehmer zu landen (erst
    // nach der Rang-Abfrage bekannt) - nicht für jeden Abschluss, sonst verliert
    // es seine Bedeutung. confettiFired verhindert ein doppeltes Auslösen, falls
    // beides zutrifft.
    let confettiFired = false;
    if (beatReference) {
      spawnConfetti();
      confettiFired = true;
    }

    recordRun(finalSeconds)
      .then((stats) => {
        if (!stats) return;
        // Rang/Vergleich beziehen sich auf die gespeicherte persönliche Bestzeit
        // (my_seconds), nicht zwingend auf diesen einen Lauf – ein schlechterer
        // Retry darf den bisherigen persönlichen Rang nicht verschlechtern.
        const mySeconds = Number(stats.my_seconds);
        const bestSeconds = Number(stats.best_seconds);
        const isTop = bestSeconds != null && mySeconds <= bestSeconds;
        // "Neue Bestzeit" darf nur stehen, wenn DIESER Lauf sie gesetzt hat - sonst
        // hält man z.B. mit einem alten 18s-Lauf weiter Platz 1, obwohl dieser
        // Versuch (20.8s) schlechter war, und "Neue Bestzeit!" wäre irreführend.
        const improvedThisRun = Math.abs(finalSeconds - mySeconds) < 0.005;
        const rank = isTop ? 1 : stats.rank;
        resultRankEl.innerHTML = `${rank}<small>/${stats.total}</small>`;
        if (!Number.isNaN(bestSeconds)) resultBestEl.textContent = fmtBestTime(bestSeconds);
        rankResultEl.hidden = false;
        if (isTop && improvedThisRun) {
          rankResultEl.textContent = "🏆 Neue Bestzeit! Du bist die Nr. 1.";
        } else if (isTop) {
          rankResultEl.textContent = `🏆 Deine Bestzeit (${mySeconds.toFixed(1)}s) ist weiterhin die schnellste.`;
        } else {
          const behind = (mySeconds - bestSeconds).toFixed(1);
          rankResultEl.textContent = improvedThisRun
            ? `🎉 Neue persönliche Bestzeit · noch ${behind}s bis Platz 1`
            : `Deine persönliche Bestzeit: ${mySeconds.toFixed(1)}s · noch ${behind}s bis Platz 1`;
        }
        if (!confettiFired && stats.rank <= 10) spawnConfetti();
      })
      .catch(() => {}); // best effort, keine Rang-Anzeige statt Fehlermeldung
  }

  // Zählt die Endzeit in ~0,9s von 0 hoch. setInterval statt requestAnimationFrame,
  // damit in Hintergrund-Tabs trotzdem der Endwert erscheint.
  let resultTimeTimer = null;
  function animateResultTime(target) {
    clearInterval(resultTimeTimer);
    const start = performance.now();
    const duration = 900;
    resultTimeEl.textContent = "0.0";
    resultTimeTimer = setInterval(() => {
      const t = Math.min(1, (performance.now() - start) / duration);
      const eased = 1 - Math.pow(1 - t, 3);
      resultTimeEl.textContent = (target * eased).toFixed(1);
      if (t >= 1) clearInterval(resultTimeTimer);
    }, 30);
  }

  const CONFETTI_COLORS =["#37874a", "#22452b", "#ff8a00", "#8fd3a2", "#f2c94c"];

  function spawnConfetti() {
    confettiLayerEl.innerHTML = "";
    const count = 70;
    for (let i = 0; i < count; i++) {
      const piece = document.createElement("span");
      piece.className = "confetti-piece";
      const left = Math.random() * 100;
      const drift = Math.round((Math.random() - 0.5) * 240);
      const spin = Math.round(360 + Math.random() * 360) * (Math.random() < 0.5 ? -1 : 1);
      const duration = 2400 + Math.random() * 1600;
      const delay = Math.random() * 500;
      const color = CONFETTI_COLORS[Math.floor(Math.random() * CONFETTI_COLORS.length)];
      piece.style.left = `${left}%`;
      piece.style.background = color;
      piece.style.borderRadius = Math.random() < 0.5 ? "50%" : "2px";
      piece.style.setProperty("--confetti-drift", `${drift}px`);
      piece.style.setProperty("--confetti-spin", `${spin}deg`);
      piece.style.animationDuration = `${duration}ms`;
      piece.style.animationDelay = `${delay}ms`;
      confettiLayerEl.appendChild(piece);
    }
  }

  /* ---------------------------------------------------------------
   * Keyboard-Handling
   * ------------------------------------------------------------- */
  const NAV_KEYS = new Set([" ", "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Home", "End", "PageUp", "PageDown", "Spacebar"]);
  const PURE_MODIFIER_KEYS = new Set(["Control", "Meta", "Alt", "Shift", "OS"]);

  function handleKeydown(e) {
    if (!challengeActive || locked) return;

    const task = TASKS[currentTaskIndex];
    if (task.customHandleKey) {
      task.customHandleKey(e, taskCtx);
      return;
    }

    const isModCombo = e.ctrlKey || e.metaKey || e.altKey;
    const isF4 = e.key === "F4";

    if (NAV_KEYS.has(e.key) || isModCombo || isF4) {
      e.preventDefault();
    }

    if (e.repeat) return;
    if (PURE_MODIFIER_KEYS.has(e.key)) return;
    if (!isModCombo && !isF4) return; // reine Navigation, kein Shortcut-Versuch

    if (task.isMatch(e)) {
      handleSuccess();
    } else {
      showWrongToast();
    }
  }

  document.addEventListener("keydown", handleKeydown);

  /* ---------------------------------------------------------------
   * Start / Retry
   * ------------------------------------------------------------- */
  function startChallenge() {
    currentTaskIndex = 0;
    taskCtx = TASKS[0].init();
    challengeActive = true;
    locked = false;
    showScreen("challenge");
    renderCurrentTask();
    startTimer();
  }

  function backToStart() {
    challengeActive = false;
    locked = false;
    stopTimer();
    clearTimeout(wrongToastTimer);
    wrongToastEl.classList.remove("show");
    hideCheckmark();
    showScreen("start");
  }

  btnStart.addEventListener("click", startChallenge);
  btnRetry.addEventListener("click", backToStart);
  btnReset.addEventListener("click", backToStart);

  document.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && screens.start.classList.contains("active")) {
      startChallenge();
    }
  });

  /* ---------------------------------------------------------------
   * Init
   * ------------------------------------------------------------- */
  loadBestTime();

  if (isTouchOnlyDevice()) {
    showScreen("touch");
  } else {
    showScreen("start");
  }
})();
