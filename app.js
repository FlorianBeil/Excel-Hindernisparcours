(() => {
  "use strict";

  /* ---------------------------------------------------------------
   * Referenzzeit von Excel.Flo – Platzhalter, hier leicht änderbar
   * ------------------------------------------------------------- */
  const REFERENCE_TIME_SECONDS = 28.0;

  /* ---------------------------------------------------------------
   * Ranking (Supabase, dasselbe Projekt wie das Übungsportal) – komplett
   * anonym: jeder abgeschlossene Lauf wird automatisch eingetragen, ohne
   * Name/E-Mail. Weder die Bestzeit- noch die Rang-Abfrage lesen die
   * Tabelle direkt – beides läuft über RPC-Funktionen, die nur
   * aggregierte Zahlen liefern (siehe hindernisparkour_runs.sql).
   * Alles best effort: schlägt die Verbindung fehl, bleibt einfach der
   * lokale Platzhalter "–" stehen bzw. die Rang-Anzeige leer.
   * ------------------------------------------------------------- */
  const SUPABASE_URL = "https://hbhagmmbowplzjzfvuao.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_Ee47HbgO5Ne8PP9Jh5ugag_iE_lZcp6";
  const RUNS_TABLE = "hindernisparkour_runs";

  let supabaseClient = null;
  function getSupabaseClient() {
    if (!window.supabase || !window.supabase.createClient) return null; // SDK nicht geladen (z.B. offline/geblockt)
    if (!supabaseClient) supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    return supabaseClient;
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
      .rpc("hindernisparkour_best_time")
      .then(({ data, error }) => {
        if (error || data == null) return;
        startBestTimeEl.textContent = fmtBestTime(Number(data));
      })
      .catch(() => {}); // best effort, Platzhalter "–" bleibt stehen
  }

  // Trägt den Lauf automatisch (anonym) ein und liefert Rang + Teilnehmerzahl
  // + aktuelle Bestzeit zurück – wird bei jedem Abschluss der Challenge aufgerufen.
  function recordRun(seconds) {
    const client = getSupabaseClient();
    if (!client) return Promise.reject(new Error("Supabase nicht verfügbar"));
    return client
      .from(RUNS_TABLE)
      .insert({ seconds: seconds })
      .then(({ error }) => {
        if (error) throw error;
        return client.rpc("hindernisparkour_stats", { p_seconds: seconds });
      })
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

  /* ---------------------------------------------------------------
   * Beispieldaten (aus dem Excel.Flo-Kursmaterial übernommen)
   * ------------------------------------------------------------- */
  const PEOPLE = [
    ["Anna", "Berlin", "Vertrieb", "ABC GmbH"],
    ["Ben", "Hamburg", "Entwicklung", "XYZ AG"],
    ["Carla", "München", "Marketing", "LMN GmbH"],
    ["David", "Köln", "Finanzen", "PQR AG"],
    ["Elena", "Stuttgart", "Personal", "RST GmbH"],
    ["Felix", "Frankfurt", "IT", "UVW AG"],
    ["Greta", "Düsseldorf", "Vertrieb", "JKL GmbH"],
    ["Hans", "Leipzig", "Entwicklung", "MNO AG"],
    ["Iris", "Dresden", "Marketing", "ABC GmbH"],
    ["Jonas", "Nürnberg", "Finanzen", "XYZ AG"],
    ["Karla", "Dortmund", "Personal", "LMN GmbH"],
    ["Lukas", "Essen", "IT", "PQR AG"],
    ["Mia", "Bremen", "Vertrieb", "RST GmbH"],
    ["Nico", "Hannover", "Entwicklung", "UVW AG"],
    ["Olivia", "Duisburg", "Marketing", "JKL GmbH"],
    ["Paul", "Bochum", "Finanzen", "MNO AG"],
    ["Quentin", "Wuppertal", "Personal", "ABC GmbH"],
    ["Rosa", "Bielefeld", "IT", "XYZ AG"],
    ["Simon", "Bonn", "Vertrieb", "LMN GmbH"],
    ["Tina", "Münster", "Entwicklung", "PQR AG"],
  ];

  const REVENUE_SERIES = [
    84000, 44000, 84000, 25000, 49000, 83000, 38000, 23000, 46000, 48000,
    85000, 63000, 82000, 71000, 23000, 19000, 18000, 24000, 27000, 47000,
    75000, 34000, 59000, 13000, 57000, 18000,
  ];

  const CURRENCY_VALUES = [
    59000, 99000, 19000, 26000, 50000, 47000, 82000, 24000, 40000, 95000,
    19000, 93000, 27000, 61000, 33000,
  ];

  const MONTHLY = [
    ["Januar", 85000], ["Februar", 84000], ["März", 84000], ["April", 83000],
    ["Mai", 82000], ["Juni", 75000], ["Juli", 71000], ["August", 63000],
  ];

  const PRODUCTS = [
    "Tastatur", "Monitor", "Laptop-Tasche", "Maus", "Dockingstation",
    "Headset", "Webcam", "USB-Hub", "Notebook-Ständer", "Grafiktablett",
    "Drucker", "Scanner", "Router", "Mikrofon", "Lautsprecher",
  ];

  const REGIONS = ["Nord", "Süd", "Ost", "West"];

  function fmt(n) {
    return typeof n === "number" ? n.toLocaleString("de-DE") : n;
  }

  function fmtCurrencyDE(n) {
    return n.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
  }

  function fmtPercentDE(n) {
    return n.toLocaleString("de-DE", { style: "percent", minimumFractionDigits: 2 });
  }

  // Excel-Datumsseriennummer (Tage seit 1899-12-30) -> deutsches Datumsformat tt.mm.jjjj
  function fmtDateDE(serial) {
    const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    return `${dd}.${mm}.${d.getUTCFullYear()}`;
  }

  function fmtCurrency0DE(n) {
    return n.toLocaleString("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
  }

  function fmtPercent0(n) {
    return Math.round(n * 100) + "%";
  }

  // Nachbildung von Excels Standard-3-Farbskala (bedingte Formatierung): Rot (Minimum) -> Gelb (Mitte) -> Grün (Maximum).
  const COLOR_SCALE_STOPS = [
    [0xf8, 0x69, 0x6b],
    [0xff, 0xeb, 0x84],
    [0x63, 0xbe, 0x7b],
  ];
  function colorScaleHex(value, min, max) {
    const t = max === min ? 0.5 : (value - min) / (max - min);
    const [from, to] = t <= 0.5 ? [COLOR_SCALE_STOPS[0], COLOR_SCALE_STOPS[1]] : [COLOR_SCALE_STOPS[1], COLOR_SCALE_STOPS[2]];
    const localT = t <= 0.5 ? t * 2 : (t - 0.5) * 2;
    const rgb = from.map((v, i) => Math.round(v + (to[i] - v) * localT));
    return "#" + rgb.map((v) => v.toString(16).padStart(2, "0")).join("");
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
      while (inBounds(rr + dr, cc + dc) && !isFilled(rr + dr, cc + dc)) { rr += dr; cc += dc; }
      if (isFilled(rr + dr, cc + dc)) { rr += dr; cc += dc; }
    }
    return { r: rr, c: cc };
  }

  function buildGridHTML(ctx) {
    const { cols, rows, data, activeCell, colSelected, rowSelected, headerFilterRow, styledRange, columnColors, rowColors, zone, orangeGrid, cutRect, selAnchor, wideCols, cellClass, cellStyle, colWidth, minVisualCols, minVisualRows } = ctx;
    // Bei Aufgaben mit Umschalt+Pfeiltasten-Markierung (selAnchor gesetzt) wird die aktuelle Markierung
    // aus Anker- und aktiver Zelle live berechnet, statt sie separat im State mitzuführen. Bei einer
    // reinen 1-Zellen-"Markierung" (kein Umschalt genutzt) wird nichts zusätzlich hervorgehoben –
    // die normale aktive-Zelle-Markierung reicht dafür aus.
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
      if (colSelected === c) headClasses.push("col-selected");
      if (activeCell && activeCell.c === c) headClasses.push("active-col-head");
      if (columnColors && columnColors[c]) headClasses.push("col-" + columnColors[c]);
      const w = typeof colWidth === "function" ? colWidth(c) : null;
      const headStyle = w ? ` style="width:${w}px;"` : "";
      html += `<th class="${headClasses.join(" ")}"${headStyle}>${colLetter(c)}</th>`;
    }
    html += "</tr></thead><tbody>";
    for (let r = 0; r < visualRows; r++) {
      const rowheadClasses = ["rowhead"];
      if (rowSelected === r) rowheadClasses.push("row-selected");
      if (activeCell && activeCell.r === r) rowheadClasses.push("active-row-head");
      if (rowColors && rowColors[r]) rowheadClasses.push("row-" + rowColors[r]);
      html += `<tr><td class="${rowheadClasses.join(" ")}">${r + 1}</td>`;
      for (let c = 0; c < visualCols; c++) {
        const classes = ["cell"];
        if (activeCell && activeCell.r === r && activeCell.c === c) classes.push("active");
        if (colSelected === c) classes.push("col-selected");
        if (columnColors && columnColors[c]) classes.push("col-" + columnColors[c]);
        if (rowSelected === r) classes.push("row-selected");
        if (rowColors && rowColors[r]) classes.push("row-" + rowColors[r]);
        if (typeof cellClass === "function") {
          const extra = cellClass(r, c);
          if (extra) classes.push(extra);
        }
        if (orangeGrid && orangeGrid[r] && orangeGrid[r][c]) {
          classes.push("piece-orange");
        } else if (zone && inRange(r, c, zone)) {
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
   * Menüband-Mockup für Aufgabe 5 – zeigt sichtbar, wie man über Alt-Tastenkürzel
   * (Key-Tips) zum Sortieren-Befehl kommt, passend zum Fortschritt in ctx.keytipStage.
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
  const SORT_TAB_LABEL = "Start"; // R = Start; hier lebt in dieser Übung "Sortieren und Filtern"

  function buildRibbonHTML(ctx) {
    const stage = ctx.keytipStage;
    const sortTabActive = stage >= 2;

    let html = '<div class="ribbon-tabs">';
    RIBBON_TABS.forEach((t) => {
      const isActive = t.label === SORT_TAB_LABEL;
      const badge = stage === 1 ? `<span class="ribbon-keytip">${t.letter}</span>` : "";
      html += `<div class="ribbon-tab${isActive ? " active" : ""}">${t.label}${badge}</div>`;
    });
    html += "</div>";

    if (!sortTabActive) {
      html += `<div class="ribbon-content"><div class="ribbon-group">
        <div class="ribbon-btn">📋 Einfügen</div>
        <div class="ribbon-btn">✂️ Ausschneiden</div>
        <div class="ribbon-btn">🖌️ Format übertragen</div>
      </div></div>`;
      return html;
    }

    const sortBadge = stage <= 3 ? '<span class="ribbon-keytip">OS</span>' : "";
    html += `<div class="ribbon-content"><div class="ribbon-group">
        <div class="ribbon-btn">🔤 Rechtschreibung</div>
        <div class="ribbon-btn">💬 Kommentare</div>
        <div class="ribbon-btn ribbon-btn-target">⇅ Sortieren und Filtern${sortBadge}</div>
      </div>`;

    if (stage === 4) {
      html += `<div class="ribbon-dropdown">
          <div class="ribbon-dropdown-item">Aufsteigend sortieren <span class="ribbon-keytip">S</span></div>
          <div class="ribbon-dropdown-item">Absteigend sortieren <span class="ribbon-keytip">O</span></div>
        </div>`;
    }
    html += "</div>";
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
   * Die 7 Aufgaben
   * ------------------------------------------------------------- */
  const TASKS = [
    // Aufgabe 1 – Strg+Leertaste (Spalte markieren) + Strg+Minus (löschen), für jede orangene Spalte
    {
      instruction: "Lösche die orangenen Spalten.",
      shortcutLabel: () => `${MOD} + Leertaste → ${MOD} + Minus (-)`,
      keys: () => [[MOD, "Leertaste"], [MOD, "Minus (-)"]],
      hint(ctx) {
        return `Navigiere mit den Pfeiltasten zur nächsten orangenen Spalte. (${ctx.clearedCount}/${ctx.totalOrange} gelöscht)`;
      },
      init() {
        const cols = 11, rows = 13; // A–K: F folgen weitere leere Spalten, wie in einem echten Excel-Blatt
        const columnColors = ["orange", "green", "orange", "green", "orange", "green"];
        const data = [];
        for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(""));
        return {
          cols, rows, data,
          activeCell: { r: 0, c: 0 },
          colSelected: null,
          headerFilterRow: null,
          styledRange: null,
          columnColors,
          clearedCount: 0,
          totalOrange: 3,
        };
      },
      // Diese Aufgabe hat einen mehrstufigen Ablauf (Spalte wählen → löschen, dreimal),
      // deshalb übernimmt sie ihre eigene Tastatur-Logik statt isMatch/applySuccess.
      customHandleKey(e, ctx) {
        const key = e.key;
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        const modOnly = (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey;
        const isAnyModCombo = e.ctrlKey || e.metaKey || e.altKey;
        const isSelectCombo = modOnly && (key === " " || e.code === "Space");
        const isDeleteCombo = modOnly && (key === "-" || e.code === "Minus" || e.code === "NumpadSubtract");

        if (isArrow || isAnyModCombo) {
          e.preventDefault();
        }
        if (PURE_MODIFIER_KEYS.has(key)) return;

        if (isArrow) {
          // Gehaltene Pfeiltaste soll weiter navigieren (Auto-Repeat), nur die Aktions-Shortcuts
          // unten sollen bei gehaltener Taste NICHT mehrfach feuern.
          ctx.colSelected = null;
          let { r, c } = ctx.activeCell;
          if (key === "ArrowLeft") c = Math.max(0, c - 1);
          if (key === "ArrowRight") c = Math.min(ctx.cols - 1, c + 1);
          if (key === "ArrowUp") r = Math.max(0, r - 1);
          if (key === "ArrowDown") r = Math.min(ctx.rows - 1, r + 1);
          ctx.activeCell = { r, c };
          renderCurrentTask();
          return;
        }

        if (e.repeat) return;

        if (isSelectCombo) {
          ctx.colSelected = ctx.activeCell.c;
          renderCurrentTask();
          return;
        }

        if (isDeleteCombo) {
          if (ctx.colSelected == null) return;
          const col = ctx.colSelected;
          if (ctx.columnColors[col] === "orange") {
            // Spalte komplett entfernen (nicht nur einfärben) – alles rechts rückt nach, wie beim echten Löschen einer Spalte
            ctx.columnColors.splice(col, 1);
            ctx.data.forEach((row) => row.splice(col, 1));
            ctx.cols -= 1;
            if (ctx.activeCell.c > col) ctx.activeCell.c -= 1;
            ctx.activeCell.c = Math.min(ctx.activeCell.c, ctx.cols - 1);
            ctx.clearedCount++;
            ctx.colSelected = null;
            if (ctx.clearedCount >= ctx.totalOrange) {
              handleSuccess();
            } else {
              renderCurrentTask();
            }
          } else {
            ctx.colSelected = null;
            showWrongToast();
            renderCurrentTask();
          }
          return;
        }

        if (isAnyModCombo) {
          showWrongToast();
        }
      },
    },

    // Aufgabe 2 – Umschalt+Leertaste (Zeile markieren) + Strg+Minus (löschen) – wie Aufgabe 1, nur horizontal
    {
      instruction: "Lösche die orangenen Zeilen.",
      shortcutLabel: () => `Umschalt + Leertaste → ${MOD} + Minus (-)`,
      keys: () => [["Umschalt", "Leertaste"], [MOD, "Minus (-)"]],
      hint(ctx) {
        return `Navigiere mit den Pfeiltasten zur nächsten orangenen Zeile. (${ctx.clearedCount}/${ctx.totalOrange} gelöscht)`;
      },
      init() {
        const cols = 6, rows = 6;
        const rowColors = ["orange", "green", "orange", "green", "orange", "green"];
        const data = [];
        for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(""));
        return {
          cols, rows, data,
          activeCell: { r: 0, c: 0 },
          colSelected: null,
          rowSelected: null,
          headerFilterRow: null,
          styledRange: null,
          rowColors,
          clearedCount: 0,
          totalOrange: 3,
          // Feste Ausgangshöhe: löschen der Zeilen darf das Grid-Fenster nicht schrumpfen lassen.
          minVisualRows: rows,
        };
      },
      // Spiegelbild von Aufgabe 1: statt Spalten werden hier Zeilen markiert und gelöscht.
      customHandleKey(e, ctx) {
        const key = e.key;
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        const isAnyModCombo = e.ctrlKey || e.metaKey || e.altKey || e.shiftKey;
        const isSelectRowCombo = e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey && (key === " " || e.code === "Space");
        const isDeleteCombo = (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && (key === "-" || e.code === "Minus" || e.code === "NumpadSubtract");

        if (isArrow || isAnyModCombo) {
          e.preventDefault();
        }
        if (PURE_MODIFIER_KEYS.has(key)) return;

        if (isArrow) {
          ctx.rowSelected = null;
          let { r, c } = ctx.activeCell;
          if (key === "ArrowLeft") c = Math.max(0, c - 1);
          if (key === "ArrowRight") c = Math.min(ctx.cols - 1, c + 1);
          if (key === "ArrowUp") r = Math.max(0, r - 1);
          if (key === "ArrowDown") r = Math.min(ctx.rows - 1, r + 1);
          ctx.activeCell = { r, c };
          renderCurrentTask();
          return;
        }

        if (e.repeat) return;

        if (isSelectRowCombo) {
          ctx.rowSelected = ctx.activeCell.r;
          renderCurrentTask();
          return;
        }

        if (isDeleteCombo) {
          if (ctx.rowSelected == null) return;
          const row = ctx.rowSelected;
          if (ctx.rowColors[row] === "orange") {
            // Zeile komplett entfernen – alles darunter rückt nach, wie beim echten Löschen einer Zeile
            ctx.rowColors.splice(row, 1);
            ctx.data.splice(row, 1);
            ctx.rows -= 1;
            if (ctx.activeCell.r > row) ctx.activeCell.r -= 1;
            ctx.activeCell.r = Math.min(ctx.activeCell.r, ctx.rows - 1);
            ctx.clearedCount++;
            ctx.rowSelected = null;
            if (ctx.clearedCount >= ctx.totalOrange) {
              handleSuccess();
            } else {
              renderCurrentTask();
            }
          } else {
            ctx.rowSelected = null;
            showWrongToast();
            renderCurrentTask();
          }
          return;
        }

        if (isAnyModCombo) {
          showWrongToast();
        }
      },
    },

    // Aufgabe 3 – Zellbereich mit Umschalt+Pfeiltasten markieren, dann Strg+X (ausschneiden) + Strg+V (einfügen):
    // orangene Bereiche in den grünen Zielbereich verschieben. Es wird immer exakt die aktuelle Markierung
    // ausgeschnitten (eine einzelne Zelle oder ein mit Umschalt+Pfeiltasten erweiterter Bereich) – kein Auto-Erkennen
    // ganzer Blöcke wie in Aufgabe 1/2.
    {
      instruction: "Verschiebe die orangenen Bereiche in den grünen Zielbereich.",
      shortcutLabel: () => `Umschalt + Pfeiltasten → ${MOD} + X → ${MOD} + V`,
      keys: () => [["Umschalt", "Pfeiltasten"], [MOD, "X"], [MOD, "V"]],
      hint(ctx) {
        let remaining = 0;
        ctx.orangeGrid.forEach((row, r) => {
          row.forEach((isOrange, c) => {
            if (isOrange && !inRange(r, c, ctx.zone)) remaining++;
          });
        });
        return `Markiere mit Umschalt+Pfeiltasten genau den Bereich, den du verschieben willst (auch einzelne Zellen gehen). Noch ${remaining} orangene Zellen außerhalb des grünen Bereichs.`;
      },
      init() {
        const cols = 6, rows = 10;
        const data = [];
        for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(""));
        // r1/c1 sind exklusiv. B2:B4, D2:F4 (0-indexiert: A=0, B=1, C=2, D=3, E=4, F=5)
        const ranges = [
          { r0: 1, r1: 4, c0: 1, c1: 2 },
          { r0: 1, r1: 4, c0: 3, c1: 6 },
        ];
        const orangeGrid = Array.from({ length: rows }, () => new Array(cols).fill(false));
        ranges.forEach((rng) => {
          for (let r = rng.r0; r < rng.r1; r++) {
            for (let c = rng.c0; c < rng.c1; c++) orangeGrid[r][c] = true;
          }
        });
        const zone = { r0: 5, r1: 8, c0: 1, c1: 5 }; // B6:E8
        return {
          cols, rows, data,
          activeCell: { r: 0, c: 0 },
          selAnchor: { r: 0, c: 0 },
          colSelected: null, rowSelected: null,
          headerFilterRow: null, styledRange: null,
          orangeGrid, zone,
          cutRect: null,
          cutPattern: null,
          // Eine weitere leere Spalte rechts neben F, damit das Blatt nicht am Rand endet.
          minVisualCols: cols + 1,
        };
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        const lower = key.toLowerCase();
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        const modOnly = (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey;
        const isAnyModCombo = e.ctrlKey || e.metaKey || e.altKey || e.shiftKey;
        const isCut = modOnly && lower === "x";
        const isPaste = modOnly && lower === "v";

        if (isArrow || isAnyModCombo) {
          e.preventDefault();
        }
        if (PURE_MODIFIER_KEYS.has(key)) return;

        if (isArrow) {
          let { r, c } = ctx.activeCell;
          if (key === "ArrowLeft") c = Math.max(0, c - 1);
          if (key === "ArrowRight") c = Math.min(ctx.cols - 1, c + 1);
          if (key === "ArrowUp") r = Math.max(0, r - 1);
          if (key === "ArrowDown") r = Math.min(ctx.rows - 1, r + 1);
          ctx.activeCell = { r, c };
          if (!e.shiftKey) ctx.selAnchor = { r, c }; // ohne Umschalt: Markierung kollabiert auf die neue Zelle
          renderCurrentTask();
          return;
        }

        if (e.repeat) return;

        if (isCut) {
          const sel = normalizeRange(ctx.selAnchor, ctx.activeCell);
          const pattern = [];
          for (let r = sel.r0; r < sel.r1; r++) {
            const row = [];
            for (let c = sel.c0; c < sel.c1; c++) row.push(!!(ctx.orangeGrid[r] && ctx.orangeGrid[r][c]));
            pattern.push(row);
          }
          ctx.cutRect = sel;
          ctx.cutPattern = pattern;
          renderCurrentTask();
          return;
        }

        if (isPaste) {
          if (!ctx.cutRect) return;
          const rect = ctx.cutRect, pattern = ctx.cutPattern;
          const h = rect.r1 - rect.r0, w = rect.c1 - rect.c0;
          const tR0 = ctx.activeCell.r, tC0 = ctx.activeCell.c;
          const tR1 = tR0 + h, tC1 = tC0 + w;
          const withinZone = tR0 >= ctx.zone.r0 && tR1 <= ctx.zone.r1 && tC0 >= ctx.zone.c0 && tC1 <= ctx.zone.c1;
          let overlap = false;
          for (let r = tR0; r < tR1 && !overlap; r++) {
            for (let c = tC0; c < tC1 && !overlap; c++) {
              const insideSource = r >= rect.r0 && r < rect.r1 && c >= rect.c0 && c < rect.c1;
              if (!insideSource && ctx.orangeGrid[r][c]) overlap = true;
            }
          }
          if (!withinZone || overlap) {
            showWrongToast();
            return;
          }
          for (let r = rect.r0; r < rect.r1; r++) {
            for (let c = rect.c0; c < rect.c1; c++) {
              if (pattern[r - rect.r0][c - rect.c0]) ctx.orangeGrid[r][c] = false;
            }
          }
          for (let r = tR0; r < tR1; r++) {
            for (let c = tC0; c < tC1; c++) {
              if (pattern[r - tR0][c - tC0]) ctx.orangeGrid[r][c] = true;
            }
          }
          ctx.cutRect = null;
          ctx.cutPattern = null;
          ctx.activeCell = { r: tR0, c: tC0 };
          ctx.selAnchor = { r: tR0, c: tC0 };

          const remainingOutside = ctx.orangeGrid.some((row, r) =>
            row.some((isOrange, c) => isOrange && !inRange(r, c, ctx.zone))
          );
          if (!remainingOutside) {
            handleSuccess();
          } else {
            renderCurrentTask();
          }
          return;
        }

        if (isAnyModCombo) {
          showWrongToast();
        }
      },
    },

    // Aufgabe 4 – Strg+Umschalt+4 (Währung), Strg+Umschalt+5 (Prozent), Strg+# (Datum) – spaltenweise formatieren
    {
      instruction: "Formatiere die Daten spalte für spalte.",
      shortcutLabel: () => `${MOD}+Umschalt+4 → ${MOD}+Umschalt+5 → ${MOD}+#`,
      keys: () => [[MOD, "Umschalt", "4"], [MOD, "Umschalt", "5"], [MOD, "#"]],
      hint(ctx) {
        const mark = (done) => (done ? "✓" : "—");
        return `B: Währung ${mark(ctx.formatted.b)} · C: Prozent ${mark(ctx.formatted.c)} · D: Datum ${mark(ctx.formatted.d)}`;
      },
      init() {
        const cols = 4, rows = 9;
        const RAW_CURRENCY = [59000, 99000, 19000, 26000, 50000, 47000, 82000];
        const RAW_PERCENT = [0.25, 0.1, 0.11, 0.16, 0.28, 0.11, 0.28];
        const RAW_DATE = [45344, 45323, 45348, 45599, 45378, 45633, 45523];
        const data = [];
        for (let r = 0; r < rows; r++) data.push(["", "", "", ""]);
        data[1] = ["", "💶 Währungsformat", "% Prozentformat", "📅 Datumsformat"];
        for (let i = 0; i < 7; i++) {
          data[2 + i] = [
            "",
            String(RAW_CURRENCY[i]),
            String(RAW_PERCENT[i]).replace(".", ","),
            String(RAW_DATE[i]),
          ];
        }
        // Pro Spalte (Index 1/2/3 = B/C/D) wird für jede der 7 Datenzeilen einzeln festgehalten,
        // ob sie schon korrekt formatiert wurde – so kann die Formatierung auch in mehreren
        // kleineren Markierungen statt in einem einzigen Rutsch pro Spalte erledigt werden.
        const done = { 1: new Array(7).fill(false), 2: new Array(7).fill(false), 3: new Array(7).fill(false) };
        return {
          cols, rows, data,
          activeCell: { r: 2, c: 1 },
          selAnchor: { r: 2, c: 1 },
          colSelected: null, rowSelected: null,
          headerFilterRow: null, styledRange: null,
          rawCurrency: RAW_CURRENCY, rawPercent: RAW_PERCENT, rawDate: RAW_DATE,
          done,
          formatted: { b: false, c: false, d: false },
          wideCols: true,
          // Leere Spalte E rechts und leere Zeile 10 unten, damit das Blatt nicht am Datenrand endet.
          minVisualCols: cols + 1,
          minVisualRows: rows + 1,
          // Spalte B ("💶 Währungsformat") ist mit Icon länger als die anderen Überschriften und
          // wurde in der 132px-Standardbreite abgeschnitten - deshalb hier gezielt breiter.
          colWidth(c) {
            return c === 1 ? 190 : null;
          },
          cellClass(r, c) {
            if (c !== 1 && c !== 2 && c !== 3) return null;
            if (r === 1) return (done[c].every(Boolean) ? "cell-done" : "cell-pending") + " cell-header-bold";
            if (r >= 2 && r <= 8) return done[c][r - 2] ? "cell-done" : "cell-pending";
            return null;
          },
        };
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const isAnyModCombo = e.ctrlKey || e.metaKey || e.altKey || e.shiftKey;
        // Bei gehaltener Umschalt-Taste liefert e.key auf so gut wie jeder Tastaturbelegung (auch
        // deutsch und US) das verschobene Symbol der Zifferntaste ("$" bzw. "%"), nicht die Ziffer
        // selbst – e.code bleibt dagegen unabhängig von Umschalt/Layout immer "Digit4"/"Digit5" und
        // ist deshalb die zuverlässige Prüfung für Strg+Umschalt+4/5.
        const isCurrency = ctrlOrMeta && e.shiftKey && !e.altKey && (e.code === "Digit4" || key === "4" || key === "$");
        const isPercent = ctrlOrMeta && e.shiftKey && !e.altKey && (e.code === "Digit5" || key === "5" || key === "%");
        // Auf deutschen Tastaturen sendet Strg+Umschalt+3 (das Datumsformat) wegen eines Tastatur-
        // Layout-Effekts "#" statt "3" – deshalb wird hier bewusst nur auf Strg+# geprüft, Umschalt
        // wird nicht vorausgesetzt.
        const isDate = ctrlOrMeta && !e.altKey && key === "#";

        if (isArrow || isAnyModCombo) {
          e.preventDefault();
        }
        if (PURE_MODIFIER_KEYS.has(key)) return;

        if (isArrow) {
          const delta = ARROW_DELTA[key];
          if (ctrlOrMeta && e.shiftKey) {
            // Strg+Umschalt+Pfeil: Markierung bis zum Rand des zusammenhängenden Datenbereichs erweitern
            ctx.activeCell = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
          } else if (ctrlOrMeta) {
            // Strg+Pfeil (ohne Umschalt): wie echtes Excel zum Rand springen, aber Markierung kollabiert
            const edge = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
            ctx.activeCell = edge;
            ctx.selAnchor = edge;
          } else {
            let r = Math.max(0, Math.min(ctx.rows - 1, ctx.activeCell.r + delta.dr));
            let c = Math.max(0, Math.min(ctx.cols - 1, ctx.activeCell.c + delta.dc));
            ctx.activeCell = { r, c };
            if (!e.shiftKey) ctx.selAnchor = { r, c }; // ohne Umschalt: Markierung kollabiert auf die neue Zelle
          }
          renderCurrentTask();
          return;
        }

        if (e.repeat) return;

        const applyToSelection = (col, flagKey, formatFn, rawArr) => {
          const sel = normalizeRange(ctx.selAnchor, ctx.activeCell);
          if (sel.c0 !== col || sel.c1 !== col + 1) {
            showWrongToast();
            return;
          }
          const rowLo = Math.max(sel.r0, 2), rowHi = Math.min(sel.r1, 9);
          if (rowLo >= rowHi) {
            showWrongToast();
            return;
          }
          for (let r = rowLo; r < rowHi; r++) {
            ctx.data[r][col] = formatFn(rawArr[r - 2]);
            ctx.done[col][r - 2] = true;
          }
          ctx.formatted[flagKey] = ctx.done[col].every(Boolean);
          if (ctx.formatted.b && ctx.formatted.c && ctx.formatted.d) {
            handleSuccess();
          } else {
            renderCurrentTask();
          }
        };

        if (isCurrency) {
          applyToSelection(1, "b", fmtCurrencyDE, ctx.rawCurrency);
          return;
        }
        if (isPercent) {
          applyToSelection(2, "c", fmtPercentDE, ctx.rawPercent);
          return;
        }
        if (isDate) {
          applyToSelection(3, "d", fmtDateDE, ctx.rawDate);
          return;
        }

        if (isAnyModCombo) {
          showWrongToast();
        }
      },
    },

    // Aufgabe 5 – Alt,R,O,S,S (aufsteigend sortieren) bzw. Alt,R,O,S,O (absteigend) – Tasten NACHEINANDER,
    // wie ein echtes Excel-Ribbon-Tastenkürzel (Key-Tips), nicht gleichzeitig gedrückt.
    {
      instruction: "Sortiere Spalte B aufsteigend und Spalte D absteigend.",
      shortcutLabel: () => `Alt → R → O → S → S (auf) bzw. → O (ab)`,
      keys: () => [["Alt"], ["R"], ["O"], ["S"], ["S (auf) / O (ab)"]],
      hint(ctx) {
        const stageNames = ["", "Alt", "Alt→R", "Alt→R→O", "Alt→R→O→S"];
        const progress = ctx.keytipStage > 0 ? ` — bisher: ${stageNames[ctx.keytipStage]}` : "";
        return `B aufsteigend ${ctx.sortedB ? "✓" : "—"} · D absteigend ${ctx.sortedD ? "✓" : "—"}${progress}`;
      },
      renderRibbon: (ctx) => buildRibbonHTML(ctx),
      init() {
        const cols = 4, rows = 11;
        const RAW_ASC = [84000, 44000, 84000, 25000, 49000, 83000, 38000];
        const RAW_DESC = [0.02, 0.16, 0.02, 0.18, 0.08, 0.02, 0.16];
        const SORTED_ASC = [...RAW_ASC].sort((a, b) => a - b);
        const SORTED_DESC = [...RAW_DESC].sort((a, b) => b - a);
        const minAsc = Math.min(...RAW_ASC), maxAsc = Math.max(...RAW_ASC);
        const minDesc = Math.min(...RAW_DESC), maxDesc = Math.max(...RAW_DESC);

        const data = [];
        for (let r = 0; r < rows; r++) data.push(["", "", "", ""]);
        data[1] = ["", "Aufsteigend", "", "Absteigend"];
        for (let i = 0; i < 7; i++) {
          data[2 + i] = ["", fmtCurrency0DE(RAW_ASC[i]), "", fmtPercent0(RAW_DESC[i])];
        }

        const state = { rawB: RAW_ASC.slice(), rawD: RAW_DESC.slice() };

        return {
          cols, rows, data,
          activeCell: { r: 2, c: 1 },
          colSelected: null, rowSelected: null,
          headerFilterRow: null, styledRange: null,
          keytipStage: 0,
          sortedB: false, sortedD: false,
          SORTED_ASC, SORTED_DESC,
          wideCols: true,
          // Leere Spalte E rechts; Zeile 10/11 bleiben ohnehin leer (nur 7 statt 9 Werte befüllt).
          minVisualCols: cols + 1,
          cellClass(r, c) {
            if (r === 1 && (c === 1 || c === 3)) return "cell-header-bold";
            return null;
          },
          cellStyle(r, c) {
            if (r < 2 || r > 8) return null;
            if (c === 1) return `background:${colorScaleHex(state.rawB[r - 2], minAsc, maxAsc)};`;
            if (c === 3) return `background:${colorScaleHex(state.rawD[r - 2], minDesc, maxDesc)};`;
            return null;
          },
          _state: state,
        };
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        const lower = key.toLowerCase();
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        const isTab = key === "Tab";

        if (isTab) {
          e.preventDefault();
          ctx.keytipStage = 0;
          let { r, c } = ctx.activeCell;
          // Tab bewegt wie in echtem Excel nach rechts, Umschalt+Tab nach links.
          c = e.shiftKey ? Math.max(0, c - 1) : Math.min(ctx.cols - 1, c + 1);
          ctx.activeCell = { r, c };
          renderCurrentTask();
          return;
        }

        if (isArrow) {
          e.preventDefault();
          ctx.keytipStage = 0; // Navigation bricht eine begonnene Tastenfolge ab, wie in echtem Excel
          if (e.ctrlKey || e.metaKey) {
            const delta = ARROW_DELTA[key];
            ctx.activeCell = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
          } else {
            let { r, c } = ctx.activeCell;
            if (key === "ArrowLeft") c = Math.max(0, c - 1);
            if (key === "ArrowRight") c = Math.min(ctx.cols - 1, c + 1);
            if (key === "ArrowUp") r = Math.max(0, r - 1);
            if (key === "ArrowDown") r = Math.min(ctx.rows - 1, r + 1);
            ctx.activeCell = { r, c };
          }
          renderCurrentTask();
          return;
        }

        if (e.repeat) return;

        if (ctx.keytipStage === 0) {
          if (key === "Alt") {
            e.preventDefault();
            ctx.keytipStage = 1;
            renderCurrentTask();
            return;
          }
          if (PURE_MODIFIER_KEYS.has(key)) return;
          if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) {
            e.preventDefault();
            showWrongToast();
          }
          return;
        }

        if (PURE_MODIFIER_KEYS.has(key)) return; // z.B. Alt weiterhin gehalten – ignorieren, nicht als Fehleingabe werten
        e.preventDefault();

        const EXPECTED = ["", "r", "o", "s"];
        if (ctx.keytipStage >= 1 && ctx.keytipStage <= 3) {
          if (lower === EXPECTED[ctx.keytipStage]) {
            ctx.keytipStage += 1;
            renderCurrentTask();
          } else {
            ctx.keytipStage = 0;
            showWrongToast();
            renderCurrentTask();
          }
          return;
        }

        // keytipStage === 4: letzte Taste entscheidet aufsteigend ("s") vs. absteigend ("o")
        if (lower === "s" || lower === "o") {
          const ascending = lower === "s";
          ctx.keytipStage = 0;
          const col = ctx.activeCell.c;
          if (ascending && col === 1) {
            ctx._state.rawB = ctx.SORTED_ASC.slice();
            for (let i = 0; i < 7; i++) ctx.data[2 + i][1] = fmtCurrency0DE(ctx._state.rawB[i]);
            ctx.sortedB = true;
          } else if (!ascending && col === 3) {
            ctx._state.rawD = ctx.SORTED_DESC.slice();
            for (let i = 0; i < 7; i++) ctx.data[2 + i][3] = fmtPercent0(ctx._state.rawD[i]);
            ctx.sortedD = true;
          } else {
            showWrongToast();
          }
          if (ctx.sortedB && ctx.sortedD) {
            handleSuccess();
          } else {
            renderCurrentTask();
          }
        } else {
          ctx.keytipStage = 0;
          showWrongToast();
          renderCurrentTask();
        }
      },
    },

    // Aufgabe 6 – Strg+Pfeiltasten: über einen gelb markierten Pfad aus isolierten Zellen zum Ziel springen.
    // Jeder Sprung nutzt findDataEdge() genau wie in Aufgabe 4/5 – der Pfad ist so angelegt, dass jeder
    // Sprung exakt am nächsten Wegpunkt landet (keine zwei Wegpunkte liegen ungewollt auf einer Linie
    // zueinander, außer der bewusste "Shortcut" E3→G3, siehe Kommentar unten).
    {
      instruction: "Navigiere über den gelben Pfad zum Ziel - in genau dieser Reihenfolge.",
      shortcutLabel: () => `${MOD} + Pfeiltasten`,
      keys: () => [[MOD, "Pfeiltasten"]],
      hint(ctx) {
        return `${ctx.currentIndex + 1}/${ctx.path.length} Wegpunkte in der richtigen Reihenfolge erreicht - Zielzelle: I6.`;
      },
      init() {
        const cols = 9, rows = 9;
        // A1, A3, A6, C6, C3, E3, E6, E9, G9, G6, G3, I3, I6 (0-indexiert: A=0,B=1,...,G=6,H=7,I=8)
        const path = [
          { r: 0, c: 0 }, { r: 2, c: 0 }, { r: 5, c: 0 }, { r: 5, c: 2 }, { r: 2, c: 2 },
          { r: 2, c: 4 }, { r: 5, c: 4 }, { r: 8, c: 4 }, { r: 8, c: 6 }, { r: 5, c: 6 }, { r: 2, c: 6 },
          { r: 2, c: 8 }, { r: 5, c: 8 },
        ];
        const target = path[path.length - 1];

        const segments = [];
        for (let i = 0; i < path.length - 1; i++) {
          const a = path[i], b = path[i + 1];
          const dr = Math.sign(b.r - a.r), dc = Math.sign(b.c - a.c);
          const cells = [];
          let r = a.r, c = a.c;
          cells.push({ r, c });
          while (r !== b.r || c !== b.c) { r += dr; c += dc; cells.push({ r, c }); }
          segments.push(cells);
        }
        const allPathCells = new Set();
        segments.forEach((seg) => seg.forEach((p) => allPathCells.add(`${p.r},${p.c}`)));

        const data = [];
        for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(""));
        // Jeder Wegpunkt braucht echten (wenn auch unsichtbaren) Zellinhalt, sonst hält findDataEdge()
        // – wie echtes Excel – nicht dort an: Strg+Pfeil reagiert auf Zellinhalt, nicht auf Formatierung/Farbe.
        const waypointSet = new Set(path.map((p) => `${p.r},${p.c}`));
        path.forEach((p) => { data[p.r][p.c] = "X"; });
        data[path[0].r][path[0].c] = "Start";
        data[target.r][target.c] = "Ziel";
        const visitedSet = new Set();
        visitedSet.add(`${path[0].r},${path[0].c}`);
        return {
          cols, rows, data,
          activeCell: { r: path[0].r, c: path[0].c },
          colSelected: null, rowSelected: null,
          headerFilterRow: null, styledRange: null,
          path, target, segments,
          visited: visitedSet,
          currentIndex: 0,
          // Leere Zeile 10 unten, damit das Blatt nicht direkt an der letzten Pfad-Zeile endet.
          minVisualRows: rows + 1,
          cellClass(r, c) {
            const k = `${r},${c}`;
            if (!allPathCells.has(k)) return null;
            const base = visitedSet.has(k) ? "cell-done" : "cell-pending";
            const isEndpoint = (r === path[0].r && c === path[0].c) || (r === target.r && c === target.c);
            return isEndpoint ? base + " cell-header-bold" : base;
          },
          cellStyle(r, c) {
            return waypointSet.has(`${r},${c}`) ? "text-align:center;" : null;
          },
        };
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        if (!isArrow) return; // in dieser Aufgabe gibt es keine weiteren Shortcuts
        e.preventDefault();

        const delta = ARROW_DELTA[key];
        const newPos = (e.ctrlKey || e.metaKey)
          ? findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc)
          : {
              r: Math.max(0, Math.min(ctx.rows - 1, ctx.activeCell.r + delta.dr)),
              c: Math.max(0, Math.min(ctx.cols - 1, ctx.activeCell.c + delta.dc)),
            };
        ctx.activeCell = newPos;

        const expected = ctx.path[ctx.currentIndex + 1];
        const landedOnExpected = expected && newPos.r === expected.r && newPos.c === expected.c;

        if (landedOnExpected) {
          ctx.segments[ctx.currentIndex].forEach((p) => ctx.visited.add(`${p.r},${p.c}`));
          ctx.currentIndex += 1;
          if (ctx.currentIndex === ctx.path.length - 1) {
            handleSuccess();
            return;
          }
        } else if (newPos.r === ctx.target.r && newPos.c === ctx.target.c) {
          // Zielzelle per Abkürzung erreicht, ohne die Reihenfolge einzuhalten - zählt nicht.
          showWrongToast();
        }
        renderCurrentTask();
      },
    },

    // Aufgabe 7 – B7:E7 markieren, dann Alt+= (bzw. auf deutscher Tastatur Alt+Umschalt+0, physisch
    // dieselbe Taste wie "=") für die Autosumme über allen vier Spalten gleichzeitig.
    {
      instruction: "Markiere C7:E7 und berechne die Autosumme.",
      shortcutLabel: () => `Umschalt+Pfeiltasten → Alt+= (bzw. Alt+Umschalt+0)`,
      keys: () => [["Umschalt", "Pfeiltasten"], ["Alt", "="]],
      hint: "Genau C7:E7 markieren, dann Alt+= drücken - auf deutschen Tastaturen ist das physisch Alt+Umschalt+0.",
      init() {
        const cols = 5, rows = 7;
        const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];
        const UMSATZ = [15000, 18500, 21000, 24500];
        const KOSTEN = [9000, 10200, 11500, 13000];
        const GEWINN = [6000, 8300, 9500, 11500];
        const data = [];
        for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(""));
        data[1] = ["", "Quartal", "Umsatz", "Kosten", "Gewinn"];
        for (let i = 0; i < 4; i++) {
          data[2 + i] = ["", QUARTERS[i], fmtCurrency0DE(UMSATZ[i]), fmtCurrency0DE(KOSTEN[i]), fmtCurrency0DE(GEWINN[i])];
        }
        const sumRow = 6; // Zeile 7
        data[sumRow][1] = "Gesamt";
        const sum = (arr) => arr.reduce((s, v) => s + v, 0);
        // cellClass wird von buildGridHTML als lose Funktion aufgerufen (kein "this"-Bezug), deshalb
        // hier über eine benannte lokale Variable per Closure auf den späteren "done"-Stand zugreifen,
        // statt "this" zu verwenden.
        const state = {
          cols, rows, data,
          activeCell: { r: sumRow, c: 2 },
          selAnchor: { r: sumRow, c: 2 },
          colSelected: null, rowSelected: null,
          headerFilterRow: null,
          // Kopfzeile fett + Rahmen + alternierende Zeilen, wie eine "professionell formatierte" Tabelle -
          // rein dekorativ von Anfang an sichtbar, kein eigener Schritt in dieser Aufgabe.
          styledRange: { r0: 1, r1: 5, c0: 1, c1: 4 },
          sumRow,
          sums: [sum(UMSATZ), sum(KOSTEN), sum(GEWINN)], // Index 0 = Spalte C ("Umsatz")
          done: false,
          // Leere Zeile 8 unten und leere Spalte F rechts, damit die Tabelle nicht am Datenrand endet.
          minVisualRows: rows + 1,
          minVisualCols: cols + 1,
          cellClass(r, c) {
            if (r !== sumRow || c < 2 || c > 4) return null;
            return state.done ? "cell-done" : "cell-pending";
          },
        };
        return state;
      },
      customHandleKey(e, ctx) {
        const key = e.key;
        const isArrow = key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown";
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const isAnyModCombo = e.ctrlKey || e.metaKey || e.altKey || e.shiftKey;
        // Alt+= ist der eigentliche Excel-Shortcut; auf deutschen Tastaturen erreicht man "=" nur über
        // Umschalt+0 (dieselbe Tastatur-Eigenheit wie bei Strg+Umschalt+4/5), deshalb hier alle drei
        // robust geprüft: e.key "=" bzw. "+", der US-Tastencode "Equal", oder Alt+Umschalt+0 per e.code.
        const isAutosum =
          e.altKey && !ctrlOrMeta &&
          (key === "=" || key === "+" || e.code === "Equal" || (e.code === "Digit0" && e.shiftKey));

        if (isArrow || isAnyModCombo) {
          e.preventDefault();
        }
        if (PURE_MODIFIER_KEYS.has(key)) return;

        if (isArrow) {
          const delta = ARROW_DELTA[key];
          if (ctrlOrMeta && e.shiftKey) {
            ctx.activeCell = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
          } else if (ctrlOrMeta) {
            const edge = findDataEdge(ctx, ctx.activeCell.r, ctx.activeCell.c, delta.dr, delta.dc);
            ctx.activeCell = edge;
            ctx.selAnchor = edge;
          } else {
            let r = Math.max(0, Math.min(ctx.rows - 1, ctx.activeCell.r + delta.dr));
            let c = Math.max(0, Math.min(ctx.cols - 1, ctx.activeCell.c + delta.dc));
            ctx.activeCell = { r, c };
            if (!e.shiftKey) ctx.selAnchor = { r, c };
          }
          renderCurrentTask();
          return;
        }

        if (e.repeat) return;

        if (isAutosum) {
          if (ctx.done) return;
          const sel = normalizeRange(ctx.selAnchor, ctx.activeCell);
          const matches = sel.r0 === ctx.sumRow && sel.r1 === ctx.sumRow + 1 && sel.c0 === 2 && sel.c1 === 5;
          if (!matches) {
            showWrongToast();
            return;
          }
          for (let c = 2; c <= 4; c++) ctx.data[ctx.sumRow][c] = fmtCurrency0DE(ctx.sums[c - 2]);
          ctx.done = true;
          handleSuccess();
          return;
        }

        if (isAnyModCombo) {
          showWrongToast();
        }
      },
    },
  ];

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
  const rankResultEl = document.getElementById("rank-result");

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
    progressLabelEl.textContent = `Aufgabe ${currentTaskIndex + 1}/7`;
    taskPipsEl.innerHTML = buildPipsHTML(currentTaskIndex, TASKS.length);
    taskInstructionEl.textContent = task.instruction;
    taskBadgesEl.innerHTML = typeof task.keys === "function" ? buildBadgesHTML(task.keys()) : "";
    const hint = typeof task.hint === "function" ? task.hint(taskCtx) : task.hint;
    taskHintEl.textContent = hint || "";

    gridContainerEl.innerHTML = buildGridHTML(taskCtx);
    ribbonContainerEl.innerHTML = typeof task.renderRibbon === "function" ? task.renderRibbon(taskCtx) : "";

    if (taskCtx.colSelected != null) {
      formulabarTextEl.textContent = `Spalte ${colLetter(taskCtx.colSelected)} markiert`;
    } else if (taskCtx.rowSelected != null) {
      formulabarTextEl.textContent = `Zeile ${taskCtx.rowSelected + 1} markiert`;
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
    progressLabelEl.textContent = "Aufgabe 7/7";
    showResultScreen(finalSeconds);
  }

  function showResultScreen(finalSeconds) {
    rankResultEl.hidden = true;
    rankResultEl.innerHTML = "";
    resultTimeEl.textContent = finalSeconds.toFixed(1) + "s";
    const diff = Math.abs(finalSeconds - REFERENCE_TIME_SECONDS).toFixed(1);
    const refLabel = REFERENCE_TIME_SECONDS.toFixed(1) + "s";
    if (finalSeconds < REFERENCE_TIME_SECONDS) {
      resultHeadlineEl.textContent = "Stark gemacht!";
      resultCompareEl.textContent = `🏆 ${diff}s schneller als Flos Referenzzeit (${refLabel})`;
      resultCompareEl.className = "result-compare faster";
    } else {
      resultHeadlineEl.textContent = "Geschafft!";
      resultCompareEl.textContent = `${diff}s langsamer als Flos Referenzzeit (${refLabel})`;
      resultCompareEl.className = "result-compare slower";
    }
    showScreen("result");
    spawnConfetti();

    recordRun(finalSeconds)
      .then((stats) => {
        if (!stats) return;
        const isNewBest = stats.best_seconds != null && finalSeconds <= Number(stats.best_seconds);
        rankResultEl.hidden = false;
        if (isNewBest) {
          rankResultEl.innerHTML = `🏆 Neue Bestzeit! Platz <span class="rank-number">1</span> von ${stats.total} Teilnehmern`;
        } else {
          const behind = (finalSeconds - Number(stats.best_seconds)).toFixed(1);
          rankResultEl.innerHTML = `${behind}s hinter der Bestzeit (${Number(stats.best_seconds).toFixed(1)}s) · Platz <span class="rank-number">${stats.rank}</span> von ${stats.total} Teilnehmern`;
        }
      })
      .catch(() => {}); // best effort, keine Rang-Anzeige statt Fehlermeldung
  }

  const CONFETTI_COLORS = ["#37874a", "#22452b", "#ff8a00", "#8fd3a2", "#f2c94c"];

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
