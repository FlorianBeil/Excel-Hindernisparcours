(() => {
  "use strict";

  /* ---------------------------------------------------------------
   * Anmeldeseite: Vorname + E-Mail -> KlickTipp -> Danke-Seite
   *
   * GitHub Pages bzw. Vercel liefern nur statische Dateien aus, verschicken
   * also selbst keine Mails. Das sichtbare Formular befüllt deshalb das
   * unveränderte KlickTipp-Formular im HTML (Opt-In-Prozess „Hindernisparcours“)
   * und löst dessen Versand aus; KlickTipp schickt die Bestätigungsmail, deren
   * Link direkt in den Parcours (/parcours/) führt.
   *
   * Abgeschickt wird in ein verstecktes iframe – die Antwort von KlickTipp
   * lässt sich von dort nicht auslesen (fremde Domain), der Erfolg wird also
   * angenommen. Sobald die Antwort da ist, geht es auf die Danke-Seite.
   * ------------------------------------------------------------- */

  // Gleiches Supabase-Projekt wie der Parcours (Bestzeit, Teilnehmer, Auswertung)
  const SUPABASE_URL = "https://hbhagmmbowplzjzfvuao.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_Ee47HbgO5Ne8PP9Jh5ugag_iE_lZcp6";
  const ANON_ID_KEY = "hindernisparkour_anon_id";
  const VORNAME_KEY = "hindernisparkour_vorname"; // für die Begrüßung auf der Danke-Seite
  const DANKE_URL = "danke/";
  // Länger als das hier warten wir nicht auf die Antwort von KlickTipp
  const MAX_WARTEN_MS = 4000;

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

  // Nutzungs-Ereignis für die Auswertung (Tabelle public.events, portal 'hindernisparcours',
  // exercise_id 'anmeldung'). Best effort: Fehler werden verschluckt.
  function trackEvent(event) {
    const client = getSupabaseClient();
    const anonId = getAnonId();
    if (!client || !anonId) return;
    client
      .from("events")
      .insert({ portal: "hindernisparcours", exercise_id: "anmeldung", event, detail: {}, session_id: anonId })
      .then(() => {}, () => {});
  }

  function fmtBestTime(seconds) {
    const m = Math.floor(seconds / 60);
    const s = seconds - m * 60;
    return m > 0 ? `${m}:${s.toFixed(0).padStart(2, "0")}` : `${seconds.toFixed(1)}s`;
  }

  // Bestzeit und Teilnehmerzahl wie auf dem Startscreen des Parcours
  function loadStats() {
    const client = getSupabaseClient();
    if (!client) return;
    client
      .rpc("hindernisparkour_best_time")
      .then(({ data, error }) => {
        if (error || data == null) return;
        document.getElementById("start-best-time").textContent = fmtBestTime(Number(data));
      })
      .catch(() => {}); // best effort, Platzhalter "–" bleibt stehen
    client
      .rpc("hindernisparkour_participant_count")
      .then(({ data, error }) => {
        if (error || data == null) return;
        document.getElementById("start-participants").textContent = Number(data).toLocaleString("de-DE");
      })
      .catch(() => {});
  }

  /* ---------------------------------------------------------------
   * Kampagnen-Felder des KlickTipp-Kontos: Feldname im Formular -> Parameter
   * in der Adresse. Damit steht beim Kontakt, über welche Kampagne er kam.
   * ------------------------------------------------------------- */
  const KAMPAGNEN_FELDER = {
    "fields[field231067]": "utm_campaign",
    "fields[field231068]": "utm_medium",
    "fields[field231069]": "utm_term",
    "fields[field231070]": "utm_content",
    "fields[field231071]": "utm_source",
    "fields[field231072]": "gclid",
    "fields[field231073]": "fbclid",
  };
  const KAMPAGNE_STORAGE_KEY = "hindernisparkour_kampagne";

  // Für die Sitzung merken, falls jemand die Seite zwischendurch neu lädt.
  function kampagnenWerte() {
    let gemerkt = {};
    try {
      gemerkt = JSON.parse(sessionStorage.getItem(KAMPAGNE_STORAGE_KEY) || "{}") || {};
    } catch (e) { /* z.B. sessionStorage blockiert */ }

    const ausAdresse = new URLSearchParams(window.location.search);
    let neu = false;
    Object.values(KAMPAGNEN_FELDER).forEach((parameter) => {
      const wert = ausAdresse.get(parameter);
      if (wert) {
        gemerkt[parameter] = wert.slice(0, 200); // KlickTipp-Felder nicht sprengen
        neu = true;
      }
    });
    if (neu) {
      try {
        sessionStorage.setItem(KAMPAGNE_STORAGE_KEY, JSON.stringify(gemerkt));
      } catch (e) { /* dann gilt es eben nur für diesen Seitenaufruf */ }
    }
    return gemerkt;
  }

  /* ---------------------------------------------------------------
   * Formular
   * ------------------------------------------------------------- */
  const formEl = document.getElementById("signup-form");
  const submitEl = document.getElementById("signup-submit");
  const errorEl = document.getElementById("signup-error");
  const firstNameEl = document.getElementById("signup-firstname");
  const emailEl = document.getElementById("signup-email");
  const sinkEl = document.getElementById("signup-sink");
  const submitText = submitEl.innerHTML;

  function showError(message) {
    errorEl.textContent = message;
    errorEl.hidden = false;
  }

  function clearError() {
    errorEl.hidden = true;
    errorEl.textContent = "";
    firstNameEl.classList.remove("is-invalid");
    emailEl.classList.remove("is-invalid");
  }

  // Bewusst großzügig: E-Mail-Adressen sind vielfältiger als jede Regex.
  // Ernst nimmt die Prüfung KlickTipp, hier geht es nur um offensichtliche Tippfehler.
  function looksLikeEmail(value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(value);
  }

  // Das KlickTipp-Formular wird nicht über seine ID gesucht, damit beim Austausch
  // des RAW-Codes (neue Formular-Nummer) nichts angepasst werden muss.
  function findKlicktippForm() {
    return document.querySelector(".unlock-kt form");
  }

  function sendToKlicktipp(firstName, email) {
    const ktForm = findKlicktippForm();
    const ktFirstName = ktForm && ktForm.querySelector('[name="fields[fieldFirstName]"]');
    const ktEmail = ktForm && ktForm.querySelector('[name="email"]');
    if (!ktForm || !ktFirstName || !ktEmail) return false;

    ktFirstName.value = firstName;
    ktEmail.value = email;

    const kampagne = kampagnenWerte();
    Object.entries(KAMPAGNEN_FELDER).forEach(([feldName, parameter]) => {
      const feld = ktForm.querySelector('[name="' + feldName + '"]');
      if (feld) feld.value = kampagne[parameter] || "";
    });

    ktForm.target = "signup-sink"; // Antwort landet im versteckten iframe

    // Über den Button bzw. requestSubmit absenden statt über form.submit():
    // so laufen etwaige submit-Handler von KlickTipp mit.
    const ktSubmit = ktForm.querySelector('[type="submit"]');
    if (ktSubmit) {
      ktSubmit.click();
    } else if (typeof ktForm.requestSubmit === "function") {
      ktForm.requestSubmit();
    } else {
      ktForm.submit();
    }
    return true;
  }

  function goToThankYou(firstName) {
    try {
      sessionStorage.setItem(VORNAME_KEY, firstName);
    } catch (e) { /* dann eben ohne Namen */ }
    window.location.href = DANKE_URL;
  }

  formEl.addEventListener("submit", (e) => {
    e.preventDefault();
    clearError();

    const firstName = firstNameEl.value.trim();
    const email = emailEl.value.trim();

    if (!firstName) {
      firstNameEl.classList.add("is-invalid");
      firstNameEl.focus();
      showError("Bitte trag deinen Vornamen ein.");
      return;
    }
    if (!looksLikeEmail(email)) {
      emailEl.classList.add("is-invalid");
      emailEl.focus();
      showError("Diese E-Mail-Adresse sieht nicht richtig aus.");
      return;
    }

    // Erst weiterleiten, wenn KlickTipp geantwortet hat - sonst bricht der
    // Seitenwechsel die Anfrage womöglich ab. Spätestens nach MAX_WARTEN_MS.
    let weiter = false;
    const weiterleiten = () => {
      if (weiter) return;
      weiter = true;
      goToThankYou(firstName);
    };
    sinkEl.addEventListener("load", weiterleiten, { once: true });

    if (!sendToKlicktipp(firstName, email)) {
      sinkEl.removeEventListener("load", weiterleiten);
      showError("Der Versand ist gerade nicht möglich. Bitte später noch einmal versuchen.");
      console.warn("Das KlickTipp-Formular fehlt im HTML (.unlock-kt form).");
      return;
    }

    submitEl.disabled = true;
    submitEl.textContent = "Wird gesendet …";
    trackEvent("signup_submit");
    setTimeout(weiterleiten, MAX_WARTEN_MS);
  });

  // Zurück-Knopf des Browsers: Seite kommt evtl. aus dem Cache mit gesperrtem Button
  window.addEventListener("pageshow", () => {
    submitEl.disabled = false;
    submitEl.innerHTML = submitText;
  });

  /* ---------------------------------------------------------------
   * Init
   * ------------------------------------------------------------- */
  kampagnenWerte(); // beim Laden merken, solange die Parameter noch in der Adresse stehen
  loadStats();
  trackEvent("signup_view");
})();
