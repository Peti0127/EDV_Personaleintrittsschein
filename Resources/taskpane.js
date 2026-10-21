/* Outlook state (set in Office.onReady further below) */
var IN_OUTLOOK = false;                 // running as Outlook add-in
var SENDER = { name: "", email: "" };   // sender of the e-mail being written
var PREVIEW = false;                    // secret preview, toggled by clicking the logo 5 times

/* =================================================================
   FORM DEFINITION
   Each section has: id, title, (desc), showFor (Eintrittsarten), fields
   Field: name, label, type, required, half, options, hint, showIf
   showIf: { field: "name", value: "x" }  or { field: "name", values: [...] }
   To add questions later, just extend this array.
   ================================================================= */
const ALL = ["neueintritt", "wiedereintritt", "befoerderung", "geringfuegig"];

const EINTRITTSART_LABELS = {
  neueintritt:    "Neueintritt",
  wiedereintritt: "Wiedereintritt (nach Karenz, etc.)",
  befoerderung:   "Beförderung zu Angestellte:r",
  geringfuegig:   "Geringfügig während Karenz"
};

const ATT_LOCKED = {
  "Kontrollor FD/SUB": ["Rückmeldungen", "SR Routen", "SR Kontrolle", "KFZ", "TWK"],
  "Fixer Kontrollor":    ["Rückmeldungen", "SR Routen", "SR Kontrolle", "KFZ", "TWK", "GF Betreuung"]
};
const ZUGAENGE = ["E-Mail", "ASE", "ELO", "BMD", "ATT App", "RAT_2026"];
const ATT_APP_MODULE = ["Zeiterfassung", "HR Bilderfassung", "SR Kontrolle", "TWK", "Zähler", "KFZ", "SR Routen",
  "NFC Tag Aktivierung", "Wegzeit", "Kontrollgang", "Kostenlose Leistungen", "Vermessungen", "Schlüsselverwaltung",
  "SB Kontrolle", "Containerverwaltung", "Vertriebsapp", "Bereitschaftsmappen", "To Dos", "Objektinformation",
  "Rückmeldungen", "Personal", "Bild erfassen", "Geräteverwaltung", "Schlüssel - Manuelle Eingabe", "Tätigkeiten",
  "Geräte", "Gerätebuchung - Personal (Manuelle)", "Leistungen", "Schlüsseltyp", "Schlüssel", "Straßen",
  "Gerätebuchung - Auftrag - Manuell", "Aktivierung - Manuell", "SB Verplanung Record Hub", "Bilderdatenbank",
  "Kontakte", "Kunden", "SG Vermessung Schlüsselliste", "Leistungsscheine", "Aufträge", "GB Zeiterfassung",
  "GF Betreuung"];

const STANDORTE = ["", "Klosterneuburg", "Kalsdorf", "Innsbruck", "Klagenfurt", "Leoben", "Mödling",
  "Pinzgau", "Salzburg", "Spillern", "St. Pölten", "Oberwart", "Rankweil", "Wr. Neustadt",
  "Viertel Zwei", "Krems", "Linz", "Neusiedl", "Sonstiges"];

/* Gebäude/Stock/Büro for all Standorte except Klosterneuburg and Sonstiges */
const STOCK_BY_STANDORT = {                       // default for all others: Erdgeschoß
  "Innsbruck":    ["1. Stock"],
  "Linz":         ["1. Stock"],
  "Regau":        ["Erdgeschoß", "1. Stock"],
  "Wr. Neustadt": ["Erdgeschoß", "1. Stock"]
};
const stockOptions = st => STOCK_BY_STANDORT[st] || ["Erdgeschoß"];

/* Büronummern per Standort (default: Büro 1–8). An object means: different rooms per Stock. */
const bueros = (from, to, prefix = "Büro ") =>
  Array.from({ length: to - from + 1 }, (_, i) => prefix + (from + i));
const ROOMS_BY_STANDORT = {
  "Kalsdorf":     bueros(1, 8, "B"),
  "Innsbruck":    bueros(1, 6),
  "Klagenfurt":   bueros(1, 5),
  "Leoben":       bueros(1, 5),
  "Rankweil":     bueros(1, 5),
  "Pinzgau":      bueros(1, 4),
  "Salzburg":     bueros(1, 7),
  "Spillern":     bueros(1, 2),
  "Linz":         bueros(1, 2),
  "Neusiedl":     bueros(1, 2),
  "Oberwart":     ["Büro"],
  "Viertel Zwei": ["Büro"],
  "Krems":        ["Büro"],
  "Wr. Neustadt": { "Erdgeschoß": bueros(1, 2), "1. Stock": bueros(3, 7) }
};
/* Klosterneuburg: floors per building and rooms per building + floor */
const KL_STOCK = {"Hauptgebäude": ["Erdgeschoß", "1. Stock", "2. Stock", "noch unbekannt"], "Neubau": ["Erdgeschoß", "1. Stock", "noch unbekannt"], "Inku": ["Erdgeschoß Links", "Erdgeschoß Rechts", "Obergeschoß Links", "Obergeschoß Rechts", "noch unbekannt"]};
const KL_ROOMS = {
  "Hauptgebäude": {
    "Erdgeschoß": [
      "A11",
      "A13",
      "A20",
      "A26",
      "A27",
      "A31",
      "A32",
      "A33",
      "A34",
      "A35",
      "A36",
      "Empfang",
      "noch unbekannt"
    ],
    "1. Stock": [
      "B02",
      "B08",
      "B09",
      "B10",
      "B14",
      "B15",
      "B30",
      "B31",
      "B32",
      "B33",
      "B35",
      "B36",
      "B37",
      "B38",
      "B39",
      "noch unbekannt"
    ],
    "2. Stock": [
      "C01",
      "C02",
      "C05",
      "C06",
      "C07",
      "C08",
      "C09",
      "C13",
      "noch unbekannt"
    ]
  },
  "Neubau": {
    "Erdgeschoß": [
      "D04",
      "D09",
      "noch unbekannt"
    ],
    "1. Stock": [
      "D08",
      "D11",
      "D12",
      "D13",
      "D14",
      "D15",
      "noch unbekannt"
    ]
  },
  "Inku": {
    "Erdgeschoß Links": [
      "IN16_EG.L01",
      "IN16_EG.L02",
      "IN16_EG.L05",
      "IN16_EG.L06",
      "IN16_EG.L07",
      "IN16_EG.L08",
      "IN16_EG.L09",
      "IN16_EG.L10",
      "IN16_EG.L11",
      "IN16_EG.L12",
      "IN16_EG.L14",
      "IN16_EG.L15",
      "IN16_EG.L16",
      "IN16_EG.L17",
      "noch unbekannt"
    ],
    "Erdgeschoß Rechts": [
      "Keine verfügbare Räume"
    ],
    "Obergeschoß Links": [
      "IN16_OG.L01",
      "IN16_OG.L02",
      "IN16_OG.L03",
      "IN16_OG.L06",
      "IN16_OG.L07",
      "IN16_OG.L08",
      "IN16_OG.L09",
      "IN16_OG.L10",
      "IN16_OG.L11",
      "IN16_OG.L12",
      "IN16_OG.L13",
      "IN16_OG.L14",
      "IN16_OG.L16",
      "IN16_OG.L17",
      "IN16_OG.L18",
      "IN16_OG.L19/20",
      "noch unbekannt"
    ],
    "Obergeschoß Rechts": [
      "I01",
      "I02",
      "I03",
      "I04",
      "I05",
      "I06",
      "I07",
      "I08",
      "I09",
      "I10",
      "I11",
      "I12",
      "I12a",
      "I13",
      "noch unbekannt"
    ]
  }
};

function gebaeudeChoices() {
  const st = getValue("standort");
  if (st === "Klosterneuburg") return ["Hauptgebäude", "Neubau", "Inku"];
  return isOtherStandort(st) ? ["Hauptgebäude"] : [];
}
function stockChoices() {
  const st = getValue("standort"), g = getValue("gebaeude");
  if (!g) return [];
  return st === "Klosterneuburg" ? (KL_STOCK[g] || []) : stockOptions(st);
}
function bueroChoices() {
  const st = getValue("standort"), g = getValue("gebaeude"), s = getValue("stock");
  if (!s) return [];
  if (st === "Klosterneuburg") {
    if (s === "noch unbekannt") return ["noch unbekannt"];
    return (KL_ROOMS[g] || {})[s] || [];
  }
  return roomOptions();
}
function currentStock() { return getValue("stock"); }
/* Plan for the selected Standort; plans split by floor only appear once a Stock is chosen */
function currentPlan() {
  const st = getValue("standort");
  const entry = LAGEPLAENE[st];
  if (!entry) return null;
  if (entry.buildings) {                       // Klosterneuburg: plan per Gebäude and Stock
    const g = getValue("gebaeude");
    const stock = getValue("stock");
    const src = entry.buildings[g] && entry.buildings[g][stock];
    return src ? { src, alt: `Lageplan ${st}, ${g}, ${stock} (${entry.code})` } : null;
  }
  if (!isOtherStandort(st)) return null;
  if (!entry.floors) return { src: entry.src, alt: `Lageplan ${st} (${entry.code})` };
  const stock = currentStock();
  const src = entry.floors[stock];
  return src ? { src, alt: `Lageplan ${st}, ${stock} (${entry.code})` } : null;
}

function roomOptions() {
  const r = ROOMS_BY_STANDORT[getValue("standort")] || bueros(1, 8);
  return Array.isArray(r) ? r : (r[currentStock()] || []);
}
/* true unless all three hardware questions are answered with "Nein" */
/* Neueintritt as Arbeiter:in: no office location, only Telefonie after Persönliche Daten */
const isArbeiter = () => currentType() === "neueintritt" && getValue("neu_beschaeftigung") === "Arbeiter:in";

// only counts when the EDV-Ausstattung section is actually shown (not for Arbeiter:in)
const keinComputer = () => edvSectionsVisible(currentType()) && getValue("edv_computer") === "Nein";

/* Verteiler added automatically (shown as gray hint and always included in the e-mail) */
function autoVerteiler() {
  const z = visibleByName("zugaenge") ? getValue("zugaenge") : [];
  const list = ["_Alle_Attensam_Mitarbeiter"];
  if (z.includes("ASE")) list.push("_Alle_ASE_User");
  if (z.includes("BMD")) list.push("_Alle_BMD_User");
  if (z.includes("BMD") && visibleByName("bmd_rechte") && getValue("bmd_rechte") === "ER Scan") list.push("_Alle_BMD_Scan_User");
  if (z.includes("ELO")) list.push("_Alle_ELO_User");
  if (["Klosterneuburg", "Mödling", "Viertel Zwei"].includes(getValue("standort"))) list.push("_Alle_Wien");
  if (visibleByName("gebaeude") && getValue("gebaeude") === "Inku") list.push("_Alle_Inku");
  return list;
}
/* automatic + typed by the user (one per line or comma-separated), without duplicates */
function verteilerListe() {
  const own = (getValue("verteiler") || "").split(/[\n,;]+/).map(x => x.trim()).filter(Boolean);
  return [...new Set([...autoVerteiler(), ...own])];
}

/* ASE variants that fit the chosen devices */
function aseOptions() {
  const tel = getValue("tel_art");
  if (tel === "Keine") return ["ASE (Nur PC)"];
  if (ohneComputer() && tel === "Handy") return ["ASE Workgroups App (Handy)"];
  return ["ASE (Nur PC)", "ASE Workgroups App (Handy)", "ASE & ASE Workgroups App (PC + Handy)"];
}

/* No computer: Angestellte:r with "Nein" or Arbeiter:in */
const ohneComputer = () => keinComputer() || isArbeiter();

function sectionVisibleById(id) {
  return sectionVisible(SECTIONS.find(s => s.id === id), currentType());
}

/* No computer and no phone → nothing for IT to do */
function keinEdvBenoetigt() {
  return keinComputer() && getValue("tel_art") === "Keine";
}
const hatFestnetz = (val) => ["Nur Festnetz", "Handy & Festnetz"].includes(val("tel_art"));
const hardwareNotAllNein = (val) => ["buero_hw", "ho_hw", "hw_benoetigt"].some(n => visibleByName(n) && val(n) !== "Nein");

const isOtherStandort = v => !!v && v !== "Klosterneuburg" && v !== "Sonstiges";

const SECTIONS = [
  /* ---------- Antragsteller:in – always first ---------- */
  {
    id: "antrag",
    title: "Antragsteller:in",
    desc: "Wer stellt den Antrag?",
    always: true,
    showIf: () => !IN_OUTLOOK,          // in Outlook the sender of the e-mail is used instead
    fields: [
      { name: "antrag_name", label: "Vor- und Nachname", type: "text", required: true, half: true,
        minWords: 2, placeholder: "z. B. Maria Muster" },
      { name: "antrag_email", label: "E-Mail", type: "email", required: true, half: true,
        placeholder: "vorname.nachname- oder Kürzel@attensam.at" }
    ]
  },

  /* ---------- Eintrittsart – after Antragsteller:in is complete ---------- */
  {
    id: "art",
    title: "Eintrittsart?",
    desc: "Je nach Auswahl erscheinen die passenden Felder.",
    className: "main-question",
    always: true,
    inEmail: false,                   // already in the e-mail header
    showIf: () => antragComplete(),
    fields: [
      { name: "eintrittsart", label: "Eintrittsart", hideLabel: true, type: "radio", required: true,
        errorMsg: "Bitte eine Eintrittsart auswählen.",
        options: Object.entries(EINTRITTSART_LABELS).map(([value, label]) => ({ value, label })) }
    ]
  },

  /* ---------- Neueintritt ---------- */
  {
    id: "neu",
    title: "Neueintritt",
    desc: "Angaben zur neuen Stelle",
    showFor: ["neueintritt"],
    fields: [
      { name: "neu_eintrittsdatum", label: "Eintrittsdatum", type: "date", required: true, half: true,
        dateRange: { pastYears: 1, futureYears: 1 } },
      { name: "neu_beschaeftigung", label: "Beschäftigungsart", type: "radio", required: true, cols: 2,
        options: ["Angestellte:r", "Arbeiter:in", "Lehrling", "Praktikant:in"] }
    ]
  },

  /* ---------- Shared: personal data ---------- */
  {
    id: "person",
    title: "Persönliche Daten",
    showFor: ALL,
    showIf: type => type !== "neueintritt" || neuComplete(),
    fields: [
      { name: "anrede", label: "Anrede", type: "select", half: true, options: ["", "Frau", "Herr"] },
      { name: "titel", label: "Titel", type: "text", half: true },
      { name: "vorname", label: "Vorname", type: "text", required: true, half: true },
      { name: "nachname", label: "Nachname", type: "text", required: true, half: true },
      { name: "stellenbezeichnung", label: "Stellenbezeichnung", type: "text", required: true, half: true },
      { name: "fuehrungskraft", label: "Führungskraft", type: "text", required: true, half: true },
      { name: "kostenstelle", label: "Kostenstelle", type: "text", required: true, half: true },
      { name: "standort", label: "Standort", type: "select", required: true, half: true, options: STANDORTE },
      { name: "standort_sonstiges", label: "Standort (Sonstiges)", type: "text", required: true, half: true,
        showIf: { field: "standort", value: "Sonstiges" } },

      /* Gebäude → Stock → Lageplan → Büronummer: one field each, options depend on the answers above.
         A "choice" with only one possible option is shown as text instead of a list. */
      { name: "gebaeude", label: "Gebäude", type: "choice", required: true, half: true, startRow: true,
        choices: () => gebaeudeChoices(),
        showIf: () => !isArbeiter() },
      { name: "stock", label: "Stock", type: "choice", required: true, half: true,
        choices: () => stockChoices(),
        showIf: (val, visible) => visible("gebaeude") && !!val("gebaeude") },
      /* Lageplan (screen only, not in the PDF) */
      { name: "lageplan", label: "Lageplan", type: "image", inEmail: false,
        showIf: (val, visible) => visible("stock") && !!currentPlan() },
      { name: "buero", label: "Büronummer", type: "choice", required: true, half: true, startRow: true,
        choices: () => bueroChoices(),
        blockOn: { value: "Keine verfügbare Räume", message: "Bitte anderen Stock/Gebäude wählen!" },
        showIf: (val, visible) => visible("stock") && !!val("stock") },

      /* Zutrittschip – only at these locations (below Gebäude/Stock/Büro) */
      { name: "zutrittschip", label: "Zutrittschip benötigt?", type: "select", required: true, half: true, startRow: true,
        options: ["", "Ja", "Nein", "Wird übernommen"],
        showIf: { field: "standort", values: ["Klosterneuburg", "Salzburg", "Wr. Neustadt"] } },
      { name: "zutritt_wie", label: "Zutrittsrechte wie Person", type: "text", required: true, half: true,
        placeholder: "z.B.: Max Mustermann oder Kürzel des Users",
        showIf: { field: "zutrittschip", value: "Ja" } },
      { name: "zutritt_uebernahme", label: "Übernahme von", type: "text", required: true, half: true,
        showIf: { field: "zutrittschip", value: "Wird übernommen" } },

      /* Klosterneuburg Inku: extra chip for the Hauptgebäude */
      { name: "zutritt_hg", label: "Zutrittschip für Hauptgebäude (schwarzer Chip)?", type: "select",
        required: true, half: true, startRow: true,
        options: ["", "Ja", "Nein"],
        showIf: (val, visible) => visible("zutrittschip") && val("zutrittschip") === "Ja"
          && val("standort") === "Klosterneuburg" && visible("gebaeude") && val("gebaeude") === "Inku" },
      { name: "zutritt_hg_wie", label: "Zutrittsrechte wie Person", type: "text", required: true, half: true,
        placeholder: "z.B.: Max Mustermann oder Kürzel des Users",
        showIf: { field: "zutritt_hg", value: "Ja" } }
    ]
  },

  /* ---------- Wiedereintritt ---------- */
  {
    id: "wieder",
    title: "Wiedereintritt",
    desc: "Rückkehr nach einer Abwesenheit",
    showFor: ["wiedereintritt"],
    fields: [
      { name: "wieder_datum", label: "Wiedereintrittsdatum", type: "date", required: true, half: true }
    ]
  },

  /* ---------- Beförderung ---------- */
  {
    id: "befoerd",
    title: "Beförderung zu Angestellte:r",
    desc: "Wechsel des Dienstverhältnisses",
    showFor: ["befoerderung"],
    fields: [
      { name: "bef_bisher", label: "Bisheriges Dienstverhältnis", type: "radio", required: true, cols: 2,
        options: ["Arbeiter:in", "Lehrling", "Leiharbeiter:in", "Sonstiges"] },
      { name: "bef_gueltig", label: "Gültig ab", type: "date", required: true, half: true },
      { name: "bef_kostenstelle", label: "Kostenstelle", type: "text", half: true },
      { name: "bef_pos_alt", label: "Bisherige Position", type: "text", half: true },
      { name: "bef_pos_neu", label: "Neue Position", type: "text", required: true, half: true },
      { name: "bef_abteilung_wechsel", label: "Wechselt die Abteilung?", type: "radio", required: true, cols: 2,
        options: ["Nein", "Ja"] },
      { name: "bef_abteilung", label: "Neue Abteilung", type: "text", half: true, required: true,
        showIf: { field: "bef_abteilung_wechsel", value: "Ja" } },
      { name: "bef_vorgesetzt", label: "Neue:r Vorgesetzte:r", type: "text", half: true,
        showIf: { field: "bef_abteilung_wechsel", value: "Ja" } },
      { name: "bef_pc_bisher", label: "Hatte bisher einen eigenen PC-Arbeitsplatz?", type: "radio", cols: 2,
        options: ["Ja", "Nein, nur Terminal / Shared PC"] }
    ]
  },

  /* ---------- Geringfügig während Karenz ---------- */
  {
    id: "gering",
    title: "Geringfügig während Karenz",
    desc: "Befristete geringfügige Beschäftigung",
    showFor: ["geringfuegig"],
    fields: [
      { name: "ger_karenz_seit", label: "In Karenz seit", type: "date", half: true },
      { name: "ger_karenz_bis", label: "Karenz voraussichtlich bis", type: "date", half: true },
      { name: "ger_beginn", label: "Beginn geringfügige Beschäftigung", type: "date", required: true, half: true },
      { name: "ger_ende", label: "Ende geringfügige Beschäftigung", type: "date", half: true },
      { name: "ger_stunden", label: "Stunden pro Woche", type: "number", half: true },
      { name: "ger_abteilung", label: "Abteilung", type: "text", required: true, half: true },
      { name: "ger_taetigkeit", label: "Tätigkeit", type: "textarea", required: true },
      { name: "ger_ort", label: "Arbeitsort", type: "radio", required: true, cols: 2,
        options: ["Vor Ort", "Homeoffice", "Gemischt"] }
    ]
  },

  /* ---------- IT equipment (all types) ---------- */
  {
    id: "edv",
    title: "EDV-Ausstattung",
    desc: "Was wird benötigt? Was muss gekauft werden?",
    showFor: ALL,
    showIf: type => edvSectionsVisible(type),
    fields: [
      { name: "edv_computer", label: "Wird der User am Computer arbeiten?", type: "select", required: true,
        half: true, startRow: true, options: ["Ja", "Nein"], default: "Nein" },

      /* Existing hardware */
      { name: "buero_hw", label: "Bestehende Büro Hardware?", type: "select", required: true, half: true, startRow: true,
        options: ["Ja", "Nein"], default: "Nein",
        showIf: { field: "edv_computer", value: "Ja" } },
      { name: "buero_hw_inv", label: "Inventarnummer(n) der bestehenden Hardware", type: "text", required: true, half: true,
        placeholder: "z.B.: 09521; 10485",
        showIf: { field: "buero_hw", value: "Ja" } },
      { name: "ho_hw", label: "Bestehende Home Office Hardware?", type: "select", required: true, half: true, startRow: true,
        options: ["Ja", "Nein"], default: "Nein",
        showIf: { field: "edv_computer", value: "Ja" } },
      { name: "ho_hw_inv", label: "Inventarnummer(n) der bestehenden Home Office Hardware", type: "text", required: true, half: true,
        placeholder: "z.B.: 09521; 10485",
        showIf: { field: "ho_hw", value: "Ja" } },

      /* New hardware */
      { name: "hw_benoetigt", label: "Wird zusätzliche Hardware benötigt?", type: "select", required: true, half: true, startRow: true,
        options: ["Ja", "Nein"], default: "Nein",
        // works at a computer → at least one of the three hardware questions must be "Ja" (blocks sending)
        check: () => ["buero_hw", "ho_hw", "hw_benoetigt"].every(n => getValue(n) === "Nein")
          ? "Der User arbeitet am Computer: mindestens eine der drei Fragen zur Hardware muss „Ja“ sein." : null,
        showIf: { field: "edv_computer", value: "Ja" } },
      { name: "hw_auswahl", label: "benötigte Hardware", type: "checkbox", required: true,
        errorMsg: "Bitte mindestens eine Hardware auswählen.",
        options: ["Laptop", "Bildschirm", "Docking Station", "Maus + Tastatur", "USB Headset", "USB Webcam"],
        showIf: { field: "hw_benoetigt", value: "Ja" } },
      /* shown directly below the "Bildschirm" checkbox */
      { name: "bildschirm_stueck", label: "Stück", type: "select", required: true,
        options: ["1", "2", "3", "4"],
        // gray info + red reminder below the list (information only, does not block sending)
        notes: () => {
          const n = getValue("bildschirm_stueck");
          const dock = getValue("hw_auswahl").includes("Docking Station");
          const out = [];
          if (n === "3") out.push({ kind: "hint", text: "Ausstattung: 2x Büro, 1x Home Office" });
          if (n === "4") out.push({ kind: "hint", text: "Ausstattung: 2x Büro, 2x Home Office" });
          if (!dock && (n === "2" || n === "3")) out.push({ kind: "warn", text: "Für 2 Bildschirme wird auch eine Docking Station benötigt!" });
          if (!dock && n === "4") out.push({ kind: "warn", text: "Für 2 Bildschirme wird auch eine Docking Station benötigt! (+ 1x auch für Home Office)" });
          if (getValue("ho_hw") === "Ja" && (n === "3" || n === "4"))
            out.push({ kind: "warn", text: "Home Office Equipment besteht, ist die benötigte Anzahl von Bildschirmen sicher korrekt?" });
          return out;
        },
        nestUnder: { field: "hw_auswahl", values: ["Bildschirm"] },
        showIf: { field: "hw_auswahl", value: "Bildschirm" } },
      { name: "bildschirm_groesse", label: "Bildschirmgröße", type: "select", required: true,
        options: ['24" (Standardgröße)', '27"'],
        nestUnder: { field: "hw_auswahl", values: ["Bildschirm"] },
        showIf: { field: "hw_auswahl", value: "Bildschirm" } },
      /* gray info directly below the "Docking Station" checkbox */
      { name: "dock_hinweis", type: "hint", inEmail: false,
        text: "2 Docking Stations: 1x Büro, 1x Home Office",
        nestUnder: { field: "hw_auswahl", values: ["Docking Station"] },
        showIf: (val, visible) => visible("hw_auswahl") && val("hw_auswahl").includes("Docking Station")
          && visible("bildschirm_stueck") && val("bildschirm_stueck") === "4" },

      { name: "laufwerke", label: "J: Ordner Berechtigungen", type: "textarea",
        placeholder: "Schreibweise z.B.: \"ATS01\\Schneeraemung\\Allgemein\" (Pfad kopieren) oder \"Ordner Wichtiges unter Finanz\" (beschreiben, wo der Ordner sich befindet)",
        showIf: (val) => hardwareNotAllNein(val) }
    ]
  },

  /* ---------- Telefonie ---------- */
  {
    id: "telefonie",
    title: "Telefonie",
    showFor: ALL,
    showIf: type => edvSectionsVisible(type) || (isArbeiter() && sectionComplete("person") && neuComplete()),
    fields: [
      { name: "tel_art", label: "Festnetz? Handy?", emailLabel: "Festnetz / Handy", type: "select", required: true, half: true, startRow: true,
        options: ["", "Keine", "Handy", "Nur Festnetz", "Handy & Festnetz"],
        // without a computer there is no Festnetz
        dynamicOptions: () => ohneComputer() ? ["", "Keine", "Handy"]
                                             : ["", "Keine", "Handy", "Nur Festnetz", "Handy & Festnetz"] },
      /* Festnetz – depends on Standort */
      { name: "festnetz_art", label: "Festnetztelefonie", type: "select", required: true, half: true,
        options: ["", "Tischtelefon", "3CX"],
        showIf: (val) => hatFestnetz(val) && val("standort") === "Klosterneuburg" },
      { name: "festnetz_info", label: "Festnetztelefonie", type: "info", half: true,
        text: "Festnetztelefonie mittels 3CX", value: "3CX",
        showIf: (val) => hatFestnetz(val) && !!val("standort") && val("standort") !== "Klosterneuburg" },
      /* Handy */
      { name: "handy_modell", label: "Handy Modell?", type: "select", required: true, half: true, startRow: true,
        options: ["", "Standardhandy", "Reinigerhandy", "spezifisches Modell"],
        showIf: (val) => ["Handy", "Handy & Festnetz"].includes(val("tel_art")) },
      { name: "handy_modell_text", label: "Modell", type: "text", required: true, half: true,
        showIf: { field: "handy_modell", value: "spezifisches Modell" } },
      { name: "tel_uebernahme", label: "Durchwahlübernahme?", type: "select", required: true, half: true, startRow: true,
        options: ["", "Ja", "Nein"],
        showIf: (val) => !!val("tel_art") && val("tel_art") !== "Keine" },
      { name: "tel_uebernahme_von", label: "Übernahme von", type: "text", required: true, half: true,
        showIf: { field: "tel_uebernahme", value: "Ja" } },
      { name: "tel_neu", label: "Durchwahl", type: "info", half: true,
        text: "EDV vergibt neue Durchwahl", value: "EDV vergibt neue Durchwahl",
        showIf: { field: "tel_uebernahme", value: "Nein" } }
    ]
  },
  /* ---------- Warning below Telefonie (no box, no header, not in the PDF) ---------- */
  {
    id: "kein_edv",
    bare: true,
    inEmail: false,
    showFor: ALL,
    showIf: type => edvSectionsVisible(type) && keinEdvBenoetigt(),
    fields: [
      { name: "kein_edv_hinweis", type: "warning",
        text: "Laut Auswahl braucht der User keine EDV-Hardware! Bitte überprüfen, ob es so richtig ist!" }
    ]
  },

  /* ---------- Zugänge ---------- */
  {
    id: "zugaenge_sec",
    title: "Zugänge",
    showFor: ALL,
    showIf: type => {
      // Arbeiter:in: once Telefonie is completely filled in (not without a phone)
      if (isArbeiter()) return sectionVisibleById("telefonie") && sectionComplete("telefonie") && getValue("tel_art") !== "Keine";
      // otherwise together with EDV-Ausstattung and Telefonie; hidden without computer and without a phone choice
      return edvSectionsVisible(type) && !(keinComputer() && ["", "Keine"].includes(getValue("tel_art")));
    },
    fields: [
      { name: "zugaenge", label: "Zugänge", type: "checkbox", cols: 2, options: ZUGAENGE,
        hideOptions: () => {
          const tel = getValue("tel_art");
          if (tel === "Keine") return ["ATT App"];                          // no phone → no ATT App
          if (ohneComputer() && tel === "Handy") return ["ELO", "BMD"]; // phone only → E-Mail, ASE, ATT App
          return [];
        },
        // works at a computer → E-Mail is always needed (ticked + greyed out)
        lockOptions: () => edvSectionsVisible(currentType()) && getValue("edv_computer") === "Ja" ? ["E-Mail"] : [] },
      /* E-Mail → Verteiler */
      { name: "verteiler", label: "Welchem Verteiler soll der User hinzugefügt werden?", type: "textarea",
        placeholder: "Standardverteiler: _Alle_Attensam_Mitarbeiter",
        // gray hint lists the Verteiler that are added automatically
        placeholderFn: () => `Standardverteiler: ${autoVerteiler().join(", ")}`,
        // PDF/e-mail: automatic Verteiler + whatever the user typed
        displayValue: () => verteilerListe().join(", "),
        showIf: { field: "zugaenge", value: "E-Mail" } },

      /* ASE */
      { name: "ase_anwendung", label: "ASE Anwendung", type: "select", required: true, half: true, startRow: true,
        options: [""], dynamicOptions: () => ["", ...aseOptions()],
        showIf: (val, visible) => visible("zugaenge") && val("zugaenge").includes("ASE") && aseOptions().length > 1 },
      /* Only one possible ASE variant → shown as text */
      { name: "ase_anwendung_text", label: "ASE Anwendung", type: "info", half: true, startRow: true,
        text: () => aseOptions()[0], value: () => aseOptions()[0],
        showIf: (val, visible) => visible("zugaenge") && val("zugaenge").includes("ASE") && aseOptions().length === 1 },
      { name: "ase_wie", label: "ASE Rechte wie User", type: "text", required: true, half: true,
        placeholder: "z.B.: Max Mustermann oder Kürzel des Users",
        showIf: { field: "zugaenge", value: "ASE" } },

      /* ELO */
      { name: "elo_rechte", label: "ELO Rechte", type: "select", required: true, half: true, startRow: true,
        options: ["", "Leserechte", "Schreiberechte"],
        noteOn: { value: "Schreiberechte", message: "verursacht Lizenzkosten!" },
        showIf: { field: "zugaenge", value: "ELO" } },

      /* BMD */
      { name: "bmd_rechte", label: "BMD Rechte", type: "select", required: true, half: true, startRow: true,
        options: ["", "ER Scan", "ER Controlle", "Kostenrechnung"],
        showIf: { field: "zugaenge", value: "BMD" } },
      { name: "bmd_wie", label: "BMD Rechte wie User", type: "text", half: true,
        placeholder: "z.B.: Max Mustermann oder Kürzel des Users",
        showIf: { field: "zugaenge", value: "BMD" } },

      /* ATT App */
      { name: "att_rechte", label: "ATT Rechte", type: "select", required: true, half: true, startRow: true,
        options: ["", "Wie User", "Auswahl", "Kontrollor FD/SUB", "Fixer Kontrollor"],
        showIf: { field: "zugaenge", value: "ATT App" } },
      { name: "att_wie", label: "ATT Rechte wie User", type: "text", required: true, half: true,
        placeholder: "z.B.: Max Mustermann oder Kürzel des Users",
        showIf: { field: "att_rechte", value: "Wie User" } },
      { name: "att_module", label: "ATT App Funktionen", type: "checkbox", required: true, cols: 3,
        errorMsg: "Bitte mindestens eine Funktion auswählen.",
        options: ATT_APP_MODULE,
        // Kontrollor roles: these functions are always included (ticked + greyed out)
        lockOptions: () => ATT_LOCKED[getValue("att_rechte")] || [],
        showIf: { field: "att_rechte", values: ["Auswahl", "Kontrollor FD/SUB", "Fixer Kontrollor"] } }
    ]
  },

  /* ---------- Visitenkarten ---------- */
  {
    id: "visitenkarten",
    title: "Visitenkarten",
    showFor: ALL,
    showIf: type => edvSectionsVisible(type),
    fields: [
      { name: "visitenkarten", label: "Visitenkarten benötigt?", type: "select", required: true, half: true,
        options: ["", "Ja", "Nein", "erst nach Probezeit"] }
    ]
  },

  /* ---------- Sonstige Anmerkungen (always last) ---------- */
  {
    id: "anmerkungen",
    title: "Sonstige Anmerkungen",
    showFor: ALL,
    // after Persönliche Daten (and Neueintritt) is complete – also for Arbeiter:in
    showIf: type => sectionComplete("person") && (type !== "neueintritt" || neuComplete()) && !keinEdvBenoetigt(),
    fields: [
      { name: "anmerkungen", label: "Sonstige Anmerkungen", hideLabel: true, type: "textarea", big: true }
    ]
  }
];

/* EDV and Telefonie: after Persönliche Daten is complete; for Neueintritt not for Arbeiter:in */
function edvSectionsVisible(type) {
  if (!sectionComplete("person")) return false;
  if (type !== "neueintritt") return true;
  return neuComplete() && getValue("neu_beschaeftigung") !== "Arbeiter:in";
}

/* =================================================================
   RENDERING
   ================================================================= */
const form = document.getElementById("form");
const dynamic = document.getElementById("dynamic");
const emptyEl = document.getElementById("empty");
const statusEl = document.getElementById("status");

/* Local-date helpers (toISOString would use UTC and can be off by one day) */
function toISODate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function toATDate(iso) { const [y, m, d] = iso.split("-"); return `${d}.${m}.${y}`; }
function shiftYears(years) { const d = new Date(); d.setFullYear(d.getFullYear() + years); return d; }

/* Allowed window for a field with dateRange, e.g. { pastYears: 1, futureYears: 1 } */
function dateBounds(f) {
  return { min: toISODate(shiftYears(-f.dateRange.pastYears)), max: toISODate(shiftYears(f.dateRange.futureYears)) };
}
function dateRangeError(f) {
  if (!f.dateRange) return null;
  const v = getValue(f.name);
  if (!v) return null;
  const { min, max } = dateBounds(f);
  if (v < min || v > max) return `Das Datum muss zwischen ${toATDate(min)} und ${toATDate(max)} liegen.`;
  return null;
}

const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function renderField(f) {
  const wrap = document.createElement("div");
  wrap.className = "field" + (f.half ? " half" : "") + (f.startRow ? " start-row" : "");
  wrap.dataset.field = f.name;
  const req = f.required ? '<span class="req" aria-hidden="true">*</span>' : "";
  const id = "f_" + f.name;
  let html = "";

  if (f.type === "hint") {
    html += `<div class="dyn-hint">${esc(f.text)}</div>`;
  } else if (f.type === "choice") {
    html += `<label for="${id}">${esc(f.label)}${req}</label>
      <select id="${id}" name="${f.name}"><option value="">Bitte wählen</option></select>
      <div class="info-text choice-text" hidden></div>`;
  } else if (f.type === "warning") {
    html += `<div class="warn-box" role="alert">${esc(f.text)}</div>`;
  } else if (f.type === "image") {
    html += `<div class="label">${esc(f.label)}</div>
      <button type="button" class="plan-btn" aria-label="Lageplan vergrößern"><img alt=""></button>
      <div class="hint">Zum Vergrößern klicken</div>`;
  } else if (f.type === "info") {
    html += `<div class="label"${f.hideLabel ? " hidden" : ""}>${esc(f.label)}</div><div class="info-text">${typeof f.text === "function" ? "" : esc(f.text)}</div>`;
  } else if (f.type === "radio" || f.type === "checkbox") {
    html += `<div class="label" id="${id}_lbl"${f.hideLabel ? " hidden" : ""}>${esc(f.label)}${req}</div>`;
    html += `<div class="options${f.cols ? " cols-" + f.cols : ""}" role="${f.type === "radio" ? "radiogroup" : "group"}" aria-labelledby="${id}_lbl">`;
    f.options.forEach(o => {
      const v = typeof o === "object" ? o.value : o;
      const l = typeof o === "object" ? o.label : o;
      html += `<label class="opt"><input type="${f.type}" name="${f.name}" value="${esc(v)}"> ${esc(l)}</label>`;
    });
    html += `</div>`;
  } else {
    html += `<label for="${id}"${f.hideLabel ? " hidden" : ""}>${esc(f.label)}${req}</label>`;
    if (f.type === "select") {
      html += `<select id="${id}" name="${f.name}">` +
        f.options.map(o => `<option value="${esc(o)}"${o === f.default ? " selected" : ""}>${o ? esc(o) : "Bitte wählen"}</option>`).join("") +
        `</select>`;
    } else if (f.type === "textarea") {
      html += `<textarea id="${id}" name="${f.name}"${f.big ? ' class="big"' : ""}${f.placeholder ? ` placeholder="${esc(f.placeholder)}"` : ""}></textarea>`;
    } else {
      let extra = f.type === "number" ? ' min="0" step="0.5"' : "";
      if (f.dateRange) { const b = dateBounds(f); extra = ` min="${b.min}" max="${b.max}"`; }
      if (f.digits) extra += ` inputmode="numeric" maxlength="${f.digits}" placeholder="${"0".repeat(f.digits)}"`;
      if (f.placeholder) extra += ` placeholder="${esc(f.placeholder)}"`;
      html += `<input id="${id}" type="${f.type}" name="${f.name}"${extra}>`;
    }
  }
  if (f.hint) html += `<div class="hint">${esc(f.hint)}</div>`;
  if (f.noteOn) html += `<div class="note-msg" hidden>${esc(f.noteOn.message)}</div>`;
  if (f.notes) html += `<div class="dyn-notes"></div>`;
  html += `<div class="error-msg">${esc(f.errorMsg || "Dieses Feld ist erforderlich.")}</div>`;
  wrap.innerHTML = html;

  if (f.defaultToday) {
    const inp = wrap.querySelector("input");
    inp.value = toISODate(new Date());
  }
  return wrap;
}

function renderSections() {
  SECTIONS.forEach(s => {
    const sec = document.createElement("section");
    sec.className = "section" + (s.className ? " " + s.className : "");
    sec.dataset.section = s.id;
    sec.hidden = true;
    if (s.bare) sec.classList.add("bare");
    else sec.innerHTML = `<div class="section-head"><h2>${esc(s.title)}</h2>${s.desc ? `<p>${esc(s.desc)}</p>` : ""}</div>`;
    const body = document.createElement("div");
    body.className = "section-body";
    s.fields.forEach(f => body.appendChild(renderField(f)));
    sec.appendChild(body);
    dynamic.appendChild(sec);
  });
}

/* =================================================================
   VALUES, VISIBILITY, VALIDATION
   ================================================================= */
/* Per-pass cache: the visibility rules call each other a lot, so values and results
   are remembered while one update/validation runs (nothing changes during it). */
let CACHE = null;
function withCache(fn) {
  if (CACHE) return fn();
  CACHE = { val: new Map(), vis: new Map(), sec: new Map() };
  try { return fn(); } finally { CACHE = null; }
}

function getValue(name) {
  if (CACHE) {
    if (!CACHE.val.has(name)) CACHE.val.set(name, readValue(name));
    return CACHE.val.get(name);
  }
  return readValue(name);
}

function readValue(name) {
  const def = FIELD_MAP[name];
  if (def && def.type === "choice") {
    const opts = def.choices();
    if (opts.length === 1) return opts[0];          // single option = fixed answer
    const sel = form.querySelector(`select[name="${name}"]`);
    return sel && opts.includes(sel.value) ? sel.value : "";
  }
  if (def && def.type === "info") return (typeof def.value === "function" ? def.value() : def.value) || "";
  const els = form.querySelectorAll(`[name="${name}"]`);
  if (!els.length) return "";
  const first = els[0];
  if (first.type === "radio") {
    const c = form.querySelector(`[name="${name}"]:checked`);
    return c ? c.value : "";
  }
  if (first.type === "checkbox") {
    return [...els].filter(e => e.checked).map(e => e.value);
  }
  return first.value.trim();
}

function isEmpty(v) { return Array.isArray(v) ? v.length === 0 : !v; }

/* Lookup of all fields by name (used for nested conditions) */
const FIELD_MAP = {};
SECTIONS.forEach(s => s.fields.forEach(f => { FIELD_MAP[f.name] = f; }));

function visibleByName(name) {
  const f = FIELD_MAP[name];
  return f ? fieldVisible(f) : true;
}

/* Secret preview: typing "SHOW ALL" as Antragsteller:in name reveals every section and field */
const SHOW_ALL_CODE = "SHOW ALL";
function showAll() {
  if (PREVIEW) return true;
  const el = form.querySelector('[name="antrag_name"]');
  return !!el && el.value.trim() === SHOW_ALL_CODE;
}
/* Name of the person creating the request */
function senderName() {
  return IN_OUTLOOK ? (SENDER.name || SENDER.email) : getValue("antrag_name");
}

function fieldVisible(f) {
  if (CACHE) {
    if (!CACHE.vis.has(f.name)) CACHE.vis.set(f.name, computeFieldVisible(f));
    return CACHE.vis.get(f.name);
  }
  return computeFieldVisible(f);
}

function computeFieldVisible(f) {
  if (showAll()) return true;
  if (f.type === "choice" && !f.choices().length && !showAll()) return false;
  if (!f.showIf) return true;
  // Custom rule: showIf can be a function (getValue, isVisible) => boolean
  if (typeof f.showIf === "function") return f.showIf(getValue, visibleByName);
  // A field is only visible if the field it depends on is visible too
  if (!visibleByName(f.showIf.field)) return false;
  const v = getValue(f.showIf.field);
  const allowed = f.showIf.values || [f.showIf.value];
  return Array.isArray(v) ? v.some(x => allowed.includes(x)) : allowed.includes(v);
}

function currentType() { return getValue("eintrittsart"); }

/* Neueintritt: Beschäftigungsart chosen and a valid Eintrittsdatum (within ±1 year) */
function neuComplete() {
  return !!getValue("neu_beschaeftigung") && !!getValue("neu_eintrittsdatum")
    && !dateRangeError(FIELD_MAP["neu_eintrittsdatum"]);
}

function antragComplete() {
  if (IN_OUTLOOK) return true;
  return fieldProblem(FIELD_MAP["antrag_name"]) === null && isValidEmail(getValue("antrag_email"));
}

function sectionVisible(s, t) {
  if (showAll()) return true;
  if (!s.always) {
    // Everything else needs a completed Antragsteller:in and a chosen Eintrittsart
    if (!antragComplete() || !t || !s.showFor.includes(t)) return false;
  }
  return s.showIf ? s.showIf(t) : true;
}

function visibleSections() {
  const t = currentType();
  return SECTIONS.filter(s => sectionVisible(s, t));
}

function updateVisibility() {
  const t = currentType();
  const done = antragComplete();
  emptyEl.hidden = done && !!t;
  emptyEl.textContent = done
    ? "Wählen Sie oben eine Eintrittsart, um fortzufahren."
    : "Bitte zuerst Vor- und Nachname sowie E-Mail (@attensam.at) der Antragsteller:in ausfüllen.";
  withCache(applyHiddenOptions);     // may untick options → must run before the cached pass
  withCache(applyLockedOptions);     // may tick/untick options → also before
  withCache(refreshDynamicOptions);  // may clear a list value → also before
  withCache(() => SECTIONS.forEach(s => {
    const secEl = dynamic.querySelector(`[data-section="${s.id}"]`);
    secEl.hidden = !sectionVisible(s, t);
    s.fields.forEach(f => {
      const el = secEl.querySelector(`[data-field="${f.name}"]`);
      el.hidden = !fieldVisible(f);
      if (f.type === "image" && !el.hidden) {
        const plan = currentPlan();
        const img = el.querySelector("img");
        if (plan && img.getAttribute("src") !== plan.src) {
          img.src = plan.src;
          img.alt = plan.alt;
        }
      }
      if (f.placeholderFn && !el.hidden) {
        const ta = el.querySelector("textarea, input"), ph = f.placeholderFn();
        if (ta.placeholder !== ph) ta.placeholder = ph;
      }
      if (f.type === "info" && typeof f.text === "function") el.querySelector(".info-text").textContent = f.text();
      if (f.noteOn) el.querySelector(".note-msg").hidden = getValue(f.name) !== f.noteOn.value;
      if (f.notes) {
        const html = el.hidden ? "" : f.notes()
          .map(n => `<div class="${n.kind === "warn" ? "note-msg" : "dyn-hint"}">${esc(n.text)}</div>`).join("");
        const box = el.querySelector(".dyn-notes");
        if (box.innerHTML !== html) box.innerHTML = html;
      }
    });
  }));
  withCache(refreshDynamicOptions);
  applyAutoCheck();
  placeNestedFields();

  // Preview mode: nothing is hidden, so creating an e-mail makes no sense
  const preview = showAll();
  if (preview) emptyEl.hidden = true;
  document.getElementById("previewBadge").hidden = !preview;
  const mailBtn = document.getElementById("mailBtn");
  mailBtn.disabled = preview;
  mailBtn.title = preview ? "Im Vorschau-Modus deaktiviert" : "";
}

/* lockOptions: options that are ticked and greyed out; when released they are unticked again */
function applyLockedOptions() {
  SECTIONS.forEach(s => s.fields.forEach(f => {
    if (!f.lockOptions) return;
    const lock = f.lockOptions();
    form.querySelectorAll(`[name="${f.name}"]`).forEach(inp => {
      if (lock.includes(inp.value)) {
        inp.checked = true;
        inp.disabled = true;
        inp.dataset.locked = "1";
      } else if (inp.dataset.locked) {
        inp.checked = false;
        inp.disabled = false;
        delete inp.dataset.locked;
      }
    });
  }));
}

/* hideOptions: checkbox options that are not available right now (they get unticked) */
function applyHiddenOptions() {
  SECTIONS.forEach(s => s.fields.forEach(f => {
    if (!f.hideOptions) return;
    const hide = showAll() ? [] : f.hideOptions();
    form.querySelectorAll(`[name="${f.name}"]`).forEach(inp => {
      const off = hide.includes(inp.value);
      inp.closest(".opt").hidden = off;
      if (off) inp.checked = false;
    });
  }));
}

/* choice: list with options from f.choices(); exactly one option → shown as text */
function refreshChoice(f) {
  const el = form.querySelector(`[data-field="${f.name}"]`);
  const sel = el.querySelector("select"), txt = el.querySelector(".choice-text");
  const opts = f.choices();
  const single = opts.length === 1;
  sel.hidden = single;
  txt.hidden = !single;
  if (single) { txt.textContent = opts[0]; return; }
  const wanted = ["", ...opts];
  if ([...sel.options].map(o => o.value).join("|") === wanted.join("|")) return;
  const keep = sel.value;
  sel.innerHTML = wanted.map(o => `<option value="${esc(o)}">${o ? esc(o) : "Bitte wählen"}</option>`).join("");
  sel.value = opts.includes(keep) ? keep : "";
}

/* dynamicOptions: lists whose entries depend on other answers (e.g. Büronummer per Standort) */
function refreshDynamicOptions() {
  SECTIONS.forEach(s => s.fields.forEach(f => {
    if (f.type === "choice") { refreshChoice(f); return; }
    if (!f.dynamicOptions) return;
    const sel = form.querySelector(`select[name="${f.name}"]`);
    const opts = f.dynamicOptions();
    const current = [...sel.options].map(o => o.value);
    if (current.join("|") === opts.join("|")) return;
    const keep = sel.value;
    sel.innerHTML = opts.map(o => `<option value="${esc(o)}">${o ? esc(o) : "Bitte wählen"}</option>`).join("");
    sel.value = opts.includes(keep) ? keep : "";
  }));
}

/* autoCheck: when option X is ticked, tick option Y and lock it */
function applyAutoCheck() {
  SECTIONS.forEach(s => s.fields.forEach(f => {
    if (!f.autoCheck) return;
    const target = form.querySelector(`[name="${f.name}"][value="${f.autoCheck.check}"]`);
    const on = getValue(f.name).includes(f.autoCheck.when);
    if (on) target.checked = true;
    target.disabled = on;
    target.closest(".opt").title = on ? `Wird bei „${f.autoCheck.when}“ automatisch benötigt` : "";
  }));
}

/* exclusive: ticking one option of the list unticks the others */
function applyExclusive(input) {
  const f = FIELD_MAP[input.name];
  if (!f || !f.exclusive || !input.checked || !f.exclusive.includes(input.value)) return;
  f.exclusive.filter(v => v !== input.value).forEach(v => {
    form.querySelector(`[name="${f.name}"][value="${v}"]`).checked = false;
  });
}

/* nestUnder: move a field directly below the ticked option it belongs to */
function placeNestedFields() {
  SECTIONS.forEach(s => s.fields.forEach(f => {
    if (f.nestInto) {
      const el = form.querySelector(`[data-field="${f.name}"]`);
      const grid = form.querySelector(`[data-field="${f.nestInto}"] .options`);
      el.classList.add("nested-top");
      if (grid.firstElementChild !== el) grid.prepend(el);
      return;
    }
    if (!f.nestUnder) return;
    const el = form.querySelector(`[data-field="${f.name}"]`);
    el.classList.add("nested");
    const v = getValue(f.nestUnder.field);
    const hit = f.nestUnder.values.find(x => v.includes(x));
    if (!hit) return;
    const opt = form.querySelector(`[name="${f.nestUnder.field}"][value="${hit}"]`).closest(".opt");
    // place after the option and after any nested fields already placed there
    let anchor = opt;
    while (anchor.nextElementSibling && anchor.nextElementSibling !== el
           && anchor.nextElementSibling.classList.contains("nested")) anchor = anchor.nextElementSibling;
    if (anchor.nextElementSibling !== el) anchor.after(el);
  }));
}

/* Show/hide a field's error. msg overrides the default text. */
function setError(name, on, msg) {
  const el = form.querySelector(`[data-field="${name}"]`);
  if (!el) return;
  const m = el.querySelector(".error-msg");
  if (m) {
    if (m.dataset.default === undefined) m.dataset.default = m.textContent;
    m.textContent = msg || m.dataset.default;
  }
  el.classList.toggle("has-error", on);
}

/* blockOn: a value that may be selected but stops the form (shows a red message) */
function isBlocked(f) {
  return f.blockOn && fieldVisible(f) && getValue(f.name) === f.blockOn.value;
}

function updateBlockMessages() {
  SECTIONS.forEach(s => s.fields.forEach(f => {
    if (f.check) {
      const msg = fieldVisible(f) ? f.check() : null;
      setError(f.name, !!msg, msg || undefined);
      return;
    }
    if (!f.blockOn) return;
    if (isBlocked(f)) setError(f.name, true, f.blockOn.message);
    else setError(f.name, false);
  }));
}

const ALLOWED_DOMAIN = "attensam.at";
/* Returns null if OK, otherwise the error message */
function emailProblem(v) {
  const m = /^([^\s@]+)@([^\s@]+)$/.exec(v || "");
  if (!m || m[2].toLowerCase() !== ALLOWED_DOMAIN) return `Bitte eine gültige @${ALLOWED_DOMAIN}-Adresse eingeben.`;
  // part before the @: at least 4 characters, or a "." (e.g. m.muster)
  if (m[1].length < 4 && !m[1].includes(".")) return "Der Teil vor dem @ muss mindestens 4 Zeichen lang sein oder einen Punkt enthalten.";
  return null;
}
function isValidEmail(v) { return emailProblem(v) === null; }

/* Returns null if OK, otherwise an error message ("" = default message) */
function fieldProblem(f) {
  const v = getValue(f.name);
  if (f.required && isEmpty(v)) return "";
  if (isBlocked(f)) return f.blockOn.message;
  if (f.check && fieldVisible(f) && f.check()) return f.check();
  if (dateRangeError(f)) return dateRangeError(f);
  if (f.type === "email" && v && !isValidEmail(v)) return emailProblem(v);
  if (f.minWords && v && v.trim().split(/\s+/).length < f.minWords) return "Bitte Vor- und Nachnamen eingeben.";
  if (f.digits && v && !new RegExp(`^\\d{${f.digits}}$`).test(v)) return `Bitte genau ${f.digits} Ziffern eingeben.`;
  return null;
}

/* All visible fields of a section are filled in correctly */
function sectionComplete(id) {
  if (CACHE && CACHE.sec.has(id)) return CACHE.sec.get(id);
  const s = SECTIONS.find(x => x.id === id);
  const ok = s.fields.every(f => !fieldVisible(f) || fieldProblem(f) === null);
  if (CACHE) CACHE.sec.set(id, ok);
  return ok;
}

function validate() { return withCache(validateNow); }
function validateNow() {
  let firstBad = null;
  const mark = (name, msg) => { setError(name, true, msg); if (!firstBad) firstBad = form.querySelector(`[data-field="${name}"]`); };

  form.querySelectorAll(".has-error").forEach(e => e.classList.remove("has-error"));


  visibleSections().forEach(s => s.fields.forEach(f => {
    if (!fieldVisible(f)) return;
    const problem = fieldProblem(f);
    if (problem !== null) mark(f.name, problem || undefined);
  }));

  if (firstBad) {
    firstBad.scrollIntoView({ behavior: "smooth", block: "center" });
    const focusable = firstBad.querySelector("input, select, textarea");
    if (focusable) focusable.focus({ preventScroll: true });
  }
  return !firstBad;
}

/* =================================================================
   PDF + E-MAIL DRAFT (.eml)
   A website cannot attach files to a mailto: link. Instead we build an
   .eml file marked as unsent (X-Unsent: 1). Opening it in Outlook shows
   an editable draft with To, Subject, text and the PDF attached.
   ================================================================= */
const RECIPIENT = "edv@attensam.at";
const CC_BMD = "buchhaltung@attensam.at";        // when BMD is ticked in Zugänge
const CC_VISITENKARTEN = "marketing@attensam.at"; // when Visitenkarten: Ja / erst nach Probezeit

/* Start date per Eintrittsart (used in the subject) */
const START_DATE_FIELD = {
  neueintritt:    "neu_eintrittsdatum",
  wiedereintritt: "wieder_datum",
  befoerderung:   "bef_gueltig",
  geringfuegig:   "ger_beginn"
};

function formatValue(f, v) {
  if (f.displayValue) return f.displayValue();
  if (Array.isArray(v)) return v.join(", ");
  if (f.type === "date" && v) return toATDate(v);
  return v;
}

function todayAT() { return toATDate(toISODate(new Date())); }

function beschaeftigungsart() {
  return currentType() === "neueintritt" ? getValue("neu_beschaeftigung") : EINTRITTSART_LABELS[currentType()];
}
function eintrittsdatumAT() {
  const d = getValue(START_DATE_FIELD[currentType()]);
  return d ? toATDate(d) : "";
}

/* e.g. "Personalaufnahme: Angestellte:r NOVAK Sylvia - 01.12.2026" */
function buildSubject() {
  return withCache(() => {
    const name = [getValue("nachname").toUpperCase(), getValue("vorname")].filter(Boolean).join(" ");
    const d = eintrittsdatumAT();
    return `Personalaufnahme: ${beschaeftigungsart()} ${name}${d ? " - " + d : ""}`;
  });
}

function buildCc() {
  return withCache(() => {
    const cc = [];
    if (visibleByName("zugaenge") && getValue("zugaenge").includes("BMD")) cc.push(CC_BMD);
    if (visibleByName("visitenkarten") && ["Ja", "erst nach Probezeit"].includes(getValue("visitenkarten"))) cc.push(CC_VISITENKARTEN);
    return cc;
  });
}

/* "Label: value" lines for the visible, filled-in fields of a section (null if the section is hidden) */
function sectionLines(id, skip = []) {
  if (!sectionVisibleById(id)) return null;
  const s = SECTIONS.find(x => x.id === id);
  return s.fields
    .filter(f => fieldVisible(f) && f.inEmail !== false && !["hint", "warning", "image"].includes(f.type) && !skip.includes(f.name))
    .map(f => [f.emailLabel || f.label.replace(/\?$/, ""), formatValue(f, getValue(f.name))])
    .filter(([, v]) => v)
    .map(([l, v]) => `${l}: ${v}`);
}

function buildBodyData() {
  return withCache(() => {
    const v = n => (visibleByName(n) ? getValue(n) : "");
    const has = (n, opt) => { const x = v(n); return Array.isArray(x) && x.includes(opt); };
    const sd = [], mdm = [];                 // ServiceDesk / MdM: [text, [sub items]]

    // E-Mail (+ Verteiler)
    if (has("zugaenge", "E-Mail"))
      sd.push(["E-Mail", [`Verteiler: ${verteilerListe().join(", ")}`]]);

    // Zutrittschip
    const chip = [];
    if (v("zutrittschip") === "Ja") chip.push(v("gebaeude") === "Inku"
      ? `Inku Chip: Ja, Zutrittsrechte wie Person: ${v("zutritt_wie")}`
      : `Neuer Chip${v("zutritt_wie") ? ", Zutrittsrechte wie " + v("zutritt_wie") : ""}`);
    if (v("zutrittschip") === "Wird übernommen") chip.push(`Übernahme von ${v("zutritt_uebernahme")}`);
    if (v("zutritt_hg") === "Ja") chip.push(`Hauptgebäude (schwarzer Chip)${v("zutritt_hg_wie") ? ", Zutrittsrechte wie " + v("zutritt_hg_wie") : ""}`);
    if (chip.length) sd.push(["Zutrittschip", chip]);

    // Hardwarebestellung
    if (v("hw_benoetigt") === "Ja") {
      sd.push(["Hardwarebestellung", v("hw_auswahl").map(item => {
        if (item === "Bildschirm") return `Bildschirm: ${v("bildschirm_stueck")} Stück, ${v("bildschirm_groesse")}`;
        if (item === "Docking Station" && v("bildschirm_stueck") === "4") return "Docking Station: 2 Stück (1x Büro, 1x Home Office)";
        return item;
      })]);
    }

    // Bestehende Hardware
    const best = [];
    if (v("buero_hw") === "Ja") best.push(`Büro: ${v("buero_hw_inv")}`);
    if (v("ho_hw") === "Ja") best.push(`Home Office: ${v("ho_hw_inv")}`);
    if (best.length) sd.push(["Bestehende Hardware", best]);

    // Telefonie: Festnetz + Durchwahl → ServiceDesk, Handy → MdM
    if (sectionVisibleById("telefonie")) {
      const tel = v("tel_art"), tl = [];
      const festnetz = ["Nur Festnetz", "Handy & Festnetz"].includes(tel);
      tl.push(`Festnetz: ${festnetz ? `Ja (${v("festnetz_art") || v("festnetz_info") || "3CX"})` : "Nein"}`);
      if (v("tel_uebernahme") === "Ja") tl.push(`Durchwahl: Übernahme von ${v("tel_uebernahme_von")}`);
      if (v("tel_uebernahme") === "Nein") tl.push("Durchwahl: EDV vergibt neue Durchwahl");
      sd.push(["Telefonie", tl]);
      if (["Handy", "Handy & Festnetz"].includes(tel)) {
        const modell = v("handy_modell") === "spezifisches Modell" ? `Spezifisches Modell: ${v("handy_modell_text")}` : v("handy_modell");
        mdm.push(["Telefonie", [`Handy: ${modell}`]]);
      }
    }

    // Zugänge: all → ServiceDesk
    const zgSd = [], zgMdm = [];
    for (const z of (visibleByName("zugaenge") ? getValue("zugaenge") : [])) {
      if (z === "ASE") {
        const ase = v("ase_anwendung") || v("ase_anwendung_text");
        zgSd.push(`ASE: ${ase}, Rechte wie ${v("ase_wie")}`);
      }
      if (z === "ELO") zgSd.push(`ELO: ${v("elo_rechte")}${v("elo_rechte") === "Schreiberechte" ? " (verursacht Lizenzkosten!)" : ""}`);
      if (z === "BMD") zgSd.push(`BMD: ${v("bmd_rechte")}${v("bmd_wie") ? ", Rechte wie " + v("bmd_wie") : ""}`);
      if (z === "RAT_2026") zgSd.push("RAT_2026 (Installation erst dann, wenn der User angemeldet ist!)");
      if (z === "ATT App") zgSd.push(v("att_rechte") === "Wie User"
        ? `ATT App: Rechte wie ${v("att_wie")}`
        : `ATT App (${v("att_rechte")}): ${v("att_module").join(", ")}`);
    }
    if (zgSd.length) sd.push(["Zugänge", zgSd]);
    if (zgMdm.length) mdm.push(["Zugänge", zgMdm]);

    const intro = [`Eintrittsschein von ${senderName()} vom ${todayAT()}`, "Alle Angaben siehe PDF im Anhang."];

    // Notes for the departments in Cc (same rules as buildCc)
    const person = [getValue("vorname"), getValue("nachname")].filter(Boolean).join(" ");
    const notes = [];
    if (has("zugaenge", "BMD"))
      notes.push(`@Buchhaltung: ${person} braucht BMD Berechtigung: ${v("bmd_rechte")}${v("bmd_wie") ? `, Rechte wie User: ${v("bmd_wie")}` : ""}`);
    if (["Ja", "erst nach Probezeit"].includes(v("visitenkarten")))
      notes.push(`@Marketing: ${person} braucht Visitenkarten${v("visitenkarten") === "erst nach Probezeit" ? " nach der Probezeit" : ""}`);

    const categories = [["ServiceDesk", sd], ["MdM", mdm]].filter(([, items]) => items.length);
    return { intro, notes, categories, anmerkungen: v("anmerkungen") };
  });
}

/* Plain-text version (fallback for mail programs without HTML) */
function buildBody() {
  const { intro, notes, categories, anmerkungen } = buildBodyData();
  const L = [...intro, ""];
  if (notes.length) L.push(...notes, "");
  L.push("To-Do Liste:");
  categories.forEach(([title, items]) => {
    L.push("", `${title}:`);
    items.forEach(([t, sub]) => { L.push(`  • ${t}`); sub.forEach(x => L.push(`      ◦ ${x}`)); });
  });
  if (anmerkungen) L.push("", "Sonstige Anmerkungen:", anmerkungen);
  return L.join("\r\n");
}

/* HTML version: real bullet lists in Outlook */
function buildBodyHtml() {
  const { intro, notes, categories, anmerkungen } = buildBodyData();
  const p = t => `<p style="margin:0">${esc(t)}</p>`;
  let h = intro.map(p).join("") + "<br>";
  if (notes.length) h += notes.map(n => `<p style="margin:0"><b>${esc(n.split(":")[0])}:</b>${esc(n.slice(n.indexOf(":") + 1))}</p>`).join("") + "<br>";
  h += `<p style="margin:0"><b>To-Do Liste:</b></p>`;
  categories.forEach(([title, items]) => {
    h += `<p style="margin:10px 0 0"><b>${esc(title)}:</b></p><ul style="margin-top:2px">`;
    items.forEach(([t, sub]) => {
      h += `<li>${esc(t)}`;
      if (sub.length) h += `<ul>${sub.map(x => `<li>${esc(x)}</li>`).join("")}</ul>`;
      h += `</li>`;
    });
    h += `</ul>`;
  });
  if (anmerkungen) h += `<p style="margin:10px 0 0"><b>Sonstige Anmerkungen:</b></p><p style="margin:0">${esc(anmerkungen).replace(/\n/g, "<br>")}</p>`;
  return `<html><head><meta charset="UTF-8"></head><body style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:12pt">${h}</body></html>`;
}

/* ASCII-safe file name: "Personalaufnahme_Huber_Anna" */
function fileBaseName() {
  const clean = t => (t || "")
    .replace(/ä/g, "ae").replace(/ö/g, "oe").replace(/ü/g, "ue")
    .replace(/Ä/g, "Ae").replace(/Ö/g, "Oe").replace(/Ü/g, "Ue").replace(/ß/g, "ss")
    .replace(/[^A-Za-z0-9-]+/g, "_").replace(/^_+|_+$/g, "");
  return ["Personalaufnahme", clean(getValue("nachname")), clean(getValue("vorname"))].filter(Boolean).join("_");
}

function hexToRgb(hex) {
  hex = hex.trim().replace("#", "");
  return [0, 2, 4].map(i => parseInt(hex.substr(i, 2), 16));
}

function buildPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const css = getComputedStyle(document.documentElement);
  const brand = hexToRgb(css.getPropertyValue("--brand"));
  const accent = hexToRgb(css.getPropertyValue("--accent"));
  const muted = [96, 94, 92];
  const W = 210, H = 297, M = 18, labelW = 62;
  const valueX = M + labelW, valueW = W - M - valueX;
  let y;

  // Header bar
  doc.setFillColor(...brand); doc.rect(0, 0, W, 26, "F");
  doc.setFillColor(43, 43, 43); doc.rect(0, 26, W, 1.2, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold"); doc.setFontSize(16);
  doc.text("Personaleintritt - EDV Formular", M, 13);
  doc.setFont("helvetica", "normal"); doc.setFontSize(10);
  doc.text("Eintrittsart: " + EINTRITTSART_LABELS[currentType()], M, 20);
  doc.text(`Erstellt ${senderName() ? "von " + senderName() + " " : ""}am ${todayAT()}`, W - M, 20, { align: "right" });
  y = 38;

  const ensureSpace = h => { if (y + h > H - 20) { doc.addPage(); y = 20; } };

  visibleSections().filter(s => s.inEmail !== false).forEach(s => {
    ensureSpace(18);
    doc.setFillColor(...brand); doc.rect(M, y - 4.5, 1.2, 6, "F");
    doc.setTextColor(...brand); doc.setFont("helvetica", "bold"); doc.setFontSize(12);
    doc.text(s.title, M + 4, y);
    y += 3; doc.setDrawColor(220, 218, 216); doc.line(M, y, W - M, y); y += 6;

    doc.setFontSize(10);
    s.fields.filter(f => fieldVisible(f) && f.inEmail !== false).forEach(f => {
      const v = formatValue(f, getValue(f.name));
      const valLines = doc.splitTextToSize(v || "-", valueW);
      const lblLines = doc.splitTextToSize(f.label, labelW - 4);
      const h = Math.max(valLines.length, lblLines.length) * 5 + 2;
      ensureSpace(h);
      doc.setFont("helvetica", "normal"); doc.setTextColor(...muted); doc.text(lblLines, M, y);
      doc.setFont("helvetica", v ? "bold" : "normal"); doc.setTextColor(32, 31, 30); doc.text(valLines, valueX, y);
      y += h;
    });
    y += 5;
  });

  const pages = doc.getNumberOfPages();
  for (let i = 1; i <= pages; i++) {
    doc.setPage(i); doc.setFontSize(8); doc.setTextColor(...muted);
    doc.text(`Seite ${i} von ${pages}`, W - M, H - 10, { align: "right" });
  }
  return doc;
}

/* --- MIME helpers --- */
function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = ""; bytes.forEach(b => bin += String.fromCharCode(b));
  return btoa(bin);
}
const wrap76 = b64 => b64.replace(/.{1,76}/g, "$&\r\n").trimEnd();
const encodeHeader = t => /^[\x20-\x7E]*$/.test(t) ? t : `=?UTF-8?B?${utf8ToBase64(t)}?=`;

function buildEml(pdfBase64, pdfName) {
  const boundary = "----=_Part_" + Date.now().toString(36);
  return [
    "X-Unsent: 1",
    `To: ${RECIPIENT}`,
    ...(buildCc().length ? [`Cc: ${buildCc().join(", ")}`] : []),
    `Subject: ${encodeHeader(buildSubject())}`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    "",
    `--${boundary}`,
    `Content-Type: multipart/alternative; boundary="${boundary}_alt"`,
    "",
    `--${boundary}_alt`,
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: base64",
    "",
    wrap76(utf8ToBase64(buildBody())),
    "",
    `--${boundary}_alt`,
    "Content-Type: text/html; charset=UTF-8",
    "Content-Transfer-Encoding: base64",
    "",
    wrap76(utf8ToBase64(buildBodyHtml())),
    "",
    `--${boundary}_alt--`,
    "",
    `--${boundary}`,
    `Content-Type: application/pdf; name="${pdfName}"`,
    "Content-Transfer-Encoding: base64",
    `Content-Disposition: attachment; filename="${pdfName}"`,
    "",
    wrap76(pdfBase64),
    "",
    `--${boundary}--`,
    ""
  ].join("\r\n");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function createEmailDraft() {
  const base = fileBaseName();
  const pdfBase64 = buildPDF().output("datauristring").split(",")[1];
  const eml = buildEml(pdfBase64, base + ".pdf");
  downloadBlob(new Blob([eml], { type: "message/rfc822" }), base + ".eml");
}


/* =================================================================
   EVENTS
   ================================================================= */
renderSections();
updateVisibility();

form.addEventListener("change", e => {
  if (e.target.type === "checkbox") applyExclusive(e.target);
  const fieldEl = e.target.closest(".field");
  if (fieldEl) fieldEl.classList.remove("has-error");
  updateVisibility();
  updateBlockMessages();
  const changed = FIELD_MAP[e.target.name];
  if (changed && dateRangeError(changed)) setError(changed.name, true, dateRangeError(changed));
  if (changed && (changed.digits || changed.minWords) && fieldProblem(changed)) setError(changed.name, true, fieldProblem(changed));
  if (changed && changed.type === "email" && getValue(changed.name) && !isValidEmail(getValue(changed.name)))
    setError(changed.name, true, emailProblem(getValue(changed.name)));
  statusEl.textContent = "";
  statusEl.classList.remove("error");
});
form.addEventListener("input", e => {
  const def = FIELD_MAP[e.target.name];
  if (def && def.digits) e.target.value = e.target.value.replace(/\D/g, "").slice(0, def.digits);
  const fieldEl = e.target.closest(".field");
  if (fieldEl) fieldEl.classList.remove("has-error");
  updateVisibility();
});

function runAction(action, doneMsg) {
  if (!validate()) {
    statusEl.textContent = "Bitte die markierten Pflichtfelder ausfüllen.";
    statusEl.classList.add("error");
    return;
  }
  if (!window.jspdf) {
    statusEl.textContent = "PDF-Bibliothek konnte nicht geladen werden. Internetverbindung prüfen.";
    statusEl.classList.add("error");
    return;
  }
  action();
  statusEl.classList.remove("error");
  statusEl.textContent = doneMsg;
}

/* =================================================================
   OUTLOOK ADD-IN
   In Outlook (Classic, New, Web) the form fills the e-mail that is being
   written: To, Cc, subject, to-do list and the PDF attachment.
   Opened in a normal browser it falls back to downloading an .eml file.
   ================================================================= */
let insertedAttachmentId = null;   // PDF added by us, replaced on the next click

const officeCall = fn => new Promise((resolve, reject) =>
  fn(res => res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error)));

async function readSender() {
  const mb = Office.context.mailbox, item = mb.item;
  try {
    if (item.from && item.from.getAsync) {
      const f = await officeCall(cb => item.from.getAsync(cb));
      if (f && (f.displayName || f.emailAddress)) { SENDER = { name: f.displayName, email: f.emailAddress }; return; }
    }
  } catch (e) { /* fall back to the signed-in user */ }
  SENDER = { name: mb.userProfile.displayName, email: mb.userProfile.emailAddress };
}

async function fillOutlookEmail() {
  const item = Office.context.mailbox.item;
  await readSender();                 // the From address may have been changed
  const pdfName = fileBaseName() + ".pdf";
  const pdfBase64 = buildPDF().output("datauristring").split(",")[1];
  const html = buildBodyHtml().replace(/^[\s\S]*<body[^>]*>|<\/body>[\s\S]*$/g, "");   // only the inner part

  await officeCall(cb => item.to.setAsync([RECIPIENT], cb));
  await officeCall(cb => item.cc.setAsync(buildCc(), cb));
  await officeCall(cb => item.subject.setAsync(buildSubject(), cb));

  // replace the whole body: the e-mail contains only the form content (no signature)
  const block = `<div style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:12pt">${html}</div>`;
  await officeCall(cb => item.body.setAsync(block, { coercionType: Office.CoercionType.Html }, cb));

  if (insertedAttachmentId) {
    try { await officeCall(cb => item.removeAttachmentAsync(insertedAttachmentId, cb)); } catch (e) { /* already removed by the user */ }
  }
  insertedAttachmentId = await officeCall(cb => item.addFileAttachmentFromBase64Async(pdfBase64, pdfName, { isInline: false }, cb));
}

async function onMailButton() {
  if (!IN_OUTLOOK) {
    runAction(createEmailDraft, "E-Mail-Datei heruntergeladen. Öffnen Sie sie, um den Entwurf in Outlook zu senden.");
    return;
  }
  if (!validate()) {
    statusEl.textContent = "Bitte die markierten Pflichtfelder ausfüllen.";
    statusEl.classList.add("error");
    return;
  }
  const btn = document.getElementById("mailBtn");
  const updating = !!insertedAttachmentId;
  btn.disabled = true;
  statusEl.classList.remove("error");
  statusEl.textContent = "E-Mail wird ausgefüllt …";
  try {
    await fillOutlookEmail();
    statusEl.textContent = updating ? "E-Mail aktualisiert. Bitte prüfen und senden." : "E-Mail ausgefüllt. Bitte prüfen und senden.";
    btn.textContent = "E-Mail aktualisieren";
  } catch (err) {
    statusEl.textContent = "Fehler beim Ausfüllen der E-Mail: " + (err && err.message ? err.message : err);
    statusEl.classList.add("error");
  } finally {
    btn.disabled = false;
  }
}

document.getElementById("mailBtn").addEventListener("click", onMailButton);

/* Secret preview (all fields): click the Attensam logo 5 times within 3 seconds */
(() => {
  let clicks = [];
  document.querySelector(".logo").addEventListener("click", () => {
    const now = Date.now();
    clicks = clicks.filter(t => now - t < 3000).concat(now);
    if (clicks.length >= 5) { clicks = []; PREVIEW = !PREVIEW; updateVisibility(); }
  });
})();

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(info => {
    if (info.host !== Office.HostType.Outlook || !Office.context.mailbox || !Office.context.mailbox.item) return;
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      statusEl.textContent = "Diese Outlook-Version ist zu alt für das Add-in (Mailbox 1.8 benötigt).";
      statusEl.classList.add("error");
      return;
    }
    IN_OUTLOOK = true;
    document.body.classList.add("in-outlook");
    readSender().then(() => updateVisibility());
    updateVisibility();
    document.getElementById("mailBtn").textContent = "In E-Mail übernehmen";
    document.getElementById("headerHint").textContent =
      "Füllt diese E-Mail aus: Empfänger, Betreff, To-Do-Liste und PDF im Anhang.";
  });
}
/* Lageplan: click to enlarge, click / Esc to close */
const lightbox = document.getElementById("lightbox");
form.addEventListener("click", e => {
  const btn = e.target.closest(".plan-btn");
  if (!btn) return;
  const img = btn.querySelector("img");
  lightbox.querySelector("img").src = img.src;
  lightbox.querySelector("img").alt = img.alt;
  lightbox.hidden = false;
});
lightbox.addEventListener("click", () => { lightbox.hidden = true; });
document.addEventListener("keydown", e => { if (e.key === "Escape") lightbox.hidden = true; });

/* ---------- Feedback page ---------- */
const FEEDBACK_RECIPIENT = "pnov@attensam.at";
const formPage = document.getElementById("formPage"), feedbackPage = document.getElementById("feedbackPage");
const fbTyp = document.getElementById("fbTyp"), fbText = document.getElementById("fbText");

function showFeedbackPage(on) {
  formPage.hidden = on;
  feedbackPage.hidden = !on;
  document.querySelectorAll(".form-only").forEach(b => b.hidden = on);
  document.querySelectorAll(".fb-only").forEach(b => b.hidden = !on);
  statusEl.textContent = "";
  statusEl.classList.remove("error");
  window.scrollTo({ top: 0 });
  if (on) fbTyp.focus();
}
document.getElementById("feedbackBtn").addEventListener("click", () => showFeedbackPage(true));
document.getElementById("fbBackBtn").addEventListener("click", () => showFeedbackPage(false));
fbText.addEventListener("input", () => {
  document.getElementById("fbCount").textContent = fbText.value.length;
  fbText.closest(".field").classList.remove("has-error");
});
fbTyp.addEventListener("change", () => fbTyp.closest(".field").classList.remove("has-error"));

document.getElementById("fbSendBtn").addEventListener("click", () => {
  const typ = fbTyp.value, text = fbText.value.trim();
  fbTyp.closest(".field").classList.toggle("has-error", !typ);
  fbText.closest(".field").classList.toggle("has-error", !text);
  if (!typ || !text) {
    statusEl.textContent = "Bitte die markierten Felder ausfüllen.";
    statusEl.classList.add("error");
    return;
  }
  const subject = `[${typ}] Personaleintrittsschein Feedback`;
  if (IN_OUTLOOK && Office.context.mailbox.displayNewMessageForm) {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [FEEDBACK_RECIPIENT],
      subject,
      htmlBody: esc(text).replace(/\r?\n/g, "<br>")
    });
  } else {
    window.location.href = `mailto:${FEEDBACK_RECIPIENT}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(text.replace(/\r?\n/g, "\r\n"))}`;
  }
  statusEl.classList.remove("error");
  statusEl.textContent = "E-Mail-Entwurf geöffnet. Danke für Ihr Feedback!";
  fbTyp.value = ""; fbText.value = ""; document.getElementById("fbCount").textContent = "0";
});

document.getElementById("resetBtn").addEventListener("click", () => {
  if (!confirm("Alle Eingaben löschen?")) return;
  form.reset();
  form.querySelectorAll(".has-error").forEach(e => e.classList.remove("has-error"));
  updateVisibility();
  statusEl.textContent = "";
  window.scrollTo({ top: 0, behavior: "smooth" });
});
