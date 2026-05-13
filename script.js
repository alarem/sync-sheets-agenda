const CONFIG = {
  SHEET_RDV: "RDV",
  SHEET_FACTURE: "Facture",
  SHEET_CONFIG: "CONFIG"
};

const CELLS = {
  FACTURE_NUMERO: "F6",
  FACTURE_EMAIL: "J6",
  FACTURE_CLIENT: "F4",
  FACTURE_DATE: "F5",
  FACTURE_MONTANT: "E19",
  FACTURE_MODE_PAIEMENT: "F25"
};

const CONFIG_CELLS = {
  RDV_NUMERO_RECHERCHE: "C2",
  STATUS_CELL: "A4",
  LAST_UPDATE_CELL: "A3",

  BOUTON_SYNC: "B1",
  BOUTON_FACTURE: "B2"
};
const STATUS = {
  OUI: "oui",
  NON: "non"
};

const PAYMENT_CODES = {
  "CB": "48",
  "Virement": "58",
  "Espèces": "10",
  "Chèque": "20"
};

const CALENDAR_TAGS = {
  FACTURE: "Facture envoyee",
  SUIVI: "Suivi envoye"
};

const SHEET_LAYOUT = {
  EVENT_ID_COLUMN: 21,
  SUIVI_COLUMN: 22,
  HEADERS: [
    "metier",
    "client",
    "date",
    "mois",
    "heure",
    "montant",
    "paye",
    "mode paiement",
    "n° de telephones",
    "adresse client",
    "adresse emails",
    "type client",
    "siren client",
    "siret client",
    "tva client",
    "pays client",
    "date echeance",
    "numero de facture",
    "facture envoyee",
    "suivi 15j",
    "eventid",
    "style"
  ]
};
const TOTAL_COLUMNS = SHEET_LAYOUT.HEADERS.length;
let COLUMN_CACHE = null;

//fonction principale
function importBusinessEvents() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // 🔹 1. Récupérer le fichier Google Sheets actif
    const ss = getSS();

    // 🔥 nettoyer les messages système
    const statusCell = ss.getSheetByName(CONFIG.SHEET_CONFIG).getRange(CONFIG_CELLS.STATUS_CELL);

    statusCell.clearContent();
    statusCell.setBackground(null);
    statusCell.setFontColor("black");

    // 🔹 2. Récupérer la feuille "RDV" 
    let sheet = ss.getSheetByName(CONFIG.SHEET_RDV);
    //si elle n'existe pas la créer
    if (!sheet) {  
      sheet = ss.insertSheet(CONFIG.SHEET_RDV);
    }

    // 🔥 toujours garantir les entêtes
    const headers = SHEET_LAYOUT.HEADERS;

    // 🔥 FORCER les headers à chaque exécution
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    let lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, TOTAL_COLUMNS).clearContent();
    }

    // 🔹 5. Choisir l'agenda
    const calendars = [
      CalendarApp.getDefaultCalendar(),
      CalendarApp.getCalendarsByName("Agenda de nous ! ")[0]
    ].filter(Boolean);
    // 🔹 6. Définir la période (modifiable)
    const startDate = new Date("2026-01-01"); // début large
    const endDate = new Date("2027-01-01");   // fin large // new Date (); aujourd'hui

    // 🔹 7. Récupérer tous les événements
    let events = [];    //crées une boîte vide

    for (const cal of calendars) {
    if (!cal) continue;
    events.push(...cal.getEvents(startDate, endDate));
    }

    const rows = [];              //crées un tableau vide
    const seenEvents = new Set(); //stocker les informations dedans
    const contactsMap = getContactsMap();

    // 🔹 8. Parcourir chaque événement
    events.forEach(event => {
      try {
      const eventid = event.getId().split("@")[0].trim(); //récupères l’identifiant unique de l’événement,assure format texte, nettoie
      // 🔁 éviter doublons dans le script
      if (seenEvents.has(eventid)) return;
      seenEvents.add(eventid);
      
      const rawTitle = event.getTitle() || "";
      const rawDescription = event.getDescription() || "";
      const title = sanitizeText(rawTitle);
      const description = sanitizeText(rawDescription);
      const desc = description
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, ""); // enlève accents
      
      // 🔸 détecter statut facture / suivi

      let factureenvoyee = "non";
      let suivienvoye = "non";

      if (desc.includes("facture envoyee")) {
        factureenvoyee = "oui";
      }

      if (desc.includes("suivi envoye")) {
        suivienvoye = "oui";
      }
      const start = event.getStartTime();

      // 🔸 9. Filtrer uniquement les événements pro
      const originalTitle = title.trim();
      let lowerTitle = originalTitle.toLowerCase().replace(/\u00A0/g, " ").trim();

        // 🔴 EXCLUSION
      if (lowerTitle.includes("pole art italy")) {    //ne pas tenir compte de "pole art italy"
        return;
      }

      let metier = "";

      const metiers = {
        "Hypno": ["hypno", "hypnose", "séance"],
        "Visite": ["visite", "guide", "getyourguide","ot", "tbl", "tour"],
        "Pole": ["pole", "stage"]
      };

      for (let key in metiers) {

        // 🔸 WT / AWT partout
        if (key === "Visite") {
          if (/\b(a?wt)\b/i.test(lowerTitle)) {
            metier = "Visite";
            break;
          }
        }

        // 🔸 cas classique
        if (metiers[key].some(keyword => lowerTitle.startsWith(keyword))) {
          metier = key;
          break;
        }
      }
      if (!metier) return;   

      // 🔸 10. Nettoyer le titre (enlever [HYPNO] etc.)
      let cleanTitle = originalTitle;

      // 🔥 enlever le mot Metier AU DÉBUT (sans dépendre de lowerTitle)
      for (let key in metiers) {
        for (let keyword of metiers[key]) {
          const regex = new RegExp("^" + keyword, "i");
          if (regex.test(cleanTitle)) {
            cleanTitle = cleanTitle.replace(regex, "").trim();
            break;
          }
        }
      }
      // 🔥 SUPPRIMER LES HEURES (version finale propre)
      cleanTitle = cleanTitle.replace(/\b\d{1,2}\s*(h\s*\d{0,2}|:\s*\d{2})?\b/gi, "");
      cleanTitle = cleanTitle.replace(/\bh\b/gi, "");
      cleanTitle = cleanTitle.replace(/\s+/g, " ").trim();
      let client = cleanTitle;

      // 🔸 12. Formater date
      const date = Utilities.formatDate(
        start,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

      // 🔸 13. Formater heure
      const time = Utilities.formatDate(
        start,
        Session.getScriptTimeZone(),
        "HH:mm"
      );
      // 🔸 13. Formater mois
      const mois = Utilities.formatDate(
        start,
        Session.getScriptTimeZone(),
        "yyyy-MM"
      );

      // 🔸 14. Extraire le montant depuis la description
      let montant = extractMontant(rawDescription);

      // 🔸 15. Détecter le mode de paiement
      let modepaiement = "";

      if (/\bespeces?\b/.test(desc)) {
        modepaiement = "Espèces";
      }
      else if (/\bvirement\b/.test(desc)) {
        modepaiement = "Virement";
      }
      else if (/\bcheque?s?\b/.test(desc)) {
        modepaiement = "Chèque";
      }
      else if (/\b(cb|carte)\b/.test(desc)) {
        modepaiement = "CB";
      }

      // 🔸 16. Extraire le statut Paye
      let paye = "non";

      // 🔥 détecte payé / payée / payés / payées MAIS PAS "heures payées"
      if (
            /\bpaye?s?\b/i.test(desc) &&
            !/heures?\s+paye[eé]s?\b/i.test(desc)
          ) {
            paye = "oui";
          }

      // 🔸 17. Détecter les numéros de téléphones
      
      const phonesRaw = rawDescription.match(/\b(?:\+33|0)[1-9](?:[\s.-]?\d{2}){4}\b/g) || [];

      const phones = phonesRaw.map(p =>
        p.replace(/\s|\./g, "").replace(/^0/, "+33")
      ).join(", ");

      // 🔸 Détecter les numéros adresse mails
      const emails = (rawDescription.match(
      /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g
      ) || []).join(", ");

      
      // 🔸 Détecter les numéros adresse
        let adresse = "";

        // 🔥 PRIORITÉ 1 : description
        const matchAdresse = rawDescription.match(/adresse\s*:\s*([^\n\r]+)/i);

        if (matchAdresse) {

          adresse = matchAdresse[1]
            .replace(/\s+/g, " ")
            .trim();

          // 🔥 nettoyage sécurité
          adresse = adresse.replace(
            /(facture envoyee|suivi envoye|paye|payee)$/i,
            ""
          ).trim();
        }

        // 🔥 PRIORITÉ 2 : carnet
        if (!adresse) {
          const firstEmail = (emails.split(",")[0] || "").trim().toLowerCase();

          if (contactsMap[firstEmail] && contactsMap[firstEmail].adresse) {
            adresse = contactsMap[firstEmail].adresse;
          }
        }

      let style = "vouvoiement"; // défaut

      if (/\btutoiement\b|\btuto\b/.test(desc)) {
        style = "tutoiement";
      } else if (desc.includes("vouvoiement")) {
        style = "vouvoiement";
      }

      // 🔥 PRIORITÉ 2 : contacts (si rien trouvé)
      if (style !== "tutoiement") {
        const firstEmail = (emails.split(",")[0] || "").trim().toLowerCase();

        if (contactsMap[firstEmail]) {
          style = contactsMap[firstEmail].style;
        }
      }

      let typeClient = "particulier";
      if (/\bsarl\b|\bsas\b|\beurl\b|\bsasu\b|\bentreprise\b/i.test(desc)) {
        typeClient = "professionnel";
      }

      let sirenClient = "";
      let siretClient = "";

      // 🔹 SIRET prioritaire (14 chiffres)
      const matchSiret = rawDescription.match(/\b\d{14}\b/);

      if (matchSiret && isValidSiret(matchSiret[0])) {
        siretClient = matchSiret[0];
        sirenClient = siretClient.substring(0, 9);
      } else {

        // 🔹 sinon SIREN seul
        const matchSiren = rawDescription.match(/\b\d{9}\b/);

        if (matchSiren) {
          sirenClient = matchSiren[0];
        }
      }

      let tvaClient = "";
      const matchTVA = rawDescription.match(/\b(FR|BE|DE|IT|ES|LU|NL)[A-Z0-9]{2,12}\b/);
      if (matchTVA && isValidTVA(matchTVA[0])) {
        tvaClient = matchTVA[0];
      }

      let paysClient = detectCountry(rawDescription, adresse);
      let dateEcheance = "";
      const echeance = new Date(start);
      echeance.setDate(echeance.getDate() + 30);

      dateEcheance = Utilities.formatDate(
        echeance,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

      let numerofacture = "";
      // 🔍 chercher un numéro type HYP-2026-001
      const matchFacture = rawDescription.match(/\b[A-Z]{2,5}[-_ ]?\d{4}[-_ ]?\d{1,4}\b/i);

      if (matchFacture) {
        numerofacture = matchFacture[0].toUpperCase();
      }

      // 🔸 Ajouter au tableau 
      rows.push([
        metier,
        client,
        date,
        mois,
        time,
        montant,
        paye,
        modepaiement,
        phones,
        adresse,
        emails,
        typeClient,
        sirenClient,
        siretClient,
        tvaClient,
        paysClient,
        dateEcheance,
        numerofacture,
        factureenvoyee,
        suivienvoye,
        eventid,
        style
      ]);

      } catch (err) {
        console.log("Erreur event: " + err + " | ID: " + event.getId());
      }
    });
    

    // 🔹 Écriture en une seule fois (🚀 GROS gain de performance)
    if (rows.length === 0) {
      console.log("Aucun nouvel événement à ajouter");
    } else {
    const startRow = 2; // TOUJOURS sous les headers

    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // 🔥 Mise à jour de lastRow
    //lastRow += rows.length;
    }
    console.log("Événements importés : " + rows.length);
    console.log("Import terminé !");

    const now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
    );

    if (rows.length > 0) {
      const finalLastRow = rows.length + 1;

      const dataRange = sheet.getRange(2, 1, finalLastRow - 1, TOTAL_COLUMNS);
      dataRange.sort([{column: 3, ascending: true}, {column: 5, ascending: true}]);
    }

    genererNumerosFacture();

    // 🔹 Affiche la dernière mise à jour
    getSheet(CONFIG.SHEET_CONFIG).getRange(CONFIG_CELLS.LAST_UPDATE_CELL)
    .setValue("Dernière mise à jour : " + now)
    .setFontWeight("bold");

    const COL = getColumnMap(sheet);
    const eventIdColumn = COL["eventid"] + 1;

    if (!sheet.isColumnHiddenByUser(eventIdColumn)) {
      sheet.hideColumns(eventIdColumn);
    }
  } finally {
    lock.releaseLock();
  }
}

// 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir d'un bouton dans le bandeau en haut
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🔄 Sync")
  .addItem("Actualiser RDV", "importBusinessEvents")
  .addToUi();

  ui.createMenu("📄 Facture")
    .addItem("Envoyer facture", "envoyerFacture")
    .addItem("Télécharger PDF", "telechargerPDF")
    .addItem("Télécharger JSON", "telechargerJSON")
    .addToUi();
}

 // 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir de la case à cocher
function boutonMobile() {
  const sheet = getSheet(CONFIG.SHEET_CONFIG);

  const actions = [
    { cell: CONFIG_CELLS.BOUTON_SYNC, func: importBusinessEvents },
    { cell: CONFIG_CELLS.BOUTON_FACTURE, func: envoyerFacture }
  ];

  actions.forEach(action => {
    try {
      const value = sheet.getRange(action.cell).getValue();

      if (value === true || value === "TRUE") {

          try {
            action.func();

          } finally {

            // 🔥 décoche TOUJOURS
            sheet.getRange(action.cell).setValue(false);
          }
        }
    } catch (error) {
      console.log(`Erreur ${action.cell}: ${error}`);
    }
  });
}

// 🔹 Permet de générer un Numero de Facture
function genererNumerosFacture() {

  const sheet = getSheet(CONFIG.SHEET_RDV);
  const COL = getColumnMap(sheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data = sheet
    .getRange(2, 1, lastRow - 1, TOTAL_COLUMNS)
    .getValues();

  const updates = [];

  // 🔥 stocke les numéros déjà utilisés
  const numerosExistants = new Set();

  // 🔹 récupérer tous les numéros déjà présents
  for (let i = 0; i < data.length; i++) {

    const numero = data[i][COL["numero de facture"]];

    if (numero) {
      numerosExistants.add(numero);
    }
  }

  let compteur = 1;

  for (let i = 0; i < data.length; i++) {

    const metier = data[i][COL["metier"]];
    const date = data[i][COL["date"]];
    let numeroFacture = data[i][COL["numero de facture"]];

    // 🔒 garder le numéro existant
    if (numeroFacture) {
      updates.push([numeroFacture]);
      continue;
    }

    // 🔹 seulement Hypno
    if (metier === "Hypno" && date) {

      const year = new Date(date).getFullYear();

      // 🔥 chercher un numéro libre
      do {

        const numero = String(compteur).padStart(3, "0");

        numeroFacture = `HYP-${year}-${numero}`;

        compteur++;

      } while (numerosExistants.has(numeroFacture));

      numerosExistants.add(numeroFacture);
    }

    updates.push([numeroFacture]);
  }

  // 🔥 UNE SEULE écriture
  sheet
    .getRange(2, COL["numero de facture"] + 1, updates.length, 1)
    .setValues(updates);
}

// 🔹 Permet de générer une facture en PDF et de l'enregistrer dans le drive
function genererPDF() {
  remplirFacture();
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEET_FACTURE);
  const numerofacture = sheet.getRange(CELLS.FACTURE_NUMERO).getValue();

  if (!numerofacture) {
    throw new Error("Numero de Facture manquant");
  }

  const url = ss.getUrl().replace(/edit$/, "");

  const exportUrl =
    url + "export?format=pdf" +
    "&gid=" + sheet.getSheetId() +
    "&portrait=true" +
    "&fitw=true" +
    "&gridlines=false" +
    "&printtitle=false";

  const token = ScriptApp.getOAuthToken();

  const blob = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
    followRedirects: true
  }).getBlob().setName(numerofacture + ".pdf");

  return blob;
}
function envoyerFacture() {
    const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const factureOk = remplirFacture();

    if (!factureOk) {
      return;
    }
    SpreadsheetApp.flush();
    const ss = getSS();
    const sheetFacture = ss.getSheetByName(CONFIG.SHEET_FACTURE);
    const numerofacture = sheetFacture.getRange(CELLS.FACTURE_NUMERO).getValue();
    const emailRaw = sheetFacture.getRange(CELLS.FACTURE_EMAIL).getValue();
    const email = safeString(emailRaw).split(",")[0].trim();
    const client = sheetFacture.getRange(CELLS.FACTURE_CLIENT).getValue();
    const montant = Number(sheetFacture.getRange(CELLS.FACTURE_MONTANT).getValue()
      );

      if (montant <= 0) {
        setStatus("❌ Montant invalide", true);
        return;
      }
    const sheetRDV = ss.getSheetByName(CONFIG.SHEET_RDV);
    const COL = getColumnMap(sheetRDV);
    const lastRow = sheetRDV.getLastRow();
      if (lastRow < 2) {
        setStatus("❌ Aucun RDV trouvé", true);
        return;
      }
    const data = sheetRDV.getRange(2, 1, lastRow - 1, TOTAL_COLUMNS).getValues();
    let factureDejaEnvoyee = false;
    let eventid = null;
    let style = "vouvoiement";
    let prenom = client;

    for (let i = 0; i < data.length; i++) {

      if (data[i][COL["numero de facture"]] === numerofacture) {

        eventid = data[i][COL["eventid"]];
        style = data[i][COL["style"]];
        prenom = data[i][COL["client"]];

        // 🔒 sécurité anti double envoi
        if (data[i][COL["facture envoyee"]] === "oui") {
          factureDejaEnvoyee = true;
        }

        break;
      }
    }

    if (!email) {
      setStatus("❌ Email manquant", true);
      return;
    }

    if (!isValidEmail(email)) {
      setStatus("❌ Email invalide", true);
      return;
    }
    const pdf = genererPDF();

    if (!pdf) {
    setStatus("❌ Erreur génération PDF", true);
    return;
    }
    const message = generateEmailContent(style, prenom, client);

    if (factureDejaEnvoyee) {

      setStatus("❌ Cette facture a déjà été envoyée.",true);

      return;
    }
    MailApp.sendEmail({
      to: email,
      subject: "Facture " + numerofacture,

      htmlBody: `
        ${message}
        ${getSignatureHtml()}
      `,

      attachments: [pdf],

      
    });

    // 👉 mettre à jour le sheet en mémoire
    const updatesFacture = [];

    for (let i = 0; i < data.length; i++) {

      let valeur = data[i][COL["facture envoyee"]];

      if (data[i][COL["numero de facture"]] === numerofacture) {
        valeur = STATUS.OUI;
      }

      updatesFacture.push([valeur]);
    }

    // 🔥 UNE SEULE écriture
    sheetRDV
      .getRange(2, COL["facture envoyee"] + 1, updatesFacture.length, 1)
      .setValues(updatesFacture);

    updateCalendarFromSheet();
    setStatus("✅ Facture envoyée");

  } finally {
    lock.releaseLock();
  }
}

function suiviHypnoJ15() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const sheet = getSheet(CONFIG.SHEET_RDV);
    const COL = getColumnMap(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLUMNS).getValues();
    const today = new Date();
    const contactsMap = getContactsMap();
    
    const updates = [];

    for (let i = 0; i < data.length; i++) {

      const metier = data[i][COL["metier"]];
      const nom = data[i][COL["client"]];
      const dateSeance = new Date(data[i][COL["date"]]);
      const email = data[i][COL["adresse emails"]];
      const suivi = data[i][COL["suivi 15j"]];

      if (metier !== "Hypno") {
        updates.push([suivi]);
        continue;
      }

      if (!email || (suivi || "").toLowerCase() === "oui") {
        updates.push([suivi]);
        continue;
      }

      const diffDays = (today - dateSeance) / (1000 * 60 * 60 * 24);

      if (diffDays >= 15 && diffDays < 16) {

        const message = generateSuiviContent(email, nom, contactsMap);

        MailApp.sendEmail({
          to: email,
          subject: "Comment vous sentez-vous depuis votre séance ?",
          htmlBody: message + getSignatureHtml(),
        });

        updates.push(["oui"]);

      } else {
        updates.push([suivi]);
      }
    }

    // 🔥 UNE SEULE écriture
    sheet.getRange(2, COL["suivi 15j"] + 1, updates.length, 1).setValues(updates);

    updateCalendarFromSheet();
  } finally {
    lock.releaseLock();
  }  
}

function updateCalendarFromSheet() {

  const sheet = getSheet(CONFIG.SHEET_RDV);
  const COL = getColumnMap(sheet);
  const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLUMNS).getValues();

  for (let i = 0; i < data.length; i++) {

    const eventid = data[i][COL["eventid"]]; // eventid
    const facture = data[i][COL["facture envoyee"]]; // Facture Envoyee
    const suivi = data[i][COL["suivi 15j"]];   // Suivi 15j

    if (!eventid) continue;

    const event = CalendarApp.getEventById(eventid + "@google.com") || CalendarApp.getEventById(eventid);
    if (!event) continue;

    let desc = event.getDescription() || "";
    const descLower = desc.toLowerCase();

    let updated = false;

    if (
      facture === STATUS.OUI &&
      !descLower.includes(CALENDAR_TAGS.FACTURE.toLowerCase())
    ) {
      desc += "\n" + CALENDAR_TAGS.FACTURE;
      updated = true;
    }

    if (
      suivi === STATUS.OUI &&
      !descLower.includes(CALENDAR_TAGS.SUIVI.toLowerCase())
    ) {
      desc += "\n" + CALENDAR_TAGS.SUIVI;
      updated = true;
    }

    if (updated) {
      event.setDescription(desc);
    }
  }
}
function remplirFacture() {
  const ss = getSS();
  const sheetFacture = ss.getSheetByName(CONFIG.SHEET_FACTURE);
  const sheetRDV = ss.getSheetByName(CONFIG.SHEET_RDV);
  const COL = getColumnMap(sheetRDV);

  const numero = getSheet(CONFIG.SHEET_CONFIG).getRange(CONFIG_CELLS.RDV_NUMERO_RECHERCHE).getValue();
  if (!numero) return false;

  const lastRow = sheetRDV.getLastRow();
    if (lastRow < 2) {
    setStatus("❌ Aucun RDV trouvé", true);
    return false;
    }
  const data = sheetRDV.getRange(2, 1, lastRow - 1, TOTAL_COLUMNS).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][COL["numero de facture"]] === numero) {

      sheetFacture.getRange(CELLS.FACTURE_CLIENT).setValue(data[i][COL["client"]]); // Client
      sheetFacture.getRange(CELLS.FACTURE_DATE).setValue(data[i][COL["date"]]); // Date
      sheetFacture.getRange(CELLS.FACTURE_NUMERO).setValue(data[i][COL["numero de facture"]]); // Numero de Facture
      sheetFacture.getRange(CELLS.FACTURE_MONTANT).setValue(data[i][COL["montant"]]); // Montant
      sheetFacture.getRange(CELLS.FACTURE_MODE_PAIEMENT).setValue(data[i][COL["mode paiement"]]); // type de paiment
      sheetFacture.getRange(CELLS.FACTURE_EMAIL).setValue(data[i][COL["adresse emails"]]); // Email

      return true;
    }
  }

  setStatus("❌ Facture introuvable", true);
  return false;
}

function onEdit(e) {
  // 🔒 sécurité si lancé manuellement
  if (!e) return;
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() === CONFIG.SHEET_CONFIG) {

    // 🔹 Si tu changes le numéro → auto
    if (e.range.getA1Notation() === CONFIG_CELLS.RDV_NUMERO_RECHERCHE) {
      remplirFacture();
      SpreadsheetApp.flush();
    }
  }
}
function getSignatureHtml() {
  return `
    <img src="https://drive.google.com/uc?export=view&id=1LxTNpm3QTY5U55c4wkIrFVonvo0ZFSHy" width="100"><br><br>
   
    <table>
      <tr>
        <td><a href="https://www.instagram.com/sandy_hypno/">
          <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" width="30">
        </a></td>
        <td width="10"></td>
        <td><a href="https://www.facebook.com/share/1aLWmSqeii/">
          <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" width="30">
        </a></td>
        <td width="10"></td>
        <td><a href="https://maps.app.goo.gl/ZbCGRXKUntZTxT1F6">
          <img src="https://cdn-icons-png.flaticon.com/512/684/684908.png" width="30">
        </a></td>
      </tr>
    </table>
  `;
}
function safeString(value) {

  if (value === null || value === undefined) {
    return "";
  }

  return value.toString().trim();
}
function getSheet(name) {
  return getSS().getSheetByName(name);
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function extractMontant(description) {
  description = safeString(description);
  const matches = description.match(/(\d+(?:[.,]\d+)?)\s*(€|eur|euros?)/gi);
  if (!matches) return 0;

  return matches.reduce((sum, m) => {
    const value = m.match(/(\d+(?:[.,]\d+)?)/);
    return sum + (value ? parseFloat(value[1].replace(",", ".")) : 0);
  }, 0);
}

function findContactByEmail(email) {

  const sheet = getSheet("Carnet");
  const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return null;
    }
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (let i = 0; i < data.length; i++) {
    const rowEmail = (data[i][3] || "").toLowerCase();

    if (rowEmail.includes(email.toLowerCase())) {
      return {
        prenom: data[i][1],
        nom: data[i][0],
        style: (data[i][4] || "vouvoiement").toLowerCase()
      };
    }
  }

  return null;
}

function generateEmailContent(style, prenom, clientFallback) {

  if (style === "tutoiement") {
    return `
      <p>Coucou ${prenom || clientFallback},</p>
      <p>Je t’envoie ta facture 🙂</p>
      <p>Merci encore 🙏</p>
    `;
  }

  return `
    <p>Bonjour ${prenom || clientFallback},</p>
    <p>Veuillez trouver votre facture en pièce jointe.</p>
    <p>Merci pour votre confiance 🙏</p>
    <p>Cordialement</p>
  `;
}

function getContactsMap() {
  const sheet = getSheet("Carnet");
  const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return {};
    }
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  const map = {};

  for (let i = 0; i < data.length; i++) {
    const emails = (data[i][3] || "").toLowerCase().split(",");

    emails.forEach(e => {
      const email = e.trim();
      if (!email) return;

      map[email] = {
        prenom: data[i][1],
        nom: data[i][0],
        style: (data[i][4] || "vouvoiement").toLowerCase(),
        adresse: data[i][5] || ""
      };
    });
  }

  return map;
}

function generateSuiviContent(email, nom, contactsMap) {

  email = safeString(email).toLowerCase();
  const contact = contactsMap[email];

  if (contact && contact.style === "tutoiement") {
    return `
      <p>Coucou ${contact.prenom},</p>
      <p>J’espère que tu vas bien 🙂</p>
      <p>Je voulais prendre de tes nouvelles après la séance.</p>
      <p>Tu ressens des changements ?</p>
      <p>À bientôt 🙏</p>
    `;
  }

  return `
    <p>Bonjour ${contact ? contact.prenom : nom},</p>
    <p>J'espère que vous allez bien 🙂</p>
    <p>Suite à votre séance d'hypnose, je souhaitais prendre de vos nouvelles.</p>
    <p>Avez-vous remarqué des changements ?</p>
    <p>Votre retour est précieux 🙏</p>
    <p>Cordialement</p>
  `;
}
function telechargerPDF() {

  const pdf = genererPDF();

  const file = DriveApp.createFile(pdf);

  const url = file.getDownloadUrl();

  const html = `
    <html>
      <script>
        window.open("${url}");
        google.script.host.close();
      </script>
    </html>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html),
    "Téléchargement PDF"
  );

  Utilities.sleep(5000);

  file.setTrashed(true);
}

function exportFactureJSON(numerofacture) {
  const sheet = getSheet(CONFIG.SHEET_RDV);
  const COL = getColumnMap(sheet);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL["numero de facture"]] === numerofacture) {

      const montantTTC = Number(data[i][COL["montant"]]) || 0;
      // 🔹 Micro-entreprise = TVA à 0
      const tauxTVA = 0;
      const montantHT = montantTTC;
      const montantTVA = 0;

      const facture = {
        facture: {
          type_document: "FACTURE",
          type_facture: "INVOICE",
          nature: "SERVICE",
          schema_version: "2026-FR",
          type_transaction: "380",
          standard: "Factur-X",
          profil_facturx: "MINIMUM",
          version: "1.0",
          uuid: data[i][COL["numero de facture"]],
          generated_at: new Date().toISOString(),          
          numero: data[i][COL["numero de facture"]],
          date: data[i][COL["date"]],
          date_prestation: data[i][COL["date"]],
          profil: "MINIMUM",
          devise: "EUR",
          format_electronique: "FACTUR-X",
          canal_transmission: "PPF",
          date_echeance: data[i][COL["date echeance"]],
        },

        emetteur: {
          nom: "E.I SANDY ROYET",
          pays: "FR",
          siren: "837870153",
          siret: "83787015300017", 
          tva: "FR26837870153",
          adresse: "ONAE 80 imp. Thomas Alva Edisson 84120 Pertuis"
        },

        client: {
          nom: data[i][COL["client"]],
          email: data[i][COL["adresse emails"]],
          adresse: data[i][COL["adresse client"]],
          type: data[i][COL["type client"]] === "professionnel"
          ? "B2B"
          : "B2C",
          siren: data[i][COL["siren client"]],
          siret: data[i][COL["siret client"]],
          tva: data[i][COL["tva client"]],
          pays: data[i][COL["pays client"]],
        },

        montant: {
          ht: montantHT.toFixed(2),
          total_ht: montantHT.toFixed(2),
          tva: montantTVA.toFixed(2),
          total_tva: montantTVA.toFixed(2),
          ttc: montantTTC.toFixed(2),
          total_ttc: montantTTC.toFixed(2),
          taux_tva: tauxTVA,
          categorie_tva: "E",
          mention_tva: "TVA non applicable - article 293 B du CGI",
          exoneration_tva: "TVA non applicable, art. 293 B du CGI"
        },        
        lignes: [
          {
            description: "Séance hypnose",
            code_service: "9619Z",
            categorie_service: "HYPNOSE",
            quantite: 1,
            prix_unitaire_ht: montantHT.toFixed(2),
            total_ht: montantHT.toFixed(2),
            tva: 0
          }
        ],
        paiement: {
          mode: data[i][COL["mode paiement"]],
          mode_paiement_code: PAYMENT_CODES[data[i][COL["mode paiement"]]] || "1",
          statut_paiement: data[i][6] === "oui"
          ? "PAID"
          : "NOT_PAID",
          date_paiement: data[i][6] === "oui"
          ? data[i][COL["date"]]
          : "",
        }
      };

      return JSON.stringify(facture, null, 2);
    }
  }

  return null;
}

function telechargerJSON() {
  const numero = getSheet(CONFIG.SHEET_CONFIG).getRange(CONFIG_CELLS.RDV_NUMERO_RECHERCHE).getValue();
  const json = exportFactureJSON(numero);

  if (!json) {
    setStatus("❌ Facture introuvable", true);
    return;
  }

  const html = `
    <html>
      <body>
        <a download="${numero}.json"
          href="data:application/json;charset=utf-8,${encodeURIComponent(json)}">
          Télécharger JSON
        </a>

        <script>
          document.querySelector('a').click();
        </script>
      </body>
    </html>
    `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html),
    "Export JSON"
  );
}

function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const map = {};

  headers.forEach((h, i) => {
    const key = h
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();

    map[key] = i;
  });

  return map;
}

function setStatus(message, isError = false) {

  const sheet = getSheet(CONFIG.SHEET_CONFIG);

  const cell = sheet.getRange(CONFIG_CELLS.STATUS_CELL);

  cell.setValue(message)
      .setFontWeight("bold");

  if (isError) {
    cell.setFontColor("red");
  } else {
    cell.setFontColor("green");
  }
}
function isValidEmail(email) {
  email = safeString(email);
  if (!email) return false;

  // 🔹 séparer les emails par virgule
  const emails = email.split(",");

  // 🔹 vérifier chaque email
  return emails.every(e => {

    const cleanEmail = e.trim();

    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail);
  });
}
// 🔹 Création automatique des triggers
function installerTriggers() {

  // 🔥 supprimer anciens triggers pour éviter doublons
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {

    const name = trigger.getHandlerFunction();

    if (
      name === "importBusinessEvents" ||
      name === "suiviHypnoJ15"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 🔹 Sync agenda toutes les heures
  ScriptApp.newTrigger("importBusinessEvents")
    .timeBased()
    .everyHours(1)
    .create();

  // 🔹 Vérification suivi hypno tous les jours à 9h
  ScriptApp.newTrigger("suiviHypnoJ15")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log("Triggers installés !");
}
function sanitizeText(text) {

  text = safeString(text);

  return text
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .trim();
}
function isValidTVA(tva) {

  tva = safeString(tva).replace(/\s/g, "").toUpperCase();

  return /^[A-Z]{2}[A-Z0-9]{2,12}$/.test(tva);
}
function isValidSiret(siret) {

  siret = safeString(siret).replace(/\s/g, "");

  if (!/^\d{14}$/.test(siret)) {
    return false;
  }

  let sum = 0;

  for (let i = 0; i < 14; i++) {

    let digit = parseInt(siret[i]);

    if (i % 2 === 0) {
      digit *= 2;

      if (digit > 9) {
        digit -= 9;
      }
    }

    sum += digit;
  }

  return sum % 10 === 0;
}
function detectCountry(description, adresse) {

  const text = (description + " " + adresse).toLowerCase();

  if (text.includes("belgique") || text.includes("belgium")) return "BE";
  if (text.includes("allemagne") || text.includes("germany")) return "DE";
  if (text.includes("italie") || text.includes("italy")) return "IT";
  if (text.includes("espagne") || text.includes("spain")) return "ES";
  if (text.includes("luxembourg")) return "LU";
  if (text.includes("pays-bas") || text.includes("netherlands")) return "NL";

  return "FR";
}
