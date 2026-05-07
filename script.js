const CONFIG = {
  SHEET_RDV: "RDV",
  SHEET_FACTURE: "Facture",
};

const CELLS = {
  FACTURE_NUMERO: "F6",
  FACTURE_EMAIL: "J6",
  FACTURE_CLIENT: "F4",
  FACTURE_DATE: "F5",
  FACTURE_MONTANT: "E19",
  FACTURE_MODE_PAIEMENT: "F25",
  RDV_NUMERO_RECHERCHE: "T4",
  STATUS_CELL: "Q4",
  LAST_UPDATE_CELL: "Q2"
};

const STATUS = {
  OUI: "oui",
  NON: "non"
};

const CALENDAR_TAGS = {
  FACTURE: "Facture envoyee",
  SUIVI: "Suivi envoye"
};

const SHEET_LAYOUT = {
  TOTAL_COLUMNS: 16,
  EVENT_ID_COLUMN: 15,
  FACTURE_COLUMN: 12
};

//fonction principale
function importBusinessEvents() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // 🔹 1. Récupérer le fichier Google Sheets actif
    const ss = getSS();

    // 🔥 nettoyer les messages système
    const statusCell = ss.getSheetByName(CONFIG.SHEET_RDV).getRange(CELLS.STATUS_CELL);

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
    const headers = [
      "metier",
      "client",
      "date",
      "mois",
      "heure",
      "montant",
      "paye",                // 07 → index 06
      "mode paiement",       // 08 → index 07
      "n° de telephones",    // 09 → index 08
      "adresse client",      // 10 → index 09
      "adresse emails",      // 11 → index 10
      "numero de facture",   // 12 → index 11
      "facture envoyee",     // 13 → index 12
      "suivi 15j",           // 14 → index 13
      "eventid",             // 15 → index 14
      "style"                // 16 → index 15
    ];

    // 🔥 FORCER les headers à chaque exécution
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    let lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
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
      const eventid = event.getId().toString().trim(); //récupères l’identifiant unique de l’événement,assure format texte, nettoie
      // 🔁 éviter doublons dans le script
      if (seenEvents.has(eventid)) return;
      seenEvents.add(eventid);
      
      const title = event.getTitle() || "";           // ex: HYPNO Dupont - Séance
      const description = event.getDescription() || ""; //  || "" évite les crash si déscription vide
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
      let montant = extractMontant(description);

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
      
      const phonesRaw = description.match(/\b(?:\+33|0)[1-9](?:[\s.-]?\d{2}){4}\b/g) || [];

      const phones = phonesRaw.map(p =>
        p.replace(/\s|\./g, "").replace(/^0/, "+33")
      ).join(", ");

      // 🔸 Détecter les numéros adresse mails
      const emails = (description.match(
      /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g
      ) || []).join(", ");

      
      // 🔸 Détecter les numéros adresse
        let adresse = "";

        // 🔥 PRIORITÉ 1 : description
        const matchAdresse = description.match(/adresse\s*:\s*(.+)/i);

        if (matchAdresse) {
          adresse = matchAdresse[1].trim();
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

      // 🔸 20. créer le Numero de Facture à partir de la date et de l'heure
      let numerofacture = "";
      // 🔍 chercher un numéro type HYP-2026-001
      const matchFacture = description.match(/\b[A-Z]{2,5}[-_ ]?\d{4}[-_ ]?\d{1,4}\b/i);

      if (matchFacture) {
        numerofacture = matchFacture[0].toUpperCase();
      }

      // 🔸 21. Ajouter au tableau 
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
    lastRow += rows.length;
    }

    console.log("Import terminé !");

    const now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
    );

    if (rows.length > 0) {
      const finalLastRow = rows.length + 1;

      const dataRange = sheet.getRange(2, 1, finalLastRow - 1, SHEET_LAYOUT.TOTAL_COLUMNS);
      dataRange.sort([{column: 3, ascending: true}, {column: 5, ascending: true}]);
    }

    genererNumerosFacture();

    // 🔹 Permet d'écrire "Dernière mise à jour : " dans la case P2
    sheet.getRange(CELLS.LAST_UPDATE_CELL)
    .setValue("Dernière mise à jour : " + now)
    .setFontWeight("bold");

    // 🔹 Permet de cacher la colonne avec les log google
    if (!sheet.isColumnHiddenByUser(15)) {
      sheet.hideColumns(SHEET_LAYOUT.EVENT_ID_COLUMN);
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
  const sheet = getSheet(CONFIG.SHEET_RDV);

  const actions = [
    { cell: "S3", func: importBusinessEvents },
    { cell: "S4", func: envoyerFacture }
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
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, SHEET_LAYOUT.TOTAL_COLUMNS).getValues();

  let compteur = 1;

  const updates = []; // 🔥 tableau pour stocker les résultats

  for (let i = 0; i < data.length; i++) {

    const metier = data[i][COL["metier"]];
    const date = data[i][COL["date"]];
    const numeroExistant = data[i][COL["numero de facture"]];

    let numerofacture = numeroExistant; 

    // 🔒 SI déjà un numéro → on garde
    if (!numeroExistant && metier === "Hypno" && date) {

      const year = new Date(date).getFullYear();
      const numero = String(compteur).padStart(3, "0");

      numerofacture = `HYP-${year}-${numero}`;
      compteur++;
    }

    updates.push([numerofacture]); // 🔥 on stocke
  }

  // 🔥 UNE SEULE écriture
  sheet.getRange(2, SHEET_LAYOUT.FACTURE_COLUMN, updates.length, 1).setValues(updates);
}

// 🔹 Permet de générer une facture en PDF et de l'enregistrer dans le drive
function genererPDF() {
  remplirFacture();
  SpreadsheetApp.flush();
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEET_FACTURE);
  const COL = getColumnMap(sheet);
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
    muteHttpExceptions: true
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
    const contactsMap = getContactsMap();
    const numerofacture = sheetFacture.getRange(CELLS.FACTURE_NUMERO).getValue();
    const email = sheetFacture.getRange(CELLS.FACTURE_EMAIL).getValue(); // adapte selon ta cellule
    const client = sheetFacture.getRange(CELLS.FACTURE_CLIENT).getValue();

    const sheetRDV = ss.getSheetByName(CONFIG.SHEET_RDV);
    const COL = getColumnMap(sheetRDV);
    const lastRow = sheetRDV.getLastRow();
    const data = sheetRDV.getRange(2, 1, lastRow - 1, SHEET_LAYOUT.TOTAL_COLUMNS).getValues();
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

    const data = sheet.getRange(2, 1, lastRow - 1, SHEET_LAYOUT.TOTAL_COLUMNS).getValues();
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
  const data = sheet.getRange(2, 1, lastRow - 1, SHEET_LAYOUT.TOTAL_COLUMNS).getValues();

  for (let i = 0; i < data.length; i++) {

    const eventid = data[i][COL["eventid"]]; // eventid
    const facture = data[i][COL["facture envoyee"]]; // Facture Envoyee
    const suivi = data[i][COL["suivi 15j"]];   // Suivi 15j

    if (!eventid) continue;

    const event = CalendarApp.getEventById(eventid);
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

  const numero = sheetRDV.getRange(CELLS.RDV_NUMERO_RECHERCHE).getValue();
  if (!numero) return false;

  const lastRow = sheetRDV.getLastRow();
  const data = sheetRDV.getRange(2, 1, lastRow - 1, SHEET_LAYOUT.TOTAL_COLUMNS).getValues();

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
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() === CONFIG.SHEET_RDV) {

    // 🔹 Si tu changes le numéro → auto
    if (e.range.getA1Notation() === CELLS.RDV_NUMERO_RECHERCHE) {
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

function getSheet(name) {
  return getSS().getSheetByName(name);
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function extractMontant(description) {
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

  const contact = contactsMap[email.toLowerCase()];

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

  const base64 = Utilities.base64Encode(pdf.getBytes());
  const html = `
    <html>
      <body>
        <a download="${pdf.getName()}" href="data:application/pdf;base64,${base64}">
          Télécharger la facture
        </a>
        <script>document.querySelector('a').click();</script>
      </body>
    </html>
  `;

  const ui = HtmlService.createHtmlOutput(html).setWidth(11).setHeight(11);
  SpreadsheetApp.getUi().showModalDialog(ui, "Téléchargement...");
}

function exportFactureJSON(numerofacture) {
  const sheet = getSheet(CONFIG.SHEET_RDV);
  const COL = getColumnMap(sheet);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL["numero de facture"]] === numerofacture) {

      const montantTTC = data[i][COL["montant"]];
      const tauxTVA = 0.20; 
      const montantHT = montantTTC / (1 + tauxTVA);
      const montantTVA = montantTTC - montantHT;

      const facture = {
        facture: {
          numero: data[i][COL["numero de facture"]],
          date: data[i][COL["date"]],
        },

        emetteur: {
          nom: "E.I SANDY ROYET",
          siret: "83787015300017", 
          adresse: "ONAE 80 imp. Thomas Alva Edisson 84120 Pertuis"
        },

        client: {
          nom: data[i][COL["client"]],
          email: data[i][COL["adresse emails"]],
          adresse: data[i][COL["adresse client"]],
        },

        montant: {
          ht: montantHT.toFixed(2),
          tva: montantTVA.toFixed(2),
          ttc: montantTTC.toFixed(2),
          taux_tva: tauxTVA
        },

        paiement: {
          mode: data[i][COL["mode paiement"]],
          statut: data[i][6]
        }
      };

      return JSON.stringify(facture, null, 2);
    }
  }

  return null;
}

function telechargerJSON() {
  const numero = getSheet(CONFIG.SHEET_RDV).getRange(CELLS.RDV_NUMERO_RECHERCHE).getValue();
  const json = exportFactureJSON(numero);

  if (!json) {
    setStatus("❌ Facture introuvable", true);
    return;
  }

  const html = `
    <a download="${numero}.json"
       href="data:application/json;charset=utf-8,${encodeURIComponent(json)}">
       Télécharger JSON
    </a>
    <script>document.querySelector('a').click();</script>
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

  const sheet = getSheet(CONFIG.SHEET_RDV);

  const cell = sheet.getRange(CELLS.STATUS_CELL);

  cell.setValue(message)
      .setFontWeight("bold");

  if (isError) {
    cell.setFontColor("red");
  } else {
    cell.setFontColor("green");
  }
}
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}
