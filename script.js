const CONFIG = {
  SHEET_RDV: "RDV",
  SHEET_FACTURE: "Facture",
  DRIVE: {
    FACTURES_FOLDER: "1SLFLRLfLy-AiRPNDleOCBz5nbEDF54dX",
    SIGNATURE: "1LxTNpm3QTY5U55c4wkIrFVonvo0ZFSHy",
    INSTA: "17kKuvOmaY76_r63cYsxe7kTiDaf0c4tc",
    FACEBOOK: "17KdX9oV8TwQUVC6LgqMyN1lCoFR29XOW",
    MAPS: "1_ojXvVn7B97v21prV5WDC4pUqNzMotom"
  }
};

let DRIVE_READY = false;
let PDF_CACHE = {};
const DRIVE_CACHE = {
  signature: null,
  insta: null,
  facebook: null,
  maps: null
};

//fonction principale
function importBusinessEvents() {

  // 🔹 1. Récupérer le fichier Google Sheets actif
  const ss = getSS();
  
  // 🔹 2. Récupérer la feuille "RDV" 
  let sheet = ss.getSheetByName(CONFIG.SHEET_RDV);
  //si elle n'existe pas la créer
  if (!sheet) {  
    sheet = ss.insertSheet(CONFIG.SHEET_RDV);
  }

  // 🔥 toujours garantir les entêtes
  const headers = [
    "Métier",
    "Client",
    "Date",
    "Mois",
    "Heure",
    "Montant",
    "Payé",                // 07 → index 06
    "mode Paiement",       // 08 → index 07
    "N° de téléphones",    // 09 → index 08
    "Adresse emails",      // 10 → index 09
    "Numéro de Facture",   // 11 → index 10
    "Facture Envoyée",     // 12 → index 11
    "Suivi 15j",           // 13 → index 12
    "EventID"              // 14 → index 13
  ];

  // 🔥 FORCER les headers à chaque exécution
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  let lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 14).clearContent();
  }

  // 🔹 5. Choisir l'agenda
  const calendars = CalendarApp.getAllCalendars();  //Récupère TOUS les agendas liés à mon compte Google

  // 🔹 6. Définir la période (modifiable)
  const startDate = new Date("2026-01-01"); // début large
  const endDate = new Date("2027-01-01");   // fin large // new Date (); aujourd'hui

  // 🔹 7. Récupérer tous les événements
  let events = [];    //crées une boîte vide

  calendars.forEach(cal => {      //Pour chaque agenda
    const calEvents = cal.getEvents(startDate, endDate);
    events = events.concat(calEvents);
  });

  const rows = [];              //crées un tableau vide
  const seenEvents = new Set(); //stocker les informations dedans

  // 🔹 DEBUG : nombre d'événements trouvés
  console.log("Nombre d'événements : " + events.length); // afficher dans la console le nombre d'événement

  // 🔹 8. Parcourir chaque événement
  events.forEach(event => {
    try {
    const eventId = event.getId().toString().trim(); //récupères l’identifiant unique de l’événement,assure format texte, nettoie
    // 🔁 éviter doublons dans le script
    if (seenEvents.has(eventId)) return;
    seenEvents.add(eventId);
    
    const title = event.getTitle() || "";           // ex: HYPNO Dupont - Séance
    const description = event.getDescription() || ""; //  || "" évite les crash si déscription vide
    const desc = description
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // enlève accents
    
    // 🔸 détecter statut facture / suivi

    let factureEnvoyee = "non";
    let suiviEnvoye = "non";

    if (desc.includes("facture envoyee")) {
      factureEnvoyee = "oui";
    }

    if (desc.includes("suivi envoye")) {
      suiviEnvoye = "oui";
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

    // 🔥 enlever le mot métier AU DÉBUT (sans dépendre de lowerTitle)
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
    let modePaiement = "";

    if (/\bespeces?\b/.test(desc)) {
      modePaiement = "Espèces";
    }
    else if (/\bvirement\b/.test(desc)) {
      modePaiement = "Virement";
    }
    else if (/\bcheque?s?\b/.test(desc)) {
      modePaiement = "Chèque";
    }
    else if (/\b(cb|carte)\b/.test(desc)) {
      modePaiement = "CB";
    }

    // 🔸 16. Extraire le statut payé
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

    // 🔸 18. Détecter les numéros adresse mails
    const emails = (description.match(
    /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g
    ) || []).join(", ");

    // 🔸 19. créer le numéro de facture à partir de la date et de l'heure

    let numeroFacture = "";

    // 🔍 chercher un numéro type HYP-2026-001
    const matchFacture = description.match(/\b[A-Z]{2,5}[-_ ]?\d{4}[-_ ]?\d{1,4}\b/i);

    if (matchFacture) {
      numeroFacture = matchFacture[0].toUpperCase();
    }

    // 🔸 20. Ajouter au tableau (🚀 plus rapide que appendRow)
    rows.push([
      metier,
      client,
      date,
      mois,
      time,
      montant,
      paye,
      modePaiement,
      phones,
      emails,
      numeroFacture,
      factureEnvoyee,
      suiviEnvoye,
      eventId
    ]);

    } catch (err) {
      console.log("Erreur event: " + err + " | ID: " + event.getId());
    }
  });
  

  // 🔹 19. Écriture en une seule fois (🚀 GROS gain de performance)
  if (rows.length === 0) {
    console.log("Aucun nouvel événement à ajouter");
  } else {
  //const startRow = lastRow + 1;
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

    const dataRange = sheet.getRange(2, 1, finalLastRow - 1, 14);
    dataRange.sort([{column: 3, ascending: true}, {column: 5, ascending: true}]);
  }

  genererNumerosFacture();

  // 🔹 Permet d'écrire "Dernière mise à jour : " dans la case P2
  sheet.getRange("P2")
  .setValue("Dernière mise à jour : " + now)
  .setFontWeight("bold");

  // 🔹 Permet de cacher la colonne avec les log google
  if (!sheet.isColumnHiddenByUser(14)) {
    sheet.hideColumns(14);
  }
}

// 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir d'un bouton dans le bandeau en haut
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🔄 Sync")
  .addItem("Actualiser RDV", "importBusinessEvents")
  .addItem("Sync Calendar", "updateCalendarFromSheet")
  .addToUi();

  ui.createMenu("📄 Facture")
    .addItem("Générer PDF", "genererPDF")
    .addItem("Envoyer facture", "envoyerFacture")
    .addToUi();
    
  ui.createMenu("👤 Contacts")
  .addItem("Sync Contacts", "syncContactsToSheet")
  .addToUi();
}

 // 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir de la case à cocher
function boutonMobile() {
  const sheet = getSheet(CONFIG.SHEET_RDV);

  const actions = [
    { cell: "S3", func: importBusinessEvents },
    { cell: "S4", func: genererPDF },
    { cell: "S5", func: envoyerFacture }
  ];

  actions.forEach(action => {
    try {
      const value = sheet.getRange(action.cell).getValue();

      if (value === true) {
        action.func();
        sheet.getRange(action.cell).setValue(false);
      }
    } catch (error) {
      Logger.log(`Erreur ${action.cell}: ${error}`);
    }
  });
}

// 🔹 Permet de générer un numéro de facture
function genererNumerosFacture() {
  const sheet = getSheet(CONFIG.SHEET_RDV);
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues();

  let compteur = 1;

  const updates = []; // 🔥 tableau pour stocker les résultats

  for (let i = 0; i < data.length; i++) {

    const metier = data[i][0];
    const date = data[i][2];
    const numeroExistant = data[i][10];

    let numeroFacture = numeroExistant; // ⚠️ important

    // 🔒 SI déjà un numéro → on garde
    if (!numeroExistant && metier === "Hypno" && date) {

      const year = new Date(date).getFullYear();
      const numero = String(compteur).padStart(3, "0");

      numeroFacture = `HYP-${year}-${numero}`;
      compteur++;
    }

    updates.push([numeroFacture]); // 🔥 on stocke
  }

  // 🔥 UNE SEULE écriture
  sheet.getRange(2, 11, updates.length, 1).setValues(updates);
}

// 🔹 Permet de générer une facture en PDF et de l'enregistrer dans le drive
function genererPDF() {
  remplirFacture(); // 🔥 sécurise les données

  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEET_FACTURE);
  const numeroFacture = sheet.getRange("F6").getValue();

  if (PDF_CACHE[numeroFacture]) {
    return PDF_CACHE[numeroFacture];
  }

  if (!numeroFacture) {
    SpreadsheetApp.getUi().alert("Numéro de facture manquant");
    return;
  }
  const folder = DriveApp.getFolderById(CONFIG.DRIVE.FACTURES_FOLDER); 
  // 🔍 Vérifier si fichier existe déjà
  const files = folder.getFilesByName(numeroFacture + ".pdf");

  if (files.hasNext()) {
    Logger.log("PDF déjà existant");
    return files.next().getBlob(); // 👉 on retourne l'existant
  }

  const url = ss.getUrl().replace(/edit$/, "");
  
  const exportUrl = url + "export?format=pdf" +
    "&gid=" + sheet.getSheetId() +
    "&portrait=true" +
    "&fitw=true" +
    "&top_margin=0.5" +
    "&bottom_margin=0.5" +
    "&left_margin=0.5" +
    "&right_margin=0.5" +
    "&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false";

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + token
    }
  });

  const blob = response.getBlob().setName(numeroFacture + ".pdf");

  // 📁 Sauvegarde dans Drive
  folder.createFile(blob);
  
  PDF_CACHE[numeroFacture] = blob;

  return blob;
}

function envoyerFacture() {
  initApp();
  remplirFacture(); // 🔥 garantit que tout est à jour
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEET_FACTURE);
  const contactsMap = getContactsMap();
  const numeroFacture = sheet.getRange("F6").getValue();
  const email = sheet.getRange("J6").getValue(); // adapte selon ta cellule
  const client = sheet.getRange("F4").getValue();

  const sheetRDV = ss.getSheetByName(CONFIG.SHEET_RDV);
  const lastRow = sheetRDV.getLastRow();
  const data = sheetRDV.getRange(2, 1, lastRow - 1, 14).getValues();

  let eventId = null;

  for (let i = 0; i < data.length; i++) {
    if (data[i][10] === numeroFacture) { // colonne facture
      eventId = data[i][13];             // colonne EventID
      break;
    }
  }

  if (!email) {
    SpreadsheetApp.getUi().alert("Email manquant");
    return;
  }

  const pdf = genererPDF();

  if (!pdf) {
  SpreadsheetApp.getUi().alert("Erreur génération PDF");
  return;
  }
  
  const image = DRIVE_CACHE.signature;
  const iconInsta = DRIVE_CACHE.insta;
  const iconFacebook = DRIVE_CACHE.facebook;
  const iconMaps = DRIVE_CACHE.maps;
  const message = generateEmailContent(email, client, contactsMap);

  MailApp.sendEmail({
    to: email,
    subject: "Facture " + numeroFacture,

    htmlBody: `
      ${message}
      ${getSignatureHtml()}
    `,

    attachments: [pdf],
    inlineImages: {
      signature: image,
      insta: iconInsta,
      facebook: iconFacebook,
      maps: iconMaps
    }
  });
  // 👉 mettre à jour le sheet aussi
  for (let i = 0; i < data.length; i++) {
    if (data[i][10] === numeroFacture) {
      sheetRDV.getRange(i + 2, 12).setValue("oui"); // colonne Facture Envoyée
      break;
    }
  }
  updateCalendarFromSheet();
}

function suiviHypnoJ15() {

  initApp();
  const sheet = getSheet(CONFIG.SHEET_RDV);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const today = new Date();
  const contactsMap = getContactsMap();
  const updates = [];
  const image = DRIVE_CACHE.signature;
  const iconInsta = DRIVE_CACHE.insta;
  const iconFacebook = DRIVE_CACHE.facebook;
  const iconMaps = DRIVE_CACHE.maps;

  for (let i = 0; i < data.length; i++) {

    const metier = data[i][0];
    const nom = data[i][1];
    const dateSeance = new Date(data[i][2]);
    const email = data[i][9];
    const suivi = data[i][12];

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
        inlineImages: {
          signature: image,
          insta: iconInsta,
          facebook: iconFacebook,
          maps: iconMaps
        }
      });

      updates.push(["oui"]);

    } else {
      updates.push([suivi]);
    }
  }

  // 🔥 UNE SEULE écriture
  sheet.getRange(2, 13, updates.length, 1).setValues(updates);

  updateCalendarFromSheet();
}

function updateCalendarFromSheet() {

  const sheet = getSheet(CONFIG.SHEET_RDV);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  for (let i = 0; i < data.length; i++) {

    const eventId = data[i][13]; // EventID
    const facture = data[i][11]; // Facture envoyée
    const suivi = data[i][12];   // Suivi 15j

    if (!eventId) continue;

    const event = CalendarApp.getEventById(eventId);
    if (!event) continue;

    let desc = event.getDescription() || "";
    const descLower = desc.toLowerCase();

    let updated = false;

    if (facture === "oui" && !descLower.includes("facture envoyee")) {
      desc += "\nFacture envoyee";
      updated = true;
    }

    if (suivi === "oui" && !descLower.includes("suivi envoye")) {
      desc += "\nSuivi envoye";
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

  const numero = sheetRDV.getRange("T4").getValue();
  if (!numero) return;

  const lastRow = sheetRDV.getLastRow();
  const data = sheetRDV.getRange(2, 1, lastRow - 1, 14).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][10] === numero) {

      sheetFacture.getRange("F4").setValue(data[i][1]); // Client
      sheetFacture.getRange("F5").setValue(data[i][2]); // Date
      sheetFacture.getRange("F6").setValue(data[i][10]); // Numéro de facture
      sheetFacture.getRange("E19").setValue(data[i][5]); // Montant
      sheetFacture.getRange("F25").setValue(data[i][7]); // type de paiment
      sheetFacture.getRange("J6").setValue(data[i][9]); // Email

      return;
    }
  }

  SpreadsheetApp.getUi().alert("Facture introuvable");
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() === CONFIG.SHEET_RDV) {

    // 🔹 Si tu changes le numéro → auto
    if (e.range.getA1Notation() === "T4") {
      remplirFacture();
    }
  }
}

function getSignatureHtml() {
  return `
    <br>
    <img src="cid:signature" style="width:200px; height:auto;"><br><br>

    <table>
      <tr>
        <td><a href="https://www.instagram.com/sandy_hypno/">
          <img src="cid:insta" width="30">
        </a></td>
        <td width="10"></td>
        <td><a href="https://www.facebook.com/share/1aLWmSqeii/">
          <img src="cid:facebook" width="30">
        </a></td>
        <td width="10"></td>
        <td><a href="https://maps.app.goo.gl/ZbCGRXKUntZTxT1F6">
          <img src="cid:maps" width="30">
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

function loadDriveAssets() {
  if (DRIVE_READY) return;

  try {
    DRIVE_CACHE.signature = DriveApp.getFileById(CONFIG.DRIVE.SIGNATURE).getBlob();
    DRIVE_CACHE.insta = DriveApp.getFileById(CONFIG.DRIVE.INSTA).getBlob();
    DRIVE_CACHE.facebook = DriveApp.getFileById(CONFIG.DRIVE.FACEBOOK).getBlob();
    DRIVE_CACHE.maps = DriveApp.getFileById(CONFIG.DRIVE.MAPS).getBlob();

    DRIVE_READY = true;

  } catch (err) {
    Logger.log("Erreur Drive (retry): " + err);

    Utilities.sleep(1000); // ⏳ pause 1 seconde

    // 🔁 retry 1 fois
    DRIVE_CACHE.signature = DriveApp.getFileById(CONFIG.DRIVE.SIGNATURE).getBlob();
    DRIVE_CACHE.insta = DriveApp.getFileById(CONFIG.DRIVE.INSTA).getBlob();
    DRIVE_CACHE.facebook = DriveApp.getFileById(CONFIG.DRIVE.FACEBOOK).getBlob();
    DRIVE_CACHE.maps = DriveApp.getFileById(CONFIG.DRIVE.MAPS).getBlob();

    DRIVE_READY = true;
  }
}

function extractMontant(description) {
  const matches = description.match(/(\d+(?:[.,]\d+)?)\s*(€|eur|euros?)/gi);
  if (!matches) return 0;

  return matches.reduce((sum, m) => {
    const value = m.match(/(\d+(?:[.,]\d+)?)/);
    return sum + (value ? parseFloat(value[1].replace(",", ".")) : 0);
  }, 0);
}

function syncContactsToSheet() {

  const ss = getSS();
  let sheet = ss.getSheetByName("Carnet");

  if (!sheet) {
    sheet = ss.insertSheet("Carnet");
  }

  const headers = ["Nom", "Prénom", "Téléphone", "Email"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  const connections = People.People.Connections.list('people/me', {
    pageSize: 1000,
    personFields: 'names,emailAddresses,phoneNumbers'
  });

  const rows = [];

  (connections.connections || []).forEach(person => {
    try {

      const names = person.names ? person.names[0] : {};
      const firstName = names.givenName || "";
      const lastName = names.familyName || "";

      const phones = (person.phoneNumbers || [])
      .map(p => {
        if (!p.value) return "";

        return p.value
          .replace(/\s|\./g, "")
          .replace(/^0/, "+33")
          .replace(/-/g, ""); // bonus : enlève tirets
      })
      .filter(p => p) // supprime les vides
      .join(", ");

      const emails = (person.emailAddresses || [])
        .map(e => e.value)
        .join(", ");

      if (!firstName && !lastName && !phones && !emails) return;

      rows.push([
        lastName,
        firstName,
        phones,
        emails
      ]);

    } catch (err) {
      console.log("Erreur contact: " + err);
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  console.log("Contacts importés : " + rows.length);
}

function findContactByEmail(email) {

  const sheet = getSheet("Carnet");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

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

function generateEmailContent(email, clientFallback, contactsMap) {

  const contact = contactsMap[email.toLowerCase()];

  if (contact) {

    if (contact.style === "tutoiement") {
      return `
        <p>Coucou ${contact.prenom},</p>
        <p>Je t’envoie ta facture 🙂</p>
        <p>Merci encore 🙏</p>
        <p>Sandy</p>
      `;
    }

    return `
      <p>Bonjour ${contact.prenom},</p>
      <p>Veuillez trouver votre facture en pièce jointe.</p>
      <p>Merci pour votre confiance 🙏</p>
      <p>Cordialement<br>Sandy ROYET</p>
    `;
  }

  return `
    <p>Bonjour ${clientFallback || ""},</p>
    <p>Veuillez trouver votre facture en pièce jointe.</p>
    <p>Merci pour votre confiance 🙏</p>
    <p>Cordialement<br>Sandy ROYET</p>
  `;
}

function getContactsMap() {
  const sheet = getSheet("Carnet");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  const map = {};

  for (let i = 0; i < data.length; i++) {
    const emails = (data[i][3] || "").toLowerCase().split(",");

    emails.forEach(e => {
      const email = e.trim();
      if (!email) return;

      map[email] = {
        prenom: data[i][1],
        nom: data[i][0],
        style: (data[i][4] || "vouvoiement").toLowerCase()
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
      <p>Sandy</p>
    `;
  }

  return `
    <p>Bonjour ${contact ? contact.prenom : nom},</p>
    <p>J'espère que vous allez bien 🙂</p>
    <p>Suite à votre séance d'hypnose, je souhaitais prendre de vos nouvelles.</p>
    <p>Avez-vous remarqué des changements ?</p>
    <p>Votre retour est précieux 🙏</p>
    <p>Cordialement<br>Sandy ROYET</p>
  `;
}
function testDrive() {
  const file = DriveApp.getFileById(CONFIG.DRIVE.SIGNATURE);
  Logger.log(file.getName());
}
function initApp() {
  loadDriveAssets();
}
