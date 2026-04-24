//fonction principale
function importBusinessEvents() {

  // 🔹 1. Récupérer le fichier Google Sheets actif
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 🔹 2. Récupérer la feuille "RDV" 
  let sheet = ss.getSheetByName("RDV");
  //si elle n'existe pas la créer
  if (!sheet) {  
    sheet = ss.insertSheet("RDV");
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

const lastDataRow = sheet.getLastRow();

if (lastDataRow > 1) {
  sheet.getRange(2, 1, lastDataRow - 1, 14).clearContent();
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

    let factureEnvoyee = "Non";
    let suiviEnvoye = "Non";

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
    let montant = 0;    //montant 

    // 🔹 détecter € OU "euros"
    const matches = description.match(/(\d+(?:[.,]\d+)?)\s*(€|eur|euros?)/gi);

    if (matches) {
      matches.forEach(m => {
        const value = m.match(/(\d+(?:[.,]\d+)?)/);
        if (value) {
          montant += parseFloat(value[1].replace(",", "."));
        }
      });
    }

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
    let paye = "Non";

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

    let suivi = "";

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

const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14);
dataRange.sort([{column: 3, ascending: true}, {column: 5, ascending: true}]);

genererNumerosFacture();

// 🔹 Permet d'écrire "Dernière mise à jour : " dans la case P2
sheet.getRange("P2")
.setValue("Dernière mise à jour : " + now)
.setFontWeight("bold ");

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
}

 // 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir de la case à cocher
function boutonMobile() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");

  const actions = [
    { cell: "S3", func: importBusinessEvents },
    { cell: "S5", func: genererPDF },
    { cell: "S6", func: envoyerFacture }
  ];

  actions.forEach(action => {
    const value = sheet.getRange(action.cell).getValue();

    if (value === true) {
      action.func(); // lance la fonction
      sheet.getRange(action.cell).setValue(false); // reset checkbox
    }
  });
}

// 🔹 Permet de générer un numéro de facture
function genererNumerosFacture() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues();

  let compteur = 1;

  for (let i = 0; i < data.length; i++) {

    const metier = data[i][0];   // colonne Métier
    const date = data[i][2];     // colonne Date
    const numeroExistant = data[i][10]; // colonne facture
    let numeroFacture = "";
    
    // 🔒 SI déjà un numéro → on ne touche pas
    if (numeroExistant) continue;

    if (metier === "Hypno"&& date) {

      const year = new Date(date).getFullYear();

      const numero = String(compteur).padStart(3, "0");

      numeroFacture = `HYP-${year}-${numero}`;

      compteur++;
    }

    // écrire en colonne 11 (Numéro de facture)
    sheet.getRange(i + 2, 11).setValue(numeroFacture);
  }
}

// 🔹 Permet de générer une facture en PDF et de l'enregistrer dans le drive
function genererPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Facture");

  const numeroFacture = sheet.getRange("F6").getValue();

  if (!numeroFacture) {
    SpreadsheetApp.getUi().alert("Numéro de facture manquant");
    return;
  }
  const folder = DriveApp.getFolderById("1SLFLRLfLy-AiRPNDleOCBz5nbEDF54dX"); 
  // 🔍 Vérifier si fichier existe déjà
  const files = folder.getFilesByName(numeroFacture + ".pdf");

  if (files.hasNext()) {
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

  return blob;
}

function envoyerFacture() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Facture");

  const numeroFacture = sheet.getRange("F6").getValue();
  const email = sheet.getRange("J6").getValue(); // adapte selon ta cellule
  const client = sheet.getRange("F4").getValue();

  const sheetRDV = ss.getSheetByName("RDV");
  const data = sheetRDV.getDataRange().getValues();

  let eventId = null;

  for (let i = 1; i < data.length; i++) {
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
  const image = DriveApp.getFileById("1LxTNpm3QTY5U55c4wkIrFVonvo0ZFSHy").getBlob();
  const iconInsta = DriveApp.getFileById("17kKuvOmaY76_r63cYsxe7kTiDaf0c4tc").getBlob();
  const iconFacebook = DriveApp.getFileById("17KdX9oV8TwQUVC6LgqMyN1lCoFR29XOW").getBlob();
  const iconMaps = DriveApp.getFileById("1_ojXvVn7B97v21prV5WDC4pUqNzMotom").getBlob();

MailApp.sendEmail({
  to: email,
  subject: "Facture " + numeroFacture,

  htmlBody: `
    <p>Bonjour ${client},</p>

    <p>Veuillez trouver en pièce jointe votre facture <strong>${numeroFacture}</strong>.</p>

    <p>Merci pour votre confiance 🙏</p>

    <br>

    <p>Cordialement<br>
    Sandy ROYET</p>

    <br>

    <img src="cid:signature" style="width:200px; height:auto;"><br><br>

<table>
  <tr>
    <td>
      <a href="https://www.instagram.com/sandy_hypno/">
        <img src="cid:insta" width="30">
      </a>
    </td>
    <td width="10"></td> <!-- espace -->
    <td>
      <a href="https://www.facebook.com/share/1aLWmSqeii/">
        <img src="cid:facebook" width="30">
      </a>
    </td>
    <td width="10"></td>
    <td>
      <a href="https://maps.app.goo.gl/ZbCGRXKUntZTxT1F6">
        <img src="cid:maps" width="30">
      </a>
    </td>
  </tr>
</table> 
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === numeroFacture) {
      sheetRDV.getRange(i + 1, 12).setValue("oui"); // colonne Facture Envoyée
      break;
    }
  }
  updateCalendarFromSheet();
}

function suiviHypnoJ15() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");
  const data = sheet.getDataRange().getValues();

  const today = new Date();

  const image = DriveApp.getFileById("1LxTNpm3QTY5U55c4wkIrFVonvo0ZFSHy").getBlob();
  const iconInsta = DriveApp.getFileById("17kKuvOmaY76_r63cYsxe7kTiDaf0c4tc").getBlob();
  const iconFacebook = DriveApp.getFileById("17KdX9oV8TwQUVC6LgqMyN1lCoFR29XOW").getBlob();
  const iconMaps = DriveApp.getFileById("1_ojXvVn7B97v21prV5WDC4pUqNzMotom").getBlob();

  for (let i = 1; i < data.length; i++) {

    const metier = data[i][0];
    const nom = data[i][1];
    const dateSeance = new Date(data[i][2]);
    const email = data[i][9];
    const suivi = data[i][12]; // ✅ bonne colonne
    const eventId = data[i][13]; // ✅ récupéré ici

    if (metier !== "Hypno") continue;
    if (!email) continue;
    if (suivi.toLowerCase() === "oui") continue;

    const diffTime = today - dateSeance;
    const diffDays = diffTime / (1000 * 60 * 60 * 24);

    if (diffDays >= 15 && diffDays < 16) {
      // 👉 ENVOI MAIL
      MailApp.sendEmail({
  to: email,
  subject: "Comment vous sentez-vous depuis votre séance ?",

  htmlBody: `
    <p>Bonjour ${nom},</p>

    <p>J'espère que vous allez bien 🙂</p>

    <p>Suite à votre séance d'hypnose, je souhaitais prendre de vos nouvelles.</p>

    <p>
    Avez-vous remarqué des changements ?<br>
    Comment vous sentez-vous aujourd’hui ?
    </p>

    <p>Votre retour est toujours précieux 🙏</p>

    <br>

    <p>Au plaisir d'échanger avec vous,</p>

    <p>Sandy ROYET</p>

    <br>

    <img src="cid:signature" style="width:200px; height:auto;"><br><br>

    <table>
      <tr>
        <td>
          <a href="https://www.instagram.com/sandy_hypno/">
            <img src="cid:insta" width="30">
          </a>
        </td>
        <td width="10"></td>
        <td>
          <a href="https://www.facebook.com/share/1aLWmSqeii/">
            <img src="cid:facebook" width="30">
          </a>
        </td>
        <td width="10"></td>
        <td>
          <a href="https://maps.app.goo.gl/ZbCGRXKUntZTxT1F6">
            <img src="cid:maps" width="30">
          </a>
        </td>
      </tr>
    </table>
  `,

  inlineImages: {
    signature: image,
    insta: iconInsta,
    facebook: iconFacebook,
    maps: iconMaps
  }
});

          // 👉 marquer dans sheet
          sheet.getRange(i + 1, 13).setValue("oui");
    }
  }
  updateCalendarFromSheet();
}

function updateCalendarFromSheet() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const eventId = data[i][13];       // EventID
    const facture = data[i][11];       // Facture envoyée
    const suivi = data[i][12];         // Suivi 15j

    if (!eventId) continue;

    const event = CalendarApp.getEventById(eventId);
    if (!event) continue;

    let desc = event.getDescription() || "";
    const descLower = desc.toLowerCase();

    let updated = false;

    // 🔹 FACTURE
    if (facture === "oui" && !descLower.includes("facture envoyee")) {
      desc += "\nFacture envoyee";
      updated = true;
    }

    // 🔹 SUIVI
    if (suivi === "oui" && !descLower.includes("suivi envoye")) {
      desc += "\nSuivi envoye";
      updated = true;
    }

    // 🔥 écrire seulement si changement
    if (updated) {
      event.setDescription(desc);
    }
  }

}
function remplirFacture() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFacture = ss.getSheetByName("Facture");
  const sheetRDV = ss.getSheetByName("RDV");

  const numero = sheetRDV.getRange("T4").getValue();
  if (!numero) return;

  const data = sheetRDV.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
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

  if (sheet.getName() === "RDV") {

    if (e.range.getA1Notation() === "S4" && e.value === "TRUE") {
      remplirFacture();
      sheet.getRange("S4").setValue(false);
    }
  }
}
