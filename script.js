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

let lastRow = sheet.getLastRow(); //Donne-moi le numéro de la dernière ligne remplie dans la feuille

if (lastRow === 0) {              //si la feuille est vide
  sheet.appendRow([               // remplir entête
    "Métier",
    "Client",
    "Date",
    "Mois",
    "Heure",
    "Montant",
    "Payé",
    "mode Paiement",
    "N° de téléphones",
    "Adresse emails",
    "Numéro de Facture",
    "EventID"
  ]);
  lastRow = 1; // 🔥 IMPORTANT, la dernière ligne remplie dans la feuille devient 1
}

let existingIds = [];

if (lastRow > 1) {      //si données présentes
  existingIds = sheet
    .getRange(2, 12, lastRow - 1, 1) //récupère les données
    .getValues()        //Retourne un tableau 2D
    .flat()             // Transforme en tableau simple :
    .filter(String); // 🔥 enlève vides
}

// 🔥 NORMALISATION (très important)

const existingIdsSet = new Set(
  existingIds.map(id => id.toString().trim())   //nettoies chaque ID, toString() → au cas où c’est un nombre (sécurité), trim() → enlève les espaces avant/après
);


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
    // 🆕 éviter doublons déjà présents dans la feuille
    if (existingIdsSet.has(eventId)) return;
    
    const title = event.getTitle() || "";           // ex: HYPNO Dupont - Séance
    const description = event.getDescription() || ""; //  || "" évite les crash si déscription vide
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
    const matches = description.match(/(\d+[.,]?\d*)\s*(€|eur|euros?)/gi);

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

    const desc = description.toLowerCase()
                            .normalize("NFD")
                           .replace(/\p{Diacritic}/gu, "");

    if (/especes?/i.test(desc)) {
      modePaiement = "Espèces";
    } else if (/virement/i.test(desc)) {
      modePaiement = "Virement";
    } else if (/cheques?/i.test(desc)) {
      modePaiement = "Chèque";
    } else if (/\b(cb|carte)\b/i.test(desc)) {
      modePaiement = "CB";
    }

    // 🔸 16. Extraire le statut payé
    let paye = "Non";

    // 🔥 détecte payé / payée / payés / payées MAIS PAS "heures payées"
    if (
        /\bpaye?s?\b/i.test(desc) &&
        !/heures?\s+payee?s?\b/i.test(desc)
          ) {
            paye = "Oui";
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
    //const numeroFacture = "F-" + date.replace(/-/g, "") + "-" + time.split(":")[0];
    const numeroFacture ="";

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
      eventId
    ]);

  });
  

  // 🔹 19. Écriture en une seule fois (🚀 GROS gain de performance)
if (rows.length === 0) {
  console.log("Aucun nouvel événement à ajouter");
} else {
const startRow = lastRow + 1;

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

sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn())
  .sort([{column: 3, ascending: true}, {column: 5, ascending: true}]);

genererNumerosFacture();

// 🔹 Permet d'écrire "Dernière mise à jour : " dans la case P1
sheet.getRange("P1").setValue("Dernière mise à jour : " + now);

// 🔹 Permet de cacher la colonne avec les log google
if (!sheet.isColumnHiddenByUser(12)) {
  sheet.hideColumns(12);
}
}

// 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir d'un bouton dans le bandeau en haut
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🔄 Sync")
    .addItem("Actualiser les RDV", "importBusinessEvents")
    .addToUi();

  ui.createMenu("📄 Facture")
    .addItem("Générer PDF", "genererPDF")
    .addItem("Envoyer facture", "envoyerFacture")
    .addToUi();
}

 // 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir de la case à cocher
function boutonMobile() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");
  const value = sheet.getRange("P4").getValue();

  if (value === true) {
    importBusinessEvents(); // ton script principal
    sheet.getRange("P4").setValue(false); // reset
  }

}

 // 🔹 Permet de lancer la fonction principale (importBusinessEvents) à partir d'un lien internet
function doGet() {
  importBusinessEvents();
  return ContentService.createTextOutput("OK");
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
}
