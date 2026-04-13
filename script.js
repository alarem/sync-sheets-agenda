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
    "Client/Lieu",
    "Prestation",
    "Date",
    "Mois",
    "Heure",
    "Montant",
    "Payé",
    "EventID"
  ]);
  lastRow = 1; // 🔥 IMPORTANT, la dernière ligne remplie dans la feuille devient 1
}

let existingIds = [];

if (lastRow > 1) {      //si données présentes
  existingIds = sheet
    .getRange(2, 9, lastRow - 1, 1) //récupère les données
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

    // 🔥 SUPPRIMER LES HEURES AVANT ANALYSE
    lowerTitle = lowerTitle.replace(/\b\d{1,2}(h\d{0,2}|:\d{2}|h)\b/gi, "");
    lowerTitle = lowerTitle.replace(/\s+/g, " ").trim();

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

    // 🔸 11. Séparer client et prestation
    const parts = cleanTitle.split(/\s*-\s*/);
    const client = parts[0] || "";
    const prestation = parts[1] || "";

    // 🔸 12. Formater date et heure
    const date = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    const time = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      "HH:mm"
    );

    const mois = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      "yyyy-MM"
    );

    // 🔸 13. Extraire le montant depuis la description
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

    // 🔸 14. Extraire le statut payé
    let paye = "Non";
    const payeMatch = description.match(/Payé:\s*(Oui|Non)/i);
    if (payeMatch) {
      paye = payeMatch[1];
    }

    // 🔹 Ajouter au tableau (🚀 plus rapide que appendRow)
    rows.push([
      metier,
      client,
      prestation,
      date,
      mois,
      time,
      montant,
      paye,
      eventId
    ]);

  });
  

  // 🔹 11. Écriture en une seule fois (🚀 GROS gain de performance)
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

sheet.getRange("K1").setValue("Dernière mise à jour : " + now);
if (!sheet.isColumnHiddenByUser(9)) {
  sheet.hideColumns(9);
}
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔄 Sync")
    .addItem("Actualiser les RDV", "importBusinessEvents")
    .addToUi();
}

function boutonMobile() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");
  const value = sheet.getRange("K4").getValue();

  if (value === true) {
    importBusinessEvents(); // ton script principal
    sheet.getRange("K4").setValue(false); // reset
  }

}

function doGet() {
  importBusinessEvents();
  return ContentService.createTextOutput("OK");
}



