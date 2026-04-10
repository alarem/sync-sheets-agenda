function importBusinessEvents() {

  // 🔹 1. Récupérer le fichier Google Sheets actif
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 🔹 2. Récupérer la feuille "RDV" (ou la créer si elle n'existe pas)
  let sheet = ss.getSheetByName("RDV");
  if (!sheet) {
    sheet = ss.insertSheet("RDV");
  }

const lastRow = sheet.getLastRow();

let existingIds = [];

if (lastRow > 1) {
  existingIds = sheet
    .getRange(2, 9, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(String); // 🔥 enlève vides
}

// 🔥 NORMALISATION (très important)
existingIds = existingIds.map(id => id.toString().trim());

const existingIdsSet = new Set(existingIds);

  // 🔹 5. Choisir l'agenda (ici agenda principal)
  //const calendar = CalendarApp.getCalendarsByName("Agenda de nous ! ")[0];
  const calendars = CalendarApp.getAllCalendars();

  // 🔹 6. Définir la période (modifiable)
  const startDate = new Date("2026-01-01"); // début large
  const endDate = new Date("2030-01-01");   // fin large

  // 🔹 7. Récupérer tous les événements
  //const events = calendar.getEvents(startDate, endDate);
  let events = [];

  calendars.forEach(cal => {
    const calEvents = cal.getEvents(startDate, endDate, {max: 500});
    events = events.concat(calEvents);
  });

  const rows = [];
  const seenEvents = new Set();

  // 🔹 DEBUG : nombre d'événements trouvés
  console.log("Nombre d'événements : " + events.length);

  // 🔹 8. Parcourir chaque événement
  events.forEach(event => {

    const eventId = event.getId().toString().trim();
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
    const lowerTitle = originalTitle.toLowerCase().replace(/\u00A0/g, " ").trim();

    // 🔴 EXCLUSION
    if (lowerTitle.includes("pole art italy")) {
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

    for (let key in metiers) {
      const keyword = metiers[key].find(k => lowerTitle.startsWith(k.toLowerCase()));
      
      if (keyword) {
        cleanTitle = originalTitle.substring(keyword.length).trim();
        break;
      }
    }
    // 🔥 SUPPRIMER LES HEURES (AJOUT IMPORTANT)
    cleanTitle = cleanTitle.replace(/\b\d{1,2}([hH:]\d{2})?\b/g, "");
    // nettoyer les espaces
    cleanTitle = cleanTitle.replace(/\s+/g, " ").trim();

    // 🔸 11. Séparer client et prestation
    const parts = cleanTitle.split(" - ");
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
    let montant = 0;

    // 🔹 détecter € OU "euros"
    //const matches = description.match(/(\d+(?:[.,]\d+)?)\s*(€|euros?|eur)/gi);
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
if (!rows || rows.length === 0) {
  console.log("Aucun nouvel événement à ajouter");
} else {
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

  //Logger.log("Import terminé !");
  console.log("Import terminé !");

  const now = Utilities.formatDate(
  new Date(),
  Session.getScriptTimeZone(),
  "dd/MM/yyyy HH:mm"
);

sheet.getRange("J1").setValue("Dernière mise à jour : " + now);
sheet.getRange("J2").setValue("Nombre de RDV : " + rows.length);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔄 Sync")
    .addItem("Actualiser les RDV", "importBusinessEvents")
    .addToUi();
}

function boutonMobile() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RDV");
  const value = sheet.getRange("J4").getValue();

  if (value === true) {
    importBusinessEvents(); // ton script principal
    sheet.getRange("J4").setValue(false); // reset
  }

}

function doGet() {
  importBusinessEvents();

  return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          body {
            font-family: Arial;
            text-align: center;
            padding: 40px;
          }
          button {
            padding: 15px;
            font-size: 16px;
            margin: 10px;
            border-radius: 8px;
            border: none;
            background: #4CAF50;
            color: white;
          }
          a {
            display: inline-block;
            margin-top: 15px;
            font-size: 16px;
          }
        </style>
      </head>
      <body>

        <h2>📱 Gestion RDV</h2>

        <p>✅ Mise à jour effectuée</p>

        <button onclick="window.location.reload()">
          🔄 Actualiser
        </button>

        <br>

        <a href="https://docs.google.com/spreadsheets/d/18MqfHk4_feB8lWY8mH0r3YuBynGJNGlbhrrvfhN_b6g/edit?gid=1226754368#gid=1226754368" target="_blank">
          📊 Voir le tableau
        </a>

      </body>
    </html>
  `);
}



