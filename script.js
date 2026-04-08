function importBusinessEvents() {

  // 🔹 1. Récupérer le fichier Google Sheets actif
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 🔹 2. Récupérer la feuille "RDV" (ou la créer si elle n'existe pas)
  let sheet = ss.getSheetByName("RDV");
  if (!sheet) {
    sheet = ss.insertSheet("RDV");
  }

  // 🔹 3. Vider la feuille (pour repartir propre)
  sheet.clearContents();

  // 🔹 4. Ajouter les en-têtes
sheet.appendRow([
  "Métier",
  "Client/Lieu",
  "Prestation",
  "Date",
  "Mois",
  "Heure",
  "Montant",
  "Payé"
]);

  // 🔹 5. Choisir l'agenda (ici agenda principal)
  const calendar = CalendarApp.getCalendarsByName("Agenda de nous ! ")[0]; //const calendar = CalendarApp.getDefaultCalendar();

  // 🔹 6. Définir la période (modifiable)
  const startDate = new Date("2025-01-01"); // début large
  const endDate = new Date("2030-01-01");   // fin large

  // 🔹 7. Récupérer tous les événements
  const events = calendar.getEvents(startDate, endDate);

  // 🔹 DEBUG : nombre d'événements trouvés
  Logger.log("Nombre d'événements : " + events.length);

  // 🔹 8. Parcourir chaque événement
  events.forEach(event => {

    const title = event.getTitle();           // ex: HYPNO Dupont - Séance
    const description = event.getDescription();
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
    cleanTitle = cleanTitle.replace(/\b\d{1,2}[hH:]?\d{0,2}\b/g, "");
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
    const matches = description.match(/(\d+(?:[.,]\d+)?)\s*(€|euros?|eur)/gi);

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

    // 🔹 15. Ajouter la ligne dans Google Sheets
    sheet.appendRow([
      metier,
      client,
      prestation,
      date,
      mois,
      time,
      montant,
      paye
    ]);

  });

  Logger.log("Import terminé !");

  const now = Utilities.formatDate(
  new Date(),
  Session.getScriptTimeZone(),
  "dd/MM/yyyy HH:mm"
);

sheet.getRange("J1").setValue("Dernière mise à jour : " + now);
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
  //sheet.getRange("K4").setValue("Actualiser");
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

        <a href="TON_LIEN_SHEET" target="_blank">
          📊 Voir le tableau
        </a>

      </body>
    </html>
  `);
}

function runSync() {
  importBusinessEvents();
  return ContentService.createTextOutput("OK");
}



