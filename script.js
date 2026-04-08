function importBusinessEvents() {

  // 🔹 1. Récupérer le fichier Google Sheets actif
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 🔹 2. Récupérer la feuille "RDV" (ou la créer si elle n'existe pas)
  let sheet = ss.getSheetByName("RDV");
  if (!sheet) {
    sheet = ss.insertSheet("RDV");
  }

  // 🔹 3. Vider la feuille (⚠️ supprime tout → OK ici car on reconstruit)
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

  // 🔹 5. Récupérer l'agenda (⚠️ vérifier le nom exact)
  const calendars = CalendarApp.getCalendarsByName("Agenda de nous ! ");
  if (calendars.length === 0) {
    Logger.log("⚠️ Agenda non trouvé !");
    return;
  }
  const calendar = calendars[0];

  // 🔹 6. Définir la période
  const startDate = new Date("2025-01-01");
  const endDate = new Date("2030-01-01");

  // 🔹 7. Récupérer tous les événements
  const events = calendar.getEvents(startDate, endDate);
  Logger.log("Nombre d'événements : " + events.length);

  // 🔹 8. Tableau pour stocker les lignes (🚀 optimisation)
  const rows = [];

  // 🔹 9. Définition des métiers (hors boucle = plus performant)
  const metiers = {
    "Hypno": ["hypno", "hypnose", "séance"],
    "Visite": ["visite", "guide", "getyourguide", "ot", "tbl", "tour"],
    "Pole": ["pole", "stage"]
  };

  // 🔹 10. Parcourir les événements
  events.forEach(event => {

    const title = event.getTitle();
    const description = event.getDescription() || ""; // ⚠️ éviter null
    const start = event.getStartTime();

    const originalTitle = title.trim();
    const lowerTitle = originalTitle.toLowerCase().replace(/\u00A0/g, " ").trim();

    // 🔴 EXCLUSION spécifique
    if (lowerTitle.includes("pole art italy")) return;

    let metier = "";

    // 🔹 Détection métier
    for (let key in metiers) {

      // 🔸 WT / AWT détecté partout
      if (key === "Visite") {
        if (/\b(a?wt)\b/i.test(lowerTitle)) {
          metier = "Visite";
          break;
        }
      }

      // 🔸 Détection classique au début
      if (metiers[key].some(keyword => lowerTitle.startsWith(keyword))) {
        metier = key;
        break;
      }
    }

    // 🔴 Ignorer si non pro
    if (!metier) return;

    // 🔹 Nettoyage du titre
    let cleanTitle = originalTitle;

    for (let key in metiers) {
      const keyword = metiers[key].find(k => lowerTitle.startsWith(k.toLowerCase()));
      if (keyword) {
        cleanTitle = originalTitle.substring(keyword.length).trim();
        break;
      }
    }

    // 🔥 Supprimer les heures du titre
    cleanTitle = cleanTitle.replace(/\b\d{1,2}[hH:]?\d{0,2}\b/g, "");
    cleanTitle = cleanTitle.replace(/\s+/g, " ").trim();

    // 🔹 Séparer client / prestation
    const parts = cleanTitle.split(" - ");
    const client = parts[0] || "";
    const prestation = parts[1] || "";

    // 🔹 Format date
    const date = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const time = Utilities.formatDate(start, Session.getScriptTimeZone(), "HH:mm");
    const mois = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM");

    // 🔹 Extraction montant
    let montant = 0;
    const matches = description.match(/(\d+(?:[.,]\d+)?)\s*(€|euros?|eur)/gi);

    if (matches) {
      matches.forEach(m => {
        const value = m.match(/(\d+(?:[.,]\d+)?)/);
        if (value) {
          montant += parseFloat(value[1].replace(",", "."));
        }
      });
    }

    // 🔹 Statut payé
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
      paye
    ]);
  });

  // 🔹 11. Écriture en une seule fois (🚀 GROS gain de performance)
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  Logger.log("Import terminé !");

  // 🔹 12. Date de mise à jour
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  sheet.getRange("J1").setValue("Dernière mise à jour : " + now);
}
