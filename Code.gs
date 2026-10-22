// À coller dans Google Sheets : Extensions > Apps Script
const ENTETES = ["Date", "Nom et prénom", "WhatsApp", "E-mail", "Ville et pays", "Statut", "Niveau", "Ordinateur et logiciels", "Connexion stable", "Jours disponibles", "Comment connu"];

function doPost(e) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (feuille.getLastRow() === 0) feuille.appendRow(ENTETES);
    const p = e.parameter;
    const ligne = feuille.getLastRow() + 1;
    feuille.getRange(ligne, 3).setNumberFormat("@"); // garde le numéro de téléphone en texte
    feuille.getRange(ligne, 1, 1, ENTETES.length).setValues([[
      new Date(), p.nom, p.whatsapp, p.email, p.ville, p.statut, p.niveau,
      p.ordinateur, p.connexion, (e.parameters.jours || []).join(", "), p.source
    ]]);
    return ContentService.createTextOutput("ok");
  } finally {
    verrou.releaseLock();
  }
}
