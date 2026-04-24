const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// === CONFIG ===
const fichierExcel = "liste_etudiants.xlsx";
const dossierImages = "../photo_etudiants";
const colonneMatricule = "matricule";

// === LIRE EXCEL ===
const workbook = XLSX.readFile(fichierExcel);
const sheetName = workbook.SheetNames[0];
const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

// === LISTER IMAGES ===
const images = fs.readdirSync(dossierImages)
  .map(f => f.toLowerCase().trim());

// === VERIFICATION ===
const resultat = data.map(etudiant => {
  let matricule = (etudiant[colonneMatricule] || "")
    .toString()
    .trim()
    .toLowerCase();

  let nomImage = `${matricule}.png`;

  return {
    ...etudiant,
    Photo: images.includes(nomImage) ? "Photo OK" : "Pas de photo"
  };
});

// === STATS ===
const nbOK = resultat.filter(e => e.Photo === "Photo OK").length;
const nbNon = resultat.filter(e => e.Photo === "Pas de photo").length;

console.log(`✔ ${nbOK} étudiants ont une photo`);
console.log(`❌ ${nbNon} étudiants n'ont pas de photo`);

// === EXPORT EXCEL ===
const newSheet = XLSX.utils.json_to_sheet(resultat);
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Resultat");

XLSX.writeFile(newWorkbook, "resultat.xlsx");

console.log("✅ Fichier resultat.xlsx créé !");