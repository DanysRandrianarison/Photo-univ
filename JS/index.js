const XLSX = require('xlsx');
const QRCode = require('qrcode');
const fs = require('fs');
const path = require('path');

// Configuration
const EXCEL_FILE = 'liste_etudiants.xlsx';
const OUTPUT_DIR = './qr_etudiant';
const COLUMN_NAME = 'nom'; // Nom de la colonne dans votre Excel
const COLUMN_LASTNAME = 'prenoms'
const COLUMN_FONCTION = 'fonction'
const COL_MATRICULE = 'matricule'
const COL_CIN = 'cin'
const COL_NIVEAU = 'niveau'
const COL_MENTION ='mention'

// Créer le dossier de sortie s'il n'existe pas
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
}

async function processExcel() {
    try {
        // 1. Lecture du fichier Excel
        const workbook = XLSX.readFile(EXCEL_FILE);
        const sheetName = workbook.SheetNames[0]; // On prend la première feuille
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        console.log(`🚀 ${data.length} lignes détectées. Début de la génération...`);

        // 2. Boucle de génération
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
           const content = 
                 `${row[COLUMN_NAME] || ''} ${row[COLUMN_LASTNAME] || ''}
                  ${row[COL_MATRICULE] || ''}
                  ${row[COL_NIVEAU] || ''}-${row[COL_MENTION] || ''}`;
           

            const matricule = row[COL_MATRICULE];

            if (!content) {
                console.warn(`⚠️ Ligne ${i + 2} : Colonne "${COLUMN_NAME}" vide, ignorée.`);
                continue;
            }

            // Nom du fichier : on utilise un ID ou l'index pour éviter les doublons
            const fileName = `${matricule || i + 1}.svg`;
            const filePath = path.join(OUTPUT_DIR, fileName);

            await QRCode.toFile(filePath, String(content), {
                type: 'svg',
                width: 400,
                margin: 2,
                color: {
                    dark: '#000000',
                    light: '#ffffff'
                }
            });
        }

        console.log(`✅ Terminé ! Les images sont dans le dossier : ${OUTPUT_DIR}`);
    } catch (error) {
        console.error('❌ Erreur lors du traitement :', error.message);
    }
}

processExcel();