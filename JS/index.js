const XLSX = require('xlsx');
const QRCode = require('qrcode');
const fs = require('fs');
const path = require('path');

const EXCEL_FILE = 'liste_etudiants.xlsx';
const OUTPUT_DIR = './qr_etudiant';
const COLUMN_NAME = 'nom';
const COLUMN_LASTNAME = 'prenoms';
const COL_MATRICULE = 'matricule';
const COL_NIVEAU = 'niveau';
const COL_MENTION = 'mention';

if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
}

async function processExcel() {
    try {
        const workbook = XLSX.readFile(EXCEL_FILE);
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        console.log(`🚀 ${data.length} lignes détectées. Génération SVG en cours...`);

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            // On simplifie le contenu pour éviter des caractères invisibles qui bloquent le scan
            const content = `${row[COLUMN_NAME] || ''} ${row[COLUMN_LASTNAME] || ''} | ${row[COL_NIVEAU] || ''}-${row[COL_MENTION] || ''}`;
            const matricule = row[COL_MATRICULE];

            if (!content) continue;

            const fileName = `${matricule || i + 1}.svg`;
            const filePath = path.join(OUTPUT_DIR, fileName);

            // 1. On génère le SVG sous forme de chaîne de caractères (string)
            let svgString = await QRCode.toString(content, {
                type: 'svg',
                margin: 4, // Augmenté à 4 pour garantir le scan (Quiet Zone)
                color: {
                    dark: '#000000',
                    light: '#ffffff'
                }
            });

            // 2. NETTOYAGE : On retire l'en-tête XML que Figma n'aime pas
            // On ne garde que ce qui commence par <svg...
            const cleanSvg = svgString.substring(svgString.indexOf('<svg'));

            // 3. Écriture du fichier nettoyé
            fs.writeFileSync(filePath, cleanSvg);
        }

        console.log(`✅ Terminé ! Les SVG nettoyés sont dans : ${OUTPUT_DIR}`);
    } catch (error) {
        console.error('❌ Erreur :', error.message);
    }
}

processExcel();