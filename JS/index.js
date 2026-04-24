const XLSX = require('xlsx');
const QRCode = require('qrcode');
const fs = require('fs');
const path = require('path');

// --- CONFIGURATION ---
const EXCEL_FILE = 'liste_etudiants.xlsx';
const OUTPUT_DIR = './qr_etudiant';

// Noms des colonnes dans votre fichier Excel
const COL_NOM = 'nom';
const COL_PRENOM = 'prenoms';
const COL_MATRICULE = 'matricule';
const COL_NIVEAU = 'niveau';
const COL_MENTION = 'mention';

// Créer le dossier de sortie s'il n'existe pas
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
}

async function processExcel() {
    try {
        // 1. Lecture du fichier Excel
        const workbook = XLSX.readFile(EXCEL_FILE);
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        console.log(`🚀 ${data.length} lignes détectées. Début de la génération SVG...`);

        // 2. Boucle de génération
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            
            // Construction du contenu du QR Code
            // Note: On ajoute une barre "|" pour séparer les infos et faciliter la lecture au scan
            const content = `${row[COL_NOM] || ''} ${row[COL_PRENOM] || ''} | ${row[COL_NIVEAU] || ''}-${row[COL_MENTION] || ''}`;
            const matricule = row[COL_MATRICULE] || `ID-${i + 1}`;

            if (!row[COL_NOM] && !row[COL_PRENOM]) {
                console.warn(`⚠️ Ligne ${i + 2} ignorée (nom/prénom vides).`);
                continue;
            }

            const fileName = `${matricule}.svg`;
            const filePath = path.join(OUTPUT_DIR, fileName);

            // 3. Génération du SVG en tant que chaîne de caractères (String)
            let svgString = await QRCode.toString(content, {
                type: 'svg',
                margin: 4, // Marge de sécurité pour garantir le scan sur smartphone
                color: {
                    dark: '#000000', // Carrés noirs
                    light: '#ffffff' // Fond blanc
                }
            });

            // 4. NETTOYAGE CRUCIAL POUR FIGMA
            // On retire l'en-tête XML et le DOCTYPE qui bloquent l'importation via URL
            // On ne garde que la balise <svg> et son contenu
            const cleanSvg = svgString.substring(svgString.indexOf('<svg'));

            // 5. Écriture du fichier sur le disque
            fs.writeFileSync(filePath, cleanSvg);
            
            if (i % 20 === 0 && i > 0) console.log(`... ${i} codes générés`);
        }

        console.log(`\n✅ Terminé ! 200+ fichiers SVG nettoyés sont dans : ${OUTPUT_DIR}`);
        console.log(`💡 Prochaine étape : Poussez ces fichiers sur GitHub et utilisez raw.githubusercontent.com dans Figma.`);

    } catch (error) {
        console.error('❌ Erreur lors du traitement :', error.message);
    }
}

processExcel();