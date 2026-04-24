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

        console.log(`🚀 ${data.length} lignes détectées. Début de la génération PNG...`);

        // 2. Boucle de génération
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            
            // Construction du contenu du QR Code
            const content = `${row[COL_NOM] || ''} ${row[COL_PRENOM] || ''} | ${row[COL_NIVEAU] || ''}-${row[COL_MENTION] || ''}`;
            const matricule = row[COL_MATRICULE] || `ID-${i + 1}`;

            if (!row[COL_NOM] && !row[COL_PRENOM]) {
                console.warn(`⚠️ Ligne ${i + 2} ignorée (nom/prénom vides).`);
                continue;
            }

            // Extension modifiée en .png
            const fileName = `${matricule}.png`;
            const filePath = path.join(OUTPUT_DIR, fileName);

            // 3. Génération directe en fichier PNG
            // QRCode.toFile gère tout seul l'écriture sur le disque
            await QRCode.toFile(filePath, content, {
                type: 'png',
                width: 1000, // Haute résolution (1000x1000 pixels) pour l'impression
                margin: 4,   // Marge de sécurité (Quiet Zone) indispensable pour le scan mobile
                color: {
                    dark: '#000000', // Carrés noirs
                    light: '#ffffff' // Fond blanc
                }
            });
            
            if (i % 20 === 0 && i > 0) console.log(`... ${i} images PNG générées`);
        }

        console.log(`\n✅ Terminé ! 200+ images PNG sont dans : ${OUTPUT_DIR}`);
        console.log(`💡 Rappel : N'oubliez pas de mettre à jour vos liens dans Excel avec l'extension .png au lieu de .svg.`);

    } catch (error) {
        console.error('❌ Erreur lors du traitement :', error.message);
    }
}

processExcel();