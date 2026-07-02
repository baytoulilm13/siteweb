const SPREADSHEET_ID = 'REMPLACER_PAR_ID_DU_GOOGLE_SHEET';
const SHEET_NAME = 'Inscriptions';

const HEADERS = [
  'Date de soumission',
  'Parent 1 type',
  'Parent 1 nom',
  'Parent 1 prenom',
  'Adresse',
  'Code postal',
  'Ville',
  'Telephone',
  'Email',
  'Parent 2 type',
  'Parent 2 nom',
  'Parent 2 prenom',
  'Parent 2 adresse',
  'Parent 2 code postal',
  'Parent 2 ville',
  'Parent 2 telephone',
  'Parent 2 email',
  'Enfant nom',
  'Enfant prenom',
  'Date de naissance',
  'Niveau demande',
  'Departement',
  'Remarque medicale',
  'Autorisation sortie',
  'Autorisation image',
  'Acceptation donnees',
  'Signature electronique',
  'Nom signature',
  'Date signature',
  'Horodatage signature'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet();
    ensureHeaders(sheet);

    const rows = (data.enfants || []).map(function(enfant) {
      return [
        data.submitted_at || new Date().toISOString(),
        data.parent_type || '',
        data.parent_nom || '',
        data.parent_prenom || '',
        data.adresse || '',
        data.code_postal || '',
        data.ville || '',
        data.telephone || '',
        data.email || '',
        data.parent2_type || '',
        data.parent2_nom || '',
        data.parent2_prenom || '',
        data.parent2_adresse || '',
        data.parent2_code_postal || '',
        data.parent2_ville || '',
        data.parent2_telephone || '',
        data.parent2_email || '',
        enfant.nom || '',
        enfant.prenom || '',
        enfant.date_naissance || '',
        enfant.niveau_demande || '',
        enfant.departement || '',
        data['remarque_médicale'] || '',
        data.autorisation_sortie || '',
        data.autorisation_image || '',
        data.accept_privacy || '',
        data.signature_electronique || '',
        data.signature_nom || '',
        data.date_signature || '',
        data.signature_timestamp || ''
      ];
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    }

    return jsonResponse({ ok: true, rows: rows.length });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function doGet() {
  return jsonResponse({ ok: true, service: 'Baytoul Ilm inscriptions' });
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
