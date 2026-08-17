// ============================================================
// BAYTOUL ILM — INSCRIPTION + RÈGLEMENT + SIGNATURE
// Année 2026-2027
// ============================================================

const SPREADSHEET_ID = '1xYej3l4qNDCh0VCg-A_aZJm_6tvhYm5U5jhcP2XM-o4';
const SHEET_NAME = 'Inscriptions';

const ADMIN_EMAIL = 'baytoulilm13@gmail.com';

const TEMPLATE_DOC_ID = '1ViLGz4poyhTd-p0W9NraYNd7CroNTXl0';
const REGLEMENT_DOC_ID = '14jhoKqu5x0qaUckb2ZWJqGz8VcT97Bqg';
const OUTPUT_FOLDER_ID = '1voXoPD7vxvzfjqD5HOK5kFBCqGvJ4dto';

const REGLEMENT_VERSION = '2026-2027';

const TURQUOISE = '#0096A0';

const NIVEAUX_AUTORISES = [
  'Classe Initiation',
  'Classe Approfondissement',
  'Parcours Hifdh',
  'Parcours Qirâa'
];

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
  'Horodatage signature',
  'Google Docs',
  'Dossier PDF',
  'Dossier Word',
  'Reglement PDF',
  'Reglement Word'
];


// ============================================================
// DO POST
// ============================================================

function doPost(e) {

  try {

    const data = JSON.parse(
      (e && e.postData && e.postData.contents) || '{}'
    );

    const now = new Date();

    const enfants = Array.isArray(data.enfants)
      ? data.enfants
      : [];

    if (!enfants.length) {
      throw new Error('Aucun enfant à inscrire n’a été reçu.');
    }

    // Nettoyage des parcours
    data.enfants = enfants.map(function(enfant) {

      const enfantClean = Object.assign({}, enfant);

      enfantClean.niveau_demande =
        normalizeNiveau(enfantClean.niveau_demande);

      return enfantClean;

    });


    // ========================================================
    // GOOGLE SHEETS
    // ========================================================

    const sheet = getSheet();

    ensureHeaders(sheet);


    const rows = data.enfants.map(function(enfant) {

      return [

        data.submitted_at || now.toISOString(),

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

        data['remarque_médicale'] ||
        data.remarque_medicale ||
        '',

        data.autorisation_sortie || '',
        data.autorisation_image || '',
        data.accept_privacy || '',

        data.signature_electronique || '',
        data.signature_nom || '',
        data.date_signature || '',
        data.signature_timestamp || '',

        '',
        '',
        '',
        '',
        ''

      ];

    });


    const startRow = sheet.getLastRow() + 1;

    sheet
      .getRange(
        startRow,
        1,
        rows.length,
        HEADERS.length
      )
      .setValues(rows);


    // ========================================================
    // GÉNÉRATION DES DOCUMENTS
    // ========================================================

    const documents = generateDocuments(
      data,
      now
    );


    // ========================================================
    // LIENS DANS GOOGLE SHEETS
    // ========================================================

    const links = data.enfants.map(function() {

      return [

        documents.googleDocUrl,
        documents.dossierPdfUrl,
        documents.dossierDocxUrl,
        documents.reglementPdfUrl,
        documents.reglementDocxUrl

      ];

    });


    sheet
      .getRange(
        startRow,
        HEADERS.length - 4,
        links.length,
        5
      )
      .setValues(links);


    // ========================================================
    // EMAIL
    // ========================================================

    sendEmailToParentsAndAdmin(
      data,
      documents
    );


    // ========================================================
    // RÉPONSE AU FORMULAIRE
    // ========================================================

    return jsonResponse({

      ok: true,

      rows: rows.length,

      downloads: {

        dossierPdf:
          documents.dossierPdfUrl,

        dossierWord:
          documents.dossierDocxUrl,

        reglementPdf:
          documents.reglementPdfUrl,

        reglementWord:
          documents.reglementDocxUrl,

        googleDoc:
          documents.googleDocUrl

      }

    });


  } catch (err) {

    console.error(err);

    return jsonResponse({

      ok: false,

      error:
        err && err.message
          ? err.message
          : String(err)

    });

  }

}


// ============================================================
// DO GET
// ============================================================

function doGet() {

  return jsonResponse({

    ok: true,

    service:
      'Baytoul Ilm — inscriptions + règlement signé',

    version:
      REGLEMENT_VERSION

  });

}


// ============================================================
// GOOGLE SHEETS
// ============================================================

function getSheet() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  return (
    ss.getSheetByName(SHEET_NAME) ||
    ss.insertSheet(SHEET_NAME)
  );

}


function ensureHeaders(sheet) {

  if (sheet.getLastRow() === 0) {

    sheet
      .getRange(
        1,
        1,
        1,
        HEADERS.length
      )
      .setValues([HEADERS]);

    sheet.setFrozenRows(1);

  }

  else if (
    sheet.getLastColumn() < HEADERS.length
  ) {

    const current =
      sheet.getLastColumn();

    sheet
      .getRange(
        1,
        current + 1,
        1,
        HEADERS.length - current
      )
      .setValues([
        HEADERS.slice(current)
      ]);

  }

}


// ============================================================
// PARCOURS
// ============================================================

function normalizeNiveau(value) {

  const v =
    String(value || '').trim();

  if (!v) return '';

  // ITQAN EST SUPPRIMÉ

  if (
    v.toLowerCase()
      .indexOf('itqan') !== -1
  ) {

    return '';

  }


  const found =
    NIVEAUX_AUTORISES.find(
      function(niveau) {

        return (
          niveau.toLowerCase() ===
          v.toLowerCase()
        );

      }
    );


  return found || '';

}


// ============================================================
// GÉNÉRATION DES DOCUMENTS
// ============================================================

function generateDocuments(
  data,
  now
) {

  const folder =
    DriveApp.getFolderById(
      OUTPUT_FOLDER_ID
    );

  const template =
    DriveApp.getFileById(
      TEMPLATE_DOC_ID
    );

  const reglement =
    DriveApp.getFileById(
      REGLEMENT_DOC_ID
    );


  const enfantsNoms =
    data.enfants
      .map(function(enfant) {

        return fullName(
          enfant.prenom,
          enfant.nom
        );

      })
      .filter(Boolean);


  const familyName =
    sanitizeFileName(

      enfantsNoms[0] ||

      fullName(
        data.parent_prenom,
        data.parent_nom
      ) ||

      'famille'

    );


  const base =
    'Baytoul Ilm - Dossier - ' +
    familyName;


  // ========================================================
  // DOSSIER D'INSCRIPTION
  // ========================================================

  const copy =
    template.makeCopy(
      base + ' - Google Docs',
      folder
    );


  const doc =
    DocumentApp.openById(
      copy.getId()
    );


  const body =
    doc.getBody();


  styleDocument(body);


  // --------------------------------------------------------
  // PARENT 1
  // --------------------------------------------------------

  replace(
    body,
    '{{parent_type}}',
    data.parent_type
  );

  replace(
    body,
    '{{parent_nom}}',
    data.parent_nom
  );

  replace(
    body,
    '{{parent_prenom}}',
    data.parent_prenom
  );

  replace(
    body,
    '{{adresse}}',
    data.adresse
  );

  replace(
    body,
    '{{code_postal}}',
    data.code_postal
  );

  replace(
    body,
    '{{ville}}',
    data.ville
  );

  replace(
    body,
    '{{telephone}}',
    data.telephone
  );

  replace(
    body,
    '{{email}}',
    data.email
  );


  // --------------------------------------------------------
  // PARENT 2
  // --------------------------------------------------------

  replace(
    body,
    '{{parent2_type}}',
    data.parent2_type
  );

  replace(
    body,
    '{{parent2_nom}}',
    data.parent2_nom
  );

  replace(
    body,
    '{{parent2_prenom}}',
    data.parent2_prenom
  );

  replace(
    body,
    '{{parent2_adresse}}',
    data.parent2_adresse
  );

  replace(
    body,
    '{{parent2_code_postal}}',
    data.parent2_code_postal
  );

  replace(
    body,
    '{{parent2_ville}}',
    data.parent2_ville
  );

  replace(
    body,
    '{{parent2_telephone}}',
    data.parent2_telephone
  );

  replace(
    body,
    '{{parent2_email}}',
    data.parent2_email
  );


  // --------------------------------------------------------
  // ENFANTS
  // --------------------------------------------------------

  for (
    let i = 1;
    i <= 5;
    i++
  ) {

    const enfant =
      data.enfants[i - 1] || {};


    replace(
      body,
      '{{nom_' + i + '}}',
      enfant.nom
    );

    replace(
      body,
      '{{prenom_' + i + '}}',
      enfant.prenom
    );

    replace(
      body,
      '{{date_naissance_' + i + '}}',
      enfant.date_naissance
    );

    replace(
      body,
      '{{niveau_demande_' + i + '}}',
      normalizeNiveau(
        enfant.niveau_demande
      )
    );

    replace(
      body,
      '{{departement_' + i + '}}',
      enfant.departement
    );

  }


  // --------------------------------------------------------
  // AUTORISATIONS
  // --------------------------------------------------------

  replace(
    body,
    '{{remarque_médicale}}',
    data['remarque_médicale'] ||
    data.remarque_medicale
  );

  replace(
    body,
    '{{autorisation_sortie}}',
    checkValue(
      data.autorisation_sortie
    )
  );

  replace(
    body,
    '{{autorisation_image}}',
    checkValue(
      data.autorisation_image
    )
  );

  replace(
    body,
    '{{accept_privacy}}',
    checkValue(
      data.accept_privacy
    )
  );


  // --------------------------------------------------------
  // SIGNATURE
  // --------------------------------------------------------

  replace(
    body,
    '{{signature_electronique}}',
    data.signature_electronique
  );

  replace(
    body,
    '{{signature_nom}}',
    data.signature_nom
  );

  replace(
    body,
    '{{date_signature}}',
    data.date_signature ||
    formatDate(now)
  );

  replace(
    body,
    '{{signature_timestamp}}',
    data.signature_timestamp ||
    formatDateTime(now)
  );


  // ========================================================
  // AJOUT DU RÈGLEMENT
  // ========================================================

  body.appendPageBreak();

  appendStyledTitle(
    body,
    'RÈGLEMENT INTÉRIEUR'
  );

  appendCentered(
    body,
    'Institut Baytoul Ilm — Année ' +
    REGLEMENT_VERSION,
    11
  );


  appendInfoLine(
    body,
    'Responsable légal 1',
    fullName(
      data.parent_prenom,
      data.parent_nom
    )
  );


  appendInfoLine(
    body,
    'Responsable légal 2',
    fullName(
      data.parent2_prenom,
      data.parent2_nom
    )
  );


  appendInfoLine(
    body,
    'Enfant(s)',
    enfantsNoms.join(', ')
  );


  appendInfoLine(
    body,
    'Date de signature',
    data.date_signature ||
    formatDate(now)
  );


  appendReglementFromGoogleDoc(
    body,
    reglement
  );


  // ========================================================
  // RÈGLEMENT SIGNÉ
  // ========================================================

  body.appendParagraph('');

  appendStyledTitle(
    body,
    'RÈGLEMENT SIGNÉ'
  );


  appendInfoLine(
    body,
    'Signataire',
    data.signature_nom ||
    fullName(
      data.parent_prenom,
      data.parent_nom
    )
  );


  appendInfoLine(
    body,
    'Date',
    data.date_signature ||
    formatDate(now)
  );


  appendInfoLine(
    body,
    'Horodatage',
    data.signature_timestamp ||
    formatDateTime(now)
  );


  addSignatureImage(
    body,
    data.signature_image
  );


  doc.saveAndClose();


  // ========================================================
  // PDF + WORD DU DOSSIER
  // ========================================================

  const dossierPdf =
    folder.createFile(

      exportGoogleDoc(
        copy.getId(),
        'pdf'
      )
      .setName(
        base + '.pdf'
      )

    );


  const dossierDocx =
    folder.createFile(

      exportGoogleDoc(
        copy.getId(),
        'docx'
      )
      .setName(
        base + '.docx'
      )

    );


  // ========================================================
  // RÈGLEMENT SIGNÉ SEUL
  // ========================================================

  const regDoc =
    DocumentApp.create(
      base +
      ' - Reglement signe'
    );


  const rb =
    regDoc.getBody();


  styleDocument(rb);


  appendStyledTitle(
    rb,
    'RÈGLEMENT INTÉRIEUR — BAYTOUL ILM'
  );


  appendCentered(
    rb,
    'Année scolaire ' +
    REGLEMENT_VERSION,
    11
  );


  appendInfoLine(
    rb,
    'Responsable légal 1',
    fullName(
      data.parent_prenom,
      data.parent_nom
    )
  );


  appendInfoLine(
    rb,
    'Responsable légal 2',
    fullName(
      data.parent2_prenom,
      data.parent2_nom
    )
  );


  appendInfoLine(
    rb,
    'Enfant(s)',
    enfantsNoms.join(', ')
  );


  appendInfoLine(
    rb,
    'Date de signature',
    data.date_signature ||
    formatDate(now)
  );


  rb.appendParagraph('');


  appendReglementFromGoogleDoc(
    rb,
    reglement
  );


  rb.appendParagraph('');


  appendStyledTitle(
    rb,
    'SIGNATURE DU RESPONSABLE LÉGAL'
  );


  appendInfoLine(
    rb,
    'Nom du signataire',
    data.signature_nom ||
    fullName(
      data.parent_prenom,
      data.parent_nom
    )
  );


  appendInfoLine(
    rb,
    'Date',
    data.date_signature ||
    formatDate(now)
  );


  appendInfoLine(
    rb,
    'Horodatage',
    data.signature_timestamp ||
    formatDateTime(now)
  );


  addSignatureImage(
    rb,
    data.signature_image
  );


  regDoc.saveAndClose();


  const regPdf =
    folder.createFile(

      exportGoogleDoc(
        regDoc.getId(),
        'pdf'
      )
      .setName(
        base +
        ' - Reglement signe.pdf'
      )

    );


  const regDocx =
    folder.createFile(

      exportGoogleDoc(
        regDoc.getId(),
        'docx'
      )
      .setName(
        base +
        ' - Reglement signe.docx'
      )

    );


  DriveApp
    .getFileById(
      regDoc.getId()
    )
    .setTrashed(true);


  return {

    googleDocUrl:
      copy.getUrl(),

    dossierPdfUrl:
      dossierPdf.getUrl(),

    dossierDocxUrl:
      dossierDocx.getUrl(),

    reglementPdfUrl:
      regPdf.getUrl(),

    reglementDocxUrl:
      regDocx.getUrl(),

    dossierPdfFile:
      dossierPdf,

    dossierDocxFile:
      dossierDocx,

    reglementPdfFile:
      regPdf,

    reglementDocxFile:
      regDocx

  };

}


// ============================================================
// IMPORT DU RÈGLEMENT
// ============================================================

function appendReglementFromGoogleDoc(
  destinationBody,
  reglementFile
) {

  const source =
    DocumentApp.openById(
      reglementFile.getId()
    );


  const sourceBody =
    source.getBody();


  for (
    let i = 0;
    i < sourceBody.getNumChildren();
    i++
  ) {

    const element =
      sourceBody.getChild(i);


    const type =
      element.getType();


    if (
      type ===
      DocumentApp.ElementType.PARAGRAPH
    ) {

      const paragraph =
        element.asParagraph();


      const text =
        paragraph.getText();


      if (!text) continue;


      const destination =
        destinationBody
          .appendParagraph(text);


      const heading =
        paragraph.getHeading();


      if (
        heading &&
        heading !==
        DocumentApp.ParagraphHeading.NORMAL
      ) {

        destination.setHeading(
          heading
        );


        const txt =
          destination.editAsText();


        txt.setForegroundColor(
          0,
          text.length - 1,
          TURQUOISE
        );

      }

    }

    else if (
      type ===
      DocumentApp.ElementType.LIST_ITEM
    ) {

      destinationBody
        .appendListItem(
          element
            .asListItem()
            .getText()
        );

    }

  }

}


// ============================================================
// STYLE
// ============================================================

function styleDocument(body) {

  const text =
    body.editAsText();


  if (
    text.getText().length > 0
  ) {

    text.setFontFamily(
      'Arial'
    );

    text.setFontSize(
      10
    );

    text.setForegroundColor(
      '#23373C'
    );

  }

}


function appendStyledTitle(
  body,
  title
) {

  const p =
    body.appendParagraph(
      title
    );


  p.setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );


  p.setSpacingBefore(
    10
  );


  p.setSpacingAfter(
    8
  );


  const text =
    p.editAsText();


  text.setFontFamily(
    'Arial'
  );

  text.setFontSize(
    20
  );

  text.setBold(
    true
  );

  text.setForegroundColor(
    TURQUOISE
  );


  return p;

}


function appendCentered(
  body,
  text,
  size
) {

  const p =
    body.appendParagraph(
      text || ''
    );


  p.setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );


  const t =
    p.editAsText();


  t.setFontFamily(
    'Arial'
  );

  t.setFontSize(
    size || 10
  );

  t.setForegroundColor(
    '#23373C'
  );


  return p;

}


function appendInfoLine(
  body,
  label,
  value
) {

  const p =
    body.appendParagraph(
      ''
    );


  const r1 =
    p.appendText(
      (label || '') +
      ' : '
    );


  r1.setBold(
    true
  );


  r1.setForegroundColor(
    TURQUOISE
  );


  const r2 =
    p.appendText(
      value || ''
    );


  r2.setForegroundColor(
    '#23373C'
  );


  return p;

}


// ============================================================
// SIGNATURE IMAGE
// ============================================================

function addSignatureImage(
  body,
  signatureImage
) {

  if (!signatureImage) return;


  try {

    const base64 =
      String(signatureImage)
        .split(',')[1] ||
      signatureImage;


    const imageBlob =
      Utilities.newBlob(
        Utilities.base64Decode(
          base64
        ),
        'image/png',
        'signature.png'
      );


    body.appendParagraph('');


    appendCentered(
      body,
      'Signature électronique',
      10
    );


    body
      .appendImage(
        imageBlob
      )
      .setWidth(200)
      .setHeight(80);


  } catch (error) {

    console.warn(
      'Signature image non ajoutée : ' +
      error
    );

  }

}


// ============================================================
// EXPORT PDF / WORD
// ============================================================

function exportGoogleDoc(
  docId,
  format
) {

  const url =
    'https://docs.google.com/document/d/' +
    docId +
    '/export?format=' +
    encodeURIComponent(format);


  const response =
    UrlFetchApp.fetch(
      url,
      {

        headers: {

          Authorization:
            'Bearer ' +
            ScriptApp.getOAuthToken()

        },

        muteHttpExceptions:
          true

      }
    );


  if (
    response.getResponseCode() !== 200
  ) {

    throw new Error(
      'Export ' +
      format +
      ' impossible. HTTP ' +
      response.getResponseCode()
    );

  }


  return response.getBlob();

}


// ============================================================
// EMAIL
// ============================================================

function sendEmailToParentsAndAdmin(
  data,
  documents
) {

  const recipients = [];


  if (
    ADMIN_EMAIL &&
    ADMIN_EMAIL.indexOf('@') > 0
  ) {

    recipients.push(
      ADMIN_EMAIL
    );

  }


  if (
    data.email &&
    String(data.email)
      .indexOf('@') > 0
  ) {

    recipients.push(
      String(data.email)
    );

  }


  if (!recipients.length) return;


  const enfants =
    data.enfants
      .map(function(enfant) {

        return fullName(
          enfant.prenom,
          enfant.nom
        );

      })
      .filter(Boolean)
      .join(', ');


  const subject =
    'Baytoul Ilm — Dossier et règlement signé — ' +
    enfants;


  const message =
    'Bonjour,\n\n' +

    'Votre dossier Baytoul Ilm pour ' +
    enfants +
    ' est prêt.\n\n' +

    'Vous trouverez en pièces jointes :\n' +

    '• le dossier complet en PDF ;\n' +
    '• le dossier complet en Word ;\n' +
    '• le règlement intérieur signé en PDF ;\n' +
    '• le règlement intérieur signé en Word.\n\n' +

    'Les documents sont également conservés dans Google Drive.\n\n' +

    'Cordialement,\n' +
    'Baytoul Ilm';


  MailApp.sendEmail({

    to:
      recipients.join(','),

    subject:
      subject,

    body:
      message,

    attachments: [

      documents
        .dossierPdfFile
        .getBlob(),

      documents
        .dossierDocxFile
        .getBlob(),

      documents
        .reglementPdfFile
        .getBlob(),

      documents
        .reglementDocxFile
        .getBlob()

    ],

    name:
      'Baytoul Ilm'

  });

}


// ============================================================
// OUTILS
// ============================================================

function replace(
  body,
  placeholder,
  value
) {

  body.replaceText(

    escapeRegExp(
      placeholder
    ),

    value == null
      ? ''
      : String(value)

  );

}


function escapeRegExp(
  text
) {

  return String(text)
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );

}


function checkValue(value) {

  if (
    value === true ||
    value === 'true' ||
    ['oui','Oui','OUI','1']
      .indexOf(
        String(value)
      ) >= 0
  ) {

    return '☑';

  }


  if (
    value === false ||
    value === 'false' ||
    ['non','Non','NON','0','']
      .indexOf(
        String(value)
      ) >= 0
  ) {

    return '☐';

  }


  return String(
    value || ''
  );

}


function fullName(
  prenom,
  nom
) {

  return [

    prenom || '',
    nom || ''

  ]
    .filter(Boolean)
    .join(' ')
    .trim();

}


function formatDate(date) {

  return Utilities.formatDate(
    date,
    'Europe/Paris',
    'dd/MM/yyyy'
  );

}


function formatDateTime(date) {

  return Utilities.formatDate(
    date,
    'Europe/Paris',
    'dd/MM/yyyy HH:mm:ss'
  );

}


function sanitizeFileName(
  name
) {

  return String(
    name || 'famille'
  )
    .replace(
      /[\\/:*?"<>|#%{}~&]/g,
      '-'
    )
    .substring(
      0,
      100
    );

}


function jsonResponse(
  payload
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(
        payload
      )
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
