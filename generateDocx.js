const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require('docx');

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const BLUE       = "1F3864";
const ACCENT     = "2E75B6";
const LIGHT_BLUE = "D6E4F7";
const LIGHT_GREY = "F5F5F5";
const ORANGE     = "C55A11";
const YELLOW_FILL= "FFF2CC";
const GREEN_FILL = "E2EFDA";

// ─── BORDURES ─────────────────────────────────────────────────────────────────
const border    = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders   = { top: border, bottom: border, left: border, right: border };
const noBorder  = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const para = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.before ?? 80, after: opts.after ?? 80 },
  children: [new TextRun({
    text: text || '',
    bold:    opts.bold    || false,
    italics: opts.italic  || false,
    size:    opts.size    || 20,
    color:   opts.color   || "000000",
    font:    "Arial",
  })]
});

const empty = (before = 60, after = 60) => new Paragraph({
  spacing: { before, after },
  children: [new TextRun({ text: "", font: "Arial", size: 20 })]
});

// ─── COMPOSANTS ──────────────────────────────────────────────────────────────

const documentHeader = (etablissement, adresse, typeDoc, sousTitre) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: etablissement || '', bold: true, size: 20, color: BLUE, font: "Arial" })]
  }),
  ...(adresse ? [new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: adresse, bold: false, size: 18, color: "888888", font: "Arial" })]
  })] : [empty(0, 160)]),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 4 } },
    children: [new TextRun({ text: typeDoc, bold: true, size: 36, color: BLUE, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text: sousTitre, bold: true, size: 24, color: ACCENT, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 280 },
    children: [new TextRun({ text: "Document confidentiel — usage interne", italics: true, size: 16, color: "888888", font: "Arial" })]
  }),
];

const infoRow = (label, value) => new TableRow({
  children: [
    new TableCell({
      borders, width: { size: 2600, type: WidthType.DXA },
      shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, font: "Arial" })] })]
    }),
    new TableCell({
      borders, width: { size: 6760, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text: value || '—', size: 18, font: "Arial" })] })]
    }),
  ]
});

const sectionTitle = (num, titre) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [600, 8760],
  rows: [new TableRow({
    children: [
      new TableCell({
        borders, shading: { fill: BLUE, type: ShadingType.CLEAR },
        width: { size: 600, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: num, bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })]
      }),
      new TableCell({
        borders, shading: { fill: BLUE, type: ShadingType.CLEAR },
        width: { size: 8760, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 160, right: 140 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          children: [new TextRun({ text: titre, bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })]
      }),
    ]
  })]
});

const decisionTableHeader = (cols) => new TableRow({
  children: cols.map((label, i) =>
    new TableCell({
      borders, shading: { fill: BLUE, type: ShadingType.CLEAR },
      width: { size: [500, 5260, 2000, 1600][i], type: WidthType.DXA },
      margins: { top: 70, bottom: 70, left: i === 0 ? 100 : 120, right: 120 },
      children: [new Paragraph({ alignment: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
        children: [new TextRun({ text: label, bold: true, size: 17, color: "FFFFFF", font: "Arial" })] })]
    })
  )
});

const decisionRow = (ref, texte, responsable, echeance, fill) => new TableRow({
  children: [
    new TableCell({
      borders, width: { size: 500, type: WidthType.DXA },
      shading: { fill: fill || YELLOW_FILL, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 100, right: 100 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: ref, bold: true, size: 17, color: ORANGE, font: "Arial" })] })]
    }),
    new TableCell({
      borders, width: { size: 5260, type: WidthType.DXA },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: texte || '', size: 17, font: "Arial" })] })]
    }),
    new TableCell({
      borders, width: { size: 2000, type: WidthType.DXA },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: responsable || '', size: 17, font: "Arial" })] })]
    }),
    new TableCell({
      borders, width: { size: 1600, type: WidthType.DXA },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: echeance || '', italics: true, size: 16, color: "555555", font: "Arial" })] })]
    }),
  ]
});

const signaturesBlock = (nomPresident, nomSecretaire) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  rows: [new TableRow({
    children: [
      new TableCell({
        borders: noBorders, width: { size: 4680, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        children: [
          para("Le Président de séance", { bold: true, size: 18, before: 0, after: 20 }),
          para(nomPresident || '', { size: 18, before: 0, after: 200 }),
          para("Signature :", { size: 18, before: 0, after: 0 }),
        ]
      }),
      new TableCell({
        borders: noBorders, width: { size: 4680, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        children: [
          para("Le Secrétaire de séance", { bold: true, size: 18, before: 0, after: 20 }),
          para(nomSecretaire || '', { size: 18, before: 0, after: 200 }),
          para("Signature :", { size: 18, before: 0, after: 0 }),
        ]
      }),
    ]
  })]
});

const documentFooter = (adresse) => [
  empty(200, 0),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
    spacing: { before: 160, after: 0 },
    children: [new TextRun({ text: adresse || '', size: 16, color: "888888", font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: "Document confidentiel — usage interne — Ne pas diffuser", italics: true, size: 16, color: "AAAAAA", font: "Arial" })]
  }),
];

// ─── PARSEUR DE CR TEXTE ─────────────────────────────────────────────────────
// Parse le texte brut du CR généré par Claude et construit le DOCX

function parseAndBuildDocx(crTexte, meta) {
  const {
    etablissement, adresse, titre, type, date, lieu,
    heureDebut, heureFin, participants, president, secretaire
  } = meta;

  const sousTitre = `${titre} — ${date}`;
  const horaires = heureDebut || heureFin
    ? `${heureDebut || ''}${heureDebut && heureFin ? ' - ' : ''}${heureFin || ''}`
    : '—';

  // Tableau d'informations générales
  const infoTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 6760],
    rows: [
      infoRow("Date", date),
      infoRow("Horaires", horaires),
      infoRow("Lieu", lieu || '—'),
      infoRow("Type de réunion", type),
      infoRow("Président de séance", president || '—'),
      infoRow("Secrétaire de séance", secretaire || '—'),
      infoRow("Participants", participants || '—'),
    ]
  });

  // Parser les sections du CR texte
  const lignes = crTexte.split('\n');
  const children = [];

  // En-tête
  children.push(...documentHeader(etablissement, adresse, "COMPTE-RENDU DE RÉUNION", sousTitre));
  children.push(empty(40, 40));
  children.push(infoTable);
  children.push(empty(100, 60));

  // Analyser le texte pour extraire les sections
  let inDecisions = false;
  let inActions = false;
  let decisionRows = [];
  let actionRows = [];
  let currentSection = null;
  let sectionContent = [];
  let sectionNum = 0;

  const flushSection = () => {
    if (!currentSection) return;
    sectionNum++;
    children.push(sectionTitle(String(sectionNum), currentSection));
    children.push(empty(40, 20));
    sectionContent.forEach(l => {
      if (l.trim()) {
        children.push(para(l.replace(/^[-•*]\s*/, '').trim(), { size: 19, before: 30, after: 30 }));
      }
    });
    children.push(empty(60, 40));
    currentSection = null;
    sectionContent = [];
  };

  for (const ligne of lignes) {
    const l = ligne.trim();
    if (!l) continue;

    // Détecter sections principales
    if (/^#{1,2}\s+OBJET|^OBJET ET CONTEXTE/i.test(l)) {
      flushSection();
      inDecisions = false; inActions = false;
      currentSection = "OBJET ET CONTEXTE DE LA RÉUNION";
      continue;
    }
    if (/^#{1,2}\s+POINTS\s+ABORDÉS|^POINTS\s+ABORDÉS/i.test(l)) {
      flushSection();
      inDecisions = false; inActions = false;
      currentSection = "POINTS ABORDÉS";
      continue;
    }
    if (/^#{1,2}\s+DÉCISIONS|^DÉCISIONS/i.test(l)) {
      flushSection();
      inDecisions = true; inActions = false;
      sectionNum++;
      children.push(sectionTitle(String(sectionNum), "DÉCISIONS"));
      children.push(empty(40, 20));
      continue;
    }
    if (/^#{1,2}\s+PLAN D.ACTIONS|^PLAN D.ACTIONS/i.test(l)) {
      // Flush décisions table
      if (decisionRows.length > 0) {
        children.push(new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [500, 5260, 2000, 1600],
          rows: [decisionTableHeader(["Réf.", "Décision", "Responsable", "Échéance"]), ...decisionRows]
        }));
        decisionRows = [];
      }
      inDecisions = false; inActions = true;
      children.push(empty(60, 40));
      sectionNum++;
      children.push(sectionTitle(String(sectionNum), "PLAN D'ACTIONS"));
      children.push(empty(40, 20));
      continue;
    }

    // Lignes de tableau (décisions/actions)
    if ((inDecisions || inActions) && l.startsWith('|')) {
      const cells = l.split('|').map(c => c.trim()).filter(Boolean);
      if (cells.length >= 2 && !cells[0].match(/^[-:]+$/) && !cells[0].match(/^Réf/i) && !cells[0].match(/^Ref/i)) {
        const ref = cells[0] || '';
        const texte = cells[1] || '';
        const responsable = cells[2] || '';
        const echeance = cells[3] || '';
        if (inDecisions) {
          decisionRows.push(decisionRow(ref, texte, responsable, echeance, YELLOW_FILL));
        } else {
          actionRows.push(decisionRow(ref, texte, responsable, echeance, GREEN_FILL));
        }
      }
      continue;
    }

    // Contenu de section
    if (currentSection && !inDecisions && !inActions) {
      // Sauter les lignes d'en-tête markdown
      if (l.startsWith('#')) {
        const titre = l.replace(/^#+\s*/, '');
        if (titre && !titre.match(/^POINTS ABORDÉS|^OBJET|^DÉCISIONS|^PLAN/i)) {
          sectionContent.push(titre);
        }
        continue;
      }
      // Sauter les lignes de tableau markdown
      if (l.startsWith('|')) continue;
      // Sauter les séparateurs
      if (l.match(/^---+$/)) continue;
      sectionContent.push(l);
    }
  }

  // Flush dernière section
  flushSection();

  // Flush actions table
  if (actionRows.length > 0) {
    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [500, 5260, 2000, 1600],
      rows: [decisionTableHeader(["Réf.", "Action", "Responsable", "Échéance"]), ...actionRows]
    }));
    children.push(empty(60, 40));
  }

  // Signatures
  children.push(empty(160, 80));
  children.push(signaturesBlock(president, secretaire));

  // Pied de page
  children.push(...documentFooter(adresse));

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 20, color: "000000" } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 900, right: 900, bottom: 900, left: 900 }
        }
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

module.exports = { parseAndBuildDocx };
