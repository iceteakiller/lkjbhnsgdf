const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, ImageRun, AlignmentType, HeightRule, VerticalAlign,
  LevelFormat, TableLayoutType,
} = require("docx");

const OUT = process.argv[2] || ".";
const photo = fs.readFileSync(path.join(__dirname, "foto-platzhalter.png"));

const VARIANTS = [
  { key: "Navy", main: "1F3A5F", soft: "D6DFEA", rule: "7088A8" },
  { key: "Salbeigruen", main: "4E6E55", soft: "DCE7DE", rule: "8DA894" },
  { key: "Anthrazit", main: "3A3F44", soft: "DDE0E3", rule: "8A9097" },
];

const FONT = "Calibri";
const PAGE_W = 11906, PAGE_H = 16838;
const LEFT_W = 3970, RIGHT_W = PAGE_W - LEFT_W;
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };

function build(v) {
  const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || "2B2F33", characterSpacing: o.spacing });
  const p = (children, o = {}) => new Paragraph({ children, alignment: o.align, spacing: o.noLine ? { before: 0, after: o.after || 0 } : { before: o.before || 0, after: o.after == null ? 60 : o.after, line: o.line || 264, lineRule: "auto" }, border: o.border, numbering: o.numbering, keepNext: o.keepNext });

  // Sidebar (white text on colour)
  const sideHead = (t) => p([run(t, { size: 20, bold: true, color: "FFFFFF", spacing: 40 })], {
    before: 380, after: 140, keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: v.rule, space: 4 } },
  });
  const sideLine = (t, o = {}) => p([run(t, { size: 19, color: o.strong ? "FFFFFF" : v.soft, bold: o.strong })], { after: o.after == null ? 50 : o.after });

  const left = [
    p([new ImageRun({ type: "png", data: photo, transformation: { width: 132, height: 170 }, altText: { title: "Foto", description: "Bewerbungsfoto", name: "Foto" } })], { align: AlignmentType.CENTER, after: 0, noLine: true }),
    sideHead("KONTAKT"),
    sideLine("Musterstraße 12"),
    sideLine("12345 Musterstadt", { after: 120 }),
    sideLine("0151 12345678", { after: 120 }),
    sideLine("max.mustermann@email.de", { after: 120 }),
    sideLine("Geboren am 01.01.2000"),
    sideLine("in Musterstadt"),
    sideHead("KENNTNISSE"),
    sideLine("MS Office", { strong: true, after: 0 }),
    sideLine("Word, Excel, Outlook", { after: 120 }),
    sideLine("SAP", { strong: true, after: 0 }),
    sideLine("Grundkenntnisse", { after: 120 }),
    sideLine("Führerschein", { strong: true, after: 0 }),
    sideLine("Klasse B"),
    sideHead("SPRACHEN"),
    sideLine("Deutsch", { strong: true, after: 0 }),
    sideLine("Muttersprache", { after: 120 }),
    sideLine("Englisch", { strong: true, after: 0 }),
    sideLine("Fließend (C1)", { after: 120 }),
    sideLine("Spanisch", { strong: true, after: 0 }),
    sideLine("Grundkenntnisse (A2)"),
    sideHead("INTERESSEN"),
    sideLine("Fußball im Verein"),
    sideLine("Fotografie"),
    sideLine("Kochen"),
  ];

  // Main column
  const head = (t) => p([run(t, { size: 23, bold: true, color: v.main, spacing: 40 })], {
    before: 420, after: 180, keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: v.main, space: 4 } },
  });
  const entry = (title, date, place, bullets = []) => [
    p([run(date, { size: 18, bold: true, color: v.main, spacing: 10 })], { after: 10, keepNext: true }),
    p([run(title, { size: 22, bold: true, color: "1E2226" })], { after: 10, keepNext: true }),
    p([run(place, { size: 20, italics: true, color: "5A6067" })], { after: bullets.length ? 70 : 240, keepNext: bullets.length > 0 }),
    ...bullets.map((b, i) => p([run(b, { size: 20 })], { numbering: { reference: "bullets", level: 0 }, after: i === bullets.length - 1 ? 240 : 40 })),
  ];

  const right = [
    p([run("MAX MUSTERMANN", { size: 54, bold: true, color: v.main, spacing: 30 })], { after: 40, line: 240 }),
    p([run("Kaufmann für Büromanagement", { size: 25, color: "5A6067" })], { after: 0 }),
    head("PROFIL"),
    p([run("Organisierter Kaufmann mit drei Jahren Erfahrung in der Auftragsabwicklung. Ich arbeite gern genau und im Team und suche eine neue Aufgabe im Vertriebsinnendienst.", { size: 20 })], { after: 0, line: 288 }),
    head("BERUFSERFAHRUNG"),
    ...entry("Sachbearbeiter Auftragsabwicklung", "06/2023 – heute", "Muster GmbH, Musterstadt", [
      "Bearbeitung von rund 40 Kundenaufträgen pro Tag",
      "Abstimmung mit Lager und Versand",
      "Pflege der Kundendaten in SAP",
    ]),
    ...entry("Aushilfe im Kundenservice", "09/2021 – 05/2023", "Beispiel AG, Musterstadt", [
      "Beantwortung von Kundenanfragen per Telefon und E-Mail",
    ]),
    head("AUSBILDUNG"),
    ...entry("Ausbildung zum Kaufmann für Büromanagement", "08/2020 – 06/2023", "Beispiel AG, Musterstadt · Abschlussnote 2,0"),
    head("WEITERBILDUNG"),
    ...entry("Excel für Fortgeschrittene", "03/2024", "IHK Musterstadt · 2 Tage"),
    head("SCHULBILDUNG"),
    ...entry("Mittlere Reife", "08/2014 – 07/2020", "Realschule Musterstadt · Abschlussnote 2,3"),
    p([run("Musterstadt, 18.11.2026", { color: "5A6067" })], { before: 400, after: 420 }),
    p([run("Max Mustermann", { color: "5A6067" })], { after: 0, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "9AA0A6", space: 4 } } }),
  ];

  const table = new Table({
    width: { size: PAGE_W, type: WidthType.DXA },
    columnWidths: [LEFT_W, RIGHT_W],
    layout: TableLayoutType.FIXED,
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    rows: [new TableRow({
      height: { value: PAGE_H - 160, rule: HeightRule.EXACT },
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: LEFT_W, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: v.main },
          margins: { top: 760, bottom: 400, left: 520, right: 420 },
          verticalAlign: VerticalAlign.TOP, borders: noBorders, children: left,
        }),
        new TableCell({
          width: { size: RIGHT_W, type: WidthType.DXA },
          margins: { top: 760, bottom: 400, left: 620, right: 700 },
          verticalAlign: VerticalAlign.TOP, borders: noBorders, children: right,
        }),
      ],
    })],
  });

  const doc = new Document({
    creator: "EasyTemplateShopDE",
    title: "Lebenslauf Modern Klar",
    styles: { default: { document: { run: { font: FONT, size: 19 } } } },
    numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 300, hanging: 220 } }, run: { color: v.main } } }] }] },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } } },
      children: [table, new Paragraph({ children: [new TextRun({ text: "", size: 2 })], spacing: { before: 0, after: 0, line: 20, lineRule: "exact" } })],
    }],
  });
  return Packer.toBuffer(doc).then((buf) => {
    const f = path.join(OUT, `Lebenslauf_Modern_Klar_${v.key}.docx`);
    fs.writeFileSync(f, buf);
    console.log("wrote", f);
  });
}

Promise.all(VARIANTS.map(build)).catch((e) => { console.error(e); process.exit(1); });
