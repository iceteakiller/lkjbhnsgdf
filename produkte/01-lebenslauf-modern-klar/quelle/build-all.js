// Baut alle Referenz-Vorlagen (Produkt 2–10) als Word-Dateien.
// Aufruf: node build-all.js <ausgabeordner>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, ImageRun, AlignmentType, HeightRule, VerticalAlign,
  LevelFormat, TableLayoutType,
} = require("docx");

const OUT = process.argv[2] || "out";
const PHOTO = fs.readFileSync(path.join(__dirname, "foto-platzhalter.png"));
const FONT = "Calibri";
const PAGE_W = 11906, PAGE_H = 16838;
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const TABLE_NO_BORDERS = { ...NO_BORDERS, insideHorizontal: NONE, insideVertical: NONE };
const INK = "2B2F33", GREY = "5A6067", LIGHT = "6A7077";

const C = {
  navy: { main: "1F3A5F", accent: "1F3A5F", soft: "D6DFEA", rule: "7088A8" },
  petrol: { main: "1E6F78", accent: "1E6F78", soft: "D5E9EB", rule: "7FB3B9" },
  pflege: { main: "2E6A9E", accent: "2E6A9E", soft: "D9E6F3", rule: "86AED3" },
  handwerk: { main: "33373D", accent: "C2610F", soft: "DDE0E3", rule: "E07A1F" },
  schwarz: { main: "111111", accent: "111111", soft: "DDDDDD", rule: "BBBBBB" },
};

function H(v) {
  const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || INK, characterSpacing: o.cs });
  const p = (children, o = {}) => new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: o.align,
    spacing: o.noLine ? { before: o.before || 0, after: o.after || 0 } : { before: o.before || 0, after: o.after == null ? 60 : o.after, line: o.line || 264, lineRule: "auto" },
    border: o.border, numbering: o.numbering, keepNext: o.keepNext, indent: o.indent,
  });
  const t = (text, o = {}) => p([run(text, o)], o);
  const photo = (w = 132, h = 170) => new ImageRun({ type: "png", data: PHOTO, transformation: { width: w, height: h }, altText: { title: "Foto", description: "Bewerbungsfoto", name: "Foto" } });
  const rule = (color, size = 8) => ({ bottom: { style: BorderStyle.SINGLE, size, color, space: 4 } });

  // Seitenleiste (weiße Schrift auf Farbe)
  const sideHead = (x) => t(x, { size: 20, bold: true, color: "FFFFFF", cs: 40, before: 380, after: 140, keepNext: true, border: rule(v.rule, 6) });
  const sideLine = (x, o = {}) => t(x, { size: 19, color: o.strong ? "FFFFFF" : v.soft, bold: o.strong, after: o.after == null ? 50 : o.after });
  const sideKV = (k, val, last) => [sideLine(k, { strong: true, after: 0 }), sideLine(val, { after: last ? 50 : 120 })];

  // Hauptspalte
  const head = (x) => t(x, { size: 23, bold: true, color: v.accent, cs: 40, before: 420, after: 180, keepNext: true, border: rule(v.accent) });
  const bullet = (x, last) => t(x, { numbering: { reference: "punkte", level: 0 }, after: last ? 240 : 40 });
  const entry = (title, date, place, bullets = []) => [
    t(date, { size: 18, bold: true, color: v.accent, cs: 10, after: 10, keepNext: true }),
    t(title, { size: 22, bold: true, color: "1E2226", after: 10, keepNext: true }),
    t(place, { italics: true, color: GREY, after: bullets.length ? 70 : 240, keepNext: bullets.length > 0 }),
    ...bullets.map((b, i) => bullet(b, i === bullets.length - 1)),
  ];
  const nameBlock = (name, role, color = v.accent) => [
    t(name, { size: 54, bold: true, color, cs: 30, after: 40, line: 240 }),
    t(role, { size: 25, color: GREY, after: 0 }),
  ];
  const signature = (place, name) => [
    t(place, { color: GREY, before: 400, after: 420 }),
    p([run(name, { color: GREY })], { after: 0, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "9AA0A6", space: 4 } } }),
  ];
  const placeholder = (x, after = 200) => t(x, { italics: true, color: "7A8088", after, line: 288 });
  return { run, p, t, photo, rule, sideHead, sideLine, sideKV, head, bullet, entry, nameBlock, signature, placeholder };
}

const tiny = () => new Paragraph({ children: [new TextRun({ text: "", size: 2 })], spacing: { before: 0, after: 0, line: 20, lineRule: "exact" } });
const zeroMargins = { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 };
const pageProps = (margin) => ({ page: { size: { width: PAGE_W, height: PAGE_H }, margin } });

function sidebarSection(v, left, right, leftW = 3970) {
  const rightW = PAGE_W - leftW;
  return {
    properties: pageProps(zeroMargins),
    children: [new Table({
      width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [leftW, rightW], layout: TableLayoutType.FIXED, borders: TABLE_NO_BORDERS,
      rows: [new TableRow({
        height: { value: PAGE_H - 160, rule: HeightRule.EXACT }, cantSplit: true,
        children: [
          new TableCell({ width: { size: leftW, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: v.main }, margins: { top: 760, bottom: 400, left: 520, right: 420 }, verticalAlign: VerticalAlign.TOP, borders: NO_BORDERS, children: left }),
          new TableCell({ width: { size: rightW, type: WidthType.DXA }, margins: { top: 760, bottom: 400, left: 620, right: 700 }, verticalAlign: VerticalAlign.TOP, borders: NO_BORDERS, children: right }),
        ],
      })],
    }), tiny()],
  };
}

// Farbiger Kopfbalken über die ganze Breite, darunter eine Spalte Text
function bandSection(v, h, bandLeft, bandRight, body) {
  const innerL = 7200, innerR = PAGE_W - 2000 - innerL;
  const bandTable = new Table({
    width: { size: PAGE_W - 2000, type: WidthType.DXA }, columnWidths: [innerL, innerR], layout: TableLayoutType.FIXED, borders: TABLE_NO_BORDERS,
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: innerL, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: NO_BORDERS, children: bandLeft }),
      new TableCell({ width: { size: innerR, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: NO_BORDERS, children: bandRight.length ? bandRight : [h.t("")] }),
    ] })],
  });
  // Seitenränder links/rechts 1000; der Balken ragt mit negativem Einzug bis an die Blattkante
  return {
    properties: pageProps({ top: 0, bottom: 900, left: 1000, right: 1000, header: 0, footer: 0 }),
    children: [new Table({
      width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [PAGE_W], layout: TableLayoutType.FIXED, borders: TABLE_NO_BORDERS,
      indent: { size: -1000, type: WidthType.DXA },
      rows: [new TableRow({ children: [new TableCell({ width: { size: PAGE_W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: v.main }, margins: { top: 560, bottom: 420, left: 1000, right: 1000 }, borders: NO_BORDERS, children: [bandTable, h.t("", { after: 0 })] })] })],
    }), ...body],
  };
}

// Zweispaltige Tabelle: Datum links, Inhalt rechts (klassischer deutscher Lebenslauf)
function dateTable(h, v, rows) {
  const W = PAGE_W - 2000, L = 2300, R = W - L;
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [L, R], layout: TableLayoutType.FIXED, borders: TABLE_NO_BORDERS,
    rows: rows.map((r) => new TableRow({ cantSplit: true, children: [
      new TableCell({ width: { size: L, type: WidthType.DXA }, borders: NO_BORDERS, margins: { top: 0, bottom: 80, left: 0, right: 200 }, children: [h.t(r[0], { size: 19, bold: true, color: v.accent })] }),
      new TableCell({ width: { size: R, type: WidthType.DXA }, borders: NO_BORDERS, margins: { top: 0, bottom: 80, left: 0, right: 0 }, children: [
        h.t(r[1], { bold: true, size: 21, color: "1E2226", after: 10 }),
        ...(r[2] ? [h.t(r[2], { italics: true, color: GREY, after: (r[3] || []).length ? 50 : 0 })] : []),
        ...(r[3] || []).map((b, i, a) => h.bullet(b, false)),
      ] }),
    ] })),
  });
}

const std = { top: 1134, bottom: 1134, left: 1417, right: 1134, header: 0, footer: 0 };

function makeDoc(v, sections, title) {
  return new Document({
    creator: "EasyTemplateShopDE", title,
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    numbering: { config: [
      { reference: "punkte", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 300, hanging: 220 } }, run: { color: v.accent } } }] },
      { reference: "zahlen", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 360 } }, run: { color: v.accent, bold: true } } }] },
    ] },
    sections,
  });
}

// ---------- Bausteine pro Produkt ----------

const MAX = {
  name: "MAX MUSTERMANN", plain: "Max Mustermann",
  contact: ["Musterstraße 12", "12345 Musterstadt", "0151 12345678", "max.mustermann@email.de"],
  line: "Musterstraße 12 · 12345 Musterstadt · 0151 12345678 · max.mustermann@email.de",
};

function cvModernKlar(v) { // Produkt 1, für Set und zweisprachige Version
  const h = H(v);
  const left = [
    h.p([h.photo()], { align: AlignmentType.CENTER, noLine: true }),
    h.sideHead("KONTAKT"),
    h.sideLine("Musterstraße 12"), h.sideLine("12345 Musterstadt", { after: 120 }), h.sideLine("0151 12345678", { after: 120 }),
    h.sideLine("max.mustermann@email.de", { after: 120 }), h.sideLine("Geboren am 01.01.2000"), h.sideLine("in Musterstadt"),
    h.sideHead("KENNTNISSE"), ...h.sideKV("MS Office", "Word, Excel, Outlook"), ...h.sideKV("SAP", "Grundkenntnisse"), ...h.sideKV("Führerschein", "Klasse B", true),
    h.sideHead("SPRACHEN"), ...h.sideKV("Deutsch", "Muttersprache"), ...h.sideKV("Englisch", "Fließend (C1)"), ...h.sideKV("Spanisch", "Grundkenntnisse (A2)", true),
    h.sideHead("INTERESSEN"), h.sideLine("Fußball im Verein"), h.sideLine("Fotografie"), h.sideLine("Kochen"),
  ];
  const right = [
    ...h.nameBlock(MAX.name, "Kaufmann für Büromanagement"),
    h.head("PROFIL"),
    h.t("Organisierter Kaufmann mit drei Jahren Erfahrung in der Auftragsabwicklung. Ich arbeite gern genau und im Team und suche eine neue Aufgabe im Vertriebsinnendienst.", { after: 0, line: 288 }),
    h.head("BERUFSERFAHRUNG"),
    ...h.entry("Sachbearbeiter Auftragsabwicklung", "06/2023 – heute", "Muster GmbH, Musterstadt", ["Bearbeitung von rund 40 Kundenaufträgen pro Tag", "Abstimmung mit Lager und Versand", "Pflege der Kundendaten in SAP"]),
    ...h.entry("Aushilfe im Kundenservice", "09/2021 – 05/2023", "Beispiel AG, Musterstadt", ["Beantwortung von Kundenanfragen per Telefon und E-Mail"]),
    h.head("AUSBILDUNG"),
    ...h.entry("Ausbildung zum Kaufmann für Büromanagement", "08/2020 – 06/2023", "Beispiel AG, Musterstadt · Abschlussnote 2,0"),
    h.head("WEITERBILDUNG"),
    ...h.entry("Excel für Fortgeschrittene", "03/2024", "IHK Musterstadt · 2 Tage"),
    h.head("SCHULBILDUNG"),
    ...h.entry("Mittlere Reife", "08/2014 – 07/2020", "Realschule Musterstadt · Abschlussnote 2,3"),
    ...h.signature("Musterstadt, 18.11.2026", MAX.plain),
  ];
  return sidebarSection(v, left, right);
}

function deckblattA(v) {
  const h = H(v);
  return { properties: pageProps(std), children: [
    h.p([h.photo(190, 244)], { align: AlignmentType.CENTER, noLine: true, before: 600 }),
    h.t("BEWERBUNG", { align: AlignmentType.CENTER, size: 24, bold: true, color: v.accent, cs: 120, before: 700, after: 160 }),
    h.t("als Sachbearbeiter im Vertriebsinnendienst", { align: AlignmentType.CENTER, size: 40, bold: true, color: "1E2226", after: 80, line: 252 }),
    h.t("bei der Muster GmbH", { align: AlignmentType.CENTER, size: 28, color: GREY, after: 700 }),
    h.t(MAX.plain, { align: AlignmentType.CENTER, size: 30, bold: true, color: v.accent, after: 80 }),
    h.t(MAX.line, { align: AlignmentType.CENTER, size: 19, color: GREY, after: 1400 }),
    h.p([h.run("ANLAGEN   ", { size: 18, bold: true, color: v.accent, cs: 60 }), h.run("Anschreiben · Lebenslauf · Zeugnisse", { size: 19, color: GREY })], { align: AlignmentType.CENTER, border: { top: { style: BorderStyle.SINGLE, size: 6, color: v.rule, space: 10 } } }),
  ] };
}

function deckblattB(v) {
  const h = H(v);
  const left = [
    h.p([h.photo(150, 193)], { align: AlignmentType.CENTER, noLine: true }),
    h.sideHead("KONTAKT"), ...MAX.contact.map((x) => h.sideLine(x, { after: 90 })),
  ];
  const right = [
    h.t("", { before: 3200 }),
    h.t("BEWERBUNG", { size: 24, bold: true, color: v.accent, cs: 120, after: 200 }),
    h.t("als Sachbearbeiter im Vertriebsinnendienst", { size: 44, bold: true, color: "1E2226", after: 120, line: 240 }),
    h.t("bei der Muster GmbH", { size: 28, color: GREY, after: 600, border: h.rule(v.accent, 12) }),
    h.t(MAX.plain, { size: 30, bold: true, color: v.accent, before: 200, after: 1800 }),
    h.t("ANLAGEN", { size: 18, bold: true, color: v.accent, cs: 60, after: 80 }),
    h.t("Anschreiben", { color: GREY, after: 30 }), h.t("Lebenslauf", { color: GREY, after: 30 }), h.t("Zeugnisse", { color: GREY }),
  ];
  return sidebarSection(v, left, right, 3600);
}

function deckblattC(v) {
  const h = H(v);
  const W = PAGE_W - std.left - std.right, R = 3000, L = W - R;
  return { properties: pageProps(std), children: [
    h.t("", { after: 0, border: h.rule(v.accent, 18) }),
    h.t("", { before: 1600 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [L, R], layout: TableLayoutType.FIXED, borders: TABLE_NO_BORDERS, rows: [new TableRow({ children: [
      new TableCell({ width: { size: L, type: WidthType.DXA }, borders: NO_BORDERS, verticalAlign: VerticalAlign.CENTER, margins: { right: 400 }, children: [
        h.t("Bewerbung", { size: 64, bold: true, color: v.accent, after: 160, line: 240 }),
        h.t("als Sachbearbeiter im Vertriebsinnendienst bei der Muster GmbH", { size: 28, color: "1E2226", line: 288, after: 0 }),
      ] }),
      new TableCell({ width: { size: R, type: WidthType.DXA }, borders: NO_BORDERS, verticalAlign: VerticalAlign.CENTER, children: [h.p([h.photo(170, 219)], { align: AlignmentType.RIGHT, noLine: true })] }),
    ] })] }),
    h.t("", { before: 3600 }),
    h.t(MAX.plain, { size: 30, bold: true, color: "1E2226", after: 80 }),
    ...MAX.contact.map((x) => h.t(x, { color: GREY, after: 20 })),
    h.t("Anlagen: Anschreiben, Lebenslauf, Zeugnisse", { size: 19, color: GREY, before: 600, after: 0, border: { top: { style: BorderStyle.SINGLE, size: 4, color: v.rule, space: 8 } } }),
  ] };
}

function anschreiben(v) {
  const h = H(v);
  return { properties: pageProps(std), children: [
    h.t(MAX.name, { size: 36, bold: true, color: v.accent, cs: 30, after: 40 }),
    h.t(MAX.line, { size: 18, color: GREY, after: 0, border: h.rule(v.accent, 8) }),
    h.t("Max Mustermann · Musterstraße 12 · 12345 Musterstadt", { size: 14, color: LIGHT, before: 900, after: 60 }),
    ...["Muster GmbH", "Frau Erika Beispiel", "Personalabteilung", "Beispielweg 5", "12345 Musterstadt"].map((x) => h.t(x, { after: 0 })),
    h.t("Musterstadt, 18.11.2026", { align: AlignmentType.RIGHT, before: 600, after: 500 }),
    h.t("Bewerbung als Sachbearbeiter im Vertriebsinnendienst", { bold: true, size: 22, after: 360 }),
    h.t("Sehr geehrte Frau Beispiel,", { after: 240 }),
    h.placeholder("[Einstieg: Warum bewirbst du dich genau bei dieser Firma? Ein konkreter Satz wirkt stärker als ein allgemeiner.]"),
    h.placeholder("[Erfahrung: Was hast du bisher gemacht, das zur Stelle passt? Ein oder zwei Beispiele, gern mit Zahlen.]"),
    h.placeholder("[Stärken: Welche zwei oder drei Eigenschaften bringst du mit, und wo hast du sie gezeigt?]"),
    h.placeholder("[Schluss: Ab wann kannst du anfangen? Bitte um ein Vorstellungsgespräch.]", 360),
    h.t("Mit freundlichen Grüßen", { after: 900 }),
    h.t(MAX.plain, { after: 700 }),
    h.t("Anlagen: Lebenslauf, Zeugnisse", { size: 18, color: GREY }),
  ] };
}

function anlagen(v) {
  const h = H(v);
  return { properties: pageProps(std), children: [
    h.t(MAX.name, { size: 36, bold: true, color: v.accent, cs: 30, after: 40 }),
    h.t(MAX.line, { size: 18, color: GREY, after: 0, border: h.rule(v.accent, 8) }),
    h.t("ANLAGENVERZEICHNIS", { size: 28, bold: true, color: v.accent, cs: 60, before: 900, after: 400 }),
    ...["Lebenslauf", "Abschlusszeugnis Ausbildung (IHK)", "Arbeitszeugnis Muster GmbH", "Arbeitszeugnis Beispiel AG", "Zeugnis der Mittleren Reife", "Zertifikat Excel für Fortgeschrittene"]
      .map((x) => h.t(x, { size: 22, numbering: { reference: "zahlen", level: 0 }, after: 200 })),
  ] };
}

function cvSchueler(v) {
  const h = H(v);
  const bandLeft = [
    h.t("LENA BEISPIEL", { size: 52, bold: true, color: "FFFFFF", cs: 30, after: 60, line: 240 }),
    h.t("Bewerbung um einen Ausbildungsplatz als Kauffrau im Einzelhandel", { size: 23, color: v.soft, line: 276, after: 0 }),
  ];
  const bandRight = [h.p([h.photo(118, 152)], { align: AlignmentType.RIGHT, noLine: true })];
  const body = [
    h.head("PERSÖNLICHE DATEN"),
    dateTable(h, v, [["Geburtsdatum", "12.03.2010 in Musterstadt"], ["Anschrift", "Beispielweg 7, 12345 Musterstadt"], ["Telefon", "0176 11223344"], ["E-Mail", "lena.beispiel@email.de"]]),
    h.head("SCHULBILDUNG"),
    dateTable(h, v, [["seit 08/2020", "Realschule Musterstadt", "Angestrebter Abschluss: Mittlere Reife (Sommer 2027)"], ["08/2016 – 07/2020", "Grundschule am Park, Musterstadt"]]),
    h.head("PRAKTIKA"),
    dateTable(h, v, [["02/2025 · 2 Wochen", "Schülerpraktikum im Modegeschäft", "Mode Muster, Musterstadt", ["Kundenberatung und Kasse", "Einräumen und Dekorieren der Ware"]], ["06/2024 · 1 Woche", "Praktikum in der Kita Sonnenschein", "Musterstadt"]]),
    h.head("KENNTNISSE"),
    dateTable(h, v, [["PC", "Word, PowerPoint, Grundkenntnisse Excel"], ["Sprachen", "Deutsch (Muttersprache), Englisch (6 Jahre Schule), Französisch (Grundkenntnisse)"]]),
    h.head("STÄRKEN UND HOBBYS"),
    dateTable(h, v, [["Stärken", "Freundlich, zuverlässig, teamfähig"], ["Hobbys", "Handball im Verein, Zeichnen, Babysitten"]]),
    h.t("Musterstadt, 18.11.2026", { color: GREY, before: 200, after: 300 }),
    h.p([h.run("Lena Beispiel", { color: GREY })], { after: 0, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "9AA0A6", space: 4 } } }),
  ];
  return bandSection(v, h, bandLeft, bandRight, body);
}

function cvMinimal(v) {
  const h = H(v);
  const m = { top: 1134, bottom: 1134, left: 1300, right: 1300, header: 0, footer: 0 };
  const head = (x) => h.t(x, { size: 21, bold: true, color: v.accent, cs: 60, before: 380, after: 160, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: v.rule, space: 4 } } });
  const job = (title, meta, bullets = []) => [
    h.t(title, { size: 22, bold: true, color: "1E2226", after: 10, keepNext: true }),
    h.t(meta, { color: GREY, after: bullets.length ? 70 : 220, keepNext: bullets.length > 0 }),
    ...bullets.map((b, i) => h.bullet(b, i === bullets.length - 1)),
  ];
  return { properties: pageProps(m), children: [
    h.t("Alex Beispiel", { size: 56, bold: true, color: v.accent, after: 40, line: 240 }),
    h.t("Projektkoordination", { size: 26, color: GREY, after: 120 }),
    h.t("Musterstadt · 0151 12345678 · alex.beispiel@email.de · linkedin.com/in/alexbeispiel", { size: 19, color: GREY, after: 0 }),
    head("PROFIL"),
    h.t("Projektkoordination mit fünf Jahren Erfahrung in Logistik und Einkauf. Ich halte Termine, Budgets und viele Beteiligte zusammen und arbeite gern mit klaren Zahlen.", { after: 0, line: 288 }),
    head("BERUFSERFAHRUNG"),
    ...job("Projektkoordination Einkauf", "Muster Logistik GmbH, Musterstadt · 04/2022 – heute", ["Steuerung von 6 bis 8 Lieferantenprojekten gleichzeitig", "Wöchentliche Berichte an die Geschäftsleitung", "Einführung eines neuen Bestellprozesses, 15 % weniger Rückfragen"]),
    ...job("Sachbearbeitung Disposition", "Beispiel Spedition AG, Musterstadt · 09/2019 – 03/2022", ["Planung von Touren und Fahrzeugeinsatz", "Abstimmung mit Kunden und Fahrern"]),
    head("AUSBILDUNG"),
    ...job("B.A. Betriebswirtschaftslehre", "Hochschule Musterstadt · 2016 – 2019 · Note 2,1"),
    head("KENNTNISSE"),
    h.t("Projektmanagement: Jira, MS Project", { after: 40 }),
    h.t("MS Office: sehr gut, Excel: Pivot-Tabellen und SVERWEIS", { after: 40 }),
    h.t("Sprachen: Deutsch (Muttersprache), Englisch (verhandlungssicher, C1)", { after: 0 }),
  ] };
}

function cvPflege(v) {
  const h = H(v);
  const left = [
    h.p([h.photo()], { align: AlignmentType.CENTER, noLine: true }),
    h.sideHead("KONTAKT"),
    h.sideLine("Beispielweg 3"), h.sideLine("12345 Musterstadt", { after: 120 }), h.sideLine("0160 98765432", { after: 120 }), h.sideLine("sarah.beispiel@email.de", { after: 120 }), h.sideLine("Geboren am 15.05.1998"),
    h.sideHead("KENNTNISSE"), ...h.sideKV("Dokumentation", "Elektronische Pflegeakte"), ...h.sideKV("Expertenstandards", "Dekubitus, Sturz, Schmerz"), ...h.sideKV("Führerschein", "Klasse B", true),
    h.sideHead("SPRACHEN"), ...h.sideKV("Deutsch", "Muttersprache"), ...h.sideKV("Englisch", "Gut (B2)", true),
    h.sideHead("INTERESSEN"), h.sideLine("Laufen"), h.sideLine("Lesen"),
  ];
  const right = [
    ...h.nameBlock("SARAH BEISPIEL", "Pflegefachfrau"),
    h.head("PROFIL"),
    h.t("Examinierte Pflegefachfrau mit vier Jahren Erfahrung auf der Inneren Station. Ich bleibe auch in hektischen Schichten ruhig und arbeite gern im Team.", { after: 0, line: 288 }),
    h.head("QUALIFIKATIONEN & EXAMEN"),
    ...h.entry("Staatliches Examen zur Pflegefachfrau", "09/2018 – 08/2021", "Pflegeschule am Klinikum Musterstadt · Note 1,7"),
    h.head("BERUFSERFAHRUNG"),
    ...h.entry("Pflegefachfrau, Innere Medizin", "09/2021 – heute", "Klinikum Musterstadt", ["Grund- und Behandlungspflege für bis zu 10 Patienten pro Schicht", "Anleitung von Auszubildenden", "Elektronische Pflegedokumentation"]),
    h.head("FORTBILDUNGEN"),
    ...h.entry("Praxisanleiterin (300 Stunden)", "2023", "Bildungszentrum Musterstadt"),
    h.head("EINSATZBEREICHE"),
    h.t("Innere Medizin · Kardiologie · Notaufnahme (während der Ausbildung)", { after: 0 }),
    ...h.signature("Musterstadt, 18.11.2026", "Sarah Beispiel"),
  ];
  return sidebarSection(v, left, right);
}

function cvHandwerk(v) {
  const h = H(v);
  const left = [
    h.p([h.photo()], { align: AlignmentType.CENTER, noLine: true }),
    h.sideHead("KONTAKT"),
    h.sideLine("Am Werk 9"), h.sideLine("12345 Musterstadt", { after: 120 }), h.sideLine("0157 55667788", { after: 120 }), h.sideLine("tim.beispiel@email.de", { after: 120 }), h.sideLine("Geboren am 22.07.2001"),
    h.sideHead("FÜHRERSCHEINE"), h.sideLine("Klasse B und BE"), h.sideLine("Staplerschein"),
    h.sideHead("SPRACHEN"), ...h.sideKV("Deutsch", "Muttersprache"), ...h.sideKV("Englisch", "Grundkenntnisse", true),
    h.sideHead("INTERESSEN"), h.sideLine("Motorrad"), h.sideLine("Heimwerken"), h.sideLine("Fußball"),
  ];
  const right = [
    ...h.nameBlock("TIM BEISPIEL", "Elektroniker für Energie- und Gebäudetechnik", v.main),
    h.head("PROFIL"),
    h.t("Elektroniker mit fünf Jahren Erfahrung in Neubau und Sanierung. Ich arbeite selbstständig und sauber und behalte auch auf großen Baustellen den Überblick.", { after: 0, line: 288 }),
    h.head("BERUFSERFAHRUNG"),
    ...h.entry("Elektroniker für Energie- und Gebäudetechnik", "02/2021 – heute", "Elektro Muster GmbH, Musterstadt", ["Installation von Elektroanlagen in Wohn- und Gewerbebauten", "Montage und Prüfung von Unterverteilungen", "Fehlersuche und Wartung beim Kunden"]),
    h.head("AUSBILDUNG"),
    ...h.entry("Ausbildung zum Elektroniker", "08/2017 – 01/2021", "Elektro Beispiel KG · Gesellenprüfung Note 2,1"),
    h.head("MASCHINEN & WERKZEUGE"),
    h.t("Installationstester nach VDE 0100 · Kernbohrgerät · Mauernutfräse · Hebebühne", { after: 0 }),
    h.head("ZERTIFIKATE"),
    ...h.entry("Fachkraft für Photovoltaikanlagen", "2023", "Handwerkskammer Musterstadt"),
    ...h.signature("Musterstadt, 18.11.2026", "Tim Beispiel"),
  ];
  return sidebarSection(v, left, right);
}

function cvEnglisch(v) {
  const h = H(v);
  const left = [
    h.sideHead("CONTACT"),
    h.sideLine("Musterstraße 12"), h.sideLine("12345 Musterstadt"), h.sideLine("Germany", { after: 120 }), h.sideLine("+49 151 12345678", { after: 120 }), h.sideLine("max.mustermann@email.de"),
    h.sideHead("SKILLS"), ...h.sideKV("MS Office", "Word, Excel, Outlook"), ...h.sideKV("SAP", "Basic knowledge"), ...h.sideKV("Driving licence", "Class B", true),
    h.sideHead("LANGUAGES"), ...h.sideKV("German", "Native"), ...h.sideKV("English", "Fluent (C1)"), ...h.sideKV("Spanish", "Basic (A2)", true),
    h.sideHead("INTERESTS"), h.sideLine("Football"), h.sideLine("Photography"), h.sideLine("Cooking"),
  ];
  const right = [
    ...h.nameBlock(MAX.name, "Office Administrator"),
    h.head("PROFILE"),
    h.t("Organised office administrator with three years of experience in order processing. I enjoy accurate work and teamwork and am looking for a new role in inside sales.", { after: 0, line: 288 }),
    h.head("EXPERIENCE"),
    ...h.entry("Order Processing Clerk", "06/2023 – present", "Muster GmbH, Musterstadt, Germany", ["Processing around 40 customer orders per day", "Coordinating with warehouse and shipping", "Maintaining customer data in SAP"]),
    ...h.entry("Customer Service Assistant (part-time)", "09/2021 – 05/2023", "Beispiel AG, Musterstadt, Germany", ["Answering customer enquiries by phone and email"]),
    h.head("EDUCATION"),
    ...h.entry("Vocational Training: Office Management Assistant", "08/2020 – 06/2023", "Beispiel AG and Chamber of Commerce (IHK) · Final grade: good (2.0)"),
    ...h.entry("Intermediate School Certificate", "08/2014 – 07/2020", "Realschule Musterstadt"),
  ];
  return sidebarSection(v, left, right);
}

function motivation(v) {
  const h = H(v);
  const bandLeft = [
    h.t("LENA BEISPIEL", { size: 46, bold: true, color: "FFFFFF", cs: 30, after: 60, line: 240 }),
    h.t("Beispielweg 7 · 12345 Musterstadt · 0176 11223344 · lena.beispiel@email.de", { size: 19, color: v.soft, after: 0 }),
  ];
  const sub = (x) => h.t(x, { size: 22, bold: true, color: v.accent, before: 280, after: 80, keepNext: true });
  const body = [
    h.t("Musterstadt, 18.11.2026", { align: AlignmentType.RIGHT, before: 300, after: 400 }),
    h.t("Motivationsschreiben", { size: 44, bold: true, color: v.accent, after: 60 }),
    h.t("für: Bachelor Soziale Arbeit, Hochschule Musterstadt", { size: 23, color: GREY, after: 120, border: h.rule(v.rule, 6) }),
    sub("Einleitung"),
    h.placeholder("[Wer bist du, und wofür bewirbst du dich? Zwei bis drei Sätze reichen.]", 120),
    sub("Warum dieses Programm"),
    h.placeholder("[Was reizt dich genau an diesem Studiengang, dieser Stelle oder diesem Programm? Nenne etwas Konkretes.]", 120),
    sub("Warum ich"),
    h.placeholder("[Welche Erfahrungen und Stärken bringst du mit? Ein Beispiel aus Schule, Praktikum, Ehrenamt oder Job.]", 120),
    sub("Schluss"),
    h.placeholder("[Was möchtest du erreichen? Bedanke dich und freu dich auf eine Rückmeldung.]", 400),
    h.t("Mit freundlichen Grüßen", { after: 800 }),
    h.t("Lena Beispiel", { after: 0 }),
  ];
  return bandSection(v, h, bandLeft, [], body);
}

const PRODUCTS = [
  { dir: "02-bewerbungsset", file: "Bewerbungsset_Modern_Klar_Navy", v: C.navy, title: "Bewerbungsset", sections: (v) => [deckblattA(v), anschreiben(v), cvModernKlar(v), anlagen(v)] },
  { dir: "03-lebenslauf-schueler", file: "Lebenslauf_Erste_Bewerbung_Petrol", v: C.petrol, title: "Lebenslauf Erste Bewerbung", sections: (v) => [cvSchueler(v)] },
  { dir: "04-lebenslauf-minimal", file: "Lebenslauf_Minimal_Schwarz", v: C.schwarz, title: "Lebenslauf Minimal", sections: (v) => [cvMinimal(v)] },
  { dir: "05-lebenslauf-pflege", file: "Lebenslauf_Pflege_Blau", v: C.pflege, title: "Lebenslauf Pflege", sections: (v) => [cvPflege(v)] },
  { dir: "06-lebenslauf-handwerk", file: "Lebenslauf_Handwerk_Anthrazit_Orange", v: C.handwerk, title: "Lebenslauf Handwerk", sections: (v) => [cvHandwerk(v)] },
  { dir: "07-deckblatt", file: "Deckblatt_3_Designs_Navy", v: C.navy, title: "Deckblatt", sections: (v) => [deckblattA(v), deckblattB(v), deckblattC(v)] },
  { dir: "08-anschreiben", file: "Anschreiben_Modern_Klar_Navy", v: C.navy, title: "Anschreiben", sections: (v) => [anschreiben(v)] },
  { dir: "09-lebenslauf-deutsch-englisch", file: "Lebenslauf_Deutsch_Englisch_Navy", v: C.navy, title: "Lebenslauf Deutsch Englisch", sections: (v) => [cvModernKlar(v), cvEnglisch(v)] },
  { dir: "10-motivationsschreiben", file: "Motivationsschreiben_Petrol", v: C.petrol, title: "Motivationsschreiben", sections: (v) => [motivation(v)] },
];

(async () => {
  for (const pr of PRODUCTS) {
    const dir = path.join(OUT, pr.dir);
    fs.mkdirSync(dir, { recursive: true });
    const buf = await Packer.toBuffer(makeDoc(pr.v, pr.sections(pr.v), pr.title));
    const f = path.join(dir, pr.file + ".docx");
    fs.writeFileSync(f, buf);
    console.log("wrote", f);
  }
})().catch((e) => { console.error(e); process.exit(1); });
