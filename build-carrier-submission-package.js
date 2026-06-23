// Builds the Carrier Submission Package:
//   1) Instructions page
//   2) Master tracking grid  (13 materials x {Aetna, Humana, UnitedHealthcare})
//   3) One pre-filled cover sheet per material (carrier + date blank to fill)
// Run:  NODE_PATH=$(npm root -g) node build-carrier-submission-package.js
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition, LevelFormat,
} = require("docx");

const NAVY = "0F2D5E", GREY = "CCCCCC", LIGHTBLUE = "D5E8F0", AMBER = "FBBF24";
const border = { style: BorderStyle.SINGLE, size: 1, color: GREY };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

const CARRIERS = ["Aetna", "Humana", "UnitedHealthcare"];

const TPMO = "We do not offer every plan available in your area. Any information we provide is limited to those plans we do offer in your area. Please contact Medicare.gov, 1-800-MEDICARE, or your local State Health Insurance Program (SHIP) to get information on all of your options.";

// Every marketing material that needs carrier sign-off before use.
const MATERIALS = [
  { id: "MULTIPLAN_MPCNativeTX26_M", name: "Texas Advertorial", type: "Native advertorial / web page", channel: "Taboola native (paid)", where: "https://www.texasplancheck.org", phone: "1-888-595-8641", status: "LIVE" },
  { id: "MULTIPLAN_MPCNativeNC26_M", name: "North Carolina Advertorial", type: "Native advertorial / web page", channel: "Taboola native (paid)", where: "https://www.northcarolinaplancheck.org", phone: "1-888-595-8709", status: "LIVE" },
  { id: "MULTIPLAN_MPCNativeFL26_M", name: "Florida Advertorial", type: "Native advertorial / web page", channel: "Taboola native (paid)", where: "https://www.floridaplancheck.org", phone: "1-888-595-8806", status: "LIVE" },
  { id: "MULTIPLAN_MPCWeb26LP1_M", name: "Flagship Website / Landing Page", type: "Website landing page", channel: "FB + StackAdapt destination", where: "https://www.medicareplancheck.org", phone: "1-877-339-1956", status: "LIVE" },
  { id: "NSBA_MPC_SOCIAL26_FA1_M", name: "Facebook — Grocery Allowance (FA1)", type: "Social display ad (static image)", channel: "Facebook / Instagram Feed", where: "→ medicareplancheck.org", phone: "1-866-659-6122", status: "LIVE" },
  { id: "NSBA_MPC_SOCIAL26_FA2_M", name: "Facebook — Part B Give-Back (FA2)", type: "Social display ad (static image)", channel: "Facebook / Instagram Feed", where: "→ medicareplancheck.org", phone: "1-866-312-0390", status: "LIVE" },
  { id: "NSBA_MPC_SOCIAL26_FA3_M", name: "Facebook — OTC Allowance (FA3)", type: "Social display ad (static image)", channel: "Facebook / Instagram Feed", where: "→ medicareplancheck.org", phone: "1-866-312-0462", status: "LIVE" },
  { id: "NSBA_MPC_SOCIAL26_FA4_M", name: "Facebook — Trust / Independent (FA4)", type: "Social display ad (static image)", channel: "Facebook / Instagram Feed", where: "→ medicareplancheck.org", phone: "1-866-312-0368", status: "LIVE" },
  { id: "NSBA_MPC_SOCIAL26_FA1_M", name: "StackAdapt — Grocery Allowance (FA1)", type: "Programmatic display banner (300x250, 320x50)", channel: "StackAdapt display", where: "→ medicareplancheck.org", phone: "(click-to-call banner)", status: "PAUSED" },
  { id: "NSBA_MPC_SOCIAL26_FA2_M", name: "StackAdapt — Part B Give-Back (FA2)", type: "Programmatic display banner (300x250, 320x50)", channel: "StackAdapt display", where: "→ medicareplancheck.org", phone: "(click-to-call banner)", status: "PAUSED" },
  { id: "NSBA_MPC_SOCIAL26_FA3_M", name: "StackAdapt — OTC Allowance (FA3)", type: "Programmatic display banner (300x250, 320x50)", channel: "StackAdapt display", where: "→ medicareplancheck.org", phone: "(click-to-call banner)", status: "PAUSED" },
  { id: "NSBA_MPC_SOCIAL26_FA4_M", name: "StackAdapt — Trust / Independent (FA4)", type: "Programmatic display banner (300x250, 320x50)", channel: "StackAdapt display", where: "→ medicareplancheck.org", phone: "(click-to-call banner)", status: "PAUSED" },
  { id: "NSBA_MPC_CTV26_TV1_M", name: "CTV / OTT :30 Video — “Benefits on the Table”", type: "Connected-TV / OTT :30 video", channel: "Tubi, Pluto, Peacock, Xumo, Samsung TV+", where: "(video) → medicareplancheck.org", phone: "1-866-659-6123", status: "STAGED (paused)" },
];

const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const p = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 100 }, alignment: o.align, children: [new TextRun({ text: t, bold: o.bold, italics: o.italics, size: o.size, color: o.color })] });
const bullet = (t) => new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun(t)] });
const pageBreak = () => new Paragraph({ children: [new TextRun({ text: "", break: 0 })], pageBreakBefore: true });

function lv(label, value, fill) {
  return new TableRow({ children: [
    new TableCell({ borders, width: { size: 2700, type: WidthType.DXA }, shading: { fill: LIGHTBLUE, type: ShadingType.CLEAR }, margins: cm, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ children: [new TextRun({ text: label, bold: true })] })] }),
    new TableCell({ borders, width: { size: 6660, type: WidthType.DXA }, margins: cm, verticalAlign: VerticalAlign.CENTER, shading: fill ? { fill, type: ShadingType.CLEAR } : undefined, children: [new Paragraph({ children: [new TextRun(value)] })] }),
  ] });
}
function cell(text, opts = {}) {
  return new TableCell({ borders, width: { size: opts.w, type: WidthType.DXA }, margins: cm, verticalAlign: VerticalAlign.CENTER, shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({ alignment: opts.align, children: [new TextRun({ text, bold: opts.bold, size: opts.size ?? 18, color: opts.color })] })] });
}

// ---------- Master tracking grid ----------
function trackerTable() {
  const head = new TableRow({ tableHeader: true, children: [
    cell("Material (SMID)", { w: 3960, bold: true, fill: NAVY, color: "FFFFFF", size: 18 }),
    ...CARRIERS.map((c) => cell(c, { w: 1800, bold: true, fill: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, size: 18 })),
  ] });
  const rows = MATERIALS.map((m) => new TableRow({ children: [
    new TableCell({ borders, width: { size: 3960, type: WidthType.DXA }, margins: cm, children: [
      new Paragraph({ children: [new TextRun({ text: m.name, bold: true, size: 18 })] }),
      new Paragraph({ children: [new TextRun({ text: m.id, size: 15, color: "666666" })] }),
    ] }),
    ...CARRIERS.map(() => new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, margins: cm, children: [
      new Paragraph({ children: [new TextRun({ text: "Sub: ____ ", size: 15, color: "888888" })] }),
      new Paragraph({ children: [new TextRun({ text: "Appr: ____", size: 15, color: "888888" })] }),
    ] })),
  ] }));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3960, 1800, 1800, 1800], rows: [head, ...rows] });
}

// ---------- Per-material cover sheet ----------
function coverSheet(m) {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "MARKETING MATERIAL — CARRIER REVIEW COVER SHEET", bold: true, size: 26, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: m.name, size: 22, color: "555555" })] }),
    new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2700, 6660], rows: [
      lv("Submitting Organization", "Medicare Finder (licensed insurance marketing organization)"),
      lv("Contact", "Shaun Hunsaker — shaun@hskinsurance.com"),
      lv("Carrier (this submission)", "    __________________________  (fill in)", "FFF7E6"),
      lv("Date Submitted", "    __________________________  (fill in)", "FFF7E6"),
      lv("Material ID (SMID)", m.id),
      lv("Material Name", m.name),
      lv("Material Type", m.type),
      lv("Channel / Placement", m.channel),
      lv("URL / Location", m.where),
      lv("Phone Disclosed", `${m.phone} (CallRail tracking → licensed agents)`),
      lv("TPMO Status", "MULTI-PLAN (we do not offer every plan in the area)"),
      lv("Plan Type", "Medicare Advantage (MA / MAPD)"),
      lv("Time of Use", "Year-Round (Special Enrollment Period — SEP)"),
    ] }),
    new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: "Request", bold: true, size: 22, color: NAVY })] }),
    p("Medicare Finder respectfully submits the attached material for your marketing-compliance review and approval prior to use. Please advise of any required edits or your approval / disposition. Upon carrier approval this material will be filed in CMS HPMS under the SMID above.", { size: 19 }),
    new Paragraph({ spacing: { before: 160, after: 80 }, children: [new TextRun({ text: "TPMO Disclaimer present on / with this material (verbatim)", bold: true, size: 20, color: NAVY })] }),
    new Paragraph({ spacing: { after: 120 }, indent: { left: 360, right: 360 }, border: borders, children: [new TextRun({ text: TPMO, size: 18, italics: true })] }),
    new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: "Carrier disposition:  ☐ Approved   ☐ Approved w/ edits   ☐ Returned   — Reviewer: ____________  Date: ________", size: 18, color: "555555" })] }),
  ];
}

const children = [
  // ---- Instructions ----
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "CARRIER SUBMISSION PACKAGE", bold: true, size: 34, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Medicare Finder — Medicare Advantage Marketing Materials", size: 22, color: "555555" })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "Prepared June 15, 2026", italics: true, size: 20, color: "777777" })] }),
  h1("How to use this package"),
  bullet("Submit each material below to each carrier's marketing-compliance review (portal or compliance inbox) BEFORE it runs. Carrier approval comes first; CMS HPMS filing under the SMID comes after."),
  bullet("Use the tracking grid on the next page to log Submitted (Sub) and Approved (Appr) dates per carrier."),
  bullet("Attach the matching one-page cover sheet (one per material, following the grid) to each submission — fill in the Carrier and Date fields (highlighted)."),
  bullet("Full creative + verbatim copy + storyboards live in the per-material submission docs already in the repo (Advertorial-Submission-*, HPMS-Website-*, HPMS-FB-*, HPMS-StackAdapt-*, HPMS-CTV-*). Attach the relevant one alongside the cover sheet."),
  p("Note: LIVE materials are already running on original baked-in SMIDs; advertorials + website carry the MULTIPLAN_ marketing SMIDs. Confirm each carrier's specific submission channel — methods differ by carrier.", { italics: true, size: 18, after: 160 }),
  h1("Master Tracking Grid"),
  trackerTable(),
  ...MATERIALS.flatMap((m) => [pageBreak(), ...coverSheet(m)]),
];

const doc = new Document({
  creator: "Medicare Finder",
  title: "Carrier Submission Package",
  styles: { default: { document: { run: { font: "Arial", size: 22 } } }, paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: NAVY }, paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 0 } },
  ] },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } }, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }], children: [new TextRun({ text: "Medicare Finder — Carrier Submission Package", bold: true, color: NAVY, size: 18 }), new TextRun({ text: "\tCONFIDENTIAL", color: "666666", size: 18 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "For carrier marketing-compliance review      Page ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }), new TextRun({ text: " of ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: "888888" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = "/Users/shaunhunsaker/medicareplancheck/Carrier-Submission-Package-MedicareFinder.docx";
  fs.writeFileSync(out, buf);
  console.log("Wrote " + out + "  (" + (buf.length / 1024).toFixed(0) + " KB, " + MATERIALS.length + " cover sheets)");
});
