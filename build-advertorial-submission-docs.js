const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition,
  LevelFormat, ImageRun,
} = require("docx");

const DIR = "/Users/shaunhunsaker/medicareplancheck";
const SD = path.join(DIR, "state-sites");
const NAVY = "0F2D5E", GREY = "CCCCCC", LIGHTBLUE = "D5E8F0", FLAGFILL = "FCE9E6";
const border = { style: BorderStyle.SINGLE, size: 1, color: GREY };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

function lv(label, value) {
  return new TableRow({ children: [
    new TableCell({ borders, width: { size: 2880, type: WidthType.DXA }, shading: { fill: LIGHTBLUE, type: ShadingType.CLEAR }, margins: cm, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ children: [new TextRun({ text: label, bold: true })] })] }),
    new TableCell({ borders, width: { size: 6480, type: WidthType.DXA }, margins: cm, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ children: [new TextRun(value)] })] }),
  ] });
}
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const p = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 100 }, children: [new TextRun({ text: t, bold: o.bold, italics: o.italics, size: o.size })] });
const quote = (t) => new Paragraph({ spacing: { after: 120 }, indent: { left: 360, right: 360 }, border: borders, children: [new TextRun({ text: t, size: 19 })] });
const bullet = (t) => new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun(t)] });

function flagBox(lines) {
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [new TableRow({ children: [new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "C0392B" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "C0392B" }, left: { style: BorderStyle.SINGLE, size: 4, color: "C0392B" }, right: { style: BorderStyle.SINGLE, size: 4, color: "C0392B" } },
    width: { size: 9360, type: WidthType.DXA }, shading: { fill: FLAGFILL, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [ new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "⚠ BEFORE PUBLISHING", bold: true, color: "C0392B" })] }), ...lines.map((l) => new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: l, size: 20 })] })) ],
  })] })] });
}

const TPMO = "We do not offer every plan available in your area. Any information we provide is limited to those plans we do offer in your area. Please contact Medicare.gov, 1-800-MEDICARE, or your local State Health Insurance Program (SHIP) to get information on all of your options.";

const STATES = [
  { code: "tx", name: "Texas", domain: "TexasPlanCheck.org", phone: "1-866-312-7162", tz: "CT", smid: "MULTIPLAN_MPCNativeTX26_M", shot: "submission-shot-tx.png" },
  { code: "nc", name: "North Carolina", domain: "NorthCarolinaPlanCheck.org", phone: "1-888-595-8709", tz: "ET", smid: "MULTIPLAN_MPCNativeNC26_M", shot: "submission-shot-nc.png" },
  { code: "fl", name: "Florida", domain: "FloridaPlanCheck.org", phone: "1-888-595-8806", tz: "ET", smid: "MULTIPLAN_MPCNativeFL26_M", shot: "submission-shot-fl.png" },
];

function buildDoc(st) {
  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "NATIVE ADVERTORIAL — WEB PAGE", bold: true, size: 30, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Marketing Material Submission", size: 24, color: "555555" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `${st.name} — https://www.${st.domain}`, italics: true, size: 22, color: "555555" })] }),

    h1("1. Submission Details"),
    new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2880, 6480], rows: [
      lv("Material ID (SMID)", `${st.smid}`),
      lv("Sponsor", "Medicare Finder"),
      lv("Marketing Organization", "Medicare Finder"),
      lv("Material Type", "Website / native advertorial (editorial article)"),
      lv("Traffic Source", "Taboola native content recommendation (paid)"),
      lv("URL", `https://www.${st.domain}`),
      lv("Time of Use", "Year-Round (Special Enrollment Period — SEP)"),
      lv("Plan Type Referenced", "Medicare Advantage (MA / MAPD)"),
      lv("Phone Number Disclosed", `${st.phone} (CallRail tracking → licensed agents; ${st.tz} hours)`),
      lv("Geography", `${st.name} residents`),
    ] }),

    h1("2. Page Capture (Above the Fold)"),
    new Paragraph({ spacing: { after: 80 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(SD, st.shot)), transformation: { width: 580, height: 741 }, altText: { title: `${st.name} advertorial`, description: "Above the fold", name: st.shot } })] }),
    p("Representative capture; full page copy transcribed verbatim below. Live page renders identically across states with state name, domain, phone, and timezone localized.", { italics: true, size: 18, after: 160 }),

    h1("3. Page Copy (Verbatim)"),
    h2("Header / Branding"),
    p('Label: "Advertisement · Paid Content"'),
    p(`Masthead: "${st.name} PlanCheck — Senior Benefits Desk"  (neutral branding; "Medicare" not used in domain or brand name)`),
    h2("Article"),
    p(`Kicker: "${st.name} · Medicare Benefits"`),
    p(`Headline: "Many ${st.name} Seniors on Medicare May Be Missing Out on Monthly Extra Benefits"`),
    p(`Dek: "Some Medicare Advantage plans available in ${st.name} may include extra benefits — like a grocery allowance, Part B give-back, dental coverage, and prescription savings — yet many members never know to ask. Here's how to check what you may qualify for."`),
    p('Tag: "Sponsored Content · Provided by [State] PlanCheck — A licensed insurance marketing resource"'),
    p("Body lists benefits some MA plans MAY include: (1) monthly grocery / healthy-food allowance (for members who qualify — see eligibility note); (2) Part B give-back; (3) dental, vision & hearing; (4) prescription drug savings; (5) monthly OTC allowance."),
    h2("Grocery / Food Allowance Eligibility Note (verbatim)"),
    quote("A note on grocery and food allowances: these are generally available only to members who meet specific eligibility requirements — for example, having a qualifying chronic health condition, or also being enrolled in Medicaid. A licensed agent can tell you whether you may or may not qualify."),
    p(`CTAs: "Call ${st.phone}" — "Speak with a licensed insurance agent · Free · No obligation" (repeated; sticky mobile call bar). Checklist of items to ask a licensed agent about.`),

    h1("4. Disclaimers (Verbatim)"),
    p("TPMO disclaimer (count-free generic form — CMS-accepted for anonymous landing pages where visitor ZIP is unknown; area-specific organization/plan counts are used in the verbal disclaimer on calls):", { bold: true, after: 100 }),
    quote(TPMO),
    p("Footer non-affiliation / full disclaimer:", { bold: true, after: 100 }),
    quote(`${st.domain} is operated by Medicare Finder, a licensed insurance marketing organization, and is not affiliated with or endorsed by the U.S. government, the federal Medicare program, or CMS. This is a solicitation for insurance. Calling the number on this page connects you with a licensed insurance agent. Medicare Advantage plans are offered by private insurance companies contracted with Medicare. Enrollment depends on the plan's contract renewal. You must continue to pay your Medicare Part B premium. Benefits, premiums, and plan availability vary by plan and location and are subject to change. Not all plans available in all areas. Plans are available to those who qualify. Enrollment may be limited to certain times of the year unless you qualify for a Special Enrollment Period. SMID: ${st.smid}`),

    h1("5. Compliance Notes"),
    bullet('FTC native-ad disclosure: page is labeled "Advertisement · Paid Content" and "Sponsored Content."'),
    bullet('No fake byline and no impersonation of any news outlet — editorial style only; publisher identified as the marketing resource.'),
    bullet('Neutral domain and brand ("[State] PlanCheck") — "Medicare" appears only as textual explanation in body content, never in the domain, brand name, or as an agent title.'),
    bullet('Staff referred to only as "licensed insurance agent" — no "Medicare expert/specialist" titles.'),
    bullet('Grocery/food allowance is gated with an explicit eligibility note (SSBCI chronic-condition / Medicaid) — not presented as universally available; consistent with CMS SSBCI marketing rules.'),
    bullet('All benefit claims qualified ("may include," "some plans," "for members who qualify").'),
    bullet('No CMS/HHS logos and no image of the Medicare card.'),
    bullet('Gov non-affiliation disclaimer is prominent (top bar), adjacent to branding — not footer-only.'),

    new Paragraph({ spacing: { before: 160, after: 80 }, children: [new TextRun("")] }),
    flagBox([
      "On-page TPMO disclaimer is the count-free anonymous-page version; the area-specific verbal disclaimer (with organization/plan counts for the caller's ZIP) is delivered by the licensed agent within the first minute of the call.",
      "Hero image is licensed stock; confirm license terms permit advertising use.",
    ]),
  ];

  return new Document({
    creator: "Medicare Finder",
    title: `Advertorial Submission — ${st.domain}`,
    styles: { default: { document: { run: { font: "Arial", size: 22 } } }, paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: NAVY }, paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 23, bold: true, font: "Arial", color: NAVY }, paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
    ] },
    numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } }, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }], children: [ new TextRun({ text: "Medicare Finder — Marketing Material Submission", bold: true, color: NAVY, size: 18 }), new TextRun({ text: `\t${st.domain}`, color: "666666", size: 18 }) ] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: "CONFIDENTIAL — For carrier / CMS-HPMS review only      Page ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }), new TextRun({ text: " of ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: "888888" }) ] })] }) },
      children,
    }],
  });
}

(async () => {
  for (const st of STATES) {
    const buf = await Packer.toBuffer(buildDoc(st));
    const fname = `Advertorial-Submission-${st.name.replace(/ /g, "")}-${st.smid}.docx`;
    fs.writeFileSync(path.join(DIR, fname), buf);
    console.log("Wrote", fname);
  }
})();
