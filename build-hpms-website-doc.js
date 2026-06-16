const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition,
  LevelFormat, ImageRun,
} = require("docx");

const DIR = "/Users/shaunhunsaker/medicareplancheck";
const NAVY = "0F2D5E", AMBER = "B8860B", GREY = "CCCCCC", LIGHTBLUE = "D5E8F0", FLAGFILL = "FCE9E6";
const border = { style: BorderStyle.SINGLE, size: 1, color: GREY };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function lv(label, value) {
  return new TableRow({ children: [
    new TableCell({ borders, width: { size: 2880, type: WidthType.DXA }, shading: { fill: LIGHTBLUE, type: ShadingType.CLEAR }, margins: cellMargins, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ children: [new TextRun({ text: label, bold: true })] })] }),
    new TableCell({ borders, width: { size: 6480, type: WidthType.DXA }, margins: cellMargins, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ children: [new TextRun(value)] })] }),
  ] });
}
function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] }); }
function p(t, o = {}) { return new Paragraph({ spacing: { after: o.after ?? 100 }, indent: o.indent ? { left: 360 } : undefined, children: [new TextRun({ text: t, bold: o.bold, italics: o.italics, size: o.size })] }); }
function quote(t) { return new Paragraph({ spacing: { after: 120 }, indent: { left: 360, right: 360 }, border: borders, children: [new TextRun({ text: t, size: 19 })] }); }
function bullet(t) { return new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun(t)] }); }

function flagBox(title, color, lines) {
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [new TableRow({ children: [new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color }, bottom: { style: BorderStyle.SINGLE, size: 4, color }, left: { style: BorderStyle.SINGLE, size: 4, color }, right: { style: BorderStyle.SINGLE, size: 4, color } },
    width: { size: 9360, type: WidthType.DXA }, shading: { fill: FLAGFILL, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [ new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: title, bold: true, color })] }), ...lines.map((l) => new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: l, size: 20 })] })) ],
  })] })] });
}

const TCPA = "By submitting this form, I provide my express written consent to be contacted by MedicarePlanCheck.org and its licensed insurance agents at the phone number provided, including via live agent, automated dialing system, pre-recorded message, and/or text/SMS, regarding Medicare Advantage and Supplement insurance plans. I understand consent is not required as a condition of purchasing any insurance product. Message & data rates may apply. Reply STOP to opt out of texts. See our Privacy Policy.";
const DISCLAIMER = "Not affiliated with or endorsed by the U.S. government, CMS, or the federal Medicare program. MedicarePlanCheck.org is a licensed insurance marketing organization (IMO). We represent multiple insurance carriers and are compensated by insurers if you enroll in a plan. This is a solicitation for insurance. Plans are insured or covered by a Medicare Advantage organization and/or a Medicare-approved Part D sponsor. Enrollment depends on the plan's contract renewal. Calling this number connects you with a licensed insurance agent. Benefits and plan availability vary by location and are subject to change. Not all plans available in all areas. We do not offer every plan available in your area. Any information we provide is limited to those plans we do offer in your area. Please contact Medicare.gov, 1-800-MEDICARE, or your local State Health Insurance Program (SHIP) to get information on all of your options.";

const SMID = "NSBA_MULTIPLAN_MPC_WEB26_LP1_M  (displayed in site footer)";

const doc = new Document({
  creator: "Medicare Finder",
  title: "HPMS Website Submission — MedicarePlanCheck.org",
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: NAVY }, paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 23, bold: true, font: "Arial", color: NAVY }, paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } }, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }], children: [ new TextRun({ text: "Medicare Finder — Marketing Material Submission", bold: true, color: NAVY, size: 18 }), new TextRun({ text: "\tWebsite / Landing Page", color: "666666", size: 18 }) ] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: "CONFIDENTIAL — For CMS/HPMS review only      Page ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }), new TextRun({ text: " of ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: "888888" }) ] })] }) },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "WEBSITE / LANDING PAGE", bold: true, size: 30, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Marketing Material Submission", size: 24, color: "555555" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "https://www.medicareplancheck.org", italics: true, size: 22, color: "555555" })] }),

      h1("1. Submission Details"),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2880, 6480], rows: [
        lv("Material ID (SMID)", SMID),
        lv("Sponsor", "Medicare Finder"),
        lv("Marketing Organization", "Medicare Finder"),
        lv("Material Type", "Website / Landing Page (lead capture form + click-to-call)"),
        lv("URL", "https://www.medicareplancheck.org"),
        lv("Time of Use", "Year-Round (Special Enrollment Period — SEP)"),
        lv("Plan Types Referenced", "Medicare Advantage (Part C), Medicare Supplement (Medigap), Part D"),
        lv("Phone Number Disclosed", "1-877-339-1956  •  TTY: 711"),
        lv("Hours of Operation", "Mon–Fri 8am–8pm ET; Sat 9am–5pm ET"),
        lv("Lead Capture", "3-step form (ZIP → Name → Phone) with TCPA express written consent checkbox"),
        lv("Tracking Pixels", "Meta Pixel (2318251708228114); StackAdapt Universal Pixel"),
      ] }),

      h1("2. Page Screenshot (Above the Fold)"),
      new Paragraph({ spacing: { after: 80 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(DIR, "hpms-lp-hero.png")), transformation: { width: 600, height: 431 }, altText: { title: "Landing page hero", description: "MedicarePlanCheck.org above the fold", name: "hpms-lp-hero.png" } })] }),
      p("Representative capture of the hero / above-the-fold view. The complete live page is available at the URL above; full page copy is transcribed verbatim in Sections 3–5.", { italics: true, size: 18, after: 160 }),

      h1("3. Page Content (Verbatim)"),
      h2("Hero"),
      p("Eyebrow: 🎁 Extra Benefits Available in Your Area"),
      p("Headline: Your Medicare Plan Could Include $150/Month in Extra Benefits"),
      p("Subhead: Many Medicare Advantage plans include monthly allowances for groceries, OTC items, dental, and vision — at $0 extra cost. Find out what you qualify for in minutes."),
      p("Primary CTA: Call Now: 1-877-339-1956   |   Hours: Mon–Fri 8am–8pm ET, Sat 9am–5pm ET"),
      p("Trust bullets: Same-day enrollment available • $0 cost consultation — no strings attached • No pushy sales — just honest answers"),
      h2("Stats Bar"),
      p("$0 Cost Consultation • 500k+ People Helped • No Obligation • Same Day Enrollment"),
      h2("Extra Benefits Section"),
      p("Heading: You May Have Unclaimed Benefits Waiting"),
      p("Sub: Many Medicare Advantage members never use all the extra perks included in their plan — or don't even know they exist."),
      p("Monthly OTC Allowance: Up to $150 per month to spend on everyday essentials — Groceries & healthy foods; Over-the-counter medications; Personal care products; Household health items. (\"Like a monthly reward card — loaded automatically.\")"),
      p("Plus many plans also include: Dental, Vision & Hearing; Free Gym Membership (SilverSneakers and similar); $0 Premium Options (availability varies by ZIP)."),
      h2("How It Works"),
      p("1) Call Our Free Helpline — Dial 1-877-339-1956 and connect with a licensed Medicare specialist. 2) Compare Your Options — review plans from 30+ carriers in your ZIP. 3) Enroll — Hassle-Free — we handle enrollment paperwork; most enrolled in under 20 minutes."),
      h2("Plans We Can Help You Compare"),
      p("Medicare Advantage (Part C) — Most Popular; Medicare Supplement (Medigap); Part D Drug Plans. Feature lists per card included on page."),
      h2("Why Choose MedicarePlanCheck?"),
      p("Unbiased Comparisons (represent multiple carriers); No Cost to You — Ever (compensated by carriers, not you); Real Humans, Real Answers (no chatbots)."),
      h2("FAQ"),
      p("Covers: enrollment periods (IEP/AEP/MA-OEP/SEP); $0 premium plans (notes Part B premium still applies); MA vs. Medigap; annual plan review; cost of service (free)."),
      h2("Closing CTA Banner"),
      p("Ready to Find Your Best Medicare Plan? Speak with a licensed agent now — free, fast, and no obligation. Call 1-877-339-1956."),

      h1("4. Lead Form & TCPA Consent"),
      p("The hero contains a 3-step lead capture form:", { after: 80 }),
      bullet("Step 1 — ZIP Code (\"Unlock Plans in Your Area\"). Fine print: \"Plans and benefits vary by location.\""),
      bullet("Step 2 — First & Last Name (\"Who Are We Looking Up?\")."),
      bullet("Step 3 — Phone Number + required TCPA consent checkbox (\"Claim Your Free Results\")."),
      p("TCPA express written consent (verbatim, appears at form Step 3 and in the footer):", { bold: true, after: 80 }),
      quote(TCPA),

      h1("5. Disclaimer (Verbatim — Footer)"),
      quote(DISCLAIMER),
      p("Footer also displays: contact phone 1-877-339-1956, TTY: 711, hours, copyright © 2026 MedicarePlanCheck.org, and links labeled Privacy Policy, Terms of Use, Accessibility, Do Not Sell My Info.", { italics: true, size: 18 }),

      h1("6. Compliance Notes"),
      bullet("TCPA express written consent is present, with not-a-condition-of-purchase language, STOP opt-out, and message/data-rate notice."),
      bullet("Non-affiliation disclaimer (\"Not affiliated with or endorsed by the U.S. government, CMS, or the federal Medicare program\") is present, plus IMO disclosure and \"solicitation for insurance.\""),
      bullet("TPMO disclaimer present (count-free anonymous-page version): \"We do not offer every plan available in your area...\" with Medicare.gov / 1-800-MEDICARE / SHIP referral, consistent with the advertorial and CTV materials."),
      bullet("Benefit claims use qualifiers (\"could include,\" \"up to,\" \"many plans,\" \"availability varies\")."),
      bullet("TTY: 711 disclosed. Hours of operation disclosed."),

      new Paragraph({ spacing: { before: 160, after: 80 }, children: [new TextRun("")] }),
      flagBox("⚠ REVIEW BEFORE SUBMITTING", "C0392B", [
        "Footer links for Terms of Use and Accessibility still point to \"#\" (Privacy Policy and Do Not Sell My Info are live).",
        "Specific dollar figures (\"$150/Month,\" \"Up to $150\") appear; qualifiers are present but confirm acceptable for HPMS.",
        "Substantiation: \"500k+ People Helped,\" \"Helping Americans … since 2018,\" and \"Same-day enrollment available\" are factual claims — ensure they are substantiated or softened.",
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(DIR, "HPMS-Website-Submission-MedicarePlanCheck.docx"), buf);
  console.log("Wrote HPMS-Website-Submission-MedicarePlanCheck.docx");
});
