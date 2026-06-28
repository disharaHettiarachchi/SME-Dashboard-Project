import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = process.env.FINAL_PPTX || path.resolve("output/BI_Dashboard_Proposal_Presentation.pptx");
const TMP = process.env.TMP_DIR || path.resolve("tmp/proposal_presentation");
const PREVIEW_DIR = process.env.PREVIEW_DIR || path.join(TMP, "preview");
const LAYOUT_DIR = process.env.LAYOUT_DIR || path.join(TMP, "layout");
const QA_DIR = process.env.QA_DIR || path.join(TMP, "qa");

const W = 1280;
const H = 720;
const C = {
  ink: "#000000",
  muted: "#555555",
  panel: "#EDEDED",
  light: "#F7F7F7",
  rule: "#B8BCC4",
  accent: "#FF6B35",
  white: "#FFFFFF",
};

const monthlyRevenue = [
  ["Dec 10", 824],
  ["Jan", 691],
  ["Feb", 524],
  ["Mar", 718],
  ["Apr", 538],
  ["May", 771],
  ["Jun", 762],
  ["Jul", 719],
  ["Aug", 759],
  ["Sep", 1059],
  ["Oct", 1155],
  ["Nov", 1509],
  ["Dec 11", 639],
];

const topCountries = [
  ["United Kingdom", 9025],
  ["Netherlands", 285],
  ["EIRE", 283],
  ["Germany", 229],
  ["France", 210],
  ["Australia", 139],
];

async function writeBlob(filePath, blob) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, Buffer.from(await blob.arrayBuffer()));
}

function addShape(slide, name, left, top, width, height, fill = C.panel, lineFill = "none") {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: lineFill, width: lineFill === "none" ? 0 : 1 },
  });
}

function addText(slide, name, text, left, top, width, height, opts = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left, top, width, height },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: opts.size ?? 22,
    bold: opts.bold ?? false,
    color: opts.color ?? C.ink,
    alignment: opts.align ?? "left",
  };
  return shape;
}

function addSlideNumber(slide, n) {
  addText(slide, `slide-${n}-footer`, String(n).padStart(2, "0"), 1184, 659, 54, 25, {
    size: 15,
    color: C.muted,
    align: "right",
  });
}

function addTitle(slide, n, title, kicker = "Proposal presentation") {
  addText(slide, `slide-${n}-kicker`, kicker, 42, 38, 360, 26, {
    size: 16,
    bold: true,
    color: C.muted,
  });
  addText(slide, `slide-${n}-title`, title, 42, 78, 990, 88, {
    size: 39,
    bold: true,
    color: C.ink,
  });
  addSlideNumber(slide, n);
}

function addBullets(slide, name, items, left, top, width, height, opts = {}) {
  const text = items.map((item) => `- ${item}`).join("\n");
  return addText(slide, name, text, left, top, width, height, {
    size: opts.size ?? 21,
    color: opts.color ?? C.ink,
  });
}

function addPanelText(slide, name, title, body, left, top, width, height, accent = false) {
  addShape(slide, `${name}-panel`, left, top, width, height, accent ? "#FFE4DA" : C.panel);
  addText(slide, `${name}-title`, title, left + 28, top + 25, width - 56, 34, {
    size: 24,
    bold: true,
    color: C.ink,
  });
  addText(slide, `${name}-body`, body, left + 28, top + 76, width - 56, height - 95, {
    size: 19,
    color: C.ink,
  });
}

function addMetric(slide, name, value, label, left, top, width, height, accent = false) {
  addShape(slide, `${name}-panel`, left, top, width, height, accent ? "#FFE4DA" : C.panel);
  addText(slide, `${name}-value`, value, left + 22, top + 24, width - 44, 60, {
    size: 38,
    bold: true,
    color: C.ink,
  });
  addText(slide, `${name}-label`, label, left + 22, top + 102, width - 44, height - 116, {
    size: 18,
    color: C.muted,
  });
}

function slide1(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addText(slide, "s1-kicker", "COM4901 Final Year Individual Research Project", 42, 42, 720, 30, {
    size: 18,
    bold: true,
    color: C.muted,
  });
  addText(slide, "s1-title", "BI Decision Support Dashboard", 42, 144, 790, 150, {
    size: 52,
    bold: true,
    color: C.ink,
  });
  addText(
    slide,
    "s1-subtitle",
    "Design and Development of a Business Intelligence-Based Decision Support Dashboard for Strategic Decision-Making in Small and Medium Enterprises",
    42,
    320,
    760,
    92,
    { size: 24, color: C.ink },
  );
  addShape(slide, "s1-side-panel", 864, 108, 374, 410, C.panel);
  addText(slide, "s1-panel-title", "Proposal focus", 896, 142, 310, 34, {
    size: 24,
    bold: true,
  });
  addBullets(
    slide,
    "s1-panel-bullets",
    [
      "Streamlit web app",
      "Interactive BI dashboard",
      "Public UCI Online Retail dataset",
      "Customer segmentation and forecasting",
      "Decision-support insight cards",
    ],
    896,
    198,
    302,
    240,
    { size: 21 },
  );
  addText(slide, "s1-meta", "Student: [Name]   ID: [Student ID]   Supervisor: [Name]", 42, 610, 850, 32, {
    size: 18,
    color: C.muted,
  });
  addSlideNumber(slide, 1);
}

function slide2(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 2, "Why this project is needed", "Project motivation");
  addText(slide, "s2-main", "SMEs often have transaction data, but not a simple decision-support layer.", 42, 212, 510, 180, {
    size: 40,
    bold: true,
  });
  addText(slide, "s2-note", "The proposed system converts raw retail records into KPIs, visual analytics, segments, forecasts, and recommendations.", 42, 440, 510, 100, {
    size: 22,
    color: C.muted,
  });
  addPanelText(slide, "s2-p1", "Manual reporting", "Spreadsheet review is slow and difficult to update when decisions need timely evidence.", 658, 213, 271, 279);
  addPanelText(slide, "s2-p2", "Weak customer insight", "Managers may not know which customers are high-value, inactive, or suitable for targeted retention.", 968, 213, 271, 279);
  addShape(slide, "s2-p3-panel", 658, 506, 581, 125, "#FFE4DA");
  addText(slide, "s2-p3-title", "Limited forecasting", 696, 538, 500, 34, {
    size: 24,
    bold: true,
  });
  addText(slide, "s2-p3-body", "Trend analysis helps plan sales and product decisions before performance drops.", 696, 588, 500, 36, {
    size: 18,
  });
}

function slide3(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 3, "Aim and objectives", "Research direction");
  addShape(slide, "s3-aim-panel", 42, 213, 581, 170, "#FFE4DA");
  addText(slide, "s3-aim-label", "Aim", 74, 242, 120, 32, { size: 24, bold: true, color: C.accent });
  addText(
    slide,
    "s3-aim",
    "Design, develop, and evaluate a BI-based web dashboard that supports strategic decision-making for SME-style retail businesses.",
    74,
    286,
    520,
    76,
    { size: 24, bold: true },
  );
  addPanelText(slide, "s3-o1", "Objective 1", "Study BI dashboards, DSS concepts, retail analytics, RFM, clustering, and forecasting.", 658, 213, 581, 172);
  addPanelText(slide, "s3-o2", "Objective 2", "Clean and preprocess the UCI Online Retail transaction dataset for repeatable analysis.", 42, 459, 581, 172);
  addPanelText(slide, "s3-o3", "Objective 3", "Build Streamlit dashboard modules for KPIs, sales, customers, products, decisions, and forecasts.", 658, 459, 581, 172);
}

function slide4(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 4, "Scope and deliverables", "What approval enables");
  addPanelText(slide, "s4-a", "Included", "A deployable Streamlit prototype with executive, sales, customer, product, decision-support, and forecasting pages.", 42, 214, 374, 238);
  addPanelText(slide, "s4-b", "Data pipeline", "Data loading, cleaning, calculated features, KPI functions, processed files, and reusable source modules.", 453, 214, 374, 238);
  addPanelText(slide, "s4-c", "Analytics", "RFM analysis, K-Means segmentation, monthly sales forecasting, and model evaluation metrics.", 864, 214, 374, 238, true);
  addText(slide, "s4-out", "Out of scope", 42, 524, 210, 34, { size: 24, bold: true });
  addBullets(slide, "s4-out-bullets", [
    "No private Sri Lankan SME data",
    "No paid APIs or enterprise BI tools",
    "No real-time POS, ERP, or accounting integration",
    "No production user account or multi-tenant security system",
  ], 270, 512, 790, 100, { size: 20 });
}

function slide5(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 5, "Dataset and ethical position", "Public transaction dataset");
  addMetric(slide, "s5-m1", "541,909", "transaction rows", 42, 213, 271, 148);
  addMetric(slide, "s5-m2", "38", "countries / markets", 351, 213, 271, 148);
  addMetric(slide, "s5-m3", "4,372", "identifiable customer IDs", 42, 393, 271, 148);
  addMetric(slide, "s5-m4", "25,900", "invoice numbers", 351, 393, 271, 148);
  addText(slide, "s5-ethics-title", "Ethics boundary", 42, 580, 130, 30, { size: 20, bold: true, color: C.accent });
  addText(slide, "s5-ethics", "Public UCI dataset only; no invented private SME data or new human participant data.", 178, 576, 440, 52, { size: 18 });
  slide.charts.add("bar", {
    position: { left: 680, top: 212, width: 530, height: 300 },
    categories: topCountries.map((row) => row[0]),
    series: [{ name: "Revenue GBP thousands", values: topCountries.map((row) => row[1]), fill: C.ink }],
    hasLegend: false,
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 52 },
    xAxis: { visible: false, majorGridlines: null },
    yAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: { style: "solid", fill: C.rule, width: 1 } },
    dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 12, bold: true } },
  });
  addText(slide, "s5-chart-note", "Top markets by positive-transaction revenue, GBP thousands", 680, 520, 530, 28, {
    size: 16,
    color: C.muted,
    align: "center",
  });
  addText(slide, "s5-cleaning", "Cleaning issues: 135,080 missing customer rows, 10,624 negative quantity rows, 2,515 zero-price rows.", 680, 570, 530, 60, {
    size: 18,
    color: C.ink,
  });
}

function slide6(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 6, "Proposed system approach", "Development-oriented methodology");
  const steps = [
    ["01", "Dataset understanding", "Inspect fields, transaction dates, countries, customers, invoices, and data quality issues."],
    ["02", "Preprocessing", "Clean invalid rows, handle cancellations, engineer revenue, month, order, customer, and product features."],
    ["03", "Dashboard analytics", "Calculate KPIs and build executive, sales, customer, product, and decision-support views."],
    ["04", "ML and evaluation", "Segment customers, forecast sales, test functions, validate KPIs, and prepare deployment."],
  ];
  steps.forEach((step, i) => {
    const left = 42 + i * 299;
    addShape(slide, `s6-step-${i}`, left, 232, 258, 310, i === 3 ? "#FFE4DA" : C.panel);
    addText(slide, `s6-num-${i}`, step[0], left + 24, 258, 70, 42, { size: 28, bold: true, color: i === 3 ? C.accent : C.ink });
    addText(slide, `s6-title-${i}`, step[1], left + 24, 322, 210, 62, { size: 24, bold: true });
    addText(slide, `s6-body-${i}`, step[2], left + 24, 408, 210, 104, { size: 18, color: C.muted });
  });
  addText(slide, "s6-note", "Method: requirements analysis -> system design -> implementation -> testing -> evaluation -> documentation.", 42, 590, 950, 34, {
    size: 20,
    color: C.muted,
  });
}

function slide7(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 7, "Dashboard modules", "Planned Streamlit application");
  addPanelText(slide, "s7-1", "Executive overview", "Revenue, orders, customers, average order value, monthly trend, top markets, and top products.", 42, 213, 374, 175);
  addPanelText(slide, "s7-2", "Sales analytics", "Sales by month, country, product, customer, quantity, and cancellation handling.", 453, 213, 374, 175);
  addPanelText(slide, "s7-3", "Customer analytics", "RFM analysis, clustering, high-value customers, low-value customers, and purchase frequency.", 864, 213, 374, 175, true);
  addPanelText(slide, "s7-4", "Product analytics", "Best-selling, slow-moving, high-revenue products, demand patterns, and product-level recommendations.", 42, 454, 374, 175);
  addPanelText(slide, "s7-5", "Decision support", "Insight cards, alerts, simple recommendation logic, and business focus suggestions.", 453, 454, 374, 175, true);
  addPanelText(slide, "s7-6", "Forecasting", "Monthly revenue or quantity forecasting with model comparison and evaluation metrics.", 864, 454, 374, 175);
}

function slide8(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 8, "Analytics and ML components", "Segmentation plus forecasting");
  addShape(slide, "s8-chart-panel", 42, 213, 580, 467, C.panel);
  slide.charts.add("line", {
    position: { left: 80, top: 275, width: 505, height: 275 },
    categories: monthlyRevenue.map((row) => row[0]),
    series: [{ name: "Revenue GBP thousands", values: monthlyRevenue.map((row) => row[1]), line: { style: "solid", fill: C.accent, width: 3 } }],
    hasLegend: false,
    yAxis: { numberFormatCode: "0", majorGridlines: { style: "solid", fill: "#D8D8D8", width: 1 }, textStyle: { fill: C.muted, fontSize: 12 } },
    xAxis: { textStyle: { fill: C.muted, fontSize: 11 }, line: { style: "solid", fill: C.rule, width: 1 } },
  });
  addText(slide, "s8-chart-title", "Monthly positive-transaction revenue", 80, 238, 505, 30, {
    size: 20,
    bold: true,
  });
  addText(slide, "s8-chart-foot", "Used as the forecasting input after cleaning and monthly aggregation.", 80, 570, 505, 42, {
    size: 17,
    color: C.muted,
  });
  addPanelText(slide, "s8-rfm", "Customer segmentation", "RFM features: recency, frequency, and monetary value. StandardScaler prepares features before K-Means clustering.", 657, 213, 556, 150, true);
  addMetric(slide, "s8-m1", "MAE", "forecast error metric", 679, 428, 224, 150);
  addMetric(slide, "s8-m2", "RMSE", "forecast error metric", 989, 428, 224, 150);
  addText(slide, "s8-note", "Forecast models: moving average and linear trend baseline, with optional exponential smoothing if feasible.", 657, 602, 556, 38, {
    size: 18,
    color: C.muted,
  });
}

function slide9(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 9, "Decision-support logic", "Turning analytics into actions");
  addPanelText(slide, "s9-a", "Sales alert", "If recent monthly revenue drops below the previous period, highlight a declining sales warning.", 42, 249, 581, 172);
  addPanelText(slide, "s9-b", "Product focus", "If a product has high revenue and strong quantity demand, recommend stock availability and promotion.", 658, 249, 581, 172, true);
  addPanelText(slide, "s9-c", "Customer retention", "If a segment has high monetary value but low recent activity, recommend re-engagement.", 42, 459, 581, 172, true);
  addPanelText(slide, "s9-d", "Low performer review", "If products show low sales and weak recent demand, recommend review, promotion, or removal.", 658, 459, 581, 172);
}

function slide10(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 10, "Evaluation plan", "How the prototype will be validated");
  addPanelText(slide, "s10-a", "Functional testing", "Check pages, filters, calculations, charts, navigation, and file loading.", 42, 213, 374, 175);
  addPanelText(slide, "s10-b", "Data validation", "Compare dashboard KPIs against manual Pandas calculations on sample outputs.", 453, 213, 374, 175);
  addPanelText(slide, "s10-c", "Model evaluation", "Use silhouette score for clusters and MAE, RMSE, MAPE where suitable for forecasts.", 864, 213, 374, 175, true);
  addPanelText(slide, "s10-d", "Usability review", "Review clarity of KPIs, chart readability, module flow, and decision-support usefulness.", 42, 454, 374, 175);
  addPanelText(slide, "s10-e", "Deployment check", "Confirm local run, GitHub structure, requirements file, and Streamlit Cloud readiness.", 453, 454, 374, 175);
  addPanelText(slide, "s10-f", "Academic evidence", "Document methodology, results, screenshots, limitations, and viva-ready explanations.", 864, 454, 374, 175);
}

function slide11(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addTitle(slide, 11, "Timeline and expected contribution", "Feasibility");
  const phases = [
    ["Weeks 1-4", "Proposal, literature, dataset study, requirements"],
    ["Weeks 5-8", "Preprocessing, KPIs, executive and sales analytics"],
    ["Weeks 9-12", "RFM, segmentation, product analytics, forecasting"],
    ["Weeks 13-16", "Integration, evaluation, deployment, documentation"],
  ];
  phases.forEach((phase, i) => {
    const left = 42 + i * 299;
    addShape(slide, `s11-phase-${i}`, left, 232, 258, 230, i === 3 ? "#FFE4DA" : C.panel);
    addText(slide, `s11-week-${i}`, phase[0], left + 22, 256, 210, 38, {
      size: 26,
      bold: true,
      color: i === 3 ? C.accent : C.ink,
    });
    addText(slide, `s11-body-${i}`, phase[1], left + 22, 326, 210, 96, { size: 19, color: C.muted });
  });
  addText(slide, "s11-contrib-title", "Expected contribution", 42, 530, 260, 32, {
    size: 24,
    bold: true,
  });
  addText(
    slide,
    "s11-contrib",
    "A practical, low-cost BI decision-support prototype that demonstrates the full path from raw retail transactions to KPIs, segmentation, forecasts, and actionable recommendations.",
    330,
    524,
    830,
    72,
    { size: 22 },
  );
}

function slide12(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  addText(slide, "s12-kicker", "Approval request", 42, 42, 300, 28, { size: 18, bold: true, color: C.muted });
  addText(slide, "s12-title", "Approve proposal to proceed with implementation", 42, 150, 980, 130, {
    size: 54,
    bold: true,
  });
  addText(slide, "s12-summary", "The topic is feasible, ethical, technically aligned with COM4901, and suitable for a development-oriented final year project.", 42, 330, 850, 72, {
    size: 26,
  });
  addShape(slide, "s12-panel", 42, 472, 1196, 120, C.panel);
  addText(slide, "s12-refs-title", "Main sources", 74, 494, 160, 28, { size: 20, bold: true });
  addText(
    slide,
    "s12-refs",
    "Chen, Chiang and Storey (2012) on BI analytics; Few (2006) on dashboard design; UCI Online Retail dataset; Fader, Hardie and Lee (2005) on RFM/CLV; Hyndman and Athanasopoulos (2021) on forecasting.",
    250,
    494,
    930,
    58,
    { size: 18, color: C.ink },
  );
  addSlideNumber(slide, 12);
}

async function main() {
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  await fs.mkdir(LAYOUT_DIR, { recursive: true });
  await fs.mkdir(QA_DIR, { recursive: true });
  await fs.mkdir(path.dirname(OUT), { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  [
    slide1,
    slide2,
    slide3,
    slide4,
    slide5,
    slide6,
    slide7,
    slide8,
    slide9,
    slide10,
    slide11,
    slide12,
  ].forEach((build) => build(presentation));

  const sourceNotes = [
    "Proposal deck source notes",
    "Primary content: docs/research_proposal.md and output/Research_Proposal_BI_Dashboard.docx.",
    "Guideline basis: Final_Year_Research_Project_Guidelines.pdf proposal sections.",
    "Dataset basis: Online Retail.xlsx public UCI Online Retail dataset.",
    "Layout reference: Codex Grid layout library slides 21, 42, 62, 64, 79, and 08.",
    "No external image assets were used.",
  ].join("\n");
  await fs.writeFile(path.join(TMP, "source-notes.txt"), sourceNotes, "utf8");

  const snapshot = await presentation.inspect({
    kind: "slide,textbox,shape,table,chart",
    maxChars: 20000,
  });
  await fs.writeFile(path.join(QA_DIR, "inspect.ndjson"), snapshot.ndjson, "utf8");

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(path.join(PREVIEW_DIR, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(LAYOUT_DIR, `${stem}.layout.json`), await layout.text(), "utf8");
  }

  await writeBlob(path.join(PREVIEW_DIR, "deck-montage.webp"), await presentation.export({ format: "webp", montage: true, scale: 1 }));

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUT);
  console.log(OUT);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
