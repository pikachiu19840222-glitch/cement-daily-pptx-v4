
const pptxgen = require("pptxgenjs");

// ====================================
// 1. 讀取 GitHub Actions 傳入的施工資料
// ====================================

function getInput(name) {
  return String(process.env[name] || "").trim();
}

const reportDate = getInput("REPORT_DATE");
const projectName = getInput("PROJECT_NAME");
const contractor = getInput("CONTRACTOR");
const workers = getInput("WORKERS");
const managers = getInput("MANAGERS");
const workContent = getInput("WORK_CONTENT");

// ====================================
// 2. 檢查必要欄位
// ====================================

const requiredFields = {
  "施工日期": reportDate,
  "工程名稱": projectName,
  "施工單位": contractor,
  "施工人數": workers,
  "管理人數": managers,
  "施工內容": workContent
};

const missingFields = Object.entries(requiredFields)
  .filter(([key, value]) => !value)
  .map(([key]) => key);

if (missingFields.length > 0) {
  console.error(
    "施工日報資料缺漏：" + missingFields.join("、")
  );
  process.exit(1);
}

// ====================================
// 3. 建立 PPT
// ====================================

const pptx = new pptxgen();

pptx.layout = "LAYOUT_WIDE";
pptx.author = "Cement Daily PPTX V4";
pptx.subject = "每日施工日報";
pptx.title = projectName + " 每日施工日報";
pptx.lang = "zh-TW";

pptx.theme = {
  headFontFace: "Microsoft JhengHei",
  bodyFontFace: "Microsoft JhengHei",
  lang: "zh-TW"
};

// ====================================
// 4. 建立投影片
// ====================================

const slide = pptx.addSlide();

slide.background = {
  color: "FFFFFF"
};

// 主標題
slide.addText("水泥廠每日施工日報", {
  x: 0.5,
  y: 0.3,
  w: 12.3,
  h: 0.6,
  fontFace: "Microsoft JhengHei",
  fontSize: 28,
  bold: true,
  color: "17365D",
  margin: 0
});

// 標題下方分隔線
slide.addShape(pptx.ShapeType.line, {
  x: 0.5,
  y: 1.05,
  w: 12.3,
  h: 0,
  line: {
    color: "4472C4",
    width: 1.5
  }
});

// ====================================
// 5. 施工基本資料
// ====================================

function addInfo(label, value, y, options = {}) {
  slide.addText(label + "：" + value, {
    x: 0.7,
    y: y,
    w: 11.9,
    h: options.h || 0.42,
    fontFace: "Microsoft JhengHei",
    fontSize: options.fontSize || 17,
    color: "222222",
    margin: 0,
    breakLine: false,
    valign: "mid",
    fit: "shrink"
  });
}

addInfo("施工日期", reportDate, 1.25);

addInfo("工程名稱", projectName, 1.82, {
  h: 0.6,
  fontSize: 16
});

addInfo("施工單位", contractor, 2.55);

addInfo("施工人數", workers + " 人", 3.12);

addInfo("管理人數", managers + " 人", 3.69);

// ====================================
// 6. 每日施工內容
// ====================================

slide.addText("當日施工內容", {
  x: 0.7,
  y: 4.32,
  w: 11.9,
  h: 0.4,
  fontFace: "Microsoft JhengHei",
  fontSize: 19,
  bold: true,
  color: "17365D",
  margin: 0
});

// 施工內容區塊
slide.addShape(pptx.ShapeType.rect, {
  x: 0.65,
  y: 4.85,
  w: 12.0,
  h: 2.05,
  rectRadius: 0,
  line: {
    color: "D9E2F3",
    width: 1
  },
  fill: {
    color: "F7F9FC"
  }
});

// 根據內容長度調整字體
let contentFontSize = 17;

if (workContent.length > 250) {
  contentFontSize = 12;
} else if (workContent.length > 150) {
  contentFontSize = 14;
} else if (workContent.length > 90) {
  contentFontSize = 15;
}

slide.addText(workContent, {
  x: 0.85,
  y: 5.0,
  w: 11.6,
  h: 1.72,
  fontFace: "Microsoft JhengHei",
  fontSize: contentFontSize,
  color: "222222",
  margin: 0.05,
  breakLine: false,
  valign: "top",
  fit: "shrink"
});

// ====================================
// 7. 輸出 PPTX
// ====================================

async function main() {
  try {
    await pptx.writeFile({
      fileName: "daily_report.pptx"
    });

    console.log("PPTX generated successfully");
    console.log("Report date:", reportDate);
    console.log("Project:", projectName);
    console.log("Contractor:", contractor);
    console.log("Output: daily_report.pptx");
  } catch (error) {
    console.error("PPTX generation failed:", error);
    process.exit(1);
  }
}

main();
