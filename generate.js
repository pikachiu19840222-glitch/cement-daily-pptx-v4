
// Cement Daily PPTX - fixed company template edition
// 保留現有 GitHub Actions 輸入參數及輸出檔名

const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");

// ====================================
// 1. 固定模板與輸出檔案
// ====================================

// 必須與 GitHub 中的模板檔名完全一致
const TEMPLATE_NAME = "每日施工日報－正式固定模板.pptx";

const TEMPLATE_PATH = path.join(__dirname, TEMPLATE_NAME);
const OUTPUT_PATH = path.join(__dirname, "daily_report.pptx");

// ====================================
// 2. 讀取 GitHub Actions 施工資料
// ====================================

function input(name) {
  return String(process.env[name] ?? "").trim();
}

const data = {
  project_name: input("PROJECT_NAME"),
  report_date: input("REPORT_DATE"),
  contractor: input("CONTRACTOR"),
  workers: input("WORKERS"),
  managers: input("MANAGERS"),
  work_content: input("WORK_CONTENT"),

  // 尚未由扣子提供的欄位
  progress: input("PROGRESS") || "未提供",
  start_date: input("START_DATE") || "未提供",
  planned_end: input("PLANNED_END") || "未提供",
  contract_end: input("CONTRACT_END") || "未提供",
  elapsed_days: input("ELAPSED_DAYS") || "未提供"
};

// ====================================
// 3. XML 特殊字元處理
// ====================================

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// ====================================
// 4. 替換 PPT 模板中的變數
// ====================================

// 此版本適用於目前已確認的公司模板：
// 每個 {{變數}} 均位於同一段 XML 文字中

function replacePlaceholders(xml, values) {
  let result = xml;

  for (const [key, value] of Object.entries(values)) {
    const token = `{{${key}}}`;

    if (!result.includes(token)) {
      throw new Error(
        `模板缺少欄位 ${token}；請確認使用的是正式固定模板`
      );
    }

    result = result.split(token).join(xmlEscape(value));
  }

  // 檢查是否仍有未替換的變數
  const remaining = result.match(/\{\{[a-z_]+\}\}/g);

  if (remaining) {
    throw new Error(
      `尚有未填寫模板欄位：${[
        ...new Set(remaining)
      ].join(", ")}`
    );
  }

  return result;
}

// ====================================
// 5. 主要 PPT 生成程序
// ====================================

async function main() {

  // 檢查必要資料
  const required = [
    "project_name",
    "report_date",
    "contractor",
    "workers",
    "managers",
    "work_content"
  ];

  const missing = required.filter((key) => !data[key]);

  if (missing.length > 0) {
    throw new Error(
      `施工日報缺少必要資料：${missing.join(", ")}`
    );
  }

  // 確認模板存在
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error(
      `找不到模板：${TEMPLATE_NAME}`
    );
  }

  console.log("開始生成每日施工日報...");
  console.log("施工日期:", data.report_date);
  console.log("工程名稱:", data.project_name);

  // 讀取 PPTX
  const zip = await JSZip.loadAsync(
    fs.readFileSync(TEMPLATE_PATH)
  );

  // 取得第一頁投影片 XML
  const slidePath = "ppt/slides/slide1.xml";
  const slide = zip.file(slidePath);

  if (!slide) {
    throw new Error("模板缺少第一頁投影片 XML");
  }

  const original = await slide.async("string");

  // 替換文字變數
  const updated = replacePlaceholders(
    original,
    data
  );

  // 寫回 PPTX
  zip.file(slidePath, updated);

  // 重新封裝 PPTX
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6
    }
  });

  // 輸出檔案
  fs.writeFileSync(OUTPUT_PATH, buffer);

  console.log("PPTX generated successfully");
  console.log("Report date:", data.report_date);
  console.log("Project:", data.project_name);
  console.log("Workers:", data.workers);
  console.log("Managers:", data.managers);
  console.log("Output:", OUTPUT_PATH);
}

// ====================================
// 6. 執行與錯誤處理
// ====================================

main().catch((error) => {
  console.error(
    "PPTX generation failed:",
    error.message
  );

  process.exitCode = 1;
});
