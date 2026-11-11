
/**
 * Cement Daily PPTX V5.3
 * 台泥施工日報正式固定模板
 *
 * 功能：
 * 1. 保留公司原有 PPT 版型
 * 2. 自動填入施工資料
 * 3. 施工項目橫向排列
 * 4. 自動下載兩張施工照片
 * 5. 照片等比例放入左右照片框
 *
 * Node.js 22
 * 套件：jszip
 */

const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");

// =====================================
// 1. 模板及輸出檔案
// =====================================

const TEMPLATE_NAME =
  "每日施工日報－正式固定模板.pptx";

const TEMPLATE_PATH = path.join(
  __dirname,
  TEMPLATE_NAME
);

const OUTPUT_PATH = path.join(
  __dirname,
  "daily_report.pptx"
);

const SLIDE_PATH =
  "ppt/slides/slide1.xml";

const RELS_PATH =
  "ppt/slides/_rels/slide1.xml.rels";

const MAX_IMAGE_BYTES = 12 * 1024 * 1024;

// =====================================
// 2. 取得 GitHub Actions 參數
// =====================================

function input(name) {
  return String(
    process.env[name] ?? ""
  ).trim();
}

const data = {
  project_name: input("PROJECT_NAME"),
  report_date: input("REPORT_DATE"),
  contractor: input("CONTRACTOR"),
  workers: input("WORKERS"),
  managers: input("MANAGERS"),
  work_content: input("WORK_CONTENT"),

  progress: input("PROGRESS") || "未提供",
  start_date: input("START_DATE") || "未提供",
  planned_end: input("PLANNED_END") || "未提供",
  contract_end: input("CONTRACT_END") || "未提供",
  elapsed_days: input("ELAPSED_DAYS") || "未提供"
};

const photoUrls = [
  input("SITE_PHOTO_1_URL"),
  input("SITE_PHOTO_2_URL")
];

// =====================================
// 3. XML 特殊字元處理
// =====================================

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// =====================================
// 4. 替換模板文字變數
// =====================================

function replacePlaceholders(xml, values) {
  let result = xml;

  for (const [key, value] of Object.entries(values)) {
    const token = `{{${key}}}`;

    if (!result.includes(token)) {
      throw new Error(
        `模板缺少欄位：${token}`
      );
    }

    result = result
      .split(token)
      .join(xmlEscape(value));
  }

  const remaining = result.match(
    /\{\{[a-z_]+\}\}/g
  );

  if (remaining) {
    throw new Error(
      "模板尚有未替換欄位：" +
      [...new Set(remaining)].join(", ")
    );
  }

  return result;
}

// =====================================
// 5. 判斷 JPG / PNG 圖片及尺寸
// =====================================

function imageInfo(buffer) {

  // PNG
  const pngSignature = Buffer.from(
    "89504e470d0a1a0a",
    "hex"
  );

  if (
    buffer.length >= 24 &&
    buffer.subarray(0, 8).equals(pngSignature)
  ) {
    const w = buffer.readUInt32BE(16);
    const h = buffer.readUInt32BE(20);

    if (!w || !h) {
      throw new Error("PNG 圖片尺寸無效");
    }

    return {
      ext: "png",
      contentType: "image/png",
      w,
      h
    };
  }

  // JPEG
  if (
    buffer.length > 4 &&
    buffer[0] === 0xff &&
    buffer[1] === 0xd8
  ) {
    let pos = 2;

    const sofMarkers = [
      0xc0, 0xc1, 0xc2, 0xc3,
      0xc5, 0xc6, 0xc7,
      0xc9, 0xca, 0xcb,
      0xcd, 0xce, 0xcf
    ];

    while (pos + 9 < buffer.length) {

      if (buffer[pos] !== 0xff) {
        pos++;
        continue;
      }

      const marker = buffer[pos + 1];

      if (marker === 0xff) {
        pos++;
        continue;
      }

      if (
        marker === 0xd9 ||
        marker === 0xda
      ) {
        break;
      }

      if (
        marker === 0x01 ||
        (marker >= 0xd0 && marker <= 0xd7)
      ) {
        pos += 2;
        continue;
      }

      const length = buffer.readUInt16BE(
        pos + 2
      );

      if (
        length < 2 ||
        pos + 2 + length > buffer.length
      ) {
        break;
      }

      if (sofMarkers.includes(marker)) {
        const h = buffer.readUInt16BE(
          pos + 5
        );

        const w = buffer.readUInt16BE(
          pos + 7
        );

        if (!w || !h) {
          throw new Error(
            "JPEG 圖片尺寸無效"
          );
        }

        return {
          ext: "jpg",
          contentType: "image/jpeg",
          w,
          h
        };
      }

      pos += 2 + length;
    }
  }

  throw new Error(
    "目前僅支援 JPG 或 PNG 照片"
  );
}

// =====================================
// 6. 從扣子圖片 URL 下載照片
// =====================================

async function downloadImage(url, label) {

  if (!url) {
    console.log(
      `${label}：未提供，保留空白照片框`
    );
    return null;
  }

  let parsed;

  try {
    parsed = new URL(url);
  } catch {
    throw new Error(
      `${label}：圖片網址無效`
    );
  }

  if (parsed.protocol !== "https:") {
    throw new Error(
      `${label}：圖片網址必須使用 HTTPS`
    );
  }

  const controller = new AbortController();

  const timeout = setTimeout(
    () => controller.abort(),
    20000
  );

  try {

    const response = await fetch(
      parsed.href,
      {
        signal: controller.signal
      }
    );

    if (!response.ok) {
      throw new Error(
        `${label}：下載失敗，HTTP ${response.status}。` +
        "請確認圖片連結尚未過期。"
      );
    }

    const declaredSize = Number(
      response.headers.get(
        "content-length"
      ) || 0
    );

    if (declaredSize > MAX_IMAGE_BYTES) {
      throw new Error(
        `${label}：圖片超過 12 MB`
      );
    }

    const chunks = [];
    let size = 0;

    for await (
      const chunk of response.body
    ) {
      size += chunk.length;

      if (size > MAX_IMAGE_BYTES) {
        throw new Error(
          `${label}：圖片超過 12 MB`
        );
      }

      chunks.push(Buffer.from(chunk));
    }

    const buffer = Buffer.concat(
      chunks
    );

    return {
      buffer,
      ...imageInfo(buffer)
    };

  } catch (error) {

    if (error.name === "AbortError") {
      throw new Error(
        `${label}：下載逾時`
      );
    }

    throw error;

  } finally {
    clearTimeout(timeout);
  }
}

// =====================================
// 7. 插入圖片到指定照片框
// =====================================

function addPictureToRectangle(
  xml,
  rectangleName,
  image,
  rid,
  picId,
  picName
) {

  const escapedName = rectangleName.replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&"
  );

  // 找到公司模板內指定矩形
  const pattern = new RegExp(
    `<p:sp>(?:(?!<p:sp>)[\\s\\S])*?` +
    `<p:cNvPr\\b[^>]*\\bname="${escapedName}"` +
    `[^>]*(?:\\/\\>|\\>)` +
    `(?:(?!<\\/p:sp>)[\\s\\S])*?` +
    `<\\/p:sp>`
  );

  const match = xml.match(pattern);

  if (!match) {
    throw new Error(
      `找不到照片框：${rectangleName}`
    );
  }

  const shape = match[0];

  // 取得矩形位置及尺寸
  const pos = shape.match(
    /<a:xfrm[^>]*>[\s\S]*?<a:off x="(\d+)" y="(\d+)"\s*\/>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/
  );

  if (!pos) {
    throw new Error(
      `${rectangleName} 缺少位置資料`
    );
  }

  const x = Number(pos[1]);
  const y = Number(pos[2]);
  const w = Number(pos[3]);
  const h = Number(pos[4]);

  // 照片四周保留少量邊距
  const margin = 20000;

  const maxW = w - margin * 2;
  const maxH = h - margin * 2;

  // 等比例縮放
  const scale = Math.min(
    maxW / image.w,
    maxH / image.h
  );

  const iw = Math.round(
    image.w * scale
  );

  const ih = Math.round(
    image.h * scale
  );

  // 置中定位
  const ix = x + Math.round(
    (w - iw) / 2
  );

  const iy = y + Math.round(
    (h - ih) / 2
  );

  // 建立 PowerPoint 圖片 XML
  const picture = `
<p:pic>
  <p:nvPicPr>
    <p:cNvPr
      id="${picId}"
      name="${xmlEscape(picName)}"
    />
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>

  <p:blipFill>
    <a:blip r:embed="${rid}"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>

  <p:spPr>
    <a:xfrm>
      <a:off x="${ix}" y="${iy}"/>
      <a:ext cx="${iw}" cy="${ih}"/>
    </a:xfrm>

    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;

  // 保留原本矩形作為底色
  return xml.replace(
    shape,
    shape + picture
  );
}

// =====================================
// 8. 增加 PPT 圖片關聯
// =====================================

function addRelationship(
  xml,
  id,
  filename
) {

  const relationship =
    `<Relationship ` +
    `Id="${id}" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
    `Target="../media/${filename}"/>`;

  if (!xml.includes("</Relationships>")) {
    throw new Error(
      "投影片關聯檔格式不正確"
    );
  }

  if (xml.includes(`Id="${id}"`)) {
    throw new Error(
      `圖片關聯 ID 已存在：${id}`
    );
  }

  return xml.replace(
    "</Relationships>",
    relationship + "</Relationships>"
  );
}

// =====================================
// 9. 增加圖片格式宣告
// =====================================

function ensureImageContentType(
  xml,
  extension,
  contentType
) {

  const pattern = new RegExp(
    `<Default\\s+Extension="${extension}"`,
    "i"
  );

  if (pattern.test(xml)) {
    return xml;
  }

  return xml.replace(
    "</Types>",
    `<Default Extension="${extension}" ` +
    `ContentType="${contentType}"/>` +
    "</Types>"
  );
}

// =====================================
// 10. 主要 PPT 產生程序
// =====================================

async function main() {

  console.log(
    "開始生成台泥每日施工日報..."
  );

  // 檢查必要施工資料
  const required = [
    "project_name",
    "report_date",
    "contractor",
    "workers",
    "managers",
    "work_content"
  ];

  const missing = required.filter(
    key => !data[key]
  );

  if (missing.length > 0) {
    throw new Error(
      "缺少必要資料：" +
      missing.join(", ")
    );
  }

  // 檢查公司模板
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error(
      `找不到公司模板：${TEMPLATE_NAME}`
    );
  }

  console.log(
    "施工日期：",
    data.report_date
  );

  console.log(
    "工程名稱：",
    data.project_name
  );

  // 確認兩張照片是否提供
  console.log(
    "施工照片狀態：",
    photoUrls.map(
      url => url ? "已提供" : "未提供"
    ).join(" / ")
  );

  // 下載兩張照片
  const images = await Promise.all(
    photoUrls.map(
      (url, index) =>
        downloadImage(
          url,
          `施工照片 ${index + 1}`
        )
    )
  );

  // 讀取公司正式模板
  const zip = await JSZip.loadAsync(
    fs.readFileSync(TEMPLATE_PATH)
  );

  const slideFile = zip.file(
    SLIDE_PATH
  );

  if (!slideFile) {
    throw new Error(
      "模板缺少第一頁投影片"
    );
  }

  // 施工內容改為橫向排列
  const formattedData = {
    ...data,

    work_content: data.work_content
      .split(/\r?\n/)
      .map(item => item.trim())
      .filter(Boolean)
      .join("   ")
  };

  // 自動替換施工文字
  let slideXml = replacePlaceholders(
    await slideFile.async("string"),
    formattedData
  );

  // 讀取 PPT 關聯檔
  const relsFile = zip.file(
    RELS_PATH
  );

  if (!relsFile) {
    throw new Error(
      "模板缺少投影片關聯檔"
    );
  }

  let relsXml = await relsFile.async(
    "string"
  );

  // 讀取內容類型檔
  const typesFile = zip.file(
    "[Content_Types].xml"
  );

  if (!typesFile) {
    throw new Error(
      "模板缺少內容類型檔"
    );
  }

  let typesXml = await typesFile.async(
    "string"
  );

  // 計算下一個圖片關聯 ID
  const existingIds = [
    ...relsXml.matchAll(
      /\bId="rId(\d+)"/g
    )
  ].map(
    match => Number(match[1])
  );

  let nextRid =
    Math.max(0, ...existingIds) + 1;

  // 正式模板照片框名稱
  const photoFrames = [
    "Rectangle 27",
    "Rectangle 30"
  ];

  // 將照片嵌入 PPT
  for (
    let i = 0;
    i < images.length;
    i++
  ) {

    const image = images[i];

    if (!image) {
      continue;
    }

    const filename =
      `construction_photo_${i + 1}.` +
      image.ext;

    const rid = `rId${nextRid++}`;

    // 加入照片檔案
    zip.file(
      `ppt/media/${filename}`,
      image.buffer
    );

    // 放入指定矩形
    slideXml = addPictureToRectangle(
      slideXml,
      photoFrames[i],
      image,
      rid,
      200 + i,
      `施工照片 ${i + 1}`
    );

    // 增加圖片關聯
    relsXml = addRelationship(
      relsXml,
      rid,
      filename
    );

    // 增加圖片格式
    typesXml = ensureImageContentType(
      typesXml,
      image.ext,
      image.contentType
    );

    console.log(
      `施工照片 ${i + 1}：嵌入成功，` +
      `${image.w} × ${image.h}`
    );
  }

  // 寫回 PPT 內部檔案
  zip.file(
    SLIDE_PATH,
    slideXml
  );

  zip.file(
    RELS_PATH,
    relsXml
  );

  zip.file(
    "[Content_Types].xml",
    typesXml
  );

  // 重新封裝 PPTX
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6
    }
  });

  // 輸出檔案
  fs.writeFileSync(
    OUTPUT_PATH,
    buffer
  );

  console.log(
    "PPTX generated successfully"
  );

  console.log(
    "Output:",
    OUTPUT_PATH
  );
}

// =====================================
// 11. 執行及錯誤處理
// =====================================

main().catch(error => {

  console.error(
    "PPTX generation failed:",
    error.message
  );

  process.exitCode = 1;
});
