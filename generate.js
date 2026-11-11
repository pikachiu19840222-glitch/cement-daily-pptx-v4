
/**
 * 台泥施工日報 generate.js V5.4
 *
 * 保留公司正式 16:9 PPT 模板
 * 施工項目依長度自動換行
 * 施工照片填滿矩形並置中裁切
 *
 * Node.js 22
 * Dependency: jszip
 */

const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");

// =====================================
// 1. 基本設定
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

const SLIDE_PATH = "ppt/slides/slide1.xml";

const RELS_PATH =
  "ppt/slides/_rels/slide1.xml.rels";

const MAX_IMAGE_BYTES = 12 * 1024 * 1024;

const PHOTO_FRAMES = [
  "Rectangle 27",
  "Rectangle 30"
];

// =====================================
// 2. 取得 GitHub Actions 資料
// =====================================

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
// 3. XML 處理
// =====================================

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeRegex(value) {
  return value.replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&"
  );
}

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
      "尚有未填寫欄位：" +
      [...new Set(remaining)].join(", ")
    );
  }

  return result;
}

// =====================================
// 4. 施工項目整理
// =====================================

function normalizeWorkItems(text) {
  return String(text)
    .replace(/\r/g, "")
    .split("\n")
    .map(item => item.trim())
    .filter(Boolean);
}

// 計算文字大約佔用的寬度
// 中文視為全形，英文數字約半形
function textWidth(text) {
  let width = 0;

  for (const char of Array.from(text)) {
    width += /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(char)
      ? 1
      : 0.55;
  }

  return width;
}

// 依文字長度分行，優先保持每項完整
function wrapWorkItems(items, maxWidth) {

  const lines = [];
  let current = "";

  for (const item of items) {

    const next = current
      ? current + "   " + item
      : item;

    if (
      current &&
      textWidth(next) > maxWidth
    ) {
      lines.push(current);
      current = item;
    } else {
      current = next;
    }

    // 單一施工項目太長時，才在項目內換行
    while (
      textWidth(current) > maxWidth
    ) {
      let cut = 0;

      for (
        let i = 1;
        i <= current.length;
        i++
      ) {
        if (
          textWidth(current.slice(0, i)) >
          maxWidth
        ) {
          break;
        }

        cut = i;
      }

      if (cut === 0) cut = 1;

      lines.push(current.slice(0, cut));

      current = current.slice(cut).trim();
    }
  }

  if (current) {
    lines.push(current);
  }

  return lines;
}

// 根據內容量選擇字體及行數
function layoutWorkContent(rawText) {

  const items = normalizeWorkItems(rawText);

  if (!items.length) {
    throw new Error("施工內容不可為空");
  }

  // 對應目前模板的施工內容文字框
  // 以較保守的寬度預留邊距
  const candidates = [
    { fontSize: 1100, width: 65, maxLines: 3 },
    { fontSize: 1000, width: 72, maxLines: 3 },
    { fontSize: 900, width: 80, maxLines: 4 },
    { fontSize: 800, width: 90, maxLines: 4 }
  ];

  for (const option of candidates) {

    const lines = wrapWorkItems(
      items,
      option.width
    );

    if (lines.length <= option.maxLines) {
      return {
        items,
        lines,
        fontSize: option.fontSize
      };
    }
  }

  throw new Error(
    "施工內容超過單頁可容納範圍，" +
    "請縮短描述或改為多頁報告；" +
    "系統不會自動刪除施工項目。"
  );
}

// 將施工內容替換為真正的 PowerPoint 換行
function replaceWorkContent(xml, layout) {

  const token = "{{work_content}}";

  // 找到包含變數的文字 Run
  const pattern =
    /<a:r>(?:(?!<\/a:r>)[\s\S])*?\{\{work_content\}\}(?:(?!<\/a:r>)[\s\S])*?<\/a:r>/;

  const match = xml.match(pattern);

  if (!match) {
    throw new Error(
      "找不到 {{work_content}} 文字物件"
    );
  }

  const originalRun = match[0];

  // 保留模板原有文字格式，
  // 僅調整字體大小與內容
  const makeRun = (text) => {

    let run = originalRun.replace(
      token,
      xmlEscape(text)
    );

    if (/<a:rPr\b/.test(run)) {

      run = run.replace(
        /<a:rPr\b[^>]*>/,
        tag => {

          if (/\bsz="\d+"/.test(tag)) {
            return tag.replace(
              /\bsz="\d+"/,
              `sz="${layout.fontSize}"`
            );
          }

          return tag.replace(
            "<a:rPr",
            `<a:rPr sz="${layout.fontSize}"`
          );
        }
      );
    }

    return run;
  };

  const replacement = layout.lines
    .map(makeRun)
    .join("<a:br/>");

  return xml.replace(
    originalRun,
    replacement
  );
}

// =====================================
// 5. 判斷圖片格式與尺寸
// =====================================

function imageInfo(buffer) {

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
      throw new Error("PNG 尺寸無效");
    }

    return {
      ext: "png",
      mime: "image/png",
      w,
      h
    };
  }

  if (
    buffer.length > 4 &&
    buffer[0] === 0xff &&
    buffer[1] === 0xd8
  ) {

    const sofMarkers = [
      0xc0, 0xc1, 0xc2, 0xc3,
      0xc5, 0xc6, 0xc7,
      0xc9, 0xca, 0xcb,
      0xcd, 0xce, 0xcf
    ];

    let pos = 2;

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
          throw new Error("JPEG 尺寸無效");
        }

        return {
          ext: "jpg",
          mime: "image/jpeg",
          w,
          h
        };
      }

      pos += 2 + length;
    }
  }

  throw new Error(
    "照片格式不支援，請使用 JPG 或 PNG"
  );
}

// =====================================
// 6. 下載施工照片
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
      `${label} URL 格式錯誤`
    );
  }

  if (parsed.protocol !== "https:") {
    throw new Error(
      `${label} 必須使用 HTTPS URL`
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
        `${label} 下載失敗：HTTP ${response.status}。` +
        "請確認扣子圖片 URL 尚未過期。"
      );
    }

    const chunks = [];
    let total = 0;

    for await (
      const chunk of response.body
    ) {

      total += chunk.length;

      if (total > MAX_IMAGE_BYTES) {
        throw new Error(
          `${label} 超過 12 MB`
        );
      }

      chunks.push(Buffer.from(chunk));
    }

    const buffer = Buffer.concat(chunks);

    return {
      buffer,
      ...imageInfo(buffer)
    };

  } catch (error) {

    if (error.name === "AbortError") {
      throw new Error(
        `${label} 下載逾時`
      );
    }

    throw error;

  } finally {
    clearTimeout(timeout);
  }
}

// =====================================
// 7. 找到模板照片框
// =====================================

function findRectangle(xml, name) {

  const pattern = new RegExp(
    `<p:sp>(?:(?!<p:sp>)[\\s\\S])*?` +
    `<p:cNvPr\\b[^>]*\\bname="${escapeRegex(name)}"` +
    `[^>]*>(?:(?!<\\/p:sp>)[\\s\\S])*?` +
    `<\\/p:sp>`
  );

  const match = xml.match(pattern);

  if (!match) {
    throw new Error(
      `找不到照片框：${name}`
    );
  }

  const shape = match[0];

  const pos = shape.match(
    /<a:xfrm[^>]*>[\s\S]*?<a:off x="(\d+)" y="(\d+)"\s*\/>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/
  );

  if (!pos) {
    throw new Error(
      `${name} 缺少位置或尺寸`
    );
  }

  return {
    shape,
    x: Number(pos[1]),
    y: Number(pos[2]),
    w: Number(pos[3]),
    h: Number(pos[4])
  };
}

// =====================================
// 8. 計算圖片置中裁切範圍
// =====================================

function calculateCrop(image, frame) {

  const imageRatio = image.w / image.h;
  const frameRatio = frame.w / frame.h;

  let left = 0;
  let right = 0;
  let top = 0;
  let bottom = 0;

  if (imageRatio > frameRatio) {

    // 圖片較寬：裁切左右兩側
    const keepRatio =
      frameRatio / imageRatio;

    const crop = Math.round(
      (1 - keepRatio) * 50000
    );

    left = crop;
    right = crop;

  } else if (imageRatio < frameRatio) {

    // 圖片較高：裁切上下兩側
    const keepRatio =
      imageRatio / frameRatio;

    const crop = Math.round(
      (1 - keepRatio) * 50000
    );

    top = crop;
    bottom = crop;
  }

  return {
    left,
    right,
    top,
    bottom
  };
}

// =====================================
// 9. 將圖片填滿指定照片框
// =====================================

function insertPicture(
  xml,
  rectangleName,
  image,
  rid,
  picId,
  picName
) {

  const frame = findRectangle(
    xml,
    rectangleName
  );

  const crop = calculateCrop(
    image,
    frame
  );

  // 圖片本身佔滿整個照片框
  // 使用 srcRect 做裁切，不拉伸變形
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

    <a:srcRect
      l="${crop.left}"
      t="${crop.top}"
      r="${crop.right}"
      b="${crop.bottom}"
    />

    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>

  <p:spPr>
    <a:xfrm>
      <a:off
        x="${frame.x}"
        y="${frame.y}"
      />
      <a:ext
        cx="${frame.w}"
        cy="${frame.h}"
      />
    </a:xfrm>

    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;

  return xml.replace(
    frame.shape,
    frame.shape + picture
  );
}

// =====================================
// 10. 增加圖片關聯
// =====================================

function addRelationship(
  xml,
  id,
  filename
) {

  if (xml.includes(`Id="${id}"`)) {
    throw new Error(
      `Relationship ID 重複：${id}`
    );
  }

  const relationship =
    `<Relationship ` +
    `Id="${id}" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
    `Target="../media/${filename}"/>`;

  if (!xml.includes("</Relationships>")) {
    throw new Error(
      "PPT 關聯檔格式錯誤"
    );
  }

  return xml.replace(
    "</Relationships>",
    relationship + "</Relationships>"
  );
}

function addContentType(
  xml,
  ext,
  mime
) {

  const pattern = new RegExp(
    `<Default\\b[^>]*Extension="${ext}"`,
    "i"
  );

  if (pattern.test(xml)) {
    return xml;
  }

  return xml.replace(
    "</Types>",
    `<Default Extension="${ext}" ` +
    `ContentType="${mime}"/>` +
    "</Types>"
  );
}

// =====================================
// 11. 主要執行程序
// =====================================

async function main() {

  console.log(
    "=== Cement Daily PPTX V5.4 ==="
  );

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

  if (missing.length) {
    throw new Error(
      "缺少施工資料：" +
      missing.join(", ")
    );
  }

  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error(
      `找不到模板：${TEMPLATE_NAME}`
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

  // A. 計算施工項目排版
  const layout = layoutWorkContent(
    data.work_content
  );

  console.log(
    `施工項目：${layout.items.length} 項`
  );

  console.log(
    `施工內容排版：${layout.lines.length} 行`
  );

  console.log(
    `施工內容字體：${layout.fontSize / 100} pt`
  );

  // B. 下載照片
  const images = await Promise.all(
    photoUrls.map(
      (url, index) =>
        downloadImage(
          url,
          `施工照片 ${index + 1}`
        )
    )
  );

  // C. 開啟公司 PPT 模板
  const zip = await JSZip.loadAsync(
    fs.readFileSync(TEMPLATE_PATH)
  );

  const slideFile = zip.file(
    SLIDE_PATH
  );

  const relsFile = zip.file(
    RELS_PATH
  );

  const typesFile = zip.file(
    "[Content_Types].xml"
  );

  if (
    !slideFile ||
    !relsFile ||
    !typesFile
  ) {
    throw new Error(
      "PPT 模板內部檔案不完整"
    );
  }

  let slideXml = await slideFile.async(
    "string"
  );

  let relsXml = await relsFile.async(
    "string"
  );

  let typesXml = await typesFile.async(
    "string"
  );

  // D. 施工內容先做多行排版
  slideXml = replaceWorkContent(
    slideXml,
    layout
  );

  // E. 替換其餘施工資料
  const {
    work_content,
    ...otherData
  } = data;

  slideXml = replacePlaceholders(
    slideXml,
    otherData
  );

  // F. 避免與模板現有圖片 ID 衝突
  const relationshipIds = [
    ...relsXml.matchAll(
      /\bId="rId(\d+)"/g
    )
  ].map(
    match => Number(match[1])
  );

  let nextRid =
    Math.max(0, ...relationshipIds) + 1;

  const pictureIds = [
    ...slideXml.matchAll(
      /<p:cNvPr\b[^>]*\bid="(\d+)"/g
    )
  ].map(
    match => Number(match[1])
  );

  let nextPicId =
    Math.max(0, ...pictureIds) + 1;

  // G. 插入施工照片
  for (let i = 0; i < 2; i++) {

    const image = images[i];

    if (!image) {
      continue;
    }

    const filename =
      `construction_photo_${i + 1}.${image.ext}`;

    const rid = `rId${nextRid++}`;

    const picId = nextPicId++;

    zip.file(
      `ppt/media/${filename}`,
      image.buffer
    );

    slideXml = insertPicture(
      slideXml,
      PHOTO_FRAMES[i],
      image,
      rid,
      picId,
      `施工照片 ${i + 1}`
    );

    relsXml = addRelationship(
      relsXml,
      rid,
      filename
    );

    typesXml = addContentType(
      typesXml,
      image.ext,
      image.mime
    );

    console.log(
      `施工照片 ${i + 1}：嵌入成功`
    );

    console.log(
      `原始尺寸：${image.w} x ${image.h}`
    );

    console.log(
      "顯示方式：填滿照片框，置中裁切"
    );
  }

  // H. 寫回 PPT
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

  // I. 輸出 PPTX
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6
    }
  });

  fs.writeFileSync(
    OUTPUT_PATH,
    buffer
  );

  console.log(
    "PPTX generated successfully"
  );

  console.log(
    "輸出檔案：",
    OUTPUT_PATH
  );
}

// =====================================
// 12. 錯誤處理
// =====================================

main().catch(error => {

  console.error(
    "PPTX generation failed:",
    error.message
  );

  process.exitCode = 1;
});
