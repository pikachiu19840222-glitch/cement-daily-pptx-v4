
const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();

pptx.layout = "LAYOUT_WIDE";
pptx.author = "Cement Daily PPTX V4";
pptx.subject = "每日施工日報";

let slide = pptx.addSlide();
slide.background = { color: "FFFFFF" };

slide.addText("水泥廠每日施工日報", {
  x: 0.5, y: 0.3, w: 12.3, h: 0.6,
  fontFace: "Microsoft JhengHei",
  fontSize: 28,
  bold: true,
  color: "17365D"
});

slide.addText("施工日期：2026/11/09", {
  x: 0.7, y: 1.3, w: 10, h: 0.4,
  fontFace: "Microsoft JhengHei",
  fontSize: 18
});

slide.addText("施工單位：測試施工單位", {
  x: 0.7, y: 2.0, w: 10, h: 0.4,
  fontFace: "Microsoft JhengHei",
  fontSize: 18
});

slide.addText("施工內容：設備安裝及現場施工", {
  x: 0.7, y: 2.7, w: 11, h: 0.6,
  fontFace: "Microsoft JhengHei",
  fontSize: 18
});

slide.addText("施工人數：20 人", {
  x: 0.7, y: 3.5, w: 10, h: 0.4,
  fontFace: "Microsoft JhengHei",
  fontSize: 18
});

pptx.writeFile({ fileName: "daily_report.pptx" });
