const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  TabStopType, TabStopPosition, Header, Footer, PageNumber, LineRuleType, PageBreak,
} = require("docx");

const md = fs.readFileSync(path.join(__dirname, "讲稿.md"), "utf8");
const lines = md.split(/\r?\n/);

const SERIF = "Noto Serif SC";
const SANS = "Noto Sans SC";
const MONO = "Consolas";
const INK = "17160F";
const SOFT = "3B392D";
const MUTED = "8A8375";
const RULE = "C9C2B2";

const CONTENT_W = 9026; // A4 11906 - 2*1440

/* ---------- parse ---------- */
const sections = [];
let appendix = [];
let inAppendix = false;
let cur = null;
for (const raw of lines) {
  const line = raw.trim();
  if (line.startsWith("## S")) {
    const m = line.match(/^## (S\d{2})\s*(.*?)\s*【([^】]*)】\s*$/);
    inAppendix = line.includes("资料来源");
    if (m && !inAppendix) {
      cur = { id: m[1], title: m[2], time: m[3], body: [] };
      sections.push(cur);
    }
    continue;
  }
  if (inAppendix) {
    if (line.startsWith("- ")) appendix.push(line.slice(2));
    continue;
  }
  if (!cur) continue;
  if (line === "---" || line === "") continue;
  cur.body.push(line);
}

/* ---------- helpers ---------- */
function runsFor(text, size = 24) {
  const parts = text.split(/(【[^】]*】)/).filter(Boolean);
  return parts.map((p) => {
    if (p.startsWith("【")) {
      return new TextRun({ text: p, font: SANS, size: 19, color: MUTED });
    }
    return new TextRun({ text: p, font: SERIF, size, color: INK });
  });
}

function sectionHeading(sec) {
  const num = sec.id.slice(1);
  return new Paragraph({
    spacing: { before: 360, after: 140 },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    keepNext: true,
    children: [
      new TextRun({ text: `${num}\u3000${sec.title}`, font: SANS, size: 26, bold: true, color: INK }),
      new TextRun({ text: `\t${sec.time}`, font: MONO, size: 18, color: MUTED }),
    ],
  });
}

function bodyPara(text) {
  return new Paragraph({
    spacing: { line: 400, lineRule: LineRuleType.EXACT, after: 60 },
    indent: { firstLine: 480 },
    children: runsFor(text),
  });
}

/* ---------- document ---------- */
const doc = new Document({
  styles: {
    default: { document: { run: { font: SERIF, size: 24, color: INK } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "学历“贬值” · 课堂展示讲稿", font: SANS, size: 16, color: MUTED })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "— ", font: SERIF, size: 18, color: MUTED }),
            new TextRun({ children: [PageNumber.CURRENT], font: SERIF, size: 18, color: MUTED }),
            new TextRun({ text: " —", font: SERIF, size: 18, color: MUTED }),
          ],
        })],
      }),
    },
    children: [
      new Paragraph({
        spacing: { before: 240, after: 80 },
        children: [new TextRun({ text: "学历“贬值”", font: SERIF, size: 44, bold: true, color: INK })],
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "怎样看待高等教育的价值", font: SERIF, size: 30, color: SOFT })],
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: "形势与政策 · 课堂展示 · 20 分钟 · 配套 18 页", font: SANS, size: 19, color: MUTED })],
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: "约 3 942 字 ｜ 语速 205 字/分钟 ｜ 版本 v7", font: SANS, size: 19, color: MUTED })],
      }),
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [new TextRun({ text: "【】内为舞台提示，不念出来。全文按 205 字/分钟设计，可按自己的语速增减。", font: SANS, size: 18, color: MUTED })],
      }),
      new Paragraph({
        spacing: { after: 200 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
        children: [new TextRun({ text: "核心句：文凭数量增加，不等于教育价值下降。", font: SERIF, size: 21, color: SOFT })],
      }),
      ...sections.flatMap((sec) => [sectionHeading(sec), ...sec.body.map(bodyPara)]),
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({
        spacing: { before: 120, after: 160 },
        children: [new TextRun({ text: "附　资料来源", font: SANS, size: 26, bold: true, color: INK })],
      }),
      ...appendix.map((t) => new Paragraph({
        spacing: { line: 300, lineRule: LineRuleType.EXACT, after: 80 },
        indent: { left: 480, hanging: 180 },
        children: [new TextRun({ text: "· ", font: SERIF, size: 19, color: MUTED }),
                   new TextRun({ text: t, font: SERIF, size: 19, color: SOFT })],
      })),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "讲稿.docx");
  fs.writeFileSync(out, buf);
  console.log("written", out, buf.length);
});
