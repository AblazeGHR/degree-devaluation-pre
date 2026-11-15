const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

/* ---------------- design tokens ---------------- */
const C = {
  paper: "F2EFE7",
  paper2: "E7E1D3",
  ink: "17160F",
  inkSoft: "3B392D",
  muted: "7A7466",
  rule: "CFC7B6",
  ruleDark: "3A382C",
  seal: "A6392B",
  sealLight: "C87F70",
};
const F = { serif: "Noto Serif SC", sans: "Noto Sans SC", mono: "Consolas" };
const W = 13.333, H = 7.5, M = 0.78, R = W - M;

/* ---------------- script notes ---------------- */
const md = fs.readFileSync(path.join(__dirname, "讲稿.md"), "utf8");
const notes = {};
{
  const parts = md.split(/^## (S\d{2})[^\n]*$/m);
  for (let i = 1; i < parts.length; i += 2) {
    let body = parts[i + 1] || "";
    body = body.replace(/^【[^】]*】\s*$/gm, "").trim();
    notes[parts[i]] = body;
  }
}

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "形势与政策课堂展示";
pres.title = "学历“贬值”：该怎样看待高等教育的价值";

const S = pres.shapes;
let page = 0;

function newSlide(dark = false) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.ink : C.paper };
  page += 1;
  return s;
}

function header(s, dark = false) {
  const line = dark ? C.ruleDark : C.rule;
  s.addShape(S.LINE, { x: M, y: 0.62, w: R - M, h: 0, line: { color: line, width: 0.5 } });
  s.addText("形势与政策  /  2026 秋", {
    x: M, y: 0.28, w: 5.2, h: 0.24, fontFace: F.mono, fontSize: 7.5,
    color: C.muted, charSpacing: 1.4, margin: 0, valign: "middle",
  });
  s.addText(String(page).padStart(2, "0"), {
    x: R - 0.5, y: 0.28, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 7.5,
    color: C.muted, align: "right", margin: 0, valign: "middle",
  });
}

function kicker(s, text, dark = false, y = 0.92) {
  s.addText(text, {
    x: M, y, w: 9, h: 0.26, fontFace: F.mono, fontSize: 9,
    color: dark ? C.sealLight : C.seal, charSpacing: 2.2, margin: 0, valign: "middle",
  });
}

function title(s, text, dark = false, opts = {}) {
  const { y = 1.22, h = 1.05, w = 11.7, size = 31 } = opts;
  s.addText(text, {
    x: M, y, w, h, fontFace: F.serif, fontSize: size, bold: true,
    color: dark ? C.paper : C.ink, lineSpacingMultiple: 1.18, margin: 0,
    valign: "top", charSpacing: 0.5,
  });
}

function footer(s, source, dark = false) {
  const line = dark ? C.ruleDark : C.rule;
  s.addShape(S.LINE, { x: M, y: 6.93, w: R - M, h: 0, line: { color: line, width: 0.5 } });
  s.addText(source || "", {
    x: M, y: 7.0, w: R - M - 0.6, h: 0.26, fontFace: F.sans, fontSize: 7.5,
    color: C.muted, margin: 0, valign: "middle",
  });
}

function seal(s, x, y, size, text, dark = false) {
  const fill = C.seal;
  s.addShape(S.RECTANGLE, { x, y, w: size, h: size, fill: { color: fill }, line: { color: fill, width: 0 }, rotate: -4 });
  s.addText(
    text.split("").map((ch, i, arr) => ({ text: ch, options: { breakLine: i < arr.length - 1 } })),
    { x, y, w: size, h: size, fontFace: F.serif, fontSize: size * 18, bold: true, color: C.paper,
      align: "center", valign: "middle", charSpacing: 2, rotate: -4, margin: 0, lineSpacingMultiple: 0.95 }
  );
}

/* ================= S01 cover ================= */
{
  const s = newSlide();
  s.addShape(S.RECTANGLE, {
    x: 0.34, y: 0.34, w: W - 0.68, h: H - 0.68,
    fill: { color: C.paper, transparency: 100 }, line: { color: C.rule, width: 0.6 },
  });
  s.addText("形势与政策 · 课堂展示", {
    x: M + 0.06, y: 0.85, w: 6, h: 0.3, fontFace: F.mono, fontSize: 9.5,
    color: C.seal, charSpacing: 3, margin: 0,
  });
  s.addText("学历", {
    x: M, y: 1.48, w: 5.4, h: 1.4, fontFace: F.serif, fontSize: 88, bold: true,
    color: C.ink, margin: 0, charSpacing: 2, lineSpacingMultiple: 0.95,
  });
  s.addText("“贬值”", {
    x: M, y: 2.68, w: 6.4, h: 1.45, fontFace: F.serif, fontSize: 88, bold: true,
    color: C.seal, margin: 0, charSpacing: 2, lineSpacingMultiple: 0.95,
  });
  s.addText("该怎样看待高等教育的价值", {
    x: M, y: 4.72, w: 7.6, h: 0.62, fontFace: F.serif, fontSize: 28,
    color: C.inkSoft, margin: 0, charSpacing: 1,
  });
  s.addText([
    { text: "17 MIN", options: { bold: true, color: C.ink } },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "2026.11", options: {} },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "3340 字讲稿", options: {} },
  ], { x: M, y: 6.12, w: 7.2, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, charSpacing: 1.2, margin: 0 });

  const lx = 9.55;
  s.addShape(S.LINE, { x: lx, y: 1.15, w: 0, h: 5.1, line: { color: C.rule, width: 0.5 } });
  const rows = [["1998", "毛入学率 9.8%"], ["2019", "毕业生 834 万"], ["2025", "毛入学率 61.3%"], ["2026", "毕业生 1270 万"]];
  rows.forEach((r, i) => {
    const y = 1.6 + i * 1.12;
    s.addText(r[0], { x: lx + 0.32, y, w: 1.1, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, margin: 0, charSpacing: 1 });
    s.addText(r[1], { x: lx + 0.32, y: y + 0.32, w: 2.55, h: 0.36, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0 });
  });
  seal(s, 11.38, 5.55, 1.05, "重估");
  s.addNotes(notes.S01 || "");
}

/* ================= S02 case ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "01  /  现象");
  s.addText("135", {
    x: M - 0.04, y: 1.35, w: 4.3, h: 2.0, fontFace: F.serif, fontSize: 148, bold: true,
    color: C.ink, margin: 0, lineSpacingMultiple: 0.9,
  });
  s.addText("人", { x: 4.22, y: 2.5, w: 0.6, h: 0.5, fontFace: F.serif, fontSize: 20, color: C.seal, margin: 0 });
  s.addText("2025 年，郑州铁路职业技术学院\n“本升专”招生人数", {
    x: M, y: 3.5, w: 4.9, h: 0.8, fontFace: F.sans, fontSize: 14, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.45,
  });
  s.addText("2022 年 30 人   ·   2024 年 70 人   ·   2025 年 135 人", {
    x: M, y: 4.32, w: 5.2, h: 0.3, fontFace: F.mono, fontSize: 8.5, color: C.muted, margin: 0, charSpacing: 0.6,
  });
  const rx = 6.15;
  s.addText("已经有本科文凭，\n为什么还来读专科？", {
    x: rx, y: 1.42, w: 6.2, h: 1.3, fontFace: F.serif, fontSize: 27, bold: true,
    color: C.ink, margin: 0, lineSpacingMultiple: 1.32,
  });
  s.addText("学校与铁路、轨道企业办订单班。校方说，每年签约国企、医院、行业龙头的毕业生接近七成。", {
    x: rx, y: 2.92, w: 6.1, h: 0.95, fontFace: F.sans, fontSize: 13, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.68,
  });
  s.addText("多读两年，换一个更稳的去向。", {
    x: rx, y: 3.98, w: 5.9, h: 0.34, fontFace: F.serif, fontSize: 15, color: C.seal, margin: 0,
  });
  s.addShape(S.LINE, { x: 5.25, y: 1.32, w: 0, h: 4.62, line: { color: C.rule, width: 0.5 } });
  s.addShape(S.RECTANGLE, { x: rx, y: 4.62, w: 6.05, h: 1.32, fill: { color: C.paper2 }, line: { color: C.paper2, width: 0 } });
  s.addShape(S.RECTANGLE, { x: rx, y: 4.62, w: 0.055, h: 1.32, fill: { color: C.seal }, line: { color: C.seal, width: 0 } });
  s.addText("2025 年三个专业各招 45 人：\n铁道交通运营管理、动车组检修技术、\n高速铁路综合维修技术。学制两年。", {
    x: rx + 0.34, y: 4.76, w: 5.4, h: 1.1, fontFace: F.sans, fontSize: 13, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
  });
  seal(s, 11.75, 1.35, 0.92, "实录");
  footer(s, "来源：环球网 2025 年报道；郑州铁路职业技术学院招生信息。");
  s.addNotes(notes.S02 || "");
}

/* ================= S03 three sayings ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "01  /  现象");
  title(s, "三种说法，一个共同的毛病", false, { size: 31, h: 0.6 });

  const rows = [
    ["01", "“大学毕业还不是送外卖。”"],
    ["02", "“本科满街走，硕士多如狗。”"],
    ["03", "“花四年、十几万，起薪不如去打工。”"],
  ];
  rows.forEach((r, i) => {
    const y = 2.05 + i * 1.12;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.22, w: 0.54, h: 0.26, fontFace: F.mono, fontSize: 9, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.42, y: y + 0.14, w: 10.6, h: 0.55, fontFace: F.serif, fontSize: 24, color: C.inkSoft, margin: 0 });
  });

  s.addText("三种说法都能找到一个例子，但例子说明不了整体。", {
    x: M, y: 5.45, w: 11.5, h: 0.4, fontFace: F.serif, fontSize: 18, color: C.inkSoft, margin: 0,
  });
  s.addText("它们都把“文凭的行情”，当成了“教育的价值”。", {
    x: M, y: 5.95, w: 11.5, h: 0.45, fontFace: F.serif, fontSize: 19, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "说明：三种说法为课堂归纳的常见论调，不是统计结论。");
  s.addNotes(notes.S03 || "");
}

/* ================= S04 expansion ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "01  /  现象");
  title(s, "文凭不再稀缺，是从扩招开始的", false, { size: 31 });

  const baseY = 5.72, maxH = 3.18, x0 = 0.98, bw = 0.58, gap = 0.31;
  const data = [["2019", 834], ["2020", 874], ["2021", 909], ["2022", 1076], ["2023", 1158], ["2024", 1179], ["2025", 1222], ["2026", 1270]];
  const scale = maxH / 1270;
  s.addShape(S.LINE, { x: x0 - 0.12, y: baseY, w: 6.95, h: 0, line: { color: C.ink, width: 0.9 } });
  data.forEach((d, i) => {
    const x = x0 + i * (bw + gap);
    const h = d[1] * scale;
    const last = i === data.length - 1;
    s.addShape(S.RECTANGLE, { x, y: baseY - h, w: bw, h, fill: { color: last ? C.seal : "D9D2C2" }, line: { color: last ? C.seal : "D9D2C2", width: 0 } });
    s.addText(String(d[1]), { x: x - 0.16, y: baseY - h - 0.31, w: bw + 0.32, h: 0.24, fontFace: F.mono, fontSize: 8, color: last ? C.seal : C.muted, align: "center", margin: 0, bold: last });
    s.addText(d[0], { x: x - 0.16, y: baseY + 0.1, w: bw + 0.32, h: 0.22, fontFace: F.mono, fontSize: 8, color: C.muted, align: "center", margin: 0 });
  });
  s.addText("全国普通高校毕业生规模（万人）", {
    x: x0 - 0.12, y: baseY + 0.46, w: 5.4, h: 0.24, fontFace: F.sans, fontSize: 9, color: C.muted, margin: 0,
  });

  const rx = 8.42;
  s.addShape(S.LINE, { x: rx - 0.32, y: 1.95, w: 0, h: 3.55, line: { color: C.rule, width: 0.5 } });
  s.addText("高等教育毛入学率", { x: rx, y: 1.92, w: 4.0, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.muted, charSpacing: 1.5, margin: 0 });
  s.addText("9.8%", { x: rx, y: 2.15, w: 3.4, h: 0.72, fontFace: F.serif, fontSize: 42, bold: true, color: C.inkSoft, margin: 0 });
  s.addText("1998 年", { x: rx, y: 2.8, w: 1.6, h: 0.24, fontFace: F.sans, fontSize: 9, color: C.muted, margin: 0 });
  s.addText("↓", { x: rx, y: 3.04, w: 0.5, h: 0.28, fontFace: F.sans, fontSize: 15, color: C.seal, margin: 0 });
  s.addText("61.3%", { x: rx, y: 3.24, w: 3.6, h: 0.82, fontFace: F.serif, fontSize: 48, bold: true, color: C.seal, margin: 0 });
  s.addText("2025 年 · 是 1998 年的 6.3 倍", { x: rx, y: 4.06, w: 3.9, h: 0.26, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0 });
  s.addText("1999 年高校扩招，2019 年高职再扩招。2022 届单年多出 167 万人，是后一次的滞后效应。", {
    x: rx, y: 4.56, w: 4.0, h: 1.25, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });
  footer(s, "来源：教育部《2025 年全国教育事业发展统计公报》；教育部历年毕业生规模数据。");
  s.addNotes(notes.S04 || "");
}

/* ================= S05 scarcity vs value ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "稀缺性下降，不等于价值下降", false, { size: 31, h: 0.6 });

  s.addText("稀缺性", { x: M, y: 1.98, w: 3.0, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, charSpacing: 2.4, margin: 0 });
  s.addText("在下降", { x: M, y: 2.34, w: 2.6, h: 0.72, fontFace: F.serif, fontSize: 36, bold: true, color: C.inkSoft, margin: 0, charSpacing: 2 });
  s.addText("文凭的筛选作用在减弱。60% 的毛入学率，意味着一张文凭不再能把大多数人挡在门外。", {
    x: M, y: 3.32, w: 4.35, h: 1.5, fontFace: F.sans, fontSize: 13.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addText("≠", { x: 5.95, y: 2.35, w: 1.5, h: 1.7, fontFace: F.serif, fontSize: 118, bold: true, color: C.seal, align: "center", margin: 0 });

  s.addText("价值", { x: 7.72, y: 1.98, w: 3.0, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.seal, charSpacing: 2.4, margin: 0 });
  s.addText("仍在", { x: 7.72, y: 2.34, w: 2.6, h: 0.72, fontFace: F.serif, fontSize: 36, bold: true, color: C.seal, margin: 0, charSpacing: 2 });
  s.addText("训练、思维、判断、人脉。它们不会因为文凭发得多就消失，也不会因为第一份工作起薪低就归零。", {
    x: 7.72, y: 3.32, w: 4.55, h: 1.5, fontFace: F.sans, fontSize: 13.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addText("贬值的是文凭的筛选功能，不是教育的能力功能。", {
    x: M, y: 5.5, w: 11.4, h: 0.45, fontFace: F.serif, fontSize: 20, bold: true, color: C.ink, margin: 0,
  });
  s.addText("一个直接的检验：把文凭拿掉，你身上还剩下什么？", {
    x: M, y: 6.15, w: 11.0, h: 0.32, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "概念：稀缺性（scarcity）与价值（value）不是同一个变量。");
  s.addNotes(notes.S05 || "");
}

/* ================= S06 demand ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "就业难，先看结构", false, { size: 31, h: 0.6 });

  const cards = [
    { x: M, w: 5.62, num: "18.0%", label: "16—24 岁城镇调查失业率\n（不含在校生，2026 年 9 月）", accent: true },
    { x: 6.92, w: 5.62, num: "−16%", label: "22—25 岁在 AI 高暴露职业中的\n相对就业变化（Stanford, 2025）", accent: false },
  ];
  cards.forEach((c) => {
    s.addText(c.num, { x: c.x, y: 2.2, w: c.w, h: 1.28, fontFace: F.serif, fontSize: 82, bold: true, color: c.accent ? C.seal : C.ink, margin: 0, charSpacing: 1 });
    s.addText(c.label, { x: c.x, y: 3.72, w: c.w - 0.2, h: 0.72, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  });
  s.addText("AI 先替掉的是刚上手的活：写报告、做表格、查资料、跑流程。这些过去正好是毕业生的第一份工作。", {
    x: M, y: 4.82, w: 11.5, h: 0.8, fontFace: F.serif, fontSize: 17, color: C.ink, margin: 0, lineSpacingMultiple: 1.5,
  });
  s.addText("同一研究：26 岁以上同职业就业 +6%～9%，调整主要走“少招人”而非裁员。耶鲁预算实验室认为总体证据尚不充分，结论存在争议。", {
    x: M, y: 5.95, w: 11.4, h: 0.62, fontFace: F.sans, fontSize: 10, color: C.muted, margin: 0, lineSpacingMultiple: 1.55,
  });
  footer(s, "来源：国家统计局（2026-10-21）；Brynjolfsson, Chandar & Chen, Canaries in the Coal Mine, Stanford Digital Economy Lab, 2025-11。");
  s.addNotes(notes.S06 || "");
}

/* ================= S07 mismatch evidence ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "2024 届，大专的 offer 率高于本科", false, { size: 31, h: 0.58 });

  const baseY = 5.45, maxH = 2.85, x0 = 1.15, bw = 1.12, gap = 0.78;
  const bars = [["大专", 56.6], ["本科", 45.4], ["硕博", 44.4]];
  s.addShape(S.LINE, { x: x0 - 0.2, y: baseY, w: 3 * bw + 2 * gap + 0.4, h: 0, line: { color: C.ink, width: 0.9 } });
  bars.forEach((b, i) => {
    const x = x0 + i * (bw + gap);
    const h = (b[1] / 60) * maxH;
    const c = i === 0 ? C.seal : "C9C2B2";
    s.addShape(S.RECTANGLE, { x, y: baseY - h, w: bw, h, fill: { color: c }, line: { color: c, width: 0 } });
    s.addText(b[1] + "%", { x: x - 0.28, y: baseY - h - 0.46, w: bw + 0.56, h: 0.4, fontFace: F.serif, fontSize: 20, bold: true, color: i === 0 ? C.seal : C.inkSoft, align: "center", margin: 0 });
    s.addText(b[0], { x: x - 0.28, y: baseY + 0.12, w: bw + 0.56, h: 0.3, fontFace: F.sans, fontSize: 11, color: C.inkSoft, align: "center", margin: 0 });
  });
  s.addText("智联招聘 · 2024 届 offer 获得率", { x: x0 - 0.2, y: baseY + 0.52, w: 5.6, h: 0.26, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0 });

  const rx = 7.25;
  s.addShape(S.LINE, { x: rx - 0.35, y: 1.92, w: 0, h: 3.55, line: { color: C.rule, width: 0.5 } });
  s.addText("麦可思 · 2025 届毕业去向落实率", { x: rx, y: 1.94, w: 5.2, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.muted, charSpacing: 1.2, margin: 0 });
  s.addText([{ text: "88.1", options: { fontSize: 44, bold: true, color: C.seal } }, { text: "%", options: { fontSize: 22, color: C.seal } }, { text: "  高职", options: { fontSize: 13, color: C.muted } }], { x: rx, y: 2.22, w: 5.3, h: 0.82, fontFace: F.serif, margin: 0 });
  s.addText([{ text: "85.1", options: { fontSize: 44, bold: true, color: C.inkSoft } }, { text: "%", options: { fontSize: 22, color: C.inkSoft } }, { text: "  本科", options: { fontSize: 13, color: C.muted } }], { x: rx, y: 3.12, w: 5.3, h: 0.82, fontFace: F.serif, margin: 0 });
  s.addText("高职第一次反超。", { x: rx, y: 4.02, w: 4.6, h: 0.34, fontFace: F.serif, fontSize: 16, color: C.ink, margin: 0 });
  s.addText("口径提醒：offer 获得率 ≠ 教育部“毕业去向落实率”，两者差十几个百分点。", {
    x: rx, y: 4.5, w: 5.1, h: 0.8, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0, lineSpacingMultiple: 1.55,
  });
  s.addText("学历层次的差别在缩小，专业和产业的匹配在变得重要。", { x: M, y: 6.32, w: 11.4, h: 0.38, fontFace: F.serif, fontSize: 17, bold: true, color: C.ink, margin: 0 });
  footer(s, "来源：智联招聘《2024 中国大学生就业力调研报告》；麦可思《2026 年中国大学生就业报告》。");
  s.addNotes(notes.S07 || "");
}

/* ================= S08 major adjustment ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "专业目录已经在改", false, { size: 31, h: 0.58 });

  s.addText("1.22", { x: M - 0.05, y: 1.9, w: 4.3, h: 1.5, fontFace: F.serif, fontSize: 108, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("万", { x: 3.9, y: 2.7, w: 0.5, h: 0.4, fontFace: F.serif, fontSize: 22, color: C.ink, margin: 0 });
  s.addText("“十四五”期间撤销或停招的\n本科专业布点", { x: M, y: 3.46, w: 4.3, h: 0.72, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  s.addText([{ text: "同期新增 1.02 万个。", options: { breakLine: true } }, { text: "不是少办教育，是换赛道。" }], {
    x: M, y: 4.32, w: 4.4, h: 0.86, fontFace: F.serif, fontSize: 14, bold: true, color: C.seal, margin: 0, lineSpacingMultiple: 1.5,
  });

  const rx = 5.55;
  s.addShape(S.LINE, { x: rx - 0.3, y: 1.98, w: 0, h: 3.45, line: { color: C.rule, width: 0.5 } });
  const items = [
    ["撤销", "信息管理与信息系统、公共事业管理、市场营销、信息与计算科学：供给过剩，和产业脱节。"],
    ["缺口", "数字经济、智能制造相关领域，高技能人才缺口 2025 年已突破千万（人社部）。"],
    ["判断", "一边是毕业生挤在一起，一边是企业招不到人。这不是教育没用，是结构没对上。"],
  ];
  items.forEach((it, i) => {
    const y = 2.02 + i * 1.16;
    s.addShape(S.LINE, { x: rx, y, w: 7.0, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(it[0], { x: rx, y: y + 0.15, w: 1.0, h: 0.32, fontFace: F.serif, fontSize: 16, bold: true, color: C.seal, margin: 0 });
    s.addText(it[1], { x: rx + 0.95, y: y + 0.14, w: 6.0, h: 0.86, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  });
  s.addText("就业难，不是教育没用，是教育结构和产业结构没对上。", {
    x: M, y: 5.95, w: 11.6, h: 0.4, fontFace: F.serif, fontSize: 18, bold: true, color: C.ink, margin: 0,
  });
  footer(s, "来源：教育部“十四五”本科专业布点调整（2026-04-28）；人力资源社会保障部（2026-10-08）。");
  s.addNotes(notes.S08 || "");
}

/* ================= S09 two inequalities ================= */
{
  const s = newSlide(true);
  header(s, true);
  kicker(s, "02  /  小结", true);

  const grid = [
    ["文凭数量增加", "教育价值下降", 1.85],
    ["就业难", "教育无用", 2.98],
  ];
  grid.forEach((g) => {
    s.addText(g[0], { x: M, y: g[2], w: 4.2, h: 0.8, fontFace: F.serif, fontSize: 40, bold: true, color: C.paper, align: "right", margin: 0, charSpacing: 0.5 });
    s.addText("≠", { x: 5.22, y: g[2] + 0.02, w: 0.85, h: 0.8, fontFace: F.serif, fontSize: 44, bold: true, color: C.sealLight, align: "center", margin: 0 });
    s.addText(g[1], { x: 6.3, y: g[2], w: 5.9, h: 0.8, fontFace: F.serif, fontSize: 40, bold: true, color: C.paper, margin: 0, charSpacing: 0.5 });
  });

  s.addText("数量影响的是稀缺性，价值取决于能力。就业难是产业结构、人才供需错配共同造成的，教育只是其中一环。", {
    x: M, y: 4.35, w: 11.2, h: 0.8, fontFace: F.sans, fontSize: 14, color: "CFC9B9", margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addText("两个极端看着相反，其实都把教育当成一张可以立刻兑现的支票。", {
    x: M, y: 5.65, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 17, color: C.paper, margin: 0,
  });
  seal(s, 11.72, 1.55, 1.0, "辨析", true);
  s.addNotes(notes.S09 || "");
}

/* ================= S10 ability & cognition ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  价值");
  title(s, "先看对个人：两样带得走的东西", false, { size: 30, h: 0.6 });

  const cols = [
    { x: M, w: 5.2, t: "能力", d: "把问题拆开、找到信息、形成判断、讲清楚。换一个岗位，这些还能用。" },
    { x: 6.95, w: 5.2, t: "认知", d: "见过不同观点，知道答案不止一个，也知道自己不知道什么" },
  ];
  cols.forEach((c, i) => {
    if (i > 0) s.addShape(S.LINE, { x: 6.45, y: 2.0, w: 0, h: 3.2, line: { color: C.rule, width: 0.5 } });
    s.addText(c.t, { x: c.x, y: 2.28, w: 4.0, h: 0.7, fontFace: F.serif, fontSize: 36, bold: true, color: C.ink, margin: 0, charSpacing: 2 });
    s.addText(c.d, { x: c.x, y: 3.18, w: c.w, h: 1.5, fontFace: F.sans, fontSize: 14, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.8, valign: "top" });
  });

  s.addText("比如给你一个陌生题目，你能在一周内找到可信资料，做出一个能讲十分钟的东西。", {
    x: M, y: 4.28, w: 5.2, h: 0.8, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0, lineSpacingMultiple: 1.6,
  });
  s.addText("比如看到一条热搜，你先问谁说的、样本多大，而不是先站队。", {
    x: 6.95, y: 4.28, w: 5.2, h: 0.8, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0, lineSpacingMultiple: 1.6,
  });
  s.addText("这两样东西，考试考不出来，文凭也证明不了，但它们决定你上手一个新工作的速度。", {
    x: M, y: 5.35, w: 11.5, h: 0.45, fontFace: F.serif, fontSize: 18, color: C.ink, margin: 0,
  });
  s.addText("而且它们不是毕业那天就定型，工作以后还会长。", {
    x: M, y: 6.0, w: 11.0, h: 0.32, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "能力指可迁移能力（transferable skills），不是某一门课的知识点。");
  s.addNotes(notes.S10 || "");
}

/* ================= S11 character & long-term ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  价值");
  title(s, "还有那些不进工资单的东西", false, { size: 31, h: 0.6 });

  s.addText("协作，是你怎么跟人分工、怎么处理有人不干活。公民素养，是你怎么看一个公共事件、怎么跟不同意见的人相处。审美和判断力，是你知道什么是好的，什么只是热闹。", {
    x: M, y: 2.0, w: 11.6, h: 1.5, fontFace: F.serif, fontSize: 21, color: C.ink, margin: 0, lineSpacingMultiple: 1.75,
  });

  s.addShape(S.RECTANGLE, { x: M, y: 4.3, w: R - M, h: 0.92, fill: { color: C.ink }, line: { color: C.ink, width: 0 } });
  const stats = [["−3.7pp", "10 年内死亡概率"], ["−22%", "入狱概率"], ["−80%", "领取社会救助"], ["+18%", "投票率"]];
  stats.forEach((st, i) => {
    const x = M + 0.35 + i * 2.95;
    s.addText(st[0], { x, y: 4.42, w: 1.1, h: 0.34, fontFace: F.serif, fontSize: 17, bold: true, color: C.sealLight, margin: 0 });
    s.addText(st[1], { x: x + 1.02, y: 4.48, w: 1.85, h: 0.26, fontFace: F.sans, fontSize: 9, color: "CFC9B9", margin: 0 });
  });
  s.addText("多受一年教育 · Oreopoulos & Salvanes, 2011（义务教育自然实验，不能外推到大学）", {
    x: M, y: 5.42, w: 11.6, h: 0.26, fontFace: F.sans, fontSize: 8.5, color: C.muted, margin: 0,
  });
  s.addText("所以教育这件事，不能只用第一份工作的工资来算账。", {
    x: M, y: 5.95, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 18, bold: true, color: C.ink, margin: 0,
  });
  footer(s, "来源：Oreopoulos & Salvanes, “Priceless”, Journal of Economic Perspectives, 2011。");
  s.addNotes(notes.S11 || "");
}

/* ================= S12 personal ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "04  /  个人");
  title(s, "那具体能做什么", false, { size: 31, h: 0.58 });

  const rows = [
    ["01", "调整预期，别把学历当座位号", "预期不是降低要求，是把“我应得什么”换成“我能提供什么”。"],
    ["02", "积累可验证的能力", "作品、实习、项目。能打开看的，比说出来的可信。"],
    ["03", "学会用工具", "AI 不是替你读书，是替你干杂活。会不会用，差距会越拉越大。"],
  ];
  rows.forEach((r, i) => {
    const y = 2.05 + i * 1.36;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.3, w: 0.54, h: 0.26, fontFace: F.mono, fontSize: 9, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.42, y: y + 0.26, w: 3.55, h: 0.78, fontFace: F.serif, fontSize: 18, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.35 });
    s.addText(r[2], { x: 5.35, y: y + 0.3, w: 7.0, h: 0.75, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  });

  s.addText("把简历上的“熟悉”，换成“做过”。", {
    x: M, y: 6.2, w: 11.4, h: 0.42, fontFace: F.serif, fontSize: 22, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "这三条不一定对，你可以只挑一条。");
  s.addNotes(notes.S12 || "");
}

/* ================= S13 society ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "04  /  社会");
  title(s, "让人和岗位对上，不只是学生的事", false, { size: 31, h: 0.6 });

  const cols = [
    ["政府", "专业目录动态调整。撤掉供给过剩的，增加产业需要的。", "比如这次 1.22 万个专业点的调整。"],
    ["高校", "专硕扩招、职教高考、产教融合。把实习和就业往前放。", "比如把实习放到大三，而不是毕业前。"],
    ["企业", "有的取消学历门槛，改成技能测试；有的跟学校办订单班。", "比如铁路局和职院合办的订单班。"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 3.95;
    if (i > 0) s.addShape(S.LINE, { x: x - 0.28, y: 1.98, w: 0, h: 2.75, line: { color: C.rule, width: 0.5 } });
    s.addText(String(i + 1).padStart(2, "0"), { x, y: 1.96, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 9, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(c[0], { x, y: 2.26, w: 3.2, h: 0.55, fontFace: F.serif, fontSize: 28, bold: true, color: C.ink, margin: 0 });
    s.addText(c[1], { x, y: 3.05, w: 3.45, h: 1.0, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.75, valign: "top" });
    s.addText(c[2], { x, y: 4.3, w: 3.45, h: 0.7, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0, lineSpacingMultiple: 1.55, valign: "top" });
  });

  s.addText("匹配不是包分配。它只是让信息和渠道更顺。", {
    x: M, y: 5.3, w: 11.4, h: 0.45, fontFace: F.serif, fontSize: 19, color: C.ink, margin: 0,
  });
  s.addText("一个学生再努力，也改不了专业设置和岗位结构。反过来，社会也不能只要求学生适应。", {
    x: M, y: 5.95, w: 11.5, h: 0.35, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "政策依据：教育部专业设置调整、职业教育法、人社部技能人才评价改革。");
  s.addNotes(notes.S13 || "");
}

/* ================= S14 closing ================= */
{
  const s = newSlide(true);
  s.addShape(S.RECTANGLE, {
    x: 0.34, y: 0.34, w: W - 0.68, h: H - 0.68,
    fill: { color: C.ink, transparency: 100 }, line: { color: C.ruleDark, width: 0.6 },
  });
  s.addText("贬值的是文凭的行情，", {
    x: M + 0.1, y: 1.95, w: 11.3, h: 1.0, fontFace: F.serif, fontSize: 58, bold: true,
    color: C.paper, margin: 0, charSpacing: 1,
  });
  s.addText("不是教育的价值", {
    x: M + 0.1, y: 3.1, w: 11.3, h: 1.0, fontFace: F.serif, fontSize: 58, bold: true,
    color: C.sealLight, margin: 0, charSpacing: 1,
  });
  s.addText("行情会涨会跌，价值需要时间。文凭只是进场的一张票，走多远，看你长出了什么本事。", {
    x: M + 0.1, y: 4.55, w: 10.9, h: 0.9, fontFace: F.sans, fontSize: 15,
    color: "CFC9B9", margin: 0, lineSpacingMultiple: 1.75,
  });
  s.addText("数据和出处都在最后一页，欢迎追问。", {
    x: M + 0.1, y: 5.85, w: 11.0, h: 0.36, fontFace: F.serif, fontSize: 14, color: C.paper, margin: 0,
  });
  s.addText("谢谢", { x: R - 2.6, y: 6.55, w: 2.6, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, align: "right", charSpacing: 2, margin: 0 });
  s.addNotes(notes.S14 || "");
}

/* ================= S15 sources ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "附录  /  资料来源");
  title(s, "主要数据与文献", false, { size: 26, h: 0.5, y: 1.22 });

  const items = [
    "教育部《2025 年全国教育事业发展统计公报》；1998 年各级教育毛入学率统计",
    "国家统计局：2026 年 9 月分年龄组城镇调查失业率（2026-10-21）",
    "智联招聘《2024 中国大学生就业力调研报告》；麦可思《2026 年中国大学生就业报告》",
    "环球网：郑州铁路职业技术学院“本升专”报道（2025）",
    "教育部：“十四五”本科专业布点调整（2026-04-28）",
    "人力资源社会保障部：2025 年数字经济、智能制造高技能人才缺口（2026-10-08）",
    "Brynjolfsson, Chandar & Chen, Canaries in the Coal Mine, Stanford Digital Economy Lab, 2025-11",
    "Becker 1964；Spence 1973；Arrow 1973；Hirsch 1977",
    "Oreopoulos & Salvanes, “Priceless”, Journal of Economic Perspectives, 2011",
  ];
  items.forEach((t, i) => {
    const y = 1.92 + i * 0.5;
    s.addText(String(i + 1).padStart(2, "0"), { x: M, y, w: 0.4, h: 0.26, fontFace: F.mono, fontSize: 7.5, color: C.seal, margin: 0 });
    s.addText(t, { x: 1.32, y, w: 10.9, h: 0.34, fontFace: F.sans, fontSize: 10.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.3 });
  });
  footer(s, "备用页，不计入 17 分钟。所有数字均标注口径，便于答辩追问。");
  s.addNotes(notes.S15 || "");
}

const out = path.join(__dirname, "学历贬值_形势与政策.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written:", out, "slides:", page));
