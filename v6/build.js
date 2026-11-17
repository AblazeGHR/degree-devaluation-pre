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
    x: M, y: 4.62, w: 7.6, h: 0.62, fontFace: F.serif, fontSize: 28,
    color: C.inkSoft, margin: 0, charSpacing: 1,
  });
  s.addText("从“文凭红利”到“能力本位”的理性回归", {
    x: M, y: 5.24, w: 7.6, h: 0.36, fontFace: F.sans, fontSize: 15,
    color: C.muted, margin: 0, charSpacing: 1,
  });
  s.addText([
    { text: "20 MIN", options: { bold: true, color: C.ink } },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "2026.11", options: {} },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "3918 字讲稿", options: {} },
  ], { x: M, y: 6.3, w: 7.2, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, charSpacing: 1.2, margin: 0 });

  /* 组员栏：给上台前填名字用 */
  s.addText("第 ？ 组", {
    x: 8.55, y: 1.05, w: 3.8, h: 0.36, fontFace: F.serif, fontSize: 16, bold: true,
    color: C.seal, margin: 0, charSpacing: 2,
  });
  for (let i = 0; i < 9; i++) {
    const y = 1.62 + i * 0.42;
    s.addText(String(i + 1).padStart(2, "0"), {
      x: 8.55, y: y - 0.04, w: 0.4, h: 0.22, fontFace: F.mono, fontSize: 7,
      color: C.muted, margin: 0,
    });
    s.addShape(S.LINE, { x: 9.02, y: y + 0.14, w: 3.55, h: 0, line: { color: C.rule, width: 0.5 } });
  }
  seal(s, 11.42, 5.62, 0.98, "重估");
  s.addNotes(notes.S01 || "");
}

/* ================= S02 contradiction ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "01  /  现象");
  title(s, "一边是新高，一边是降温", false, { size: 31, h: 0.6 });

  s.addText([
    { text: "1270", options: { fontSize: 104, bold: true, color: C.ink } },
    { text: "万", options: { fontSize: 20, color: C.seal } },
  ], { x: M - 0.04, y: 2.05, w: 4.3, h: 1.45, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("2026 届高校毕业生\n同比增加 48 万，历史新高", {
    x: M, y: 3.68, w: 5.2, h: 0.8, fontFace: F.sans, fontSize: 13, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(S.LINE, { x: 6.45, y: 1.95, w: 0, h: 3.2, line: { color: C.rule, width: 0.5 } });

  const rx = 6.95;
  s.addText("474", { x: rx, y: 2.22, w: 1.9, h: 0.9, fontFace: F.serif, fontSize: 56, bold: true, color: C.inkSoft, margin: 0 });
  s.addText("→", { x: rx + 1.85, y: 2.42, w: 0.6, h: 0.5, fontFace: F.sans, fontSize: 26, color: C.seal, align: "center", margin: 0 });
  s.addText("343", { x: rx + 2.45, y: 2.22, w: 1.9, h: 0.9, fontFace: F.serif, fontSize: 56, bold: true, color: C.seal, margin: 0 });
  s.addText("万", { x: rx + 3.78, y: 2.62, w: 0.4, h: 0.32, fontFace: F.serif, fontSize: 15, color: C.muted, margin: 0 });
  s.addText("考研报名人数：2023 年峰值 → 2026 年\n连续三年下降，少了 131 万", {
    x: rx, y: 3.3, w: 5.3, h: 0.8, fontFace: F.sans, fontSize: 13, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addText("年轻人开始算投入产出账。", {
    x: M, y: 5.55, w: 11.4, h: 0.5, fontFace: F.serif, fontSize: 22, bold: true, color: C.ink, margin: 0,
  });
  s.addText("学历的滤镜在褪色，但褪色的方式值得深究。", {
    x: M, y: 6.15, w: 11.2, h: 0.32, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "来源：教育部 2026 届毕业生规模；中国教育在线 2000—2026 年考研报名数据。");
  s.addNotes(notes.S02 || "");
}

/* ================= S03 the question ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "01  /  问题");
  s.addText([
    { text: "文凭数量增加", options: { color: C.ink } },
    { text: "   ≠   ", options: { color: C.seal } },
    { text: "教育价值下降", options: { color: C.ink } },
  ], {
    x: M, y: 1.85, w: 11.7, h: 0.9, fontFace: F.serif, fontSize: 38, bold: true, margin: 0, charSpacing: 0.5,
  });
  s.addText("那“贬值”的到底是什么？", {
    x: M, y: 2.95, w: 9.0, h: 0.5, fontFace: F.serif, fontSize: 24, color: C.seal, margin: 0,
  });

  const cols = [
    ["01", "贬的是什么", "稀缺性，不是知识"],
    ["02", "不贬的是什么", "能力、认知、综合素质"],
    ["03", "该积累什么", "能验证的本事"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 3.95;
    if (i > 0) s.addShape(S.LINE, { x: x - 0.3, y: 4.05, w: 0, h: 1.02, line: { color: C.rule, width: 0.5 } });
    s.addText(c[0], { x, y: 4.05, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 8.5, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(c[1], { x, y: 4.4, w: 3.3, h: 0.42, fontFace: F.serif, fontSize: 19, bold: true, color: C.ink, margin: 0 });
    s.addText(c[2], { x, y: 4.9, w: 3.4, h: 0.3, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0 });
  });

  s.addText("三个问题，一条线：现象 → 本质 → 行动。", {
    x: M, y: 6.15, w: 11.0, h: 0.32, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "这一页是全场路线图，后面三段依次回答这三个问题。");
  s.addNotes(notes.S03 || "");
}

/* ================= S04 scarcity ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "贬的是稀缺性，不是知识", false, { size: 31, h: 0.6 });

  const baseY = 5.5, maxH = 2.55, x0 = 1.0, bw = 1.15, gap = 0.78;
  const bars = [["1990", 60.88, "本专科招生"], ["1998", 108.36, "本专科招生"], ["2025", 143.85, "研究生招生"]];
  s.addShape(S.LINE, { x: x0 - 0.15, y: baseY, w: 3 * bw + 2 * gap + 0.3, h: 0, line: { color: C.ink, width: 0.9 } });
  bars.forEach((b, i) => {
    const x = x0 + i * (bw + gap);
    const h = (b[1] / 150) * maxH;
    const last = i === 2;
    s.addShape(S.RECTANGLE, { x, y: baseY - h, w: bw, h, fill: { color: last ? C.seal : "D9D2C2" }, line: { color: last ? C.seal : "D9D2C2", width: 0 } });
    s.addText([
      { text: String(b[1]), options: { fontSize: 19, bold: true, color: last ? C.seal : C.inkSoft } },
      { text: " 万", options: { fontSize: 9, color: C.muted } },
    ], { x: x - 0.22, y: baseY - h - 0.44, w: bw + 0.44, h: 0.36, fontFace: F.serif, align: "center", margin: 0 });
    s.addText(b[0], { x: x - 0.22, y: baseY + 0.1, w: bw + 0.44, h: 0.26, fontFace: F.mono, fontSize: 9.5, color: C.muted, align: "center", margin: 0 });
    s.addText(b[2], { x: x - 0.28, y: baseY + 0.38, w: bw + 0.56, h: 0.24, fontFace: F.sans, fontSize: 8.5, color: C.muted, align: "center", margin: 0 });
  });

  const rx = 7.15;
  s.addShape(S.LINE, { x: rx - 0.35, y: 1.95, w: 0, h: 3.55, line: { color: C.rule, width: 0.5 } });
  s.addText("仅“研究生”一项，\n就超过了 1998 年\n本专科招生的总和。", {
    x: rx, y: 1.95, w: 5.1, h: 1.35, fontFace: F.serif, fontSize: 21, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.35,
  });
  s.addText("高等教育从精英化走向普及化，文凭的标签价值被稀释。这是扩容的必然结果，不是教育的失败。", {
    x: rx, y: 3.5, w: 5.0, h: 1.0, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addText("把“文凭不再稀缺”等同于“读书无用”，是逻辑上的偷换。", {
    x: M, y: 6.24, w: 11.6, h: 0.4, fontFace: F.serif, fontSize: 17, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "来源：教育部《全国教育事业发展统计公报》（1990、1998、2025 年招生数据）。");
  s.addNotes(notes.S04 || "");
}

/* ================= S05 structure ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "就业难是结构问题，不是学历问题", false, { size: 30, h: 0.6 });

  s.addText([
    { text: "200", options: { fontSize: 66, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 24, color: C.seal } },
  ], { x: M - 0.04, y: 2.12, w: 3.2, h: 0.9, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("AI 相关职位需求增长\n（招聘平台在招职位同比）", {
    x: M, y: 3.02, w: 4.3, h: 0.62, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.45,
  });

  s.addText([
    { text: "73.65", options: { fontSize: 50, bold: true, color: C.ink } },
    { text: "%", options: { fontSize: 20, color: C.ink } },
  ], { x: M - 0.04, y: 3.72, w: 3.4, h: 0.8, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("具身智能领域新发职位同比增长\n（2025 年，同比口径）", {
    x: M, y: 4.5, w: 4.4, h: 0.62, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.45,
  });

  const bx = 5.2;
  s.addShape(S.RECTANGLE, { x: bx, y: 1.85, w: 7.35, h: 2.95, fill: { color: C.paper2 }, line: { color: C.paper2, width: 0 } });
  s.addText("但另一面", { x: bx + 0.35, y: 2.02, w: 2.4, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.muted, charSpacing: 1.8, margin: 0 });
  s.addText("传统白领岗位的吸纳能力在收缩。", {
    x: bx + 0.35, y: 2.35, w: 6.5, h: 0.4, fontFace: F.serif, fontSize: 19, bold: true, color: C.ink, margin: 0,
  });
  s.addText("岗位在变，能力要求在变。只带着一张文凭、没有匹配新产业需求的能力，自然会感到“贬值”。", {
    x: bx + 0.35, y: 2.9, w: 6.5, h: 0.7, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.65,
  });
  s.addShape(S.RECTANGLE, { x: bx + 0.35, y: 3.72, w: 0.05, h: 0.62, fill: { color: C.seal }, line: { color: C.seal, width: 0 } });
  s.addText("问题不在学历，在于学历背后有没有真东西。", {
    x: bx + 0.62, y: 3.78, w: 6.3, h: 0.5, fontFace: F.serif, fontSize: 15, bold: true, color: C.seal, margin: 0, lineSpacingMultiple: 1.4,
  });

  s.addShape(S.RECTANGLE, { x: M, y: 5.35, w: R - M, h: 0.82, fill: { color: C.ink }, line: { color: C.ink, width: 0 } });
  s.addText([
    { text: "研究生 offer 率 44.4%", options: { color: C.sealLight } },
    { text: "  ＜  ", options: { color: "8B8677" } },
    { text: "本科 45.4%", options: { color: C.paper } },
    { text: "  ＜  ", options: { color: "8B8677" } },
    { text: "大专 56.6%", options: { color: C.paper } },
    { text: "      ｜      宇树科技：不唯学历唯能力", options: { color: "CFC9B9" } },
  ], { x: M + 0.32, y: 5.52, w: 11.6, h: 0.48, fontFace: F.sans, fontSize: 13.5, margin: 0, valign: "middle" });

  footer(s, "来源：招聘平台公开报告；智联招聘 2024 届 offer 获得率；企业用人标准援引自公开报道。");
  s.addNotes(notes.S05 || "");
}

/* ================= S06 three dimensions ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "学历的价值，不止于求职", false, { size: 31, h: 0.6 });

  const cols = [
    ["01", "能力", "系统的训练，让你会判断、会协作、会把复杂问题拆开。"],
    ["02", "认知", "见过不同的解释框架，知道答案不止一个，也知道自己不知道什么。"],
    ["03", "综合素质", "在多元环境里与人共处、自我管理、持续学习。"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 3.95;
    s.addText(c[0], { x, y: 2.0, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 8.5, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(c[1], { x, y: 2.32, w: 3.4, h: 0.55, fontFace: F.serif, fontSize: 27, bold: true, color: C.ink, margin: 0, charSpacing: 2 });
    s.addShape(S.LINE, { x, y: 3.0, w: 3.45, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(c[2], { x, y: 3.18, w: 3.45, h: 0.95, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.65 });
  });

  s.addShape(S.RECTANGLE, { x: M, y: 4.65, w: 0.05, h: 0.95, fill: { color: C.seal }, line: { color: C.seal, width: 0 } });
  s.addText("大学最大的产品，从来不是“就业率”，\n而是一个具备长期成长能力的人。", {
    x: M + 0.3, y: 4.62, w: 11.0, h: 1.0, fontFace: F.serif, fontSize: 20, color: C.ink, margin: 0, lineSpacingMultiple: 1.55,
  });
  s.addText("职业技能可以短期培训；认知框架与人格养成，需要系统性的土壤。", {
    x: M, y: 6.05, w: 11.4, h: 0.32, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0,
  });
  footer(s, "三个维度决定一个人能走多远：能力、认知、综合素质。");
  s.addNotes(notes.S06 || "");
}

/* ================= S07 reverse signal ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "02  /  辨析");
  title(s, "技能价值正在重估", false, { size: 31, h: 0.6 });

  s.addText([
    { text: "84", options: { fontSize: 118, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 24, color: C.seal } },
  ], { x: M - 0.05, y: 2.05, w: 3.05, h: 1.35, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("高职毕业生就业满意度\n首次反超本科生", {
    x: M, y: 3.58, w: 4.0, h: 0.72, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addText("▲ 反超 3 个百分点", {
    x: 4.95, y: 2.72, w: 2.8, h: 0.34, fontFace: F.sans, fontSize: 12, color: C.seal, align: "center", margin: 0,
  });
  s.addShape(S.LINE, { x: 4.95, y: 2.45, w: 2.8, h: 0, line: { color: C.rule, width: 0.5 } });

  s.addText([
    { text: "81", options: { fontSize: 118, bold: true, color: C.inkSoft } },
    { text: "%", options: { fontSize: 24, color: C.inkSoft } },
  ], { x: 7.55, y: 2.05, w: 3.05, h: 1.35, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("本科毕业生就业满意度\n同一调查口径", {
    x: 7.55, y: 3.58, w: 4.0, h: 0.72, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addText("“本升技”“硕升技”：从个案走向常态化。", {
    x: M, y: 4.85, w: 11.4, h: 0.45, fontFace: F.serif, fontSize: 19, bold: true, color: C.ink, margin: 0,
  });
  s.addText("市场正在从“看标签”转向“看实效”。能解决问题的人，无论来自哪种教育路径，都在获得认可。", {
    x: M, y: 5.45, w: 11.3, h: 0.5, fontFace: F.sans, fontSize: 12, color: C.muted, margin: 0, lineSpacingMultiple: 1.5,
  });
  footer(s, "来源：麦可思研究院《中国高职生就业报告》等相关调查。注意：这是满意度，不是薪资。");
  s.addNotes(notes.S07 || "");
}

/* ================= S08 inert credential ================= */
{
  const s = newSlide(true);
  header(s, true);
  kicker(s, "02  /  小结", true);
  s.addText([
    { text: "贬值的是 ", options: { color: C.paper } },
    { text: "“惰性学历”", options: { color: C.sealLight } },
  ], {
    x: M, y: 1.55, w: 11.6, h: 0.95, fontFace: F.serif, fontSize: 44, bold: true, margin: 0, charSpacing: 1,
  });

  s.addText("真正被淘汰的，是“混一张文凭就能躺赢”的旧逻辑；\n真正值钱的，是文凭背后可验证的能力。", {
    x: M, y: 2.95, w: 11.2, h: 1.1, fontFace: F.sans, fontSize: 18, color: "D9D3C3", margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addShape(S.LINE, { x: M, y: 4.55, w: 4.2, h: 0, line: { color: C.seal, width: 1.2 } });
  s.addText("稀缺性会稀释，能力不会。", {
    x: M, y: 4.85, w: 10.5, h: 0.6, fontFace: F.serif, fontSize: 28, bold: true, color: C.paper, margin: 0,
  });

  s.addText("接下来是第三部分：我们能做什么。", {
    x: M, y: 6.05, w: 10.0, h: 0.32, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  seal(s, 11.55, 1.5, 1.0, "小结", true);
  s.addNotes(notes.S08 || "");
}

/* ================= S09 expectation ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  行动");
  title(s, "从“我是什么学历”到“我能解决什么”", false, { size: 29, h: 0.6 });

  s.addText([
    { text: "60", options: { fontSize: 124, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 30, color: C.seal } },
  ], { x: M - 0.06, y: 2.15, w: 3.3, h: 1.7, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("2025 届本科毕业生\n选择地级及以下城市就业", {
    x: M, y: 4.0, w: 4.4, h: 0.8, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
  });

  const rx = 5.55;
  s.addShape(S.LINE, { x: rx - 0.35, y: 1.95, w: 0, h: 2.85, line: { color: C.rule, width: 0.5 } });
  s.addText("一线城市的虹吸效应在减弱。", {
    x: rx, y: 1.98, w: 6.9, h: 0.4, fontFace: F.serif, fontSize: 21, bold: true, color: C.ink, margin: 0,
  });
  s.addText("这不是退而求其次，是机会分布变了。县城、地级市、制造业基地，都在要人。大城市工资高，房租也高；地级市工资低一些，生活成本低更多。这笔账，每个人都可以自己算。", {
    x: rx, y: 2.55, w: 6.8, h: 1.35, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.75,
  });
  s.addText("关键不是机会有没有，是你愿不愿意去、能不能接得住。", {
    x: rx, y: 4.15, w: 6.8, h: 0.4, fontFace: F.serif, fontSize: 15, color: C.seal, margin: 0,
  });

  s.addText("调整预期不是躺平，是把“我应得什么”换成“市场需要什么”。", {
    x: M, y: 5.95, w: 11.5, h: 0.4, fontFace: F.serif, fontSize: 17, bold: true, color: C.ink, margin: 0,
  });
  footer(s, "来源：麦可思 2025 届本科毕业生就业地分布。");
  s.addNotes(notes.S09 || "");
}

/* ================= S10 tracks ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  行动");
  title(s, "看哪些领域在进水", false, { size: 31, h: 0.6 });

  const rows = [
    ["01", "AI 与具身智能", "算法工程师、智能体开发需求爆发；具身智能平均年薪 33 万以上。缺会调模型的，也缺能把模型落到场景里的。"],
    ["02", "先进制造", "工艺、制程工程师需求高速增长；工业自动化、新能源汽车产业链持续扩张。"],
    ["03", "现代服务业", "新职业不断出现；县域经济同样提供成长空间。"],
  ];
  rows.forEach((r, i) => {
    const y = 2.0 + i * 1.28;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.24, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 8.5, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.32, y: y + 0.2, w: 3.5, h: 0.4, fontFace: F.serif, fontSize: 20, bold: true, color: C.ink, margin: 0 });
    s.addText(r[2], { x: 4.95, y: y + 0.2, w: 7.5, h: 0.85, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.6 });
  });

  s.addText("不是所有人都要挤 AI，但要知道水往哪流。", {
    x: M, y: 6.05, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 17, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "选赛道不是追热点，是看需求能不能持续、能力能不能复用。");
  s.addNotes(notes.S10 || "");
}

/* ================= S11 capabilities ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  行动");
  title(s, "积累抗贬值的能力", false, { size: 31, h: 0.6 });

  s.addText([
    { text: "215", options: { fontSize: 56, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 20, color: C.seal } },
  ], { x: M - 0.04, y: 2.02, w: 2.0, h: 0.72, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("要求会 AI 的岗位同比增长", {
    x: 2.72, y: 2.05, w: 5.0, h: 0.34, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0,
  });
  s.addText("AI 正在从一门专业技能，变成职场通用能力。", {
    x: 2.72, y: 2.42, w: 8.2, h: 0.32, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0,
  });

  const rows = [
    ["01", "AI 工具运用", "不是会聊天，是能用它把活干完。"],
    ["02", "场景化解决问题", "实习、项目、竞赛里，那些“做过的事”。"],
    ["03", "跨领域协作与表达", "判断谁可信，把模糊的问题变成可执行的方案。"],
  ];
  rows.forEach((r, i) => {
    const y = 3.05 + i * 0.9;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.2, w: 0.45, h: 0.22, fontFace: F.mono, fontSize: 8, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.25, y: y + 0.16, w: 3.6, h: 0.34, fontFace: F.serif, fontSize: 17, bold: true, color: C.ink, margin: 0 });
    s.addText(r[2], { x: 5.05, y: y + 0.2, w: 7.4, h: 0.34, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0 });
  });

  s.addText("一段扎实的实习，一个完整的项目，一项可展示的硬技能。", {
    x: M, y: 5.85, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 17, color: C.ink, margin: 0,
  });
  s.addText("文凭是入场券，作品集才是通行证。", {
    x: M, y: 6.32, w: 11.4, h: 0.34, fontFace: F.serif, fontSize: 15, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "来源：招聘平台公开报告（AI 技能岗位同比）；能力清单为课堂归纳。");
  s.addNotes(notes.S11 || "");
}

/* ================= S12 mindset ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  行动");
  title(s, "第一份工作，不一定是终身岗位", false, { size: 30, h: 0.6 });

  s.addText([
    { text: "15.9", options: { fontSize: 88, bold: true, color: C.ink } },
    { text: "%", options: { fontSize: 24, color: C.ink } },
  ], { x: M - 0.04, y: 2.2, w: 3.2, h: 1.1, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("应届生离职率", { x: M, y: 3.4, w: 4.0, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });

  s.addText([
    { text: "30.6", options: { fontSize: 88, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 24, color: C.seal } },
  ], { x: 5.35, y: 2.2, w: 3.2, h: 1.1, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("在入职三个月内离职", { x: 5.35, y: 3.4, w: 4.5, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });

  s.addText("这个数不好看，但它正常。第一份工作，本来就不一定是终身岗位。", {
    x: M, y: 4.2, w: 11.5, h: 0.4, fontFace: F.serif, fontSize: 18, color: C.ink, margin: 0,
  });
  s.addText("刚毕业的时候，你对自己、对行业、对岗位的理解，都可能是错的。试错成本最低的时候，就是现在。调整预期不是降低标准，是把试错纳入规划。", {
    x: M, y: 4.85, w: 11.4, h: 0.8, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });
  s.addText("先上场，再调整。", {
    x: M, y: 5.95, w: 11.0, h: 0.5, fontFace: F.serif, fontSize: 24, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "来源：麦可思应届生离职率与离职时间分布。");
  s.addNotes(notes.S12 || "");
}

/* ================= S13 institutions ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "04  /  社会");
  title(s, "制度层面正在变", false, { size: 31, h: 0.6 });

  const stats = [
    ["10", "%", "2026 年本科专业调整比例\n首次突破 10%"],
    ["1.22", "万", "“十四五”撤销或停招的\n本科专业布点"],
    ["16", "条", "人社部稳就业举措\n访企拓岗、先进制造、现代服务"],
  ];
  stats.forEach((st, i) => {
    const x = M + i * 3.95;
    if (i > 0) s.addShape(S.LINE, { x: x - 0.3, y: 2.0, w: 0, h: 2.3, line: { color: C.rule, width: 0.5 } });
    s.addText([
      { text: st[0], options: { fontSize: 60, bold: true, color: i === 1 ? C.seal : C.ink } },
      { text: " " + st[1], options: { fontSize: 18, color: i === 1 ? C.seal : C.ink } },
    ], { x, y: 2.1, w: 3.2, h: 1.0, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
    s.addText(st[2], { x, y: 3.3, w: 3.45, h: 0.8, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.6 });
  });

  s.addText("专业动态调整 · 产教融合 · 就业成为专业设置的重要考量", {
    x: M, y: 5.35, w: 11.5, h: 0.38, fontFace: F.serif, fontSize: 16, color: C.ink, margin: 0,
  });
  s.addText("匹配不是包分配。它做的是把信息打通、把渠道铺好；最后选哪条路，还是自己定。", {
    x: M, y: 5.95, w: 11.4, h: 0.36, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "来源：教育部“十四五”本科专业布点调整；人社部稳就业举措。");
  s.addNotes(notes.S13 || "");
}

/* ================= S14 plural rulers ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "04  /  社会");
  title(s, "评价需要“多元尺”", false, { size: 31, h: 0.6 });

  s.addShape(S.LINE, { x: M, y: 1.95, w: 2.4, h: 0, line: { color: C.seal, width: 0.9 } });
  s.addText("有尺子量学问，也有尺子量手艺。", {
    x: M, y: 2.12, w: 11.4, h: 0.6, fontFace: F.serif, fontSize: 28, bold: true, color: C.ink, margin: 0, charSpacing: 1,
  });
  s.addText("—— 人民日报评论", {
    x: M, y: 2.78, w: 5.0, h: 0.3, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0,
  });

  const cols = [
    ["学问的尺子", "论文、学历、学校", "衡量一个人受过多少系统训练"],
    ["手艺的尺子", "作品、工龄、解决过什么问题", "衡量一个人能把什么事做成"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 5.85;
    s.addShape(S.LINE, { x, y: 3.55, w: 5.3, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(c[0], { x, y: 3.72, w: 5.0, h: 0.36, fontFace: F.serif, fontSize: 19, bold: true, color: C.seal, margin: 0 });
    s.addText(c[1], { x, y: 4.22, w: 5.2, h: 0.32, fontFace: F.sans, fontSize: 12, color: C.ink, margin: 0 });
    s.addText(c[2], { x, y: 4.62, w: 5.2, h: 0.5, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0, lineSpacingMultiple: 1.55 });
  });

  s.addText("问题不在于哪把尺子更高级，而在于我们过去只认一把。", {
    x: M, y: 5.65, w: 11.4, h: 0.42, fontFace: F.serif, fontSize: 18, bold: true, color: C.ink, margin: 0,
  });
  s.addText("社会的进步，不在于让所有人都去读研，而在于让不同禀赋的人都有被看见的通道。", {
    x: M, y: 6.2, w: 11.4, h: 0.34, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0,
  });
  footer(s, "引文：人民日报评论。“多元尺”指评价标准不应只有学历一项。");
  s.addNotes(notes.S14 || "");
}

/* ================= S15 closing ================= */
{
  const s = newSlide(true);
  s.addShape(S.RECTANGLE, {
    x: 0.34, y: 0.34, w: W - 0.68, h: H - 0.68,
    fill: { color: C.ink, transparency: 100 }, line: { color: C.ruleDark, width: 0.6 },
  });
  s.addText("学历没有作废。", {
    x: M + 0.1, y: 1.85, w: 11.3, h: 1.0, fontFace: F.serif, fontSize: 52, bold: true,
    color: C.paper, margin: 0, charSpacing: 1,
  });
  s.addText("只是不再承诺保底。", {
    x: M + 0.1, y: 2.95, w: 11.3, h: 1.0, fontFace: F.serif, fontSize: 52, bold: true,
    color: C.sealLight, margin: 0, charSpacing: 1,
  });
  s.addText("真正保底的，是你持续长出能力的速度。", {
    x: M + 0.1, y: 4.4, w: 10.9, h: 0.5, fontFace: F.serif, fontSize: 20,
    color: C.paper, margin: 0,
  });
  s.addText("高等教育的价值，不在证书的厚度，在它是否让你有了面对不确定性的底气。", {
    x: M + 0.1, y: 5.1, w: 10.9, h: 0.4, fontFace: F.sans, fontSize: 13,
    color: "CFC9B9", margin: 0,
  });
  s.addText("谢谢", { x: R - 2.6, y: 6.55, w: 2.6, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, align: "right", charSpacing: 2, margin: 0 });
  s.addNotes(notes.S15 || "");
}

/* ================= S16 sources ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "附录  /  资料来源");
  title(s, "主要数据与文献", false, { size: 26, h: 0.5, y: 1.22 });

  const items = [
    "教育部《全国教育事业发展统计公报》：1990、1998、2025 年招生数据；1998、2025 年毛入学率",
    "教育部：2026 届高校毕业生规模；中国教育在线：2000—2026 年考研报名人数",
    "智联招聘《2024 中国大学生就业力调研报告》：2024 届 offer 获得率",
    "麦可思研究院《中国高职生就业报告》等：高职/本科就业满意度",
    "麦可思：2025 届本科毕业生就业地分布；应届生离职率与离职时间分布",
    "招聘平台公开报告：AI 相关职位、具身智能、AI 技能岗位同比数据",
    "教育部：“十四五”本科专业布点调整；2026 年本科专业调整比例",
    "人力资源社会保障部：稳就业相关举措（16 条）",
    "人民日报评论：“有尺子量学问，也有尺子量手艺”",
    "企业用人标准：宇树科技“不唯学历唯能力”，援引自公开报道",
  ];
  items.forEach((t, i) => {
    const y = 1.92 + i * 0.46;
    s.addText(String(i + 1).padStart(2, "0"), { x: M, y, w: 0.4, h: 0.24, fontFace: F.mono, fontSize: 7.5, color: C.seal, margin: 0 });
    s.addText(t, { x: 1.32, y, w: 10.9, h: 0.32, fontFace: F.sans, fontSize: 10, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.3 });
  });
  footer(s, "备用页，不计入 20 分钟。部分平台数据为组员整理，答辩时按原始报告口径回答。");
  s.addNotes(notes.S16 || "");
}

const out = path.join(__dirname, "学历贬值_形势与政策.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written:", out, "slides:", page));
