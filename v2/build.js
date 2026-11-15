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
  white: "FFFFFF",
};
const F = {
  serif: "Noto Serif SC",
  sans: "Noto Sans SC",
  mono: "Consolas",
};
const W = 13.333, H = 7.5, M = 0.78, R = W - M;

/* ---------------- script parsing ---------------- */
const md = fs.readFileSync(path.join(__dirname, "讲稿.md"), "utf8");
const notes = {};
{
  const parts = md.split(/^## (S\d{2})[^\n]*$/m);
  for (let i = 1; i < parts.length; i += 2) {
    const id = parts[i];
    let body = parts[i + 1] || "";
    body = body.replace(/^【[^】]*】\s*$/gm, "").trim();
    notes[id] = body;
  }
}

/* ---------------- deck ---------------- */
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "形势与政策课堂展示";
pres.title = "学历“贬值”：怎样看待高等教育的价值";
pres.subject = "高等教育回报、文凭信号与匹配溢价";

const S = pres.shapes;
let page = 0;

function newSlide(dark = false) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.ink : C.paper };
  page += 1;
  return s;
}

function header(s, dark = false, noPage = false) {
  const line = dark ? C.ruleDark : C.rule;
  s.addShape(S.LINE, { x: M, y: 0.62, w: R - M, h: 0, line: { color: line, width: 0.5 } });
  s.addText("形势与政策  /  2026 秋", {
    x: M, y: 0.28, w: 5.2, h: 0.24, fontFace: F.mono, fontSize: 7.5,
    color: dark ? C.muted : C.muted, charSpacing: 1.4, margin: 0, valign: "middle",
  });
  if (!noPage) {
    s.addText(String(page).padStart(2, "0"), {
      x: R - 0.5, y: 0.28, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 7.5,
      color: dark ? C.muted : C.muted, align: "right", margin: 0, valign: "middle",
    });
  }
}

function kicker(s, text, dark = false, y = 0.92) {
  s.addText(text, {
    x: M, y, w: 9, h: 0.26, fontFace: F.mono, fontSize: 9,
    color: dark ? C.sealLight : C.seal, charSpacing: 2.2, margin: 0, valign: "middle",
  });
}

function title(s, text, dark = false, opts = {}) {
  const { y = 1.22, h = 1.05, w = 11.6, size = 33 } = opts;
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
  const fill = dark ? C.seal : C.seal;
  s.addShape(S.RECTANGLE, {
    x, y, w: size, h: size, fill: { color: fill }, line: { color: fill, width: 0 },
    rotate: -4,
  });
  s.addText(
    text.split("").map((ch, i, arr) => ({ text: ch, options: { breakLine: i < arr.length - 1 } })),
    {
      x, y, w: size, h: size, fontFace: F.serif, fontSize: size * 18, bold: true,
      color: C.paper, align: "center", valign: "middle", charSpacing: 2,
      rotate: -4, margin: 0, lineSpacingMultiple: 0.95,
    }
  );
}

/* ================= S01 cover ================= */
{
  const s = newSlide();
  // archival frame
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

  s.addText("怎样看待高等教育的价值", {
    x: M, y: 4.72, w: 7.6, h: 0.62, fontFace: F.serif, fontSize: 28,
    color: C.inkSoft, margin: 0, charSpacing: 1,
  });

  s.addText([
    { text: "17 MIN", options: { bold: true, color: C.ink } },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "2026.11", options: {} },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "3310 字讲稿", options: {} },
  ], {
    x: M, y: 6.12, w: 7.2, h: 0.3, fontFace: F.mono, fontSize: 9,
    color: C.muted, charSpacing: 1.2, margin: 0,
  });

  // right-side ledger column
  const lx = 9.55;
  s.addShape(S.LINE, { x: lx, y: 1.15, w: 0, h: 5.1, line: { color: C.rule, width: 0.5 } });
  const rows = [
    ["1998", "毛入学率 9.8%"],
    ["2019", "毕业生 834 万"],
    ["2025", "毛入学率 61.3%"],
    ["2026", "毕业生 1270 万"],
  ];
  rows.forEach((r, i) => {
    const y = 1.6 + i * 1.12;
    s.addText(r[0], {
      x: lx + 0.32, y, w: 1.1, h: 0.3, fontFace: F.mono, fontSize: 9,
      color: C.muted, margin: 0, charSpacing: 1,
    });
    s.addText(r[1], {
      x: lx + 0.32, y: y + 0.32, w: 2.55, h: 0.36, fontFace: F.sans, fontSize: 12.5,
      color: C.inkSoft, margin: 0,
    });
  });

  seal(s, 11.38, 5.55, 1.05, "重估");
  s.addNotes(notes.S01 || "");
}

/* ================= S02 hook ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "01  /  一个具体的案例");

  s.addText("135", {
    x: M - 0.04, y: 1.35, w: 4.3, h: 2.0, fontFace: F.serif, fontSize: 148, bold: true,
    color: C.ink, margin: 0, lineSpacingMultiple: 0.9,
  });
  s.addText("人", {
    x: 4.22, y: 2.5, w: 0.6, h: 0.5, fontFace: F.serif, fontSize: 20,
    color: C.seal, margin: 0,
  });

  s.addText("2025 年，郑州铁路职业技术学院\n“本升专”招生人数", {
    x: M, y: 3.5, w: 4.9, h: 0.8, fontFace: F.sans, fontSize: 14,
    color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.45,
  });

  s.addText("2022 年 30 人   ·   2024 年 70 人   ·   2025 年 135 人", {
    x: M, y: 4.32, w: 5.2, h: 0.3, fontFace: F.mono, fontSize: 8.5,
    color: C.muted, margin: 0, charSpacing: 0.6,
  });

  // right column
  const rx = 6.15;
  s.addText("已经有本科文凭，\n为什么还来读专科？", {
    x: rx, y: 1.42, w: 6.2, h: 1.3, fontFace: F.serif, fontSize: 27, bold: true,
    color: C.ink, margin: 0, lineSpacingMultiple: 1.32,
  });
  s.addText("学校与铁路、轨道企业办订单班。校方说，每年签约国企、医院、行业龙头的毕业生接近七成。", {
    x: rx, y: 2.92, w: 6.1, h: 0.95, fontFace: F.sans, fontSize: 13,
    color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.68,
  });
  s.addText("多读两年，换一个更稳的去向。", {
    x: rx, y: 3.98, w: 5.9, h: 0.34, fontFace: F.serif, fontSize: 15,
    color: C.seal, margin: 0,
  });

  s.addShape(S.LINE, { x: 5.25, y: 1.32, w: 0, h: 4.62, line: { color: C.rule, width: 0.5 } });

  s.addShape(S.RECTANGLE, { x: rx, y: 4.62, w: 6.05, h: 1.32, fill: { color: C.paper2 }, line: { color: C.paper2, width: 0 } });
  s.addShape(S.RECTANGLE, { x: rx, y: 4.62, w: 0.055, h: 1.32, fill: { color: C.seal }, line: { color: C.seal, width: 0 } });
  s.addText("2025 年三个专业各招 45 人：\n铁道交通运营管理、动车组检修技术、\n高速铁路综合维修技术。学制两年。", {
    x: rx + 0.34, y: 4.76, w: 5.4, h: 1.1, fontFace: F.sans, fontSize: 13,
    color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
  });

  seal(s, 11.75, 1.35, 0.92, "实录");
  footer(s, "来源：环球网 2025 年报道；郑州铁路职业技术学院招生信息。");
  s.addNotes(notes.S02 || "");
}

/* ================= S03 thesis ================= */
{
  const s = newSlide(true);
  header(s, true);
  kicker(s, "02  /  我的解释", true);

  s.addShape(S.LINE, { x: M, y: 1.42, w: 0, h: 4.5, line: { color: C.seal, width: 2 } });

  s.addText("学历没有失效，", {
    x: M + 0.46, y: 1.62, w: 10.6, h: 0.95, fontFace: F.serif, fontSize: 52, bold: true,
    color: C.paper, margin: 0, charSpacing: 1,
  });
  s.addText("只是定价方式变了", {
    x: M + 0.46, y: 2.62, w: 10.6, h: 1.0, fontFace: F.serif, fontSize: 52, bold: true,
    color: C.sealLight, margin: 0, charSpacing: 1,
  });

  s.addText("1998 年，它能换干部身份、城市户口和分房。2026 年，它主要是一张面试入场券。", {
    x: M + 0.46, y: 4.05, w: 9.6, h: 1.15, fontFace: F.sans, fontSize: 14.5,
    color: "BDB7A7", margin: 0, lineSpacingMultiple: 1.75,
  });

  s.addText("知识没用，和文凭不好使，是两件事。前者是情绪，后者可以查数据。", {
    x: M + 0.46, y: 5.55, w: 9.8, h: 0.4, fontFace: F.serif, fontSize: 17,
    color: C.paper, margin: 0,
  });

  seal(s, 11.62, 5.15, 1.05, "存疑", true);
  s.addNotes(notes.S03 || "");
}

/* ================= S04 supply ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "03  /  供给");
  title(s, "上大学的人，从不到一成变成六成", false, { size: 31 });

  // manual bar chart
  const baseY = 5.72, maxH = 3.18, x0 = 0.98, bw = 0.58, gap = 0.31;
  const data = [
    ["2019", 834], ["2020", 874], ["2021", 909], ["2022", 1076],
    ["2023", 1158], ["2024", 1179], ["2025", 1222], ["2026", 1270],
  ];
  const scale = maxH / 1270;
  s.addShape(S.LINE, { x: x0 - 0.12, y: baseY, w: 6.95, h: 0, line: { color: C.ink, width: 0.9 } });
  data.forEach((d, i) => {
    const x = x0 + i * (bw + gap);
    const h = d[1] * scale;
    const last = i === data.length - 1;
    s.addShape(S.RECTANGLE, {
      x, y: baseY - h, w: bw, h, fill: { color: last ? C.seal : "D9D2C2" },
      line: { color: last ? C.seal : "D9D2C2", width: 0 },
    });
    s.addText(String(d[1]), {
      x: x - 0.16, y: baseY - h - 0.31, w: bw + 0.32, h: 0.24, fontFace: F.mono,
      fontSize: 8, color: last ? C.seal : C.muted, align: "center", margin: 0,
      bold: last,
    });
    s.addText(d[0], {
      x: x - 0.16, y: baseY + 0.1, w: bw + 0.32, h: 0.22, fontFace: F.mono,
      fontSize: 8, color: C.muted, align: "center", margin: 0,
    });
  });
  s.addText("全国普通高校毕业生规模（万人）", {
    x: x0 - 0.12, y: baseY + 0.46, w: 5.4, h: 0.24, fontFace: F.sans, fontSize: 9,
    color: C.muted, margin: 0,
  });

  // right column
  const rx = 8.42;
  s.addShape(S.LINE, { x: rx - 0.32, y: 1.95, w: 0, h: 3.55, line: { color: C.rule, width: 0.5 } });
  s.addText("高等教育毛入学率", {
    x: rx, y: 1.92, w: 4.0, h: 0.26, fontFace: F.mono, fontSize: 8.5,
    color: C.muted, charSpacing: 1.5, margin: 0,
  });
  s.addText("9.8%", {
    x: rx, y: 2.15, w: 3.4, h: 0.72, fontFace: F.serif, fontSize: 42, bold: true,
    color: C.inkSoft, margin: 0,
  });
  s.addText("1998 年", {
    x: rx, y: 2.8, w: 1.6, h: 0.24, fontFace: F.sans, fontSize: 9,
    color: C.muted, margin: 0,
  });
  s.addText("↓", {
    x: rx, y: 3.04, w: 0.5, h: 0.28, fontFace: F.sans, fontSize: 15, color: C.seal, margin: 0,
  });
  s.addText("61.3%", {
    x: rx, y: 3.24, w: 3.6, h: 0.82, fontFace: F.serif, fontSize: 48, bold: true,
    color: C.seal, margin: 0,
  });
  s.addText("2025 年 · 是 1998 年的 6.3 倍", {
    x: rx, y: 4.06, w: 3.9, h: 0.26, fontFace: F.sans, fontSize: 9.5,
    color: C.muted, margin: 0,
  });
  s.addText("1999 年高校扩招，2019 年高职再扩招。2022 届单年多出 167 万人，是后一次的滞后效应。", {
    x: rx, y: 4.56, w: 4.0, h: 1.25, fontFace: F.sans, fontSize: 11.5,
    color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });

  footer(s, "来源：教育部《2025 年全国教育事业发展统计公报》；教育部历年毕业生规模数据。");
  s.addNotes(notes.S04 || "");
}

/* ================= S05 demand ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "04  /  需求");
  title(s, "AI 先替掉的是入门岗位", false, { size: 28, h: 0.62, w: 12.2 });

  const cards = [
    { x: M, w: 5.62, num: "18.0%", label: "16—24 岁城镇调查失业率\n（不含在校生，2026 年 9 月）", accent: true },
    { x: 6.92, w: 5.62, num: "−16%", label: "22—25 岁在 AI 高暴露职业中的\n相对就业变化（Stanford, 2025）", accent: false },
  ];
  cards.forEach((c) => {
    s.addText(c.num, {
      x: c.x, y: 2.2, w: c.w, h: 1.28, fontFace: F.serif, fontSize: 82, bold: true,
      color: c.accent ? C.seal : C.ink, margin: 0, charSpacing: 1,
    });
    s.addText(c.label, {
      x: c.x, y: 3.72, w: c.w - 0.2, h: 0.72, fontFace: F.sans, fontSize: 11.5,
      color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
    });
  });

  s.addText("AI 先替掉的是刚上手的活：写报告、做表格、查资料、跑流程。这些过去正好是毕业生的第一份工作。", {
    x: M, y: 4.82, w: 11.5, h: 0.8, fontFace: F.serif, fontSize: 17,
    color: C.ink, margin: 0,
  });
  s.addText("同一研究：26 岁以上同职业就业 +6%～9%，调整主要走“少招人”而非裁员。耶鲁预算实验室认为总体证据尚不充分，结论存在争议。", {
    x: M, y: 5.8, w: 11.4, h: 0.62, fontFace: F.sans, fontSize: 10,
    color: C.muted, margin: 0, lineSpacingMultiple: 1.55,
  });

  footer(s, "来源：国家统计局（2026-10-21）；Brynjolfsson, Chandar & Chen, Canaries in the Coal Mine, Stanford Digital Economy Lab, 2025-11。");
  s.addNotes(notes.S05 || "");
}

/* ================= S06 mechanism ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "05  /  机制");
  title(s, "为什么人人都上大学，大学反而不值钱了", false, { size: 30, h: 0.62 });

  const steps = ["少数人\n读大学", "学历成为\n筛选标准", "更多人\n被迫读大学", "标准继续\n上移"];
  const bw = 2.52, gapx = 0.46, x0 = 0.92, y0 = 2.02, bh = 1.02;
  steps.forEach((t, i) => {
    const x = x0 + i * (bw + gapx);
    s.addShape(S.RECTANGLE, { x, y: y0, w: bw, h: bh, fill: { color: C.paper2 }, line: { color: C.paper2, width: 0 } });
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.16, y: y0 + 0.1, w: 0.4, h: 0.22, fontFace: F.mono, fontSize: 8,
      color: C.seal, margin: 0,
    });
    s.addText(t, {
      x: x + 0.16, y: y0 + 0.36, w: bw - 0.3, h: 0.58, fontFace: F.serif, fontSize: 14.5,
      bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.3,
    });
    if (i < steps.length - 1) {
      s.addShape(S.LINE, {
        x: x + bw + 0.08, y: y0 + bh / 2, w: gapx - 0.16, h: 0,
        line: { color: C.muted, width: 1, endArrowType: "triangle" },
      });
    }
  });

  // return loop
  const loopL = x0 + 0.12;
  const loopR = x0 + 3 * (bw + gapx) + bw - 0.12;
  s.addShape(S.LINE, {
    x: loopL, y: y0 + bh + 0.42, w: loopR - loopL, h: 0,
    line: { color: C.seal, width: 1.1, dashType: "dash", beginArrowType: "triangle" },
  });
  s.addText("个体理性  →  集体贬值", {
    x: 4.55, y: y0 + bh + 0.62, w: 4.2, h: 0.3, fontFace: F.mono, fontSize: 9,
    color: C.seal, align: "center", charSpacing: 2, margin: 0,
  });

  s.addText("“有些好东西，一旦人人都有，就不再是好东西。”", {
    x: M, y: 4.35, w: 9.9, h: 0.5, fontFace: F.serif, fontSize: 23,
    color: C.ink, margin: 0,
  });
  s.addText("—— 弗雷德·赫希《增长的社会限制》，1977", {
    x: M, y: 4.86, w: 7, h: 0.3, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0,
  });

  s.addText([{ text: "位置性商品（positional goods）", options: { bold: true, color: C.ink } },
  { text: "：它的价值取决于相对位置，而相对位置的总和是零。", options: {} }], {
    x: M, y: 5.42, w: 11.7, h: 0.5, fontFace: F.sans, fontSize: 13.5,
    color: C.inkSoft, margin: 0,
  });

  s.addText("考研报名人数　2023 年 474 万　→　2026 年 343 万　→　2027 年 347 万", {
    x: M, y: 6.05, w: 11.0, h: 0.28, fontFace: F.mono, fontSize: 9,
    color: C.muted, charSpacing: 0.8, margin: 0,
  });
  footer(s, "理论：Fred Hirsch, Social Limits to Growth, 1977；Randall Collins, The Credential Society, 1979。");
  s.addNotes(notes.S06 || "");
}

/* ================= S07 three theories ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "06  /  理论");
  title(s, "教育一边在生产你，一边在证明你", false, { size: 31, h: 0.62 });

  const cols = [
    { n: "01", t: "人力资本", a: "贝克尔，1964", d: "教育提高生产率，所以工资更高。\n解释的是“绝对收益”。你真的变强了。" },
    { n: "02", t: "信号", a: "斯宾塞，1973", d: "文凭是能力的可信信号。教育未必提高生产率，但它证明了你有能力。解释“门槛”。" },
    { n: "03", t: "筛选", a: "阿罗，1973", d: "大学把能力分布切成段，文凭是那个分段标签。解释的是“排序”。" },
  ];
  cols.forEach((c, i) => {
    const x = M + i * 3.78;
    if (i > 0) s.addShape(S.LINE, { x: x - 0.32, y: 1.96, w: 0, h: 2.2, line: { color: C.rule, width: 0.5 } });
    s.addText(c.n, { x, y: 1.94, w: 0.6, h: 0.24, fontFace: F.mono, fontSize: 9, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(c.t, { x, y: 2.22, w: 3.3, h: 0.46, fontFace: F.serif, fontSize: 25, bold: true, color: C.ink, margin: 0 });
    s.addText(c.a, { x, y: 2.72, w: 3.3, h: 0.24, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0 });
    s.addText(c.d, { x, y: 3.06, w: 3.28, h: 1.65, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.68, valign: "top" });
  });

  s.addShape(S.RECTANGLE, { x: M, y: 5.06, w: R - M, h: 1.02, fill: { color: C.ink }, line: { color: C.ink, width: 0 } });
  s.addText("供给一多，先掉价的是“证明”；生产能力的价格掉得慢。", {
    x: M + 0.42, y: 5.26, w: 10.9, h: 0.62, fontFace: F.serif, fontSize: 21, bold: true,
    color: C.paper, margin: 0, charSpacing: 0.5,
  });

  s.addText("卡普兰认为 80% 是信号，多数学者估 20—25%。比例有争议，下面的推论不太受影响。", {
    x: M, y: 6.28, w: 11.6, h: 0.3, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0,
  });

  footer(s, "Becker 1964；Spence 1973；Arrow 1973；Bryan Caplan, The Case Against Education, 2018。");
  s.addNotes(notes.S07 || "");
}

/* ================= S08 premium ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "07  /  反直觉之一");
  title(s, "文凭的整体溢价，还没有消失", false, { size: 29, h: 0.58 });

  s.addText("68%", {
    x: M - 0.05, y: 1.96, w: 4.6, h: 1.72, fontFace: F.serif, fontSize: 132, bold: true,
    color: C.seal, margin: 0, lineSpacingMultiple: 0.9,
  });
  s.addText("美国大学毕业生相对高中毕业生的\n工资溢价（2024 年，中位年薪）", {
    x: M, y: 3.62, w: 4.5, h: 0.86, fontFace: F.sans, fontSize: 12,
    color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.6,
  });
  s.addText("仍接近历史高位", {
    x: M, y: 4.52, w: 3.0, h: 0.32, fontFace: F.mono, fontSize: 9.5,
    color: C.seal, charSpacing: 2, margin: 0,
  });

  // horizontal bars
  const bx = 6.55, baseW = 4.35;
  const bars = [
    { label: "大学毕业生", val: "$80,000", w: baseW, y: 2.3, c: C.seal, strong: true },
    { label: "高中毕业生", val: "$47,000", w: baseW * (47000 / 80000), y: 3.42, c: "C9C2B2", strong: false },
  ];
  bars.forEach((b) => {
    s.addText(b.label, {
      x: bx, y: b.y - 0.32, w: 3.0, h: 0.24, fontFace: F.sans, fontSize: 9.5,
      color: b.strong ? C.ink : C.muted, margin: 0,
    });
    s.addShape(S.RECTANGLE, { x: bx, y: b.y, w: b.w, h: 0.56, fill: { color: b.c }, line: { color: b.c, width: 0 } });
    s.addText(b.val, {
      x: bx + b.w + 0.14, y: b.y + 0.1, w: 1.3, h: 0.36, fontFace: F.serif, fontSize: 16,
      bold: true, color: b.strong ? C.seal : C.inkSoft, margin: 0,
    });
  });

  s.addText("68% 是平均值。计算机和文史哲、名校和普通院校，都平均在里面。", {
    x: bx, y: 4.45, w: 5.7, h: 0.72, fontFace: F.sans, fontSize: 13,
    color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.6,
  });
  s.addText("整体没跌，是内部拉开了。", {
    x: bx, y: 5.28, w: 5.4, h: 0.4, fontFace: F.serif, fontSize: 19, bold: true,
    color: C.ink, margin: 0,
  });

  footer(s, "来源：Federal Reserve Bank of New York, Liberty Street Economics, 2025（2024 年美元，25 岁以上全职全年工作者）。");
  s.addNotes(notes.S08 || "");
}

/* ================= S09 inversion ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "08  /  反直觉之二");
  title(s, "2024 届，大专的 offer 率高于本科", false, { size: 31, h: 0.58 });

  const baseY = 5.45, maxH = 2.85, x0 = 1.15, bw = 1.12, gap = 0.78;
  const bars = [["大专", 56.6], ["本科", 45.4], ["硕博", 44.4]];
  s.addShape(S.LINE, { x: x0 - 0.2, y: baseY, w: 3 * bw + 2 * gap + 0.4, h: 0, line: { color: C.ink, width: 0.9 } });
  bars.forEach((b, i) => {
    const x = x0 + i * (bw + gap);
    const h = (b[1] / 60) * maxH;
    const c = i === 0 ? C.seal : "C9C2B2";
    s.addShape(S.RECTANGLE, { x, y: baseY - h, w: bw, h, fill: { color: c }, line: { color: c, width: 0 } });
    s.addText(b[1] + "%", {
      x: x - 0.28, y: baseY - h - 0.46, w: bw + 0.56, h: 0.4, fontFace: F.serif, fontSize: 20,
      bold: true, color: i === 0 ? C.seal : C.inkSoft, align: "center", margin: 0,
    });
    s.addText(b[0], {
      x: x - 0.28, y: baseY + 0.12, w: bw + 0.56, h: 0.3, fontFace: F.sans, fontSize: 11,
      color: C.inkSoft, align: "center", margin: 0,
    });
  });
  s.addText("智联招聘 · 2024 届 offer 获得率", {
    x: x0 - 0.2, y: baseY + 0.52, w: 5.6, h: 0.26, fontFace: F.sans, fontSize: 9.5,
    color: C.muted, margin: 0,
  });

  // right panel
  const rx = 7.25;
  s.addShape(S.LINE, { x: rx - 0.35, y: 1.92, w: 0, h: 3.55, line: { color: C.rule, width: 0.5 } });
  s.addText("麦可思 · 2025 届毕业去向落实率", {
    x: rx, y: 1.94, w: 5.2, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.muted,
    charSpacing: 1.2, margin: 0,
  });
  s.addText([
    { text: "88.1", options: { fontSize: 44, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 22, color: C.seal } },
    { text: "  高职", options: { fontSize: 13, color: C.muted } },
  ], { x: rx, y: 2.22, w: 5.3, h: 0.82, fontFace: F.serif, margin: 0 });
  s.addText([
    { text: "85.1", options: { fontSize: 44, bold: true, color: C.inkSoft } },
    { text: "%", options: { fontSize: 22, color: C.inkSoft } },
    { text: "  本科", options: { fontSize: 13, color: C.muted } },
  ], { x: rx, y: 3.12, w: 5.3, h: 0.82, fontFace: F.serif, margin: 0 });
  s.addText("高职第一次反超。", {
    x: rx, y: 4.02, w: 4.6, h: 0.34, fontFace: F.serif, fontSize: 16, color: C.ink, margin: 0,
  });
  s.addText("口径提醒：offer 获得率 ≠ 教育部“毕业去向落实率”，两者差十几个百分点。", {
    x: rx, y: 4.5, w: 5.1, h: 0.8, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0,
    lineSpacingMultiple: 1.55,
  });

  s.addText("学历层次的差别在缩小，专业和产业的匹配在变得重要。", {
    x: M, y: 6.32, w: 11.4, h: 0.38, fontFace: F.serif, fontSize: 17, bold: true,
    color: C.ink, margin: 0,
  });

  footer(s, "来源：智联招聘《2024 中国大学生就业力调研报告》；麦可思《2026 年中国大学生就业报告》。");
  s.addNotes(notes.S09 || "");
}

/* ================= S10 international ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "09  /  国际比较");
  title(s, "别的国家也走过这条路", false, { size: 31, h: 0.58 });

  const rows = [
    ["韩国", "70.9%  vs  61.9%", "专科大学 vs 四年制大学就业率，差 9.1 个百分点，十年最大（2025）"],
    ["德国", "89%", "职业教育背景的 25—34 岁人群就业率，比高等教育群体高 1 个百分点（OECD 2026）"],
    ["瑞士", "58%", "高中生走双元制；但职业教育的收入优势在 35 岁左右被反超（2016）"],
  ];
  rows.forEach((r, i) => {
    const y = 2.16 + i * 1.1;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.16, w: 1.7, h: 0.44, fontFace: F.serif, fontSize: 25, bold: true, color: C.ink, margin: 0 });
    s.addText(r[1], { x: 2.42, y: y + 0.2, w: 3.9, h: 0.4, fontFace: F.serif, fontSize: 20, bold: true, color: C.seal, margin: 0 });
    s.addText(r[2], { x: 6.0, y: y + 0.18, w: 6.4, h: 0.7, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  });

  s.addShape(S.LINE, { x: M, y: 5.34, w: R - M, h: 0, line: { color: C.ink, width: 0.9 } });
  s.addText("早就业有优势，但优势会衰减。比学历层次更重要的，是专业和产业对不对得上。", {
    x: M, y: 5.72, w: 11.6, h: 0.4, fontFace: F.serif, fontSize: 17.5,
    color: C.ink, margin: 0,
  });

  footer(s, "来源：韩国大学信息公开 2025（经《首尔经济》2026-02 分析）；OECD, Education at a Glance 2026；Korber & Oesch 2016。");
  s.addNotes(notes.S10 || "");
}

/* ================= S11 pivot ================= */
{
  const s = newSlide(true);
  header(s, true);
  kicker(s, "10  /  重新定价", true);

  s.addText("从“层次溢价”", {
    x: M, y: 1.55, w: 10.2, h: 1.0, fontFace: F.serif, fontSize: 58, bold: true,
    color: C.paper, margin: 0, charSpacing: 1,
  });
  s.addText("到“匹配溢价”", {
    x: M, y: 2.6, w: 10.2, h: 1.0, fontFace: F.serif, fontSize: 58, bold: true,
    color: C.sealLight, margin: 0, charSpacing: 1,
  });

  const cols = [
    ["过去", "学历主要用来分层。学校越好，起点越高。"],
    ["现在", "分层还在，但专业、项目、实习和岗位的匹配度，权重在上升。"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 6.15;
    s.addShape(S.LINE, { x, y: 4.12, w: 5.55, h: 0, line: { color: C.ruleDark, width: 0.8 } });
    s.addText(c[0], { x, y: 4.3, w: 2.0, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.sealLight, charSpacing: 2.4, margin: 0 });
    s.addText(c[1], { x, y: 4.66, w: 5.4, h: 0.95, fontFace: F.sans, fontSize: 13.5, color: "CFC9B9", margin: 0, lineSpacingMultiple: 1.65, valign: "top" });
  });
  s.addText("面试里更常被问的是“你做过什么”。“你学过什么”反而问得少。", {
    x: M, y: 5.62, w: 11.2, h: 0.4, fontFace: F.serif, fontSize: 16,
    color: C.paper, margin: 0,
  });

  seal(s, 11.72, 1.55, 1.0, "重估", true);
  s.addNotes(notes.S11 || "");
}

/* ================= S12 three values ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "11  /  价值重估");
  title(s, "教育还在起作用的地方", false, { size: 31, h: 0.58 });

  const rows = [
    ["谋生", "工具价值", "起薪、门槛、考证。这一层确实在缩水，得承认。"],
    ["谋事", "能力价值", "把问题拆开、迁移到新场景、跟人协作。这一层在升值，但文凭证明不了，得靠做过的事。"],
    ["成人", "内在价值", "判断力、审美、怎么做公民。这一层不进工资单。"],
  ];
  rows.forEach((r, i) => {
    const y = 2.18 + i * 1.14;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.1, w: 1.5, h: 0.66, fontFace: F.serif, fontSize: 30, bold: true, color: C.seal, margin: 0 });
    s.addText(r[1], { x: 2.35, y: y + 0.32, w: 1.5, h: 0.3, fontFace: F.mono, fontSize: 8.5, color: C.muted, charSpacing: 1.5, margin: 0 });
    s.addText(r[2], { x: 3.6, y: y + 0.28, w: 8.8, h: 0.6, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.5 });
  });

  // data strip
  s.addShape(S.RECTANGLE, { x: M, y: 5.62, w: R - M, h: 0.92, fill: { color: C.ink }, line: { color: C.ink, width: 0 } });
  const stats = [["−3.7pp", "10 年内死亡概率"], ["−22%", "入狱概率"], ["−80%", "领取社会救助"], ["+18%", "投票率"]];
  stats.forEach((st, i) => {
    const x = M + 0.35 + i * 2.95;
    s.addText(st[0], { x, y: 5.74, w: 1.1, h: 0.34, fontFace: F.serif, fontSize: 17, bold: true, color: C.sealLight, margin: 0 });
    s.addText(st[1], { x: x + 1.02, y: 5.8, w: 1.85, h: 0.26, fontFace: F.sans, fontSize: 9, color: "CFC9B9", margin: 0 });
  });
  s.addText("注：基于义务教育年限的自然实验，不能直接外推到“再读一个硕士”。", {
    x: M, y: 6.6, w: 11.6, h: 0.24, fontFace: F.sans, fontSize: 8, color: C.muted, margin: 0,
  });

  footer(s, "来源：Oreopoulos & Salvanes, JEP 2011。");
  s.addNotes(notes.S12 || "");
}

/* ================= S13 policy ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "12  /  政策");
  title(s, "专业目录在改，改的是供给结构", false, { size: 30, h: 0.58 });

  s.addText("1.22", {
    x: M - 0.05, y: 2.02, w: 4.3, h: 1.5, fontFace: F.serif, fontSize: 108, bold: true,
    color: C.ink, margin: 0, lineSpacingMultiple: 0.9,
  });
  s.addText("万", { x: 3.9, y: 2.82, w: 0.5, h: 0.4, fontFace: F.serif, fontSize: 22, color: C.ink, margin: 0 });
  s.addText("“十四五”期间撤销或停招的\n本科专业布点", {
    x: M, y: 3.58, w: 4.3, h: 0.72, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
  });
  s.addText([
    { text: "同期新增 1.02 万个。", options: { breakLine: true } },
    { text: "不是少办教育，是换赛道。" },
  ], {
    x: M, y: 4.44, w: 4.4, h: 0.86, fontFace: F.serif, fontSize: 14, bold: true, color: C.seal,
    margin: 0, lineSpacingMultiple: 1.5,
  });

  const rx = 5.55;
  s.addShape(S.LINE, { x: rx - 0.3, y: 2.02, w: 0, h: 3.4, line: { color: C.rule, width: 0.5 } });
  const items = [
    ["专业", "撤掉供给过剩的：信息管理与信息系统、公共事业管理、市场营销、信息与计算科学"],
    ["培养", "专硕扩招、职教高考、产教融合；往集成电路、人工智能、智能制造加"],
    ["评价", "教育评价改革“破五唯”，推进技能等级与学历互认"],
  ];
  items.forEach((it, i) => {
    const y = 2.08 + i * 1.12;
    s.addShape(S.LINE, { x: rx, y, w: 7.0, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(it[0], { x: rx, y: y + 0.16, w: 0.9, h: 0.32, fontFace: F.serif, fontSize: 16, bold: true, color: C.seal, margin: 0 });
    s.addText(it[1], { x: rx + 0.86, y: y + 0.15, w: 6.05, h: 0.86, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  });

  s.addText("真正的变量，是专业、培养方式和出口。", {
    x: M, y: 5.75, w: 11.6, h: 0.4, fontFace: F.serif, fontSize: 19, bold: true, color: C.ink, margin: 0,
  });

  footer(s, "来源：教育部 2026-04-28 新闻发布会；“十四五”本科专业布点调整数据。");
  s.addNotes(notes.S13 || "");
}

/* ================= S14 actions ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "13  /  怎么办");
  title(s, "那具体能做什么", false, { size: 31, h: 0.58 });

  const rows = [
    ["01", "把学历当地基，别当护身符", "能拿出手的作品，比绩点更能说明问题。招聘方越来越看你做过什么。"],
    ["02", "选赛道时，三样东西放在一起看", "专业、产业、城市。“回炉”不是退步。那个本升专的例子，本质是一次重新匹配。"],
    ["03", "毕业以后还得学", "这件事以前是加分项，现在是必需项。把 AI 当同事，别当搜索引擎。"],
  ];
  rows.forEach((r, i) => {
    const y = 2.05 + i * 1.36;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.3, w: 0.54, h: 0.26, fontFace: F.mono, fontSize: 9, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.42, y: y + 0.26, w: 3.55, h: 0.78, fontFace: F.serif, fontSize: 18, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.35 });
    s.addText(r[2], { x: 5.35, y: y + 0.3, w: 7.0, h: 0.75, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55 });
  });

  s.addText("学历该拿还是要拿，只是别把它当成唯一的准备。", {
    x: M, y: 6.06, w: 11.4, h: 0.42, fontFace: F.serif, fontSize: 22, bold: true, color: C.seal, margin: 0,
  });

  footer(s, "这三条不一定对，你可以只挑一条。");
  s.addNotes(notes.S14 || "");
}

/* ================= S15 closing ================= */
{
  const s = newSlide(true);
  s.addShape(S.RECTANGLE, {
    x: 0.34, y: 0.34, w: W - 0.68, h: H - 0.68,
    fill: { color: C.ink, transparency: 100 }, line: { color: C.ruleDark, width: 0.6 },
  });

  s.addText("那张纸会旧，", {
    x: M + 0.1, y: 1.95, w: 11.3, h: 1.05, fontFace: F.serif, fontSize: 60, bold: true,
    color: C.paper, margin: 0, charSpacing: 1,
  });
  s.addText("学到的东西不会那么快旧", {
    x: M + 0.1, y: 3.32, w: 11.3, h: 1.05, fontFace: F.serif, fontSize: 62, bold: true,
    color: C.sealLight, margin: 0, charSpacing: 1,
  });

  s.addText("十七分钟讲不完一个人的选择。能说清的是：文凭只是进场的一张票，坐哪儿、走多远，看你会什么、做过什么。", {
    x: M + 0.1, y: 4.8, w: 10.9, h: 1.0, fontFace: F.sans, fontSize: 15,
    color: "CFC9B9", margin: 0, lineSpacingMultiple: 1.75,
  });

  s.addText("数据和出处都在最后一页，欢迎追问。", {
    x: M + 0.1, y: 5.98, w: 11.0, h: 0.36, fontFace: F.serif, fontSize: 14,
    color: C.paper, margin: 0,
  });

  s.addText("谢谢", {
    x: R - 2.6, y: 6.55, w: 2.6, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted,
    align: "right", charSpacing: 2, margin: 0,
  });
  s.addNotes(notes.S15 || "");
}

/* ================= S16 sources ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "附录  /  资料来源");
  title(s, "主要数据与文献", false, { size: 26, h: 0.5, y: 1.22 });

  const items = [
    "教育部《2025 年全国教育事业发展统计公报》；1998 年各级教育毛入学率统计",
    "教育部：2027 年全国硕士研究生招生考试报名人数 347 万（2026-11-13）",
    "国家统计局：2026 年 9 月分年龄组城镇调查失业率（2026-10-21）",
    "国家公务员局：2027 年度国考报名过审 345 万人（2026-10-26）",
    "智联招聘《2024 中国大学生就业力调研报告》；麦可思《2026 年中国大学生就业报告》",
    "环球网：郑州铁路职业技术学院“本升专”报道（2025）；教育部“十四五”专业布点调整（2026-04-28）",
    "Federal Reserve Bank of New York, Liberty Street Economics, 2025",
    "Brynjolfsson, Chandar & Chen, Canaries in the Coal Mine, Stanford Digital Economy Lab, 2025-11",
    "韩国大学信息公开 2025；OECD, Education at a Glance 2026；Korber & Oesch 2016",
    "Becker 1964；Spence 1973；Arrow 1973；Hirsch 1977；Collins 1979；Caplan 2018",
    "Oreopoulos & Salvanes, “Priceless”, Journal of Economic Perspectives, 2011",
  ];
  items.forEach((t, i) => {
    const y = 1.92 + i * 0.43;
    s.addText(String(i + 1).padStart(2, "0"), {
      x: M, y, w: 0.4, h: 0.26, fontFace: F.mono, fontSize: 7.5, color: C.seal, margin: 0,
    });
    s.addText(t, {
      x: 1.32, y, w: 10.9, h: 0.32, fontFace: F.sans, fontSize: 10.5, color: C.inkSoft,
      margin: 0, lineSpacingMultiple: 1.3,
    });
  });

  footer(s, "备用页，不计入 17 分钟。所有数字均标注口径，便于答辩追问。");
  s.addNotes(notes.S16 || "");
}

/* ---------------- write ---------------- */
const out = path.join(__dirname, "学历贬值_形势与政策.pptx");
pres.writeFile({ fileName: out }).then(() => {
  console.log("written:", out, "slides:", page);
});
