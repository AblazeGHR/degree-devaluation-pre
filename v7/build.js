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
pres.title = "学历“贬值”：怎样看待高等教育的价值";

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
    x: M, y, w: 9.5, h: 0.26, fontFace: F.mono, fontSize: 9,
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
  s.addText("怎样看待高等教育的价值", {
    x: M, y: 4.72, w: 7.6, h: 0.62, fontFace: F.serif, fontSize: 28,
    color: C.inkSoft, margin: 0, charSpacing: 1,
  });
  s.addText([
    { text: "20 MIN", options: { bold: true, color: C.ink } },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "2026.11", options: {} },
    { text: "    ／    ", options: { color: C.rule } },
    { text: "3942 字讲稿", options: {} },
  ], { x: M, y: 6.12, w: 7.2, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, charSpacing: 1.2, margin: 0 });

  s.addText("第？组：", {
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

/* ================= S02 contents + question ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "目录  /  问题");
  title(s, "我们今天讲什么", false, { size: 30, h: 0.6 });

  const items = ["一.学历贬值的表现", "二.学历贬值的本质", "三.如何应对学历贬值"];
  items.forEach((t, i) => {
    const y = 2.05 + i * 0.98;
    s.addShape(S.LINE, { x: M, y, w: 5.35, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(t, { x: M, y: y + 0.2, w: 5.2, h: 0.5, fontFace: F.serif, fontSize: 24, bold: true, color: C.ink, margin: 0, charSpacing: 0.5 });
  });

  s.addShape(S.LINE, { x: 6.42, y: 1.95, w: 0, h: 2.75, line: { color: C.rule, width: 0.5 } });

  const rx = 6.95;
  s.addText("先给一个判断", { x: rx, y: 1.95, w: 5.4, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.muted, charSpacing: 1.8, margin: 0 });
  s.addText([
    { text: "文凭数量增加", options: { color: C.ink } },
    { text: "  ≠  ", options: { color: C.seal } },
    { text: "教育价值下降", options: { color: C.ink } },
  ], { x: rx, y: 2.3, w: 5.6, h: 0.9, fontFace: F.serif, fontSize: 23, bold: true, margin: 0, lineSpacingMultiple: 1.4 });
  s.addText("那“贬值”的到底是什么？", {
    x: rx, y: 3.45, w: 5.4, h: 0.42, fontFace: F.serif, fontSize: 18, color: C.seal, margin: 0,
  });
  s.addText("贬的是什么、不贬的是什么、真正该积累的是什么。", {
    x: rx, y: 4.0, w: 5.4, h: 0.62, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.6,
  });

  footer(s, "三个问题，对应后面的三个部分。");
  s.addNotes(notes.S02 || "");
}

/* ================= S03 credential inflation ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第一部分  /  表现");
  title(s, "学历膨胀", false, { size: 34, h: 0.6 });

  const data = [
    ["2025", "1222 万", "+43 万"], ["2024", "1179 万", "+21 万"], ["2023", "1158 万", "+82 万"],
    ["2022", "1076 万", "+167 万"], ["2021", "909 万", "+35 万"], ["2020", "874 万", "+40 万"],
    ["2019", "834 万", "+14 万"], ["2018", "820 万", "+25 万"], ["2017", "795 万", "+30 万"],
    ["2016", "765 万", "+16 万"], ["2015", "749 万", "+22 万"], ["2014", "727 万", "+28 万"],
    ["2013", "699 万", "+19 万"], ["2012", "680 万", "+20 万"], ["2011", "660 万", "+30 万"],
    ["2010", "630 万", "+20 万"],
  ];
  const rows = [
    [
      { text: "年份", options: { bold: true, color: C.ink } },
      { text: "高校毕业生人数", options: { bold: true, color: C.ink } },
      { text: "变化", options: { bold: true, color: C.ink } },
    ],
    ...data.map((r) => [
      { text: r[0], options: { fontFace: F.mono, color: C.muted } },
      { text: r[1], options: { color: C.inkSoft, bold: r[0] === "2022" || r[0] === "2025" } },
      { text: r[2], options: { color: C.muted } },
    ]),
  ];
  s.addTable(rows, {
    x: 0.9, y: 1.95, w: 5.5, colW: [1.1, 2.5, 1.4],
    fontFace: F.sans, fontSize: 7.8, color: C.inkSoft,
    border: { type: "solid", color: C.rule, pt: 0.4 },
    fill: { color: C.paper },
    rowH: 0.205,
    margin: 0.03,
    valign: "middle",
  });

  const rx = 7.0;
  s.addText("自 2000 年以来，我国的硕士招生规模扩张了 20 倍，博士扩招 10 倍。", {
    x: rx, y: 1.95, w: 5.55, h: 1.35, fontFace: F.serif, fontSize: 21, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.5,
  });
  const tags = ["毕业生数量创新高", "竞争更激烈", "岗位要求提升"];
  tags.forEach((t, i) => {
    const y = 3.55 + i * 0.62;
    s.addShape(S.LINE, { x: rx, y, w: 5.4, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(t, { x: rx, y: y + 0.14, w: 5.2, h: 0.32, fontFace: F.serif, fontSize: 16, color: C.seal, margin: 0 });
  });
  s.addText("2026 届：1270 万，同比 +48 万", {
    x: rx, y: 5.62, w: 5.3, h: 0.3, fontFace: F.sans, fontSize: 10, color: C.muted, margin: 0,
  });
  footer(s, "来源：组员整理；教育部历年高校毕业生规模数据。");
  s.addNotes(notes.S03 || "");
}

/* ================= S04 job competition ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第一部分  /  表现");
  title(s, "就业竞争加剧", false, { size: 34, h: 0.55 });

  s.addText("数据显示，2023 年硕士和博士毕业生总数突破 100 万，可研究生就业率（44.4%）竟然比本科生（45.4%）和大专生（56.6%）都低。而且，就算找到了工作，薪资也没有想象中的那么高。", {
    x: M, y: 1.86, w: 11.6, h: 0.72, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
  });

  const rows = [
    [
      { text: "学历层次", options: { bold: true, color: C.ink } },
      { text: "2023 年就业率", options: { bold: true, color: C.ink } },
      { text: "起薪中位数（2021 年数据）", options: { bold: true, color: C.ink } },
      { text: "面临的主要就业难题", options: { bold: true, color: C.ink } },
    ],
    [
      { text: "本科生", options: { bold: true, color: C.ink } },
      { text: "45.4%", options: { bold: true, color: C.seal } },
      { text: "5000 元", options: { color: C.inkSoft } },
      { text: "本科生数量庞大，竞争激烈。在求职过程中，面临着来自研究生和专科生的双重竞争压力。一方面，与研究生相比，在一些对专业知识要求较高、发展潜力较大的岗位竞争中处于劣势；另一方面，专科生凭借较强的专业技能和较低的薪资要求，在一些技能型岗位上更具竞争力。此外，很多本科生在校所学知识与实际工作需求脱节，缺乏实践经验，导致难以满足企业对人才的实际需求。", options: { color: C.inkSoft } },
    ],
    [
      { text: "研究生", options: { bold: true, color: C.ink } },
      { text: "44.4%", options: { bold: true, color: C.seal } },
      { text: "/", options: { color: C.muted } },
      { text: "研究生虽然学历较高，但就业市场对高学历人才的需求增长速度跟不上供给速度。部分研究生在选择专业时，没有充分考虑市场需求，导致所学专业就业面狭窄。而且，一些企业认为研究生期望薪资过高，性价比不如本科生，更倾向于招聘本科生培养。同时，研究生自身对就业岗位的期望也较高，往往不愿意屈就一些看似不太理想的岗位，这也在一定程度上加剧了就业难度。", options: { color: C.inkSoft } },
    ],
    [
      { text: "博士生", options: { bold: true, color: C.ink } },
      { text: "/", options: { color: C.muted } },
      { text: "/", options: { color: C.muted } },
      { text: "博士生就业面相对较窄，主要集中在科研机构、高校以及一些大型企业的研发部门。然而，这些岗位数量有限，竞争异常激烈。而且，博士培养周期长，部分博士生在学术研究上花费了大量时间，导致实践能力和社会适应能力相对不足，在进入企业等非学术领域就业时，面临较大挑战。另外，社会对博士的期望较高，博士自身也背负着较大的就业压力，一旦就业不理想，容易遭受外界质疑。", options: { color: C.inkSoft } },
    ],
  ];
  s.addTable(rows, {
    x: M, y: 2.62, w: 11.75, colW: [1.0, 1.35, 1.75, 7.65],
    fontFace: F.sans, fontSize: 8.8, color: C.inkSoft,
    border: { type: "solid", color: C.rule, pt: 0.4 },
    fill: { color: C.paper },
    rowH: 0.36,
    margin: 0.06,
    valign: "top",
  });
  s.addText("学历更高，反而更难。这不是学历本身的问题，是结构的问题。", {
    x: M, y: 6.28, w: 11.5, h: 0.38, fontFace: F.serif, fontSize: 17, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "数据来源：组员整理（2023 年就业率、2021 年起薪中位数）。");
  s.addNotes(notes.S04 || "");
}

/* ================= S05 postgraduate exam ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第一部分  /  表现");
  title(s, "考研人数下降", false, { size: 34, h: 0.6 });

  const data = [["2020", 341], ["2021", 377], ["2022", 457], ["2023", 474], ["2024", 438], ["2025", 388], ["2026", 343]];
  const baseY = 5.35, maxH = 2.35, x0 = 0.95, bw = 0.62, gap = 0.31;
  s.addShape(S.LINE, { x: x0 - 0.15, y: baseY, w: 7 * bw + 6 * gap + 0.3, h: 0, line: { color: C.ink, width: 0.9 } });
  data.forEach((d, i) => {
    const x = x0 + i * (bw + gap);
    const h = (d[1] / 500) * maxH;
    const hot = d[0] === "2023" || d[0] === "2026";
    s.addShape(S.RECTANGLE, { x, y: baseY - h, w: bw, h, fill: { color: hot ? C.seal : "D3CCBC" }, line: { color: hot ? C.seal : "D3CCBC", width: 0 } });
    s.addText(String(d[1]), { x: x - 0.18, y: baseY - h - 0.3, w: bw + 0.36, h: 0.24, fontFace: F.mono, fontSize: 8.5, bold: hot, color: hot ? C.seal : C.muted, align: "center", margin: 0 });
    s.addText(d[0], { x: x - 0.18, y: baseY + 0.08, w: bw + 0.36, h: 0.22, fontFace: F.mono, fontSize: 8, color: C.muted, align: "center", margin: 0 });
  });
  s.addText("全国硕士研究生报名人数（万人）", {
    x: x0 - 0.15, y: baseY + 0.42, w: 5.4, h: 0.24, fontFace: F.sans, fontSize: 8.5, color: C.muted, margin: 0,
  });

  const rx = 7.75;
  s.addShape(S.LINE, { x: rx - 0.35, y: 1.95, w: 0, h: 3.3, line: { color: C.rule, width: 0.5 } });
  s.addText("•", { x: rx, y: 2.12, w: 0.25, h: 0.3, fontFace: F.sans, fontSize: 14, color: C.seal, margin: 0 });
  s.addText("不再盲目追逐更高学历，而是开始算“投入产出账”。学历的“滤镜”确实在褪色，但褪色的方式值得深究。", {
    x: rx + 0.28, y: 1.98, w: 4.4, h: 1.1, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.65,
  });
  s.addText("•", { x: rx, y: 3.62, w: 0.25, h: 0.3, fontFace: F.sans, fontSize: 14, color: C.seal, margin: 0 });
  s.addText("“贬值”的到底是什么？", {
    x: rx + 0.28, y: 3.6, w: 4.4, h: 0.34, fontFace: F.serif, fontSize: 15, bold: true, color: C.ink, margin: 0,
  });
  s.addText("贬的是什么、不贬的是什么、真正该积累的是什么", {
    x: rx + 0.28, y: 3.98, w: 4.4, h: 0.78, fontFace: F.sans, fontSize: 11.5, color: C.muted, margin: 0, lineSpacingMultiple: 1.55,
  });

  s.addText("2026 届毕业生 1270 万，同比 +48 万。", {
    x: M, y: 6.1, w: 11.4, h: 0.36, fontFace: F.serif, fontSize: 16, color: C.ink, margin: 0,
  });
  footer(s, "来源：中国教育在线 2000—2026 年考研报名数据；教育部 2026 届毕业生规模。");
  s.addNotes(notes.S05 || "");
}

/* ================= S06 scarcity ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第二部分  /  本质");
  title(s, "贬的是“稀缺性”，不是“知识”", false, { size: 31, h: 0.6 });

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
  s.addText("仅“研究生”一项，就超过了 1998 年本专科招生的总和。", {
    x: rx, y: 1.95, w: 5.1, h: 1.0, fontFace: F.serif, fontSize: 20, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.4,
  });
  s.addText("高等教育从精英化走向普及化，文凭的“标签价值”被稀释——这是扩容的必然结果，不是教育的失败。", {
    x: rx, y: 3.15, w: 5.0, h: 1.0, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addText("把“文凭不再稀缺”等同于“读书无用”，是逻辑上的偷换。", {
    x: M, y: 6.28, w: 11.6, h: 0.4, fontFace: F.serif, fontSize: 17, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "数据来源：教育部《全国教育事业发展统计公报》（1990、1998、2025 年招生数据）。");
  s.addNotes(notes.S06 || "");
}

/* ================= S07 structural ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第二部分  /  本质");
  title(s, "就业难是结构问题，不是学历问题", false, { size: 30, h: 0.6 });

  s.addText([
    { text: "200", options: { fontSize: 66, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 24, color: C.seal } },
  ], { x: M - 0.04, y: 2.12, w: 3.2, h: 0.9, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("AI 相关职位需求增长", { x: M, y: 3.05, w: 4.3, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });
  s.addText("招聘平台在招职位同比", { x: M, y: 3.36, w: 4.3, h: 0.26, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0 });

  s.addText([
    { text: "73.65", options: { fontSize: 50, bold: true, color: C.ink } },
    { text: "%", options: { fontSize: 20, color: C.ink } },
  ], { x: M - 0.04, y: 3.85, w: 3.4, h: 0.8, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("具身智能领域新发职位同比增长", { x: M, y: 4.68, w: 4.4, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });
  s.addText("同比口径，2025 年", { x: M, y: 4.99, w: 4.4, h: 0.26, fontFace: F.sans, fontSize: 9.5, color: C.muted, margin: 0 });

  const bx = 5.2;
  s.addShape(S.RECTANGLE, { x: bx, y: 1.95, w: 7.35, h: 2.9, fill: { color: C.paper2 }, line: { color: C.paper2, width: 0 } });
  s.addText("但另一面", { x: bx + 0.35, y: 2.1, w: 2.4, h: 0.26, fontFace: F.mono, fontSize: 8.5, color: C.muted, charSpacing: 1.8, margin: 0 });
  s.addText("传统白领岗位的吸纳能力在收缩。", {
    x: bx + 0.35, y: 2.42, w: 6.5, h: 0.4, fontFace: F.serif, fontSize: 19, bold: true, color: C.ink, margin: 0,
  });
  s.addText("岗位在变，能力要求在变。只带着一张文凭、没有匹配新产业需求的能力，自然会感到“贬值”。", {
    x: bx + 0.35, y: 2.98, w: 6.1, h: 0.7, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.65,
  });
  s.addShape(S.RECTANGLE, { x: bx + 0.35, y: 3.8, w: 0.05, h: 0.62, fill: { color: C.seal }, line: { color: C.seal, width: 0 } });
  s.addText("问题不在学历，在于学历背后有没有真东西。", {
    x: bx + 0.62, y: 3.86, w: 6.3, h: 0.5, fontFace: F.serif, fontSize: 15, bold: true, color: C.seal, margin: 0, lineSpacingMultiple: 1.4,
  });

  s.addShape(S.RECTANGLE, { x: M, y: 5.35, w: R - M, h: 0.86, fill: { color: C.ink }, line: { color: C.ink, width: 0 } });
  s.addText("企业招聘正从“唯学历”转向“重能力”；宇树科技明确“不唯学历唯能力”。", {
    x: M + 0.32, y: 5.52, w: 11.6, h: 0.5, fontFace: F.sans, fontSize: 13.5, color: "EDE8DA", margin: 0, valign: "middle",
  });
  footer(s, "数据来源：招聘平台公开报告；企业用人标准援引自公开报道。");
  s.addNotes(notes.S07 || "");
}

/* ================= S08 three dimensions ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第二部分  /  本质");
  title(s, "学历的真正价值：不止于求职", false, { size: 31, h: 0.55 });
  s.addText("三个维度，决定一个人能走多远", {
    x: M, y: 1.9, w: 8.0, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.muted, margin: 0,
  });

  const cols = [
    ["01", "能力", "系统训练判断力、协作力，\n以及拆解复杂问题的能力。"],
    ["02", "认知", "拓展理解世界的框架，\n而不是停留在记忆知识点。"],
    ["03", "综合素质", "在多元环境中与人共处、\n自我管理、持续学习。"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 3.95;
    s.addText(c[0], { x, y: 2.25, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 8.5, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(c[1], { x, y: 2.55, w: 3.4, h: 0.55, fontFace: F.serif, fontSize: 26, bold: true, color: C.ink, margin: 0, charSpacing: 2 });
    s.addShape(S.LINE, { x, y: 3.2, w: 3.45, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(c[2], { x, y: 3.36, w: 3.45, h: 0.9, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.65 });
  });

  s.addShape(S.RECTANGLE, { x: M, y: 4.55, w: 0.05, h: 1.0, fill: { color: C.seal }, line: { color: C.seal, width: 0 } });
  s.addText("大学最大的产品，从来不是“就业率”，\n而是一个具备长期成长能力的人。", {
    x: M + 0.3, y: 4.5, w: 11.0, h: 1.0, fontFace: F.serif, fontSize: 20, color: C.ink, margin: 0, lineSpacingMultiple: 1.55,
  });
  s.addText("职业技能可以短期培训；认知框架与人格养成，需要系统性的土壤。", {
    x: M, y: 6.02, w: 11.4, h: 0.32, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0,
  });
  footer(s, "三个维度：能力、认知、综合素质。");
  s.addNotes(notes.S08 || "");
}

/* ================= S09 reverse signal ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第二部分  /  本质");
  title(s, "一个反向信号：技能价值正在重估", false, { size: 30, h: 0.6 });

  s.addText([
    { text: "84", options: { fontSize: 118, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 24, color: C.seal } },
  ], { x: M - 0.05, y: 2.18, w: 3.05, h: 1.35, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("高职毕业生就业满意度", { x: M, y: 3.66, w: 4.0, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });
  s.addText("首次反超本科生", { x: M, y: 3.99, w: 4.0, h: 0.3, fontFace: F.sans, fontSize: 10.5, color: C.seal, margin: 0 });

  s.addText("▲ 反超 3 个百分点", {
    x: 4.95, y: 2.72, w: 2.8, h: 0.34, fontFace: F.sans, fontSize: 12, color: C.seal, align: "center", margin: 0,
  });
  s.addShape(S.LINE, { x: 4.95, y: 2.45, w: 2.8, h: 0, line: { color: C.rule, width: 0.5 } });

  s.addText([
    { text: "81", options: { fontSize: 118, bold: true, color: C.inkSoft } },
    { text: "%", options: { fontSize: 24, color: C.inkSoft } },
  ], { x: 7.55, y: 2.18, w: 3.05, h: 1.35, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("本科毕业生就业满意度", { x: 7.55, y: 3.66, w: 4.0, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });
  s.addText("同一调查口径", { x: 7.55, y: 3.99, w: 4.0, h: 0.3, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0 });

  s.addText("“本升技”“硕升技”：从个案走向常态化。", {
    x: M, y: 4.9, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 18, bold: true, color: C.ink, margin: 0,
  });
  s.addText("市场正在从“看标签”转向“看实效”——能解决问题的人，无论来自哪种教育路径，都在获得认可。", {
    x: M, y: 5.45, w: 11.3, h: 0.5, fontFace: F.sans, fontSize: 12, color: C.muted, margin: 0, lineSpacingMultiple: 1.5,
  });
  footer(s, "数据来源：麦可思研究院《中国高职生就业报告》等相关调查。注意：这是满意度，不是薪资。");
  s.addNotes(notes.S09 || "");
}

/* ================= S10 inert credential ================= */
{
  const s = newSlide(true);
  header(s, true);
  kicker(s, "第二部分  /  小结", true);
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

  s.addText("下面进入第三部分：我们能做什么。", {
    x: M, y: 6.05, w: 10.0, h: 0.32, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  seal(s, 11.55, 1.5, 1.0, "小结", true);
  s.addNotes(notes.S10 || "");
}

/* ================= S11 expectation ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第三部分  /  行动");
  title(s, "调整预期：从“我是什么学历”到“我能解决什么”", false, { size: 26, h: 0.62 });

  s.addText([
    { text: "60", options: { fontSize: 124, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 30, color: C.seal } },
  ], { x: M - 0.06, y: 2.15, w: 3.3, h: 1.7, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("2025 届本科毕业生\n选择地级及以下城市就业", {
    x: M, y: 4.0, w: 4.4, h: 0.8, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.55,
  });

  const rx = 5.55;
  s.addShape(S.LINE, { x: rx - 0.35, y: 1.95, w: 0, h: 2.9, line: { color: C.rule, width: 0.5 } });
  s.addText("一线城市的“虹吸效应”在弱化。", {
    x: rx, y: 1.98, w: 6.9, h: 0.38, fontFace: F.serif, fontSize: 20, bold: true, color: C.ink, margin: 0,
  });
  s.addText("机会在扩散。县城、地级市、制造业基地，都在要人。\n关键是愿不愿意去、能不能接得住。", {
    x: rx, y: 2.55, w: 6.8, h: 0.8, fontFace: F.sans, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.75,
  });
  s.addText("就业预期要从“学历兑换岗位”，转向“能力匹配需求”。", {
    x: rx, y: 3.55, w: 6.8, h: 0.4, fontFace: F.serif, fontSize: 15, color: C.seal, margin: 0,
  });

  s.addText("调整预期不是躺平，是把“我应得什么”换成“市场需要什么”。", {
    x: M, y: 5.95, w: 11.5, h: 0.4, fontFace: F.serif, fontSize: 17, bold: true, color: C.ink, margin: 0,
  });
  footer(s, "来源：麦可思 2025 届本科毕业生就业地分布。");
  s.addNotes(notes.S11 || "");
}

/* ================= S12 tracks ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第三部分  /  行动");
  title(s, "选对赛道：关注扩张性领域", false, { size: 31, h: 0.6 });

  const rows = [
    ["01", "AI 与具身智能", "算法工程师、智能体开发等岗位需求爆发，具身智能平均年薪 33 万以上。"],
    ["02", "先进制造", "工艺、制程工程师需求高速增长，工业自动化、新能源汽车产业链持续扩张。"],
    ["03", "现代服务业", "新职业不断涌现，县域经济同样提供成长空间。"],
  ];
  rows.forEach((r, i) => {
    const y = 2.0 + i * 1.28;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.24, w: 0.5, h: 0.24, fontFace: F.mono, fontSize: 8.5, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.32, y: y + 0.2, w: 3.5, h: 0.4, fontFace: F.serif, fontSize: 20, bold: true, color: C.ink, margin: 0 });
    s.addText(r[2], { x: 4.95, y: y + 0.2, w: 7.5, h: 0.85, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0, lineSpacingMultiple: 1.6 });
  });

  s.addText("不是所有人都要挤 AI，但要知道哪些领域在“进水”，哪些在“出水”。", {
    x: M, y: 6.05, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 16, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "选赛道不是追热点，是看需求能不能持续、能力能不能复用。");
  s.addNotes(notes.S12 || "");
}

/* ================= S13 capabilities ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第三部分  /  行动");
  title(s, "积累“抗贬值”的能力", false, { size: 31, h: 0.6 });

  s.addText([
    { text: "215", options: { fontSize: 56, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 20, color: C.seal } },
  ], { x: M - 0.04, y: 2.02, w: 2.0, h: 0.72, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("要求会 AI 的岗位同比增长", { x: 2.72, y: 2.05, w: 5.0, h: 0.34, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });
  s.addText("AI 正在成为职场通用能力。", { x: 2.72, y: 2.42, w: 8.2, h: 0.32, fontFace: F.sans, fontSize: 10.5, color: C.muted, margin: 0 });

  const rows = [
    ["01", "AI 工具运用能力", "不是会聊天，是能用它把活干完。"],
    ["02", "场景化解决问题能力", "实习、项目、竞赛中积累的“做过的事”。"],
    ["03", "跨领域协作与表达", "判断谁可信，把模糊问题变成可执行方案。"],
  ];
  rows.forEach((r, i) => {
    const y = 3.05 + i * 0.9;
    s.addShape(S.LINE, { x: M, y, w: R - M, h: 0, line: { color: C.rule, width: 0.5 } });
    s.addText(r[0], { x: M, y: y + 0.2, w: 0.45, h: 0.22, fontFace: F.mono, fontSize: 8, color: C.seal, charSpacing: 2, margin: 0 });
    s.addText(r[1], { x: 1.25, y: y + 0.16, w: 3.8, h: 0.34, fontFace: F.serif, fontSize: 17, bold: true, color: C.ink, margin: 0 });
    s.addText(r[2], { x: 5.25, y: y + 0.2, w: 7.2, h: 0.34, fontFace: F.sans, fontSize: 11, color: C.inkSoft, margin: 0 });
  });

  s.addText("一段扎实的实习、一个完整的项目、一项可展示的硬技能。", {
    x: M, y: 5.85, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 16, color: C.ink, margin: 0,
  });
  s.addText("文凭是入场券，作品集才是通行证。", {
    x: M, y: 6.3, w: 11.4, h: 0.34, fontFace: F.serif, fontSize: 15, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "来源：招聘平台公开报告（AI 技能岗位同比）；能力清单来自 PPT 大纲。");
  s.addNotes(notes.S13 || "");
}

/* ================= S14 mindset ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第三部分  /  行动");
  title(s, "一个务实的心态：第一份工作不一定是终身岗位", false, { size: 27, h: 0.6 });

  s.addText([
    { text: "15.9", options: { fontSize: 88, bold: true, color: C.ink } },
    { text: "%", options: { fontSize: 24, color: C.ink } },
  ], { x: M - 0.04, y: 2.32, w: 3.2, h: 1.1, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("应届生离职率", { x: M, y: 3.52, w: 4.0, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });

  s.addText([
    { text: "30.6", options: { fontSize: 88, bold: true, color: C.seal } },
    { text: "%", options: { fontSize: 24, color: C.seal } },
  ], { x: 5.35, y: 2.32, w: 3.2, h: 1.1, fontFace: F.serif, margin: 0, lineSpacingMultiple: 0.9 });
  s.addText("在入职前三个月内离职", { x: 5.35, y: 3.52, w: 4.5, h: 0.3, fontFace: F.sans, fontSize: 12, color: C.inkSoft, margin: 0 });

  s.addText("调整预期不是降低标准，而是把“试错”纳入规划。", {
    x: M, y: 4.45, w: 11.4, h: 0.4, fontFace: F.serif, fontSize: 18, color: C.ink, margin: 0,
  });
  s.addText("刚毕业，你对自己、对行业、对岗位的理解，都可能是错的。试错成本最低的时候，就是现在。", {
    x: M, y: 5.02, w: 11.4, h: 0.34, fontFace: F.sans, fontSize: 11.5, color: C.inkSoft, margin: 0,
  });
  s.addText("先上场，再调整。", {
    x: M, y: 5.95, w: 11.0, h: 0.5, fontFace: F.serif, fontSize: 24, bold: true, color: C.seal, margin: 0,
  });
  footer(s, "来源：麦可思应届生离职率与离职时间分布。");
  s.addNotes(notes.S14 || "");
}

/* ================= S15 institutions ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第四部分  /  社会");
  title(s, "制度层面正在发生的变化", false, { size: 30, h: 0.6 });

  const stats = [
    ["10", "%", "2026 年本科专业调整比例\n首破 10%"],
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
  s.addText("匹配不是包分配。它做的是把信息打通、把渠道铺好。", {
    x: M, y: 5.95, w: 11.4, h: 0.36, fontFace: F.sans, fontSize: 11, color: C.muted, margin: 0,
  });
  footer(s, "来源：教育部“十四五”本科专业布点调整；人社部稳就业举措。");
  s.addNotes(notes.S15 || "");
}

/* ================= S16 plural rulers ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "第四部分  /  社会");
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
  footer(s, "引文：人民日报评论。");
  s.addNotes(notes.S16 || "");
}

/* ================= S17 closing ================= */
{
  const s = newSlide(true);
  s.addShape(S.RECTANGLE, {
    x: 0.34, y: 0.34, w: W - 0.68, h: H - 0.68,
    fill: { color: C.ink, transparency: 100 }, line: { color: C.ruleDark, width: 0.6 },
  });
  s.addText("学历没有作废，", {
    x: M + 0.1, y: 1.85, w: 11.3, h: 1.0, fontFace: F.serif, fontSize: 50, bold: true,
    color: C.paper, margin: 0, charSpacing: 1,
  });
  s.addText("只是不再承诺保底。", {
    x: M + 0.1, y: 2.95, w: 11.3, h: 1.0, fontFace: F.serif, fontSize: 50, bold: true,
    color: C.sealLight, margin: 0, charSpacing: 1,
  });
  s.addText("真正保底的，是你持续长出能力的速度。", {
    x: M + 0.1, y: 4.4, w: 10.9, h: 0.5, fontFace: F.serif, fontSize: 20,
    color: C.paper, margin: 0,
  });
  s.addText("高等教育的价值，从来不在一纸证书的厚度，而在它是否让一个人具备了面对不确定性的底气。", {
    x: M + 0.1, y: 5.1, w: 10.9, h: 0.4, fontFace: F.sans, fontSize: 13,
    color: "CFC9B9", margin: 0,
  });
  s.addText("谢谢", { x: R - 2.6, y: 6.55, w: 2.6, h: 0.3, fontFace: F.mono, fontSize: 9, color: C.muted, align: "right", charSpacing: 2, margin: 0 });
  s.addNotes(notes.S17 || "");
}

/* ================= S18 sources ================= */
{
  const s = newSlide();
  header(s);
  kicker(s, "附录  /  资料来源");
  title(s, "主要数据与文献", false, { size: 26, h: 0.5, y: 1.22 });

  const items = [
    "教育部《全国教育事业发展统计公报》：1990、1998、2025 年招生数据；1998、2025 年毛入学率",
    "教育部：2026 届高校毕业生规模；中国教育在线：2000—2026 年考研报名人数",
    "智联招聘《2024 中国大学生就业力调研报告》",
    "麦可思研究院《中国高职生就业报告》等相关调查",
    "麦可思：2025 届本科毕业生就业地分布；应届生离职率",
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
  s.addNotes(notes.S18 || "");
}

const out = path.join(__dirname, "学历贬值_形势与政策.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written:", out, "slides:", page));
