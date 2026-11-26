// 效果图场景：把 pptjs / wordjs 真实渲染出的页面拼成对比图，截图由 build.sh 用 --file 传入。
// PPT、Word 场景里的数字（溢出量、修正轮数、检查结果）是 build.sh 里各脚本的实际运行结果。
// Excel 场景是示意图：表格界面和图表由本脚本画出，不是 Excel 引擎的渲染结果。
ICONS.chat = "M4 4 H20 V16 H9 L4 20 Z";

function request(slide, str, t) {
  box(slide, { l: L, t, w: W, h: 36, fill: C.accentTint, radius: 0.3 });
  icon(slide, "chat", L + 14, t + 9, 18, C.accent);
  text(slide, str, { l: L + 42, t: t + 8, w: W - 56, h: 20, size: tr(13, 14), bold: true, color: C.accentDark });
}
function caption(slide, n, h, d, l, t, w) {
  label(box(slide, { shape: "Ellipse", l, t: t + 1, w: 20, h: 20, fill: C.accentFill }), n, { size: 10.5, color: C.white });
  text(slide, h, { l: l + 28, t, w: w - 28, h: 20, size: 13, bold: true });
  if (d) text(slide, d, { l: l + 28, t: t + 22, w: w - 28, h: 36, size: 11, color: C.mute, lineSpacing: 1.15 });
}
function console_(slide, h, lines, o) {
  box(slide, { l: o.l, t: o.t, w: o.w, h: o.h, fill: C.code, lineColor: C.line, radius: 0.05 });
  text(slide, h, { l: o.l + 18, t: o.t + 12, w: o.w - 36, h: 18, size: 11.5, bold: true, color: C.accentLight });
  text(slide, lines.join("\r"), { l: o.l + 18, t: o.t + 36, w: o.w - 36, h: o.h - 44, size: 11, font: MONO, color: "#E8EBF2", lineSpacing: 1.15 });
}
// 中间一栏：AI 做了哪些修改，最后一条是检查结果
function changes(slide, title, items, o) {
  box(slide, { l: o.l, t: o.t, w: o.w, h: 46 + items.length * 29, fill: C.card, lineColor: C.line, radius: 0.06 });
  text(slide, title, { l: o.l + 14, t: o.t + 12, w: o.w - 28, h: 18, size: 12, bold: true });
  items.forEach((str, i) => {
    const last = i === items.length - 1;
    icon(slide, "check", o.l + 14, o.t + 42 + i * 29, 16, last ? C.ok : C.accent);
    text(slide, str, { l: o.l + 36, t: o.t + 42 + i * 29, w: o.w - 44, h: 18, size: 11, color: last ? C.okDark : C.fg, bold: last });
  });
}

// 表格界面（示意）：编辑栏、列标、行号、单元格、工作表标签
const XL = { grid: "#D9DCE1", head: "#F3F4F6", headText: "#5D6678", good: "#C6EFCE", goodText: "#006100" };
function sheetMock(slide, o) {
  const fb = 22, ch = 15, rh = o.rowHeight || 18, rw = 22, tabH = 20;
  const gridT = o.t + fb + 4, cellsT = gridT + ch;
  box(slide, { shape: "Rectangle", l: o.l, t: o.t, w: o.w, h: o.h, fill: C.white, lineColor: XL.grid });
  // 编辑栏
  label(box(slide, { shape: "Rectangle", l: o.l + 4, t: o.t + 4, w: 40, h: fb - 4, fill: C.white, lineColor: XL.grid }), o.ref, { size: 9, bold: false, color: C.ink });
  text(slide, "fx", { l: o.l + 50, t: o.t + 6, w: 16, h: 14, size: 9.5, color: XL.headText, font: "Georgia" });
  box(slide, { shape: "Rectangle", l: o.l + 68, t: o.t + 4, w: o.w - 72, h: fb - 4, fill: C.white, lineColor: XL.grid });
  text(slide, o.formula, { l: o.l + 74, t: o.t + 7, w: o.w - 84, h: 14, size: 9, font: MONO, color: C.ink });
  // 列标和行号
  box(slide, { shape: "Rectangle", l: o.l, t: gridT, w: rw, h: ch, fill: XL.head, lineColor: XL.grid, lineWeight: 0.5 });
  let x = o.l + rw;
  const xs = o.cols.map((w) => { const v = x; x += w; return v; });
  o.cols.forEach((w, c) => label(box(slide, { shape: "Rectangle", l: xs[c], t: gridT, w, h: ch, fill: c === o.activeCol ? "#E3E8F0" : XL.head, lineColor: XL.grid, lineWeight: 0.5 }), "ABCDEFGH"[c], { size: 8, bold: false, color: XL.headText }));
  const rows = Math.floor((o.h - fb - 4 - ch - tabH) / rh);
  for (let r = 0; r < rows; r++) {
    const t = cellsT + r * rh;
    label(box(slide, { shape: "Rectangle", l: o.l, t, w: rw, h: rh, fill: XL.head, lineColor: XL.grid, lineWeight: 0.5 }), String(r + 1), { size: 8, bold: false, color: XL.headText });
    o.cols.forEach((w, c) => {
      const v = (o.data[r] || [])[c];
      const st = (o.style && o.style(r, c, v)) || {};
      const b = box(slide, { shape: "Rectangle", l: xs[c], t, w, h: rh, fill: st.fill || C.white, lineColor: XL.grid, lineWeight: 0.5 });
      if (v !== undefined && v !== "") {
        const num = /^[-+]?[\d,.]+%?$/.test(v);
        label(b, v, { size: 9, bold: !!st.bold, color: st.color || C.ink, align: st.align || (num ? "Right" : "Left"), margin: 4 });
      }
    });
  }
  // 工作表标签
  const tabT = o.t + o.h - tabH;
  box(slide, { shape: "Rectangle", l: o.l, t: tabT, w: o.w, h: tabH, fill: XL.head, lineColor: XL.grid });
  o.tabs.forEach((name, i) => {
    const active = i === o.activeTab;
    label(box(slide, { shape: "Rectangle", l: o.l + 28 + i * 56, t: tabT, w: 52, h: tabH - 2, fill: active ? C.white : XL.head, lineColor: active ? XL.grid : XL.head }), name, { size: 9, bold: active, color: active ? C.excel : XL.headText });
  });
  return { xs, cellsT, rh };
}

const SCENE_TRANSLATE = {
  name: "translate", stages: 3, draw(s) {
    header(s, tr("Result · PowerPoint", "效果 · PowerPoint"), tr("Longer translation overflows: the AI finds it and fixes it", "译文变长、文字溢出：AI 自己发现，自己修好"));
    const crop = { left: 0.03, right: 0.03, top: 0.05, bottom: 0.17 };
    pic(s, "en", { l: L, t: 126, w: 404, h: 190, crop });
    caption(s, "1", tr("Chinese slide → English: inspect reports 3 overflows", "中文页换成英文：inspect 报 3 处溢出"), tr("Each box is 3.2 pt over: needs 4 lines, fits 3 (red boxes are annotations)", "每处超出 3.2 磅：需要 4 行，只放得下 3 行（红框是标注）"), L, 330, 404);
    if (at(2)) {
      pic(s, "fix", { l: L + 428, t: 126, w: 404, h: 190, crop });
      arrowLine(s, L + 408, 221, L + 424, 221, { color: C.accent, weight: 2 });
      caption(s, "2", tr("The AI fixes it from the report", "AI 按检查结果修正"), tr("Body text shrunk on all three cards; round 4 finds 0 overflows", "三张卡的正文统一缩小，第 4 轮检查 0 处溢出"), L + 428, 330, 404);
    }
    if (at(3)) console_(s, tr("Log returned by the AI's fix script", "AI 的修正脚本返回的记录"), [
      "round 1: 3 overflow(s), 20 -> 19 pt",
      "round 2: 2 overflow(s), 19 -> 18 pt",
      "round 3: 2 overflow(s), 18 -> 17 pt",
      "round 4: clean at 17 pt",
    ], { l: L, t: 392, w: W, h: 112 });
  },
};

const SCENE_EDIT = {
  name: "edit", stages: 2, draw(s) {
    header(s, tr("Result · PowerPoint", "效果 · PowerPoint"), tr("Edits the user's own file; layout and brand elements stay", "在用户自己的文件上改，版式和品牌元素原样保留"));
    request(s, tr("User: make this slide two columns, with a quarterly revenue chart on the right", "用户：把这页改成两栏，右边放季度收入图"), 120);
    pic(s, "eb", { l: L, t: 180, w: 380, h: 213.75 });
    caption(s, "1", tr("User's original file", "用户原文件"), tr("Office theme, Title and Content layout, brand bar on top", "Office 主题的「标题和内容」版式，顶部是品牌色条"), L, 408, 380);
    if (!at(2)) return;
    pic(s, "ea", { l: L + 452, t: 180, w: 380, h: 213.75 });
    arrowLine(s, L + 392, 287, L + 440, 287, { color: C.accent, weight: 2.5 });
    caption(s, "2", tr("After the AI's edit: inspect finds 0 issues", "AI 修改后：inspect 0 个问题"), tr("Title, layout and brand bar untouched; body at half width, with a native chart you can edit in PowerPoint", "标题、版式、色条不动；正文缩成半栏，右边是原生 PowerPoint 图表，在 PowerPoint 里能直接改数据"), L + 452, 408, 380);
  },
};

const SCENE_WORD = {
  name: "word", stages: 3, draw(s) {
    header(s, tr("Result · Word", "效果 · Word"), tr("One sentence edits the report; everything else stays as it was", "按一句话修改报告，没提到的内容原样保留"));
    request(s, tr("User: three-line table, right-align numbers, highlight the top-YoY city, add key points", "用户：表格改成三线表，数字右对齐，标出同比最高的城市；总体情况后面加几条要点"), 116);
    const crop = { bottom: 0.28 };
    pic(s, "wb", { l: L, t: 168, w: 290, h: 270, crop });
    caption(s, "1", tr("User's report", "用户的报告"), null, L, 452, 290);
    const mid = L + 324, mw = 184;
    if (at(2)) {
      changes(s, tr("Edits to the original", "AI 在原文件上做的修改"), [tr("Three-line table", "表格改成三线表"), tr("Numbers right-aligned", "数字列右对齐"), tr("Top-YoY row highlighted", "同比最高的上海整行标出"), tr("Key points after summary", "总体情况后加要点列表"), tr("inspect: 0 issues", "改完 inspect：0 个问题")], { l: mid, t: 196, w: mw });
    }
    if (at(3)) {
      arrowLine(s, mid + mw / 2 - 30, 410, mid + mw / 2 + 30, 410, { color: C.accent, weight: 2.5 });
      pic(s, "wa", { l: L + 542, t: 168, w: 290, h: 270, crop });
      caption(s, "2", tr("After", "修改后"), null, L + 542, 452, 290);
    }
  },
};

// 示意数据：与 Word 示例报告的城市和同比一致（南京 15.0% 未超过 15%，不标绿）
const XL_RAW = tr([
  ["Month", "City", "Product", "Rev. (¥10k)", "Last year"],
  ["2026-07", "Shanghai", "Cloud", "1,820", "1,505"],
  ["2026-07", "Hangzhou", "Cloud", "1,060", "912"],
  ["2026-07", "Nanjing", "Services", "790", "702"],
  ["2026-08", "Shanghai", "Services", "1,710", "1,402"],
  ["2026-08", "Suzhou", "Services", "610", "512"],
  ["2026-08", "Nanjing", "Cloud", "820", "715"],
  ["2026-09", "Hangzhou", "Services", "1,050", "905"],
  ["2026-09", "Shanghai", "Cloud", "1,790", "1,490"],
  ["2026-09", "Suzhou", "Cloud", "640", "538"],
  ["…", "", "", "", ""],
], [
  ["月份", "城市", "产品", "收入（万元）", tr("Last year", "去年同期")],
  ["2026-07", "上海", "云服务", "1,820", "1,505"],
  ["2026-07", "杭州", "云服务", "1,060", "912"],
  ["2026-07", "南京", "实施", "790", "702"],
  ["2026-08", "上海", "实施", "1,710", "1,402"],
  ["2026-08", "苏州", "实施", "610", "512"],
  ["2026-08", "南京", "云服务", "820", "715"],
  ["2026-09", "杭州", "实施", "1,050", "905"],
  ["2026-09", "上海", "云服务", "1,790", "1,490"],
  ["2026-09", "苏州", "云服务", "640", "538"],
  ["…", "", "", "", ""],
]);
const XL_SUM = tr([
  ["City", "Revenue", "Last year", "YoY"],
  ["Shanghai", "5,320", "4,397", "21.0%"],
  ["Hangzhou", "3,180", "2,741", "16.0%"],
  ["Nanjing", "2,460", "2,140", "15.0%"],
  ["Suzhou", "1,840", "1,546", "19.0%"],
  ["Total", "12,800", "10,824", "18.3%"],
], [
  ["城市", tr("This year", "本季收入"), tr("Last year", "去年同期"), "同比"],
  ["上海", "5,320", "4,397", "21.0%"],
  ["杭州", "3,180", "2,741", "16.0%"],
  ["南京", "2,460", "2,140", "15.0%"],
  ["苏州", "1,840", "1,546", "19.0%"],
  ["合计", "12,800", "10,824", "18.3%"],
]);
// 浮在表格上的柱形图（示意）
function miniChart(slide, o) {
  const cats = tr(["Shanghai", "Hangzhou", "Nanjing", "Suzhou"], ["上海", "杭州", "南京", "苏州"]), cur = [5320, 3180, 2460, 1840], prev = [4397, 2741, 2140, 1546];
  box(slide, { shape: "Rectangle", l: o.l, t: o.t, w: o.w, h: o.h, fill: C.white, lineColor: XL.grid });
  text(slide, tr("Revenue by city (¥10k)", "分城市收入（万元）"), { l: o.l, t: o.t + 5, w: o.w, h: 14, size: 9, bold: true, align: "Center", color: C.ink });
  const base = o.t + o.h - 30, top = o.t + 24, slot = (o.w - 24) / cats.length;
  rule(slide, o.l + 12, base, o.w - 24, XL.grid);
  cats.forEach((c, i) => {
    const x = o.l + 12 + i * slot + slot / 2;
    [[cur[i], C.excel, -15], [prev[i], "#A9D18E", 1]].forEach(([v, color, dx]) => {
      const h = (base - top) * v / 5320;
      box(slide, { shape: "Rectangle", l: x + dx, t: base - h, w: 14, h, fill: color });
    });
    text(slide, c, { l: x - 20, t: base + 2, w: 40, h: 12, size: 8, color: XL.headText, align: "Center" });
  });
  [[tr("This year", "本季收入"), C.excel], [tr("Last year", "去年同期"), "#A9D18E"]].forEach(([name, color], i) => {
    const x = o.l + o.w / 2 - 62 + i * 66;
    box(slide, { shape: "Rectangle", l: x, t: o.t + o.h - 12, w: 7, h: 7, fill: color });
    text(slide, name, { l: x + 10, t: o.t + o.h - 14, w: 50, h: 11, size: 7.5, color: XL.headText });
  });
}

const SCENE_EXCEL = {
  name: "excel", stages: 3, draw(s) {
    header(s, tr("Result · Excel", "效果 · Excel"), tr("One request, a native summary sheet and chart", "按一句话整理报表，公式、条件格式和图表都是原生的"));
    request(s, tr("User: sum revenue by city, add YoY, mark over 15% green, insert a column chart", "用户：按城市汇总收入，加一列同比，超过 15% 的标绿，再插一张柱形图"), 116);
    sheetMock(s, {
      l: L, t: 168, w: 290, h: 270, ref: "D2", formula: "1820", cols: tr([46, 52, 46, 62, 62], [52, 34, 44, 72, 66]), activeCol: 3,
      data: XL_RAW, tabs: [tr("Detail", "明细")], activeTab: 0,
      style: (r) => (r === 0 ? { bold: true, fill: XL.head } : null),
    });
    caption(s, "1", tr("User's monthly detail sheet", "用户的明细表（每月流水）"), null, L, 452, 290);
    const mid = L + 324, mw = 184;
    if (at(2)) {
      changes(s, tr("Edits to the original", "AI 在原文件上做的修改"), [tr("New Summary sheet", "新建「汇总」工作表"), tr("SUMIFS by city", "SUMIFS 按城市汇总"), tr("YoY column: =B2/C2-1", "同比列：=B2/C2-1"), tr("Over 15% shaded green", "超过 15% 条件格式标绿"), tr("Native column chart", "插入原生柱形图"), tr("Check: 0 formula errors", "检查：0 个公式错误")], { l: mid, t: 182, w: mw });
    }
    if (!at(3)) return;
    arrowLine(s, mid + mw / 2 - 30, 418, mid + mw / 2 + 30, 418, { color: C.accent, weight: 2.5 });
    const g = sheetMock(s, {
      l: L + 542, t: 168, w: 290, h: 270, ref: "D2", formula: "=B2/C2-1", cols: [50, 70, 70, 76], activeCol: 3, rowHeight: 15,
      data: XL_SUM, tabs: [tr("Detail", "明细"), tr("Summary", "汇总")], activeTab: 1,
      style: (r, c, v) => {
        if (r === 0) return { bold: true, fill: C.excel, color: C.white, align: c === 0 ? "Left" : "Right" };
        if (r === 5) return { bold: true, fill: XL.head };
        if (c === 3 && parseFloat(v) > 15) return { fill: XL.good, color: XL.goodText };
        return null;
      },
    });
    // 选中的单元格
    box(s, { shape: "Rectangle", l: g.xs[3], t: g.cellsT + g.rh, w: 76, h: g.rh, fill: "none", lineColor: C.excel, lineWeight: 2 });
    miniChart(s, { l: L + 542 + 30, t: g.cellsT + 6 * g.rh + 6, w: 252, h: 106 });
    caption(s, "2", tr("After: the new Summary sheet", "修改后：新的「汇总」表"), null, L + 542, 452, 290);
    text(s, tr("Illustration", "示意图"), { l: L + W - 100, t: 504, w: 100, h: 16, size: 10, color: C.mute, align: "Right" });
  },
};

// README 用的效果图（顺序对应 build.sh 里的文件名）
const EFFECTS = [SCENE_TRANSLATE, SCENE_EDIT, SCENE_WORD, SCENE_EXCEL];
