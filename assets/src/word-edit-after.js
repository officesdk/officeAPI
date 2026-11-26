// 用户：「表格改成三线表，数字右对齐，标出同比最高的城市；总体情况后面加一段要点」
// 在原文件上改：标题、页眉、列表、其他段落不动
await Word.run(async (context) => {
  const body = context.document.body;
  const t = body.tables.getFirst();
  t.load("values,rowCount");
  const paras = body.paragraphs;
  paras.load("items/text");
  await context.sync();

  // 三线表：顶线、底线粗，表头下一条细线，其余边框去掉
  for (const side of ["Left", "Right", "InsideVertical", "InsideHorizontal"]) t.getBorder(side).type = "None";
  for (const side of ["Top", "Bottom"]) { const b = t.getBorder(side); b.type = "Single"; b.width = 1.5; b.color = "#1F3864"; }
  const head = t.rows.getFirst();
  head.font.bold = true;
  head.font.color = "#1F3864";
  for (let c = 0; c < 6; c++) {
    const b = t.getCell(0, c).getBorder("Bottom"); b.type = "Single"; b.width = 0.75; b.color = "#1F3864";
  }
  // 数字列右对齐
  for (let r = 0; r < t.rowCount; r++) for (let c = 1; c < 6; c++) t.getCell(r, c).horizontalAlignment = "Right";
  // 同比最高的城市：整行浅色底、同比加粗标红
  const yoy = t.values.slice(1).map((row) => parseFloat(row[2]));
  const best = 1 + yoy.indexOf(Math.max(...yoy));
  for (let c = 0; c < 6; c++) t.getCell(best, c).shadingColor = "#FDF1DB";
  const cell = t.getCell(best, 2).body;
  cell.font.bold = true; cell.font.color = "#C00000";

  // 总体情况后面加要点
  const summary = paras.items.find((p) => p.text.startsWith(tr("In Q3 the East Region", "第三季度华东区")));
  const intro = summary.insertParagraph(tr("Key points:", "要点："), "After");
  intro.font.bold = true;
  const items = tr(["Revenue ¥128M, +18% YoY, 6% above target", "Renewal rate 92%, +3 points on Q2", "Gross margin 41%, −2 points on hardware costs"], ["收入 1.28 亿元，同比 +18%，超目标 6%", "续约率 92%，环比 +3 个百分点", "毛利率 41%，受硬件成本影响 −2 个百分点"]);
  let last = intro.insertParagraph(items[0], "After");
  last.font.bold = false;
  const list = last.startNewList();
  for (const s of items.slice(1)) last = last.insertParagraph(s, "After");
  await context.sync();
  const r = context.document.ext.inspect();
  await context.sync();
  return { best: t.values[best][0], issues: r.value.summary };
});
