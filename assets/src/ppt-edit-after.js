// 「把这页改成两栏，右边放季度收入图」：在原文件上改，标题、版式、品牌色条不动
await PowerPoint.run(async (context) => {
  const slide = context.presentation.slides.getItemAt(0);
  slide.shapes.load("items/name,items/left,items/top,items/width,items/height"); await context.sync();
  const body = slide.shapes.items.find((s) => s.name.startsWith("Content Placeholder"));
  const right = body.left + body.width;
  body.width = body.width * 0.5 - 12;
  body.textFrame.textRange.font.size = 18;
  slide.shapes.ext.addChart("ColumnClustered",
    { categories: ["Q1", "Q2", "Q3"], series: [{ name: tr("Revenue (¥M)", "收入（百万元）"), values: [96, 108, 128] }] },
    { left: body.left + body.width + 24, top: body.top, width: right - body.left - body.width - 24, height: body.height, title: tr("Quarterly revenue (¥M)", "季度收入（百万元）"), legendPosition: "None" });
  await context.sync();
});
