// 给溢出的文本框画红框（只用于截图）
await PowerPoint.run(async (context) => {
  const slide = context.presentation.slides.getItemAt(0);
  const r = context.presentation.ext.inspect({ rules: ["text-overflow"] });
  slide.shapes.load("items/name,items/id,items/left,items/top,items/width,items/height"); await context.sync();
  for (const x of r.value.issues) {
    const s = slide.shapes.items.find((s) => s.name === x.shapeName);
    const h = s.height + (x.detail.overflowPt || 0);
    const m = slide.shapes.addGeometricShape("Rectangle", { left: s.left - 4, top: s.top - 4, width: s.width + 8, height: h + 8 });
    m.fill.clear(); m.lineFormat.color = "#D93025"; m.lineFormat.weight = 2.5; m.lineFormat.dashStyle = "Dash";
  }
  await context.sync();
  return r.value.issues.length;
});
