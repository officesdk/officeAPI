// 按检查结果修：有溢出就把三张卡的正文一起缩小 1 磅（保持一致），直到没有溢出
await PowerPoint.run(async (context) => {
  const slide = context.presentation.slides.getItemAt(0);
  slide.shapes.load("items/name"); await context.sync();
  const bodies = slide.shapes.items.filter((s) => s.name.startsWith("Body "));
  const log = [];
  for (let round = 1; round <= 8; round++) {
    const r = context.presentation.ext.inspect({ rules: ["text-overflow"] });
    const f = bodies[0].textFrame.textRange.font; f.load("size");
    await context.sync();
    if (!r.value.issues.length) { log.push(`round ${round}: clean at ${f.size} pt`); break; }
    log.push(`round ${round}: ${r.value.issues.length} overflow(s), ${f.size} -> ${f.size - 1} pt`);
    for (const b of bodies) b.textFrame.textRange.font.size = f.size - 1;
    await context.sync();
  }
  return log;
});
