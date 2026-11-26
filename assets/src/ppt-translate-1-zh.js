// 译前：三张卡片的中文产品页
await PowerPoint.run(async (context) => {
  const ls = context.presentation.slideMasters.getItemAt(0).layouts;
  ls.load("items/name,items/id"); await context.sync();
  context.presentation.slides.add({ layoutId: ls.items.find((l) => l.name === "Title Only").id });
  await context.sync();
  const slide = context.presentation.slides.getItemAt(0);
  slide.shapes.load("items/name"); await context.sync();
  slide.shapes.items[0].textFrame.textRange.text = "为什么选择我们";
  slide.shapes.items[0].textFrame.textRange.font.name = "Arial";
  const cards = [
    ["交付快", "标准方案两周上线，提供本地化实施团队"],
    ["成本低", "按需计费，无需一次性采购硬件"],
    ["更安全", "通过等保三级认证，数据不出境"],
  ];
  cards.forEach(([h, d], i) => {
    const l = 60 + i * 290;
    const box = slide.shapes.addGeometricShape("RoundRectangle", { left: l, top: 170, width: 260, height: 260 });
    box.name = `Card ${i + 1}`; box.fill.setSolidColor("#EEF2F7"); box.lineFormat.visible = false;
    const t = slide.shapes.addTextBox(h, { left: l + 20, top: 195, width: 220, height: 44 });
    t.name = `Heading ${i + 1}`; t.textFrame.autoSizeSetting = "AutoSizeNone";
    t.textFrame.textRange.font.size = 28; t.textFrame.textRange.font.name = "Arial"; t.textFrame.textRange.font.bold = true; t.textFrame.textRange.font.color = "#1F3864";
    const b = slide.shapes.addTextBox(d, { left: l + 20, top: 250, width: 220, height: 100 });
    b.name = `Body ${i + 1}`; b.textFrame.autoSizeSetting = "AutoSizeNone"; b.textFrame.wordWrap = true;
    b.textFrame.textRange.font.size = 20; b.textFrame.textRange.font.name = "Arial"; b.textFrame.textRange.font.color = "#333333";
  });
  await context.sync();
});
