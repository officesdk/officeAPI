// 译文填回：只换文字，格式不动
await PowerPoint.run(async (context) => {
  const slide = context.presentation.slides.getItemAt(0);
  slide.shapes.load("items/name"); await context.sync();
  const by = (n) => slide.shapes.items.find((s) => s.name === n);
  slide.shapes.items[0].textFrame.textRange.text = "Why Choose Us";
  const en = [
    ["Fast Delivery", "Standard solutions go live in two weeks, with a local implementation team on site"],
    ["Lower Cost", "Pay-as-you-go pricing with no up-front hardware purchase required"],
    ["More Secure", "Certified to MLPS Level 3, and your data never leaves the country"],
  ];
  en.forEach(([h, d], i) => { by(`Heading ${i + 1}`).textFrame.textRange.text = h; by(`Body ${i + 1}`).textFrame.textRange.text = d; });
  await context.sync();
  const r = context.presentation.ext.inspect({ rules: ["text-overflow"] });
  await context.sync();
  return r.value.issues.map((x) => `${x.shapeName}: ${JSON.stringify(x.detail)}`);
});
