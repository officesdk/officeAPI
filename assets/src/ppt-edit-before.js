// 示例「客户文件」：一页单栏的经营回顾（用 Office 主题的标题和内容版式）
await PowerPoint.run(async (context) => {
  const ls = context.presentation.slideMasters.getItemAt(0).layouts;
  ls.load("items/name,items/id"); await context.sync();
  context.presentation.slides.add({ layoutId: ls.items.find((l) => l.name === "Title and Content").id });
  await context.sync();
  const slide = context.presentation.slides.getItemAt(0);
  slide.shapes.load("items/name"); await context.sync();
  const [title, body] = slide.shapes.items;
  title.textFrame.textRange.text = tr("East Region: Q3 2026 Review", "华东区 2026 Q3 经营回顾");
  body.textFrame.textRange.text = tr([
    "Revenue ¥128M, up 18% YoY, 6% above target",
    "46 new clients, 19 of them in finance",
    "Renewal rate 92%, up 3 points on Q2",
    "Gross margin 41%, down 2 points on hardware costs",
    "Q4 focus: grow delivery teams in Suzhou and Hefei",
    "Risk: two large clients pushed budgets to Q1",
  ], [
    "收入 1.28 亿元，同比增长 18%，超出目标 6%",
    "新签客户 46 家，其中金融行业 19 家",
    "续约率 92%，较上季度提升 3 个百分点",
    "毛利率 41%，受硬件成本上涨影响下降 2 个百分点",
    "Q4 重点：苏州、合肥两地交付团队扩编",
    "风险：两家大客户预算审批推迟到明年一季度",
  ]).join("\r");
  title.textFrame.textRange.font.name = "Arial"; body.textFrame.textRange.font.name = "Arial";
  const bar = slide.shapes.addGeometricShape("Rectangle", { left: 0, top: 0, width: 960, height: 8 });
  bar.name = "Brand Bar"; bar.fill.setSolidColor("#C0392B"); bar.lineFormat.visible = false;
  await context.sync();
});
