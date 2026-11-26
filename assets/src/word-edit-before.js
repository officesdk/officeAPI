// 示例「用户文件」：一份普通的经营报告（默认表格样式）
await Word.run(async (context) => {
  const doc = context.document;
  const body = doc.body;
  const title = body.insertParagraph(tr("East Region Q3 2026 Business Report", "华东区 2026 年第三季度经营报告"), "Start");
  title.styleBuiltIn = Word.BuiltInStyleName.title;
  if (LANG === "zh") title.font.name = "DengXian Light";
  const p0 = body.insertParagraph(tr("Prepared by East Region Operations · October 12, 2026", "编制：华东区运营部　　日期：2026 年 10 月 12 日"), "End");
  p0.styleBuiltIn = Word.BuiltInStyleName.subtitle;

  body.insertParagraph(tr("1. Overview", "一、总体情况"), "End").styleBuiltIn = Word.BuiltInStyleName.heading1;
  const p1 = body.insertParagraph(tr("In Q3 the East Region earned ¥128M in revenue, up 18% year on year and 6% above target. It signed 46 new clients, 19 of them in finance, and renewals reached 92%, up 3 points on Q2. Higher hardware costs cut gross margin by 2 points to 41%.", "第三季度华东区实现收入 1.28 亿元，同比增长 18%，超出季度目标 6%。新签客户 46 家，其中金融行业 19 家；续约率 92%，较上季度提升 3 个百分点。受硬件成本上涨影响，毛利率为 41%，下降 2 个百分点。"), "End");
  p1.styleBuiltIn = Word.BuiltInStyleName.normal;

  body.insertParagraph(tr("2. Results by city", "二、分城市业绩"), "End").styleBuiltIn = Word.BuiltInStyleName.heading1;
  const t = body.insertTable(5, 6, "End", tr([
    ["City", "Revenue (¥10k)", "YoY", "New clients", "Renewal", "Margin"],
    ["Shanghai", "5,320", "+21%", "18", "94%", "43%"],
    ["Hangzhou", "3,180", "+16%", "12", "91%", "40%"],
    ["Nanjing", "2,460", "+15%", "9", "90%", "39%"],
    ["Suzhou", "1,840", "+19%", "7", "93%", "41%"],
  ], [
    ["城市", "收入（万元）", "同比", "新签客户", "续约率", "毛利率"],
    ["上海", "5,320", "+21%", "18", "94%", "43%"],
    ["杭州", "3,180", "+16%", "12", "91%", "40%"],
    ["南京", "2,460", "+15%", "9", "90%", "39%"],
    ["苏州", "1,840", "+19%", "7", "93%", "41%"],
  ]));
  t.headerRowCount = 1;
  for (let r = 0; r < 5; r++) tr([66, 96, 60, 78, 72, 66], [60, 96, 78, 78, 78, 78]).forEach((w, c) => (t.getCell(r, c).columnWidth = w));

  body.insertParagraph(tr("3. Next quarter", "三、下季度重点"), "End").styleBuiltIn = Word.BuiltInStyleName.heading1;
  const first = body.insertParagraph(tr("Grow delivery teams in Suzhou and Hefei: 24 hires by year end", "苏州、合肥两地交付团队扩编，年底前到岗 24 人"), "End");
  first.styleBuiltIn = Word.BuiltInStyleName.normal;
  const list = first.startNewList();
  list.insertParagraph(tr("Move two large clients through budget approval and sign in Q1", "推进两家大客户的预算审批，争取在一季度签约"), "End");
  list.insertParagraph(tr("Switch hardware buying to a framework deal that locks H1 prices", "硬件采购改为框架协议，锁定明年上半年价格"), "End");

  doc.sections.getFirst().getHeader("Primary").insertParagraph(tr("East Region Operations · Internal", "华东区运营部 · 内部资料"), "Start");
  await context.sync();
  const r = doc.ext.inspect();
  await context.sync();
  return r.value;
});
