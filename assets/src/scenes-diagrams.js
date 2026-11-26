// 说明图场景：能力概览（只用于视频）、架构、格式矩阵、AI 工作流程。at(n) 控制视频里逐步出现的元素。
ICONS.terminal = "M3 4 H21 V20 H3 Z M7 9 L10 12 L7 15 M12 15 H17";
ICONS.phone = "M7 2 H17 V22 H7 Z M11 18 H13";
ICONS.bot = "M5 8 H19 V20 H5 Z M12 4 V8 M9 13 H9.01 M15 13 H15.01 M9 16.5 H15 M2 13 V16 M22 13 V16";

const PLATFORMS = ["Windows", "macOS", "Linux", "Android", "iOS", "HarmonyOS"];

function card(slide, o) {
  return box(slide, { l: o.l, t: o.t, w: o.w, h: o.h, fill: o.fill || C.card, lineColor: o.lineColor || C.line, radius: o.radius ?? 0.05, name: o.name });
}
function platformRow(slide, l, t) {
  let x = l;
  for (const p of PLATFORMS) {
    const w = p.length * 8 + 28;
    pill(slide, p, x, t, C.accentDark, C.accentTint, w);
    x += w + 8;
  }
  return x;
}

const SCENE_HERO = {
  name: "hero", stages: 5, draw(s) {
    pill(s, "OfficeAPI", L, 80, C.white, C.accent, 96);
    text(s, "一个二进制，读写所有 Office 文档", { l: L, t: 120, w: W, h: 56, size: 38, bold: true });
    text(s, "格式转换 + 为 AI 而生的无界面后端编辑器。不装 Office，不要界面，不连网络。", { l: L, t: 188, w: W, h: 26, size: 17, color: C.mute });
    const cards = [
      ["refresh", tr("Format conversion", "格式转换"), "Word、Excel、PPT、WPS、\rUOF、ODF、PDF 等 33 种格式\r转成 Office、PDF、OFD、图片"],
      ["sparkles", tr("AI editor", "AI 编辑器"), "AI 写标准 Office.js 脚本\r直接在原文件上修改\r看预览、查问题，自己修"],
      ["box", "单一二进制", "没有运行时依赖\r服务器、桌面、手机上都能跑\r内网部署，完全离线"],
    ];
    cards.forEach(([ic, h, d], i) => {
      if (!at(2 + i)) return;
      const l = L + i * 288;
      card(s, { l, t: 248, w: 256, h: 180, fill: C.soft, lineColor: C.soft });
      badge(s, ic, l + 24, 270, 44);
      text(s, h, { l: l + 24, t: 328, w: 208, h: 26, size: 19, bold: true });
      text(s, d, { l: l + 24, t: 360, w: 208, h: 64, size: 12.5, color: C.mute, lineSpacing: 1.25 });
    });
    if (at(5)) {
      text(s, "运行在", { l: L, t: 459, w: 56, h: 20, size: 13, color: C.mute });
      platformRow(s, L + 56, 457);
    }
  },
};

const SCENE_ARCH = {
  name: "arch", stages: 3, draw(s) {
    header(s, tr("Architecture", "架构"), tr("One process, two capabilities, one interface", "一个进程，两类能力，一套接口"));
    const callers = [["bot", "AI Agent", tr("Claude, agent platforms", "Claude、企业 Agent 平台")], ["server", tr("Business systems", "业务系统"), tr("Batch jobs, pipelines", "批处理、文档流水线")], ["phone", tr("Desktop & mobile", "桌面与移动 App"), tr("Offline, on the device", "在用户设备上离线处理")]];
    callers.forEach(([ic, h, d], i) => {
      const t = 150 + i * 84;
      card(s, { l: L, t, w: 184, h: 64 });
      icon(s, ic, L + 14, t + 20, 24, C.accent);
      text(s, h, { l: L + 48, t: t + 13, w: 130, h: 20, size: 14, bold: true });
      text(s, d, { l: L + 48, t: t + 35, w: 130, h: 18, size: 10.5, color: C.mute });
      arrowLine(s, L + 184, t + 32, 290, t + 32, { color: C.accentLight });
    });
    box(s, { l: 290, t: 128, w: 384, h: 276, fill: C.soft, lineColor: C.accent, lineWeight: 2, radius: 0.04 });
    text(s, tr("OfficeAPI (single binary)", "OfficeAPI（单个二进制）"), { l: 306, t: 140, w: 352, h: 22, size: 15, bold: true, color: C.accentDark });
    label(box(s, { l: 306, t: 168, w: 352, h: 26, fill: C.accentTint, radius: 0.2 }), tr("Interfaces: MCP (stdio) · CLI", "接口：MCP（stdio） · 命令行"), { size: 12, color: C.accentDark });
    const panels = [
      [tr("Format conversion", "格式转换"), [tr("Reads 33 input formats", "读取 33 种输入格式"), tr("Unified document model", "统一的文档模型"), tr("Writes Office · PDF · OFD", "Office · PDF · OFD 写出"), tr("Renders PNG · SVG", "页面渲染 PNG · SVG")]],
      [tr("AI editor", "AI 编辑器"), [tr("Office.js script sandbox", "Office.js 脚本沙箱"), tr("Sessions, full rollback", "会话与整段回滚"), tr("Layout and preview", "排版与预览渲染"), tr("Layout checks (inspect)", "版面检查（inspect）")]],
    ];
    panels.forEach(([h, items], i) => {
      const l = 306 + i * 180;
      box(s, { l, t: 204, w: 172, h: 186, fill: C.card, radius: 0.05 });
      text(s, h, { l: l + 12, t: 214, w: 148, h: 20, size: 13.5, bold: true });
      items.forEach((str, j) => label(box(s, { l: l + 12, t: 242 + j * 36, w: 148, h: 28, fill: C.soft, lineColor: C.line, radius: 0.15 }), str, { size: 11, bold: false }));
    });
    if (at(2)) {
      const outs = [["file", "docx · xlsx · pptx", tr("Native, editable", "原生、可编辑")], ["layers", "PDF · OFD", tr("Vector / raster PDF, OFD", "矢量 / 纯图 PDF，版式文档")], ["image", "PNG · SVG", tr("Page previews", "页面预览")], ["list", tr("Check report (JSON)", "检查报告（JSON）"), tr("Per-element issue list", "精确到元素的问题清单")]];
      outs.forEach(([ic, h, d], i) => {
        const t = 134 + i * 68;
        card(s, { l: 716, t, w: 180, h: 56 });
        icon(s, ic, 728, t + 16, 22, C.ok);
        text(s, h, { l: 758, t: t + 10, w: 134, h: 18, size: 12, bold: true });
        text(s, d, { l: 758, t: t + 30, w: 134, h: 16, size: 10.5, color: C.mute });
        arrowLine(s, 674, t + 28, 716, t + 28, { color: C.okDark });
      });
    }
    if (at(3)) {
      box(s, { l: L, t: 430, w: W, h: 50, fill: C.soft, radius: 0.15 });
      text(s, tr("Zero deps, one engine", "无运行时依赖，同一个引擎"), { l: L + 20, t: 446, w: 190, h: 20, size: 13, bold: true });
      platformRow(s, L + 220, 443);
    }
  },
};

const SCENE_FORMATS = {
  name: "formats", stages: 3, draw(s) {
    header(s, tr("Format conversion", "格式转换"), tr("33 input formats; Office, fixed-layout and image outputs", "33 种输入格式，Office、版式文档和图片三类输出"));
    const suites = [
      ["Word", C.word, C.wordTint, C.wordLight, ["doc", "docx", "dot", "dotx", "docm", "dotm", "wps", "wpt", "uot", "odt", "md", "txt", "pdf"], ["docx", tr("Vector PDF", "矢量 PDF"), tr("Raster PDF", "纯图 PDF"), "ofd", "png", "svg", "odt"]],
      ["Excel", C.excel, C.excelTint, C.excelLight, ["xls", "xlsx", "xlt", "xltx", "xlsm", "et", "ett", "uos", "ods", "csv", "pdf"], ["xlsx", tr("Vector PDF", "矢量 PDF"), tr("Raster PDF", "纯图 PDF"), "ofd", "png", "svg", "ods"]],
      ["PowerPoint", C.ppt, C.pptTint, C.pptLight, ["ppt", "pptx", "pot", "potx", "pptm", "potm", "dps", "dpt", "uop", "odp", "pdf"], ["pptx", tr("Vector PDF", "矢量 PDF"), tr("Raster PDF", "纯图 PDF"), "ofd", "png", "svg", "odp"]],
    ];
    suites.forEach(([name, fg, tint, light, ins, outs], i) => {
      if (!at(1 + i)) return;
      const l = L + i * 288, w = 256;
      card(s, { l, t: 124, w, h: 360 });
      label(box(s, { l, t: 124, w, h: 40, fill: fg, radius: 0.15 }), name, { size: 16, color: C.white });
      text(s, tr("Input", "输入"), { l: l + 16, t: 176, w: 100, h: 18, size: 12, bold: true, color: C.mute });
      ins.forEach((str, j) => pill(s, str, l + 16 + (j % 4) * 58, 198 + Math.floor(j / 4) * 28, light, tint, 50));
      arrowLine(s, l + w / 2, 314, l + w / 2, 340, { color: light, weight: 2 });
      text(s, tr("Output", "输出"), { l: l + 16, t: 346, w: 100, h: 18, size: 12, bold: true, color: C.mute });
      outs.forEach((str, j) => pill(s, str, l + 16 + (j % 3) * 75, 368 + Math.floor(j / 3) * 30, C.white, fg, tr(73, 68)));
    });
    if (at(3)) foot(s, tr("Vector PDF keeps text and graphics: searchable, copyable, lossless zoom. Raster PDF makes each page an image; text can't be selected. WPS, UOF and OFD are native.", "矢量 PDF 保留文字和图形，可搜索、可复制、无损缩放；纯图 PDF 每页是一张图片，文字不能被选中或提取。WPS、UOF（标文通）、OFD 等国产格式原生支持。"));
  },
};

const SCENE_LOOP = {
  name: "loop", stages: 4, draw(s) {
    header(s, tr("AI editor", "AI 编辑器"), tr("Edit, look, check, fix: the AI reviews its own work", "改、看、查、再改：AI 自己把关"));
    const steps = [
      ["1", "open", tr("Open the original file\rget a structure summary", "打开原文件\r返回结构摘要")],
      ["2", "run", tr("Run an Office.js script\rrolled back on error", "执行 Office.js 脚本\r出错整段撤销")],
      ["3", "render · inspect", tr("Preview and issue list\rdown to page and element", "出预览图和问题清单\r精确到页和元素")],
      ["4", "save", tr("No issues left\rsave as native Office file", "没有问题\r保存为原生 Office 文件")],
    ];
    steps.forEach(([n, tool, d], i) => {
      if (!at(i < 2 ? 1 : 2)) return;
      const l = L + i * 216;
      card(s, { l, t: 124, w: 184, h: 84, fill: i === 2 ? C.accentTint : C.card, lineColor: i === 2 ? C.accentLight : C.line });
      label(box(s, { shape: "Ellipse", l: l + 12, t: 136, w: 22, h: 22, fill: C.accentFill }), n, { size: 11, color: C.white });
      text(s, tool, { l: l + 42, t: 137, w: 136, h: 20, size: 13.5, bold: true, font: MONO, color: C.accentDark });
      text(s, d, { l: l + 12, t: 166, w: 160, h: 36, size: 11.5, color: C.mute });
      if (i < 3 && at(i < 1 ? 1 : 2)) arrowLine(s, l + 184, 166, l + 216, 166, { color: C.accent, weight: 2 });
    });
    if (at(3)) {
      // 回路：3 → 2
      const y0 = 208, y1 = 232, xFrom = L + 2 * 216 + 92, xTo = L + 216 + 92;
      arrowLine(s, xFrom, y0, xFrom, y1, { color: C.warn, weight: 1.75, dash: "Dash" }).lineFormat.ext.endArrowheadStyle = "None";
      arrowLine(s, xFrom, y1, xTo, y1, { color: C.warn, weight: 1.75, dash: "Dash" }).lineFormat.ext.endArrowheadStyle = "None";
      arrowLine(s, xTo, y1, xTo, y0, { color: C.warn, weight: 1.75, dash: "Dash" });
      label(box(s, { l: xTo + 30, t: y1 - 11, w: tr(196, 156), h: 22, fill: C.warnTint, radius: 0.5 }), tr("Issues: fix per hint, re-check", "有问题：按 hint 修改，再查"), { size: 11, color: C.warn });
    }
    if (!at(4)) return;
    const code = [
      "await PowerPoint.run(async (context) => {",
      "  const slide = context.presentation",
      "    .slides.getItemAt(2);",
      "  const title = slide.shapes.getItem(\"Title 1\");",
      "  title.textFrame.textRange.text =",
      tr("    \"Q3 2026: overseas revenue up 42% year on year\";", "    \"2026 年第三季度：海外收入同比增长 42%\";"),
      "  await context.sync();",
      "});",
    ].join("\r");
    const json = [
      "{ \"rule\": \"text-overflow\",",
      "  \"severity\": \"error\",",
      "  \"slideIndex\": 2,",
      "  \"shapeName\": \"Title 1\",",
      "  \"detail\": { \"overflowPt\": 18.4,",
      "              \"lines\": 3, \"fitsLines\": 2 },",
      "  \"hint\": \"Shorten the text, reduce",
      "           font.size, or autofit\" }",
    ].join("\r");
    [[tr("② The AI's script: standard Office.js", "② AI 写的脚本：标准 Office.js"), code], [tr("③ What inspect returns", "③ inspect 返回的问题"), json]].forEach(([h, body], i) => {
      const l = L + i * 428;
      box(s, { l, t: 262, w: 404, h: 196, fill: C.code, lineColor: C.line, radius: 0.04 });
      text(s, h, { l: l + 20, t: 278, w: 364, h: 18, size: 12, bold: true, color: C.accentLight });
      text(s, body, { l: l + 20, t: 306, w: 372, h: 142, size: 11.5, font: MONO, color: "#E8EBF2", lineSpacing: 1.2 });
    });
  },
};

// README 用的说明图（顺序对应 build.sh 里的文件名）
const DIAGRAMS = [SCENE_ARCH, SCENE_FORMATS, SCENE_LOOP];
