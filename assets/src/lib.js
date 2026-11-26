// 说明图、效果图和视频帧共用的绘图函数，由 build.sh 拼在各脚本前面（pptjs 脚本不能 import）。
// 改编自 pptjs 的 decks/pptjs-product-intro/build.js。
const C = {
  // 与介绍视频同一套视觉（officekit-promo/src/look.ts）：石墨暗场、纸白文档是场上唯一亮物、机器蓝做强调、Office 三色只标三件套
  stage0: "#0B0E14", stage1: "#161B25", fg: "#F2F4F8", mute: "#A9B3C4", line: "#2A3241", soft: "#1B212C", card: "#222937", white: "#FFFFFF",
  ink: "#1D2330", code: "#0F1218",
  accent: "#4C7DFF", accentFill: "#2F5BE0", accentTint: "#1D2A4A", accentDark: "#9DB8FF", accentLight: "#7FA2FF",
  ok: "#3FB97A", okDark: "#5FD394", okTint: "#173326", err: "#E5534B", errTint: "#3A1E1E", warn: "#E0A43A", warnTint: "#3A2C14",
  word: "#2B579A", wordLight: "#7FA6E6", wordTint: "#1A2740", excel: "#217346", excelLight: "#5CC08A", excelTint: "#16301F",
  ppt: "#C4441F", pptLight: "#F08A66", pptTint: "#3A1D14",
};
// 语言：build.sh 在拼接时最前面写入 const LANG（"en" | "zh"）；没写时按中文
const tr = (en, zh) => (typeof LANG !== "undefined" && LANG === "en" ? en : zh);
const FONT = tr("Arial", "Microsoft YaHei"), MONO = "Menlo";
const L = 64, W = 832;

// 新建 n 张空白页，返回这 n 张 slide
async function blankSlides(context, n) {
  const layouts = context.presentation.slideMasters.getItemAt(0).layouts;
  layouts.load("items/name,items/id");
  await context.sync();
  const blank = layouts.items.find((l) => l.name === "Blank").id;
  for (let i = 0; i < n; i++) context.presentation.slides.add({ layoutId: blank });
  await context.sync();
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides.items.slice(-n);
}

function text(slide, str, o) {
  const s = slide.shapes.addTextBox(str, { left: o.l, top: o.t, width: o.w, height: o.h });
  if (o.name) s.name = o.name;
  const tf = s.textFrame;
  tf.autoSizeSetting = "AutoSizeNone";
  tf.wordWrap = true;
  tf.leftMargin = tf.rightMargin = tf.topMargin = tf.bottomMargin = 0;
  tf.verticalAlignment = o.valign || "Top";
  const r = tf.textRange;
  r.font.name = o.font || FONT;
  r.font.size = o.size || 16;
  r.font.color = o.color || C.fg;
  r.font.bold = !!o.bold;
  r.paragraphFormat.horizontalAlignment = o.align || "Left";
  if (o.lineSpacing) r.paragraphFormat.ext.lineSpacing = o.lineSpacing;
  if (o.spaceAfter) r.paragraphFormat.ext.spaceAfter = o.spaceAfter;
  // 加粗或改色的片段：[[子串, { bold, color }], ...]
  for (const [sub, f] of o.marks || []) {
    const i = str.indexOf(sub);
    if (i < 0) throw new Error(`mark not found: ${sub}`);
    const rr = r.getSubstring(i, sub.length);
    if (f.bold !== undefined) rr.font.bold = f.bold;
    if (f.color) rr.font.color = f.color;
  }
  return s;
}
function box(slide, o) {
  const shape = o.shape || "RoundRectangle";
  const s = slide.shapes.addGeometricShape(shape, { left: o.l, top: o.t, width: o.w, height: o.h });
  if (o.name) s.name = o.name;
  if (o.fill === "none") s.fill.clear(); else s.fill.setSolidColor(o.fill || C.soft);
  if (o.transparency) s.fill.transparency = o.transparency;
  if (o.lineColor) {
    s.lineFormat.color = o.lineColor;
    s.lineFormat.weight = o.lineWeight || 1;
    if (o.dash) s.lineFormat.dashStyle = o.dash;
  } else s.lineFormat.visible = false;
  if (shape === "RoundRectangle") s.adjustments.set(0, o.radius ?? 0.06);
  return s;
}
// 形状里的文字
function label(s, str, o) {
  const tf = s.textFrame;
  tf.leftMargin = tf.rightMargin = o.margin ?? 0;
  tf.topMargin = tf.bottomMargin = 0;
  tf.verticalAlignment = "Middle";
  tf.wordWrap = true;
  const r = tf.textRange;
  r.text = str;
  r.font.name = FONT; r.font.size = o.size || 12; r.font.bold = o.bold ?? true; r.font.color = o.color || C.fg;
  r.paragraphFormat.horizontalAlignment = o.align || "Center";
  return s;
}
function rule(slide, l, t, w, color) {
  const s = slide.shapes.addLine("Straight", { left: l, top: t, width: w, height: 0 });
  s.height = 0;
  s.lineFormat.color = color || C.line;
  s.lineFormat.weight = 1;
  return s;
}
// 带箭头的直线，从 (x1, y1) 指向 (x2, y2)
function arrowLine(slide, x1, y1, x2, y2, o = {}) {
  const s = slide.shapes.addLine("Straight", { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) });
  if (y1 === y2) s.height = 0;
  if (x1 === x2) s.width = 0;
  if (o.name) s.name = o.name;
  s.lineFormat.color = o.color || C.mute;
  s.lineFormat.weight = o.weight || 1.5;
  if (o.dash) s.lineFormat.dashStyle = o.dash;
  s.lineFormat.ext.endArrowheadStyle = "Triangle";
  if (x2 < x1) s.ext.flipHorizontal = true;
  if (y2 < y1) s.ext.flipVertical = true;
  return s;
}
function pill(slide, str, l, t, fg, bg, w) {
  const s = box(slide, { l, t, w: w || str.length * 13 + 24, h: 24, fill: bg, radius: 0.5 });
  return label(s, str, { color: fg, size: 12 });
}
function header(slide, kicker, title) {
  text(slide, kicker, { l: L, t: 40, w: W, h: 20, size: 14, bold: true, color: C.accent, name: "Kicker" });
  text(slide, title, { l: L, t: 64, w: W, h: 44, size: 28, bold: true, name: "Title" });
}
function foot(slide, str) {
  return text(slide, str, { l: L, t: 504, w: W, h: 16, size: 10, color: C.mute, name: "Footnote" });
}
function notes(slide, str) { slide.ext.notes.textRange.text = str; }
function transition(slide, effect, duration) {
  slide.ext.transition.effect = effect;
  slide.ext.transition.duration = duration;
}

// ---- 图标：24×24 网格上的线条路径（M L H V C Z），转成 PowerPoint 的自定义几何，在 PowerPoint 里可以改色 ----
const circ = (x, y, r) => {
  const k = 0.5523 * r;
  return `M${x} ${y - r} C${x + k} ${y - r} ${x + r} ${y - k} ${x + r} ${y} C${x + r} ${y + k} ${x + k} ${y + r} ${x} ${y + r} ` +
    `C${x - k} ${y + r} ${x - r} ${y + k} ${x - r} ${y} C${x - r} ${y - k} ${x - k} ${y - r} ${x} ${y - r} Z `;
};
const ICONS = {
  file: "M14 2 H6 C4.9 2 4 2.9 4 4 V20 C4 21.1 4.9 22 6 22 H18 C19.1 22 20 21.1 20 20 V8 Z M14 2 V8 H20 M8 13 H16 M8 17 H13",
  layers: "M12 2 L22 7 L12 12 L2 7 Z M2 12 L12 17 L22 12 M2 17 L12 22 L22 17",
  search: circ(11, 11, 7) + "M21 21 L16.5 16.5",
  devices: "M2 4 H16 V14 H2 Z M6 18 H12 M9 14 V18 M18 8 H22 V20 H18 Z",
  grid: "M3 3 H21 V21 H3 Z M3 9 H21 M9 9 V21",
  sparkles: "M12 3 L13.8 10.2 L21 12 L13.8 13.8 L12 21 L10.2 13.8 L3 12 L10.2 10.2 Z M19 2 V6 M17 4 H21",
  user: circ(12, 8, 4) + "M4 21 C4 16.6 7.6 14 12 14 C16.4 14 20 16.6 20 21",
  server: "M3 3 H21 V10 H3 Z M3 14 H21 V21 H3 Z M7 6.5 H7.01 M7 17.5 H7.01",
  code: "M16 18 L22 12 L16 6 M8 6 L2 12 L8 18",
  eye: "M2 12 C5 6 9 4 12 4 C15 4 19 6 22 12 C19 18 15 20 12 20 C9 20 5 18 2 12 Z " + circ(12, 12, 3),
  globe: circ(12, 12, 10) + "M2 12 H22 M12 2 C15 5 16 8.5 16 12 C16 15.5 15 19 12 22 C9 19 8 15.5 8 12 C8 8.5 9 5 12 2 Z",
  box: "M12 2 L21 7 V17 L12 22 L3 17 V7 Z M3 7 L12 12 L21 7 M12 12 V22",
  zap: "M13 2 L4 14 H12 L11 22 L20 10 H12 Z",
  image: "M3 3 H21 V21 H3 Z " + circ(8.5, 8.5, 1.5) + "M21 15 L16 10 L5 21",
  lock: "M5 11 H19 V22 H5 Z M8 11 V7 C8 4.8 9.8 3 12 3 C14.2 3 16 4.8 16 7 V11",
  check: circ(12, 12, 10) + "M8 12 L11 15 L16 9",
  cross: circ(12, 12, 10) + "M15 9 L9 15 M9 9 L15 15",
  chart: "M4 20 V10 M10 20 V4 M16 20 V13 M2 20 H22",
  list: "M10 6 H21 M10 12 H21 M10 18 H21 M3 6 L4.5 7.5 L7 5 M3 12 L4.5 13.5 L7 11 M3 18 L4.5 19.5 L7 17",
  scale: "M12 3 V21 M7 21 H17 M5 7 H19 M5 7 L2 14 C2 15.7 3.3 17 5 17 C6.7 17 8 15.7 8 14 L5 7 M19 7 L16 14 C16 15.7 17.3 17 19 17 C20.7 17 22 15.7 22 14 L19 7",
  monitor: "M3 4 H21 V16 H3 Z M8 20 H16 M12 16 V20",
  plug: "M9 2 V6 M15 2 V6 M6 6 H18 V10 C18 13.3 15.3 16 12 16 C8.7 16 6 13.3 6 10 Z M12 16 V22",
  shield: "M12 2 L20 5 V11 C20 16 16.5 20 12 22 C7.5 20 4 16 4 11 V5 Z M9 12 L11 14 L15 10",
  refresh: "M21 12 C21 7 17 3 12 3 C9.2 3 6.7 4.3 5 6.3 M3 12 C3 17 7 21 12 21 C14.8 21 17.3 19.7 19 17.7 M5 2 V6.3 H9.3 M19 22 V17.7 H14.7",
  alert: "M12 3 L22 20 H2 Z M12 9 V13 M12 17 H12.01",
  cpu: "M5 7 H19 V20 H5 Z M12 3 V7 M9 12 H9.01 M15 12 H15.01 M9 16 H15 M2 12 V15 M22 12 V15",
};
function pathXml(d) {
  const toks = d.match(/[MLCHVZ]|-?[\d.]+/g);
  const pt = (x, y) => `<a:pt x="${Math.round(x * 100)}" y="${Math.round(y * 100)}"/>`;
  let out = "", i = 0, cx = 0, cy = 0, cmd = "";
  while (i < toks.length) {
    if (/[MLCHVZ]/.test(toks[i])) cmd = toks[i++];
    if (cmd === "M") { cx = +toks[i++]; cy = +toks[i++]; out += `<a:moveTo>${pt(cx, cy)}</a:moveTo>`; cmd = "L"; }
    else if (cmd === "L") { cx = +toks[i++]; cy = +toks[i++]; out += `<a:lnTo>${pt(cx, cy)}</a:lnTo>`; }
    else if (cmd === "H") { cx = +toks[i++]; out += `<a:lnTo>${pt(cx, cy)}</a:lnTo>`; }
    else if (cmd === "V") { cy = +toks[i++]; out += `<a:lnTo>${pt(cx, cy)}</a:lnTo>`; }
    else if (cmd === "C") {
      const a = toks.slice(i, i + 6).map(Number);
      i += 6;
      out += `<a:cubicBezTo>${pt(a[0], a[1])}${pt(a[2], a[3])}${pt(a[4], a[5])}</a:cubicBezTo>`;
      cx = a[4]; cy = a[5];
    } else if (cmd === "Z") { out += "<a:close/>"; cmd = ""; }
  }
  return out;
}
const EMU = (pt) => Math.round(pt * 12700);
function icon(slide, key, l, t, size, color, name) {
  const s = slide.shapes.addGeometricShape("Rectangle", { left: l, top: t, width: size, height: size });
  s.ext.setXml(
    `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<p:nvSpPr><p:cNvPr id="1" name="${name || `Icon ${key}`}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>` +
    `<a:xfrm><a:off x="${EMU(l)}" y="${EMU(t)}"/><a:ext cx="${EMU(size)}" cy="${EMU(size)}"/></a:xfrm>` +
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
    `<a:pathLst><a:path w="2400" h="2400" fill="none">${pathXml(ICONS[key])}</a:path></a:pathLst></a:custGeom>` +
    `<a:noFill/><a:ln w="${EMU(size / 24 * 1.9)}" cap="rnd"><a:solidFill><a:srgbClr val="${(color || C.accent).slice(1)}"/></a:solidFill><a:round/></a:ln>` +
    `</p:spPr></p:sp>`);
  return s;
}
// 圆形底色上的图标
function badge(slide, key, l, t, size, fg, bg, name) {
  return [
    box(slide, { shape: "Ellipse", l, t, w: size, h: size, fill: bg || C.accentTint, name: name && `${name} Badge` }),
    icon(slide, key, l + size * 0.22, t + size * 0.22, size * 0.56, fg || C.accent, name && `${name} Icon`),
  ];
}
// 截图：细边框 + 柔和阴影；o.crop 是各边裁掉的比例（0–1）
function pic(slide, key, o) {
  const pos = { left: o.l, top: o.t, width: o.w, height: o.h };
  const s = o.crop ? slide.shapes.ext.addPicture(FILES[key], { ...pos, crop: o.crop }) : slide.shapes.addPicture(FILES[key], pos);
  s.name = o.name || key;
  s.lineFormat.color = C.line;
  s.lineFormat.weight = 1;
  const sh = s.ext.shadow;
  sh.visible = true; sh.color = "#000000"; sh.transparency = 0.45; sh.blur = 22; sh.distance = 8; sh.angle = 90;
  if (o.rotation) s.rotation = o.rotation;
  return s;
}

// 暗场：上亮下暗的石墨底 + 顶部一抹极淡的机器蓝主光（同视频 Stage）
function stage(slide) {
  const bg = slide.shapes.addGeometricShape("Rectangle", { left: 0, top: 0, width: 960, height: 540 });
  bg.name = "Stage"; bg.lineFormat.visible = false;
  bg.fill.ext.setGradient({ type: "Linear", angle: 90, stops: [{ position: 0, color: C.stage1 }, { position: 1, color: C.stage0 }] });
  const key = slide.shapes.addGeometricShape("Rectangle", { left: 0, top: 0, width: 960, height: 220 });
  key.name = "Key Light"; key.lineFormat.visible = false;
  key.fill.ext.setGradient({ type: "Linear", angle: 90, stops: [{ position: 0, color: C.accent, transparency: 0.9 }, { position: 1, color: C.accent, transparency: 1 }] });
}
// ---- 分步出现 ----
// 视频里同一页按步骤渲染成多帧：at(n) 为真时才画第 n 步的元素。README 的静态图画全部步骤。
let STAGE = Infinity;
const at = (n) => STAGE >= n;
// scenes：[{ name, stages, draw(slide) }]。video 为真时每一步一页，否则每个场景一页。
// 返回每页的名字（「场景:步骤」），build.sh 交给 Blender 排时间线。
async function renderScenes(context, scenes, video) {
  const plan = [];
  for (const s of scenes) {
    if (video) for (let k = 1; k <= s.stages; k++) plan.push([s, k]);
    else plan.push([s, Infinity]);
  }
  const slides = await blankSlides(context, plan.length);
  plan.forEach(([s, k], i) => { STAGE = k; stage(slides[i]); s.draw(slides[i]); });
  STAGE = Infinity;
  await context.sync();
  return plan.map(([s, k]) => `${s.name}:${video ? k : 0}`);
}
// 估算一行文字的宽度（磅）：中文按整字宽，其余按半字宽
const textWidth = (str, size) => [...str].reduce((a, ch) => a + (/[\u2E80-\uFFFF]/.test(ch) ? size : size * 0.58), 0);
