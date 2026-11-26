// 介绍视频的开场、结尾和封面。视频的其他场景来自 scenes-diagrams.js、scenes-effects.js。
function darkBg(s) {
  box(s, { shape: "Rectangle", l: 0, t: 0, w: 960, h: 540, fill: C.fg });
}

const SCENE_INTRO = {
  name: "intro", stages: 3, draw(s) {
    darkBg(s);
    text(s, "OfficeAPI", { l: 0, t: 160, w: 960, h: 104, size: 72, bold: true, color: C.white, align: "Center" });
    if (at(2)) text(s, "一个二进制，读写所有 Office 文档", { l: 0, t: 270, w: 960, h: 40, size: 28, color: "#C9D3E8", align: "Center" });
    if (at(3)) {
      const tags = ["格式转换", "AI 编辑器", "Word · Excel · PowerPoint"];
      const ws = tags.map((t) => textWidth(t, 13) + 36);
      let x = 480 - (ws.reduce((a, w) => a + w, 0) + 16 * (tags.length - 1)) / 2;
      for (const [i, t] of tags.entries()) {
        const w = ws[i];
        label(box(s, { l: x, t: 340, w, h: 30, fill: "#2A3346", radius: 0.5 }), t, { size: 13, color: C.accentLight });
        x += w + 16;
      }
    }
  },
};

const SCENE_OUTRO = {
  name: "outro", stages: 3, draw(s) {
    darkBg(s);
    text(s, "OfficeAPI", { l: 0, t: 150, w: 960, h: 80, size: 60, bold: true, color: C.white, align: "Center" });
    text(s, "不装 Office，不要界面，不连网络", { l: 0, t: 240, w: 960, h: 34, size: 22, color: "#C9D3E8", align: "Center" });
    if (at(2)) {
      let x = 480 - (PLATFORMS.reduce((a, p) => a + p.length * 8 + 28, 0) + 8 * (PLATFORMS.length - 1)) / 2;
      for (const p of PLATFORMS) {
        const w = p.length * 8 + 28;
        pill(s, p, x, 300, C.white, "#2A3346", w);
        x += w + 8;
      }
    }
    if (at(3)) {
      text(s, "officesdk.com  ·  github.com/officesdk/OfficeAPI", { l: 0, t: 380, w: 960, h: 24, size: 16, color: C.accentLight, align: "Center" });
    }
  },
};

// README 里的视频封面：开场画面加播放按钮
const SCENE_POSTER = {
  name: "poster", stages: 1, draw(s) {
    SCENE_INTRO.draw(s);
    box(s, { shape: "Ellipse", l: 452, t: 404, w: 56, h: 56, fill: C.accent });
    const tri = s.shapes.addGeometricShape("Triangle", { left: 471, top: 418, width: 24, height: 28 });
    tri.rotation = 90; tri.fill.setSolidColor(C.white); tri.lineFormat.visible = false;
    text(s, "观看 1 分钟介绍视频", { l: 0, t: 470, w: 960, h: 22, size: 13, color: "#9AA5BC", align: "Center" });
  },
};

const VIDEO = [SCENE_INTRO, SCENE_HERO, SCENE_FORMATS, SCENE_ARCH, SCENE_LOOP, SCENE_TRANSLATE, SCENE_EDIT, SCENE_WORD, SCENE_EXCEL, SCENE_OUTRO];
