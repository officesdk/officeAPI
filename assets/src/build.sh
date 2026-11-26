#!/bin/sh
# 生成 OfficeAPI/assets/ 下的说明图和效果图，英文、中文各一套（介绍视频 intro*.mp4 / intro-poster*.png 不由本脚本生成）。
# 英文图用原文件名（README.md 用），中文图加 -zh 后缀（README.zh-CN.md 用）；只出一种语言：./build.sh en 或 ./build.sh zh。
# 依赖 pptjs 与 wordjs 的命令行（默认取同级目录下的 release 构建，可用 PPTJS / WORDJS 覆盖）。
set -e
SRC=$(cd "$(dirname "$0")" && pwd)
OUT=$(dirname "$SRC")
WS=$(cd "$SRC/../../.." && pwd)
PPTJS=${PPTJS:-$WS/pptjs/target/release/pptjs}
WORDJS=${WORDJS:-$WS/wordjs/target/release/wordjs}
BLANK=${BLANK_PPTX:-$WS/pptjs/crates/pptjs-mcp/assets/blank.pptx}
LANGS=${*:-en zh}
TMP0=$(mktemp -d)
if [ -n "$KEEP" ]; then echo "tmp: $TMP0"; else trap 'rm -rf "$TMP0"' EXIT; fi

# 脚本里的文案写成 tr(英文, 中文)：拼接时最前面写入 const LANG
prelude() { echo "const LANG = \"$OKL\";"; }

# ppt <输入 pptx> <输出 pptx> <脚本...> [-- --file 名称=路径 ...]
# pptjs 脚本不能 import：语言、lib.js 和各脚本按顺序拼成一个文件运行，返回值写到 <输出>.json
ppt() {
  in=$1; o=$2; shift 2
  { prelude; cat "$SRC/lib.js"; } > "$TMP/run.js"
  while [ $# -gt 0 ] && [ "$1" != "--" ]; do cat "$SRC/$1" >> "$TMP/run.js"; shift; done
  [ "$1" = "--" ] && shift
  "$PPTJS" run "$TMP/run.js" --in "$in" --out "$o" --json "$o.json" "$@" > /dev/null
  python3 -c "import json,sys;e=json.load(open(sys.argv[1]))['error'];sys.exit(e and print(e) or 0)" "$o.json"
}
# word <脚本> <输出 docx> [--in 输入 docx]：同样在前面拼上语言和 tr()
word() {
  s=$1; o=$2; shift 2
  { prelude; echo 'const tr = (en, zh) => (LANG === "en" ? en : zh);'; cat "$SRC/$s"; } > "$TMP/word.js"
  "$WORDJS" run "$TMP/word.js" "$@" --out "$o" > /dev/null
}
render() { # render <pptx> <目录> <宽度>
  "$PPTJS" render "$1" --width "$3" --out-dir "$2" > /dev/null
}

for OKL in $LANGS; do
  TMP=$TMP0/$OKL; mkdir -p "$TMP"
  if [ "$OKL" = en ]; then SUF=; else SUF=-$OKL; fi

  # ---- 效果图素材：用户文件 → AI 修改 → 渲染 ----
  # PPT：中文页 → 译文填回（标出溢出，仅用于截图）→ AI 按检查结果修正（两种语言都演示中译英）
  ppt "$BLANK" "$TMP/zh.pptx" ppt-translate-1-zh.js
  ppt "$TMP/zh.pptx" "$TMP/en.pptx" ppt-translate-2-en.js
  ppt "$TMP/en.pptx" "$TMP/en-marked.pptx" ppt-translate-mark.js
  ppt "$TMP/en.pptx" "$TMP/fix.pptx" ppt-translate-3-fix.js
  # PPT：用户的单栏页 → 改成两栏加图表
  ppt "$BLANK" "$TMP/eb.pptx" ppt-edit-before.js
  ppt "$TMP/eb.pptx" "$TMP/ea.pptx" ppt-edit-after.js
  for f in en-marked fix eb ea; do render "$TMP/$f.pptx" "$TMP/png-$f" 1280; done
  # Word：用户的报告 → 按一句话修改
  word word-edit-before.js "$TMP/wb.docx"
  word word-edit-after.js "$TMP/wa.docx" --in "$TMP/wb.docx"
  for f in wb wa; do
    "$WORDJS" render "$TMP/$f.docx" --pages 1 --width 1200 --out-dir "$TMP/png-$f" > /dev/null
  done
  FILES="--file en=$TMP/png-en-marked/slide_1.png --file fix=$TMP/png-fix/slide_1.png
    --file eb=$TMP/png-eb/slide_1.png --file ea=$TMP/png-ea/slide_1.png
    --file wb=$TMP/png-wb/page-1.png --file wa=$TMP/png-wa/page-1.png"

  # ---- README 的静态图 ----
  ppt "$BLANK" "$TMP/diagrams.pptx" scenes-diagrams.js main-diagrams.js
  render "$TMP/diagrams.pptx" "$TMP/diagrams" 1600
  cp "$TMP/diagrams/slide_1.png" "$OUT/architecture$SUF.png"
  cp "$TMP/diagrams/slide_2.png" "$OUT/formats$SUF.png"
  cp "$TMP/diagrams/slide_3.png" "$OUT/ai-loop$SUF.png"

  # shellcheck disable=SC2086
  ppt "$BLANK" "$TMP/effects.pptx" scenes-effects.js main-effects.js -- $FILES
  render "$TMP/effects.pptx" "$TMP/effects" 1600
  cp "$TMP/effects/slide_1.png" "$OUT/effect-ppt-translate$SUF.png"
  cp "$TMP/effects/slide_2.png" "$OUT/effect-ppt-edit$SUF.png"
  cp "$TMP/effects/slide_3.png" "$OUT/effect-word-edit$SUF.png"
  cp "$TMP/effects/slide_4.png" "$OUT/effect-excel-edit$SUF.png"

  # 自查：所有图（含效果图里的示例文件）的版面检查不能有 error / warning（字体回退除外）
  for f in diagrams effects eb ea; do
    "$PPTJS" inspect "$TMP/$f.pptx" --json > "$TMP/$f-inspect.json" || true
    python3 -c "import json,sys;d=json.load(open(sys.argv[1]));[print(sys.argv[2],i['slideIndex']+1,i['rule'],i['shapeName'],i.get('detail')) for i in d['issues'] if i['severity']!='info' and i['rule']!='font-fallback']" "$TMP/$f-inspect.json" "$OKL/$f"
  done
done

echo "done: $OUT"
