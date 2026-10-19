#!/usr/bin/env bash
# Turn one SVG page template of ai-ppt-server into a real .pptx that can be used as a template deck,
# next to a copy of the template's description (design_spec.md).
#
#   tools/svg-template-to-pptx.sh 科技蓝商务 [输出目录]
#
# It runs ai-ppt-server's own svg_to_pptx.py (read only: nothing in that project is changed) in a
# temporary folder. Set AI_PPT_SERVER if the project is not in the default place.
set -euo pipefail

name="${1:?用法：$0 模板名 [输出目录]}"
server="${AI_PPT_SERVER:-$HOME/Documents/project/ai-ppt-server}"
templates="$server/resource/ai_skills/ppt-master/templates/layouts/$name"
out="${2:-tools/templates/imported/$name}"
python="${PPT_PYTHON:-$server/.runtime/ppt-python/bin/python}"
[ -d "$templates" ] || { echo "找不到模板目录 $templates" >&2; exit 1; }

work="$(mktemp -d)"
trap 'rm -rf "$work"' EXIT
mkdir -p "$work/svg_output" "$out"

# Pages in reading order: cover, contents, chapter, content, ending (the template names them 01..04 with two 02s).
order=(01_cover 02_toc 02_chapter 03_content 04_ending)
n=1
for page in "${order[@]}"; do
  [ -f "$templates/$page.svg" ] || continue
  cp "$templates/$page.svg" "$work/svg_output/$(printf '%02d' "$n")_${page#*_}.svg"
  n=$((n + 1))
done

(cd "$server/resource/ai_skills/ppt-master/scripts" && "$python" svg_to_pptx.py "$work" -s output --only native -t none --no-notes -q -o "$work/template.pptx") >/dev/null
cp "$work/template.pptx" "$out/template.pptx"
[ -f "$templates/design_spec.md" ] && cp "$templates/design_spec.md" "$out/design_spec.md"
echo "$out/template.pptx"
