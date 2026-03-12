#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

if ! command -v pdftoppm >/dev/null 2>&1; then
  echo "pdftoppm is required"
  exit 1
fi

if ! command -v convert >/dev/null 2>&1; then
  echo "ImageMagick convert is required"
  exit 1
fi

mkdir -p assets/gallery

declare -A previews=(
  ["board-report"]="rendered/board-report.pdf"
  ["executive-dashboard"]="rendered/executive-dashboard.pdf"
  ["product-launch-brief"]="rendered/product-launch-brief.pdf"
  ["talent-profile"]="rendered/talent-profile.pdf"
  ["invoice"]="rendered/invoice.pdf"
  ["meeting-notes"]="rendered/meeting-notes.pdf"
)

for name in "${!previews[@]}"; do
  pdf="${previews[$name]}"
  if [[ ! -f "$pdf" ]]; then
    echo "missing PDF: $pdf"
    exit 1
  fi

  pdftoppm -png -singlefile -f 1 -scale-to 1400 "$pdf" "assets/gallery/$name"
  convert "assets/gallery/$name.png" \
    -background "#F6F7FB" \
    -gravity center \
    -resize 900x1200 \
    -extent 900x1200 \
    -bordercolor "#D8DFEA" \
    -border 1 \
    "assets/gallery/$name.png"
done

convert -size 1400x920 xc:"#F4F7FB" \
  \( assets/gallery/executive-dashboard.png -resize 292x389 \) -geometry +688+184 -composite \
  \( assets/gallery/board-report.png -resize 252x336 \) -geometry +1032+134 -composite \
  \( assets/gallery/product-launch-brief.png -resize 252x336 \) -geometry +1018+492 -composite \
  \( assets/gallery/talent-profile.png -resize 252x336 \) -geometry +724+532 -composite \
  -fill "#0F172A" -font Helvetica-Bold -pointsize 50 -annotate +76+98 "RusDox Template Gallery" \
  -fill "#475569" -font Helvetica -pointsize 24 -annotate +78+144 "Real previews generated from YAML -> DOCX + PDF" \
  -fill "#0F172A" -font Helvetica-Bold -pointsize 28 -annotate +78+254 "Start with examples/*.yaml" \
  -fill "#334155" -font Helvetica -pointsize 23 -annotate +78+296 "Board reports, dashboards, launch briefs," \
  -fill "#334155" -font Helvetica -pointsize 23 -annotate +78+332 "invoices, notes, and talent profiles." \
  -fill "#2563EB" -font Helvetica-Bold -pointsize 24 -annotate +78+390 "rusdox examples" \
  -fill "#64748B" -font Helvetica -pointsize 22 -annotate +78+432 "Then copy a template and change the content." \
  assets/template-gallery.png

convert -size 1280x640 xc:"#0F172A" \
  -fill "#F8FAFC" -font Helvetica-Bold -pointsize 44 -annotate +72+92 "Generate DOCX + PDF" \
  -fill "#F8FAFC" -font Helvetica-Bold -pointsize 44 -annotate +72+144 "from YAML" \
  -fill "#93C5FD" -font Helvetica-Bold -pointsize 22 -annotate +74+196 "Pure Rust engine" \
  -fill "#CBD5E1" -font Helvetica -pointsize 24 -annotate +72+246 "Readable templates. Fast rendering. Real output files." \
  -fill "#E2E8F0" -font Helvetica -pointsize 23 -annotate +72+350 "Write YAML" \
  -fill "#38BDF8" -font Helvetica-Bold -pointsize 30 -annotate +72+392 "rusdox mydoc.yaml" \
  -fill "#E2E8F0" -font Helvetica -pointsize 23 -annotate +72+468 "Get both" \
  -fill "#F8FAFC" -font Helvetica-Bold -pointsize 29 -annotate +72+510 "generated/*.docx" \
  -fill "#F8FAFC" -font Helvetica-Bold -pointsize 29 -annotate +72+552 "rendered/*.pdf" \
  \( assets/gallery/executive-dashboard.png -resize 304x405 \) -geometry +714+128 -composite \
  \( assets/gallery/board-report.png -resize 226x301 \) -geometry +1004+84 -composite \
  \( assets/gallery/talent-profile.png -resize 226x301 \) -geometry +1004+352 -composite \
  assets/social-preview-rusdox.png

echo "Generated gallery assets:"
echo "  assets/template-gallery.png"
echo "  assets/social-preview-rusdox.png"
echo "  assets/gallery/"
