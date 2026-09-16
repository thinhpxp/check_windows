#!/usr/bin/env bash
# Script gop tat ca file <COMPUTERNAME>.csv thanh 1 file duy nhat tren Fedora
TARGET_DIR="/srv/samba/share/Results"  # Chinh sua thanh duong dan local share cua ban tren Fedora
OUTPUT_FILE="$TARGET_DIR/_TongHop_BanQuyen.csv"

if [ ! -d "$TARGET_DIR" ]; then
    echo "Thu muc khong ton tai: $TARGET_DIR"
    exit 1
fi

first=1
for f in "$TARGET_DIR"/*.csv; do
    [ -e "$f" ] || continue
    [ "$f" = "$OUTPUT_FILE" ] && continue
    
    if [ $first -eq 1 ]; then
        head -n 1 "$f" > "$OUTPUT_FILE"
        first=0
    fi
    tail -n +2 "$f" >> "$OUTPUT_FILE"
done

echo "Da gop xong tat ca CSV vao: $OUTPUT_FILE"
