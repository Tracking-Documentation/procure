#!/bin/sh
# ใช้: sh .bump.sh 86 "22 Oct 2026"  — เปลี่ยนเลขเวอร์ชันทั้ง sw.js และข้อความท้ายหน้า
cd "$(dirname "$0")"
sed -i "s/boq-nego-v[0-9]*/boq-nego-v$1/" sw.js
sed -i "s/const APP_VER='[0-9]*',APP_DATE='[^']*'/const APP_VER='$1',APP_DATE='$2'/" index.html
sed -i "s/<span id=\"appVer\">[0-9]*<\/span> · <span id=\"appDate\">[^<]*<\/span>/<span id=\"appVer\">$1<\/span> · <span id=\"appDate\">$2<\/span>/" index.html
