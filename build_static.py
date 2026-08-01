# -*- coding: utf-8 -*-
"""GASテンプレートのHTMLを、GitHub Pages用の静的HTMLに変換する。"""
import os, re

EXEC_URL = "https://script.google.com/macros/s/AKfycbxiSxFn4_sJ4vYphaC0cyt77fhf400fwxVNLY56NtkOLBMtbwft6ERj4oaLERGBn49b/exec"

LIFF_IDS = {
    "LIFF_ATTENDANCE_ID": "2010927150-jRsKoCGX",
    "LIFF_LEAVE_ID":      "2010927150-3L8ATtLP",
    "LIFF_ADMIN_ID":      "2010927150-G3v0pMLx",
    "LIFF_CALENDAR_ID":   "2010927150-Fshq2q0M",
    "LIFF_KIBOU_ID":      "2010927150-ZVsHVBWW",
    "LIFF_TEAM_ID":       "2010927150-9HJyfw05",
    "LIFF_SHIFT_ID":      "2010927150-A7us46OW",
}

# 入力ファイル -> 出力ファイル名
FILES = {
    "Attendance.html":     "attendance.html",
    "LeaveRequest.html":   "leave.html",
    "AdminDashboard.html": "admin.html",
    "Calendar.html":       "calendar.html",
    "KibouRequest.html":   "kibou.html",
    "TeamCalendar.html":   "team.html",
    "ShiftRequest.html":   "shift.html",
}

OUT_DIR = "docs"
os.makedirs(OUT_DIR, exist_ok=True)

for src, dst in FILES.items():
    with open(src, encoding="utf-8") as f:
        html = f.read()

    # ScriptApp.getService().getUrl() を実URLに
    html = re.sub(r"<\?=\s*ScriptApp\.getService\(\)\.getUrl\(\)\s*\?>", EXEC_URL, html)

    # PropertiesService...getProperty("LIFF_X_ID") を実IDに
    def repl(m):
        key = m.group(1)
        return LIFF_IDS.get(key, "")
    html = re.sub(r'<\?=\s*PropertiesService\.getScriptProperties\(\)\.getProperty\("(LIFF_[A-Z_]+)"\)\s*\?>', repl, html)

    # 念のため、残った <?= ... ?> を空に
    leftover = re.findall(r"<\?=.*?\?>", html)
    if leftover:
        print("  WARNING leftover templates in", dst, ":", leftover)
        html = re.sub(r"<\?=.*?\?>", "", html)

    with open(os.path.join(OUT_DIR, dst), "w", encoding="utf-8") as f:
        f.write(html)
    print("built", dst, len(html), "bytes")

print("done ->", OUT_DIR)
