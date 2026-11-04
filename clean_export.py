#!/usr/bin/env python3
"""Turn a TikTok Shop 'Video Analysis Video List' export (.xlsx) into the slim
data/videos.csv the leaderboard reads. Keeps only what the page needs.
Usage: python3 clean_export.py export.xlsx data/videos.csv
"""
import sys, csv, datetime, openpyxl
PRODUCT_ID = "1732737544674382545"   # SOUR PATCH KIDS + SWEDISH FISH Holiday Candy Variety Pack
def money(s): return float(str(s or 0).replace("$", "").replace(",", "") or 0)
def num(s):
    try: return int(float(str(s).replace(",", "")))
    except ValueError: return 0
def iso(s):
    try: return datetime.datetime.strptime(s, "%m/%d/%Y %H:%M").strftime("%Y-%m-%d %H:%M")
    except (ValueError, TypeError): return ""
ws = openpyxl.load_workbook(sys.argv[1]).active
rows = list(ws.iter_rows(values_only=True)); head = rows[0]
seen = {}
for r in rows[2:]:
    d = dict(zip(head, r))
    if str(d.get("Product ID")).strip() != PRODUCT_ID: continue
    seen[str(d["Video ID"])] = [d["Video ID"], d["Creator name"], iso(d["Post date"]), d["Video link"],
        money(d["Creator video-attributed GMV"]), num(d["Video-attributed orders"]),
        num(d["Video-attributed items sold"]), num(d["Video views"]), num(d["Likes"])]
with open(sys.argv[2], "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["video_id", "creator", "posted", "link", "gmv", "orders", "items", "views", "likes"])
    for v in seen.values(): w.writerow(v)
print(len(seen), "videos written")
