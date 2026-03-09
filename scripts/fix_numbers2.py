with open("C:/Users/User/createslide/presentation.html", "r", encoding="utf-8") as f:
    content = f.read()

# Current state after manual edits:
# 01/15, 02/15, 03/15, 04/15, 05/15 → keep but change /15 to /14
# 06/15 → was deleted
# 07/15 (雲端) → 06/14
# 07/14 (天母) → already correct
# 08/15 → was already replaced manually above as 07/14
# 09/15 → 08/14
# 10/15 → 09/14
# 11/15 → 10/14
# 12/15 → 11/14
# 13/15 → 12/14
# 14/15 → 13/14
# 15/15 → 14/14

replacements = [
    ("07 / 15", "06 / 14"),  # 雲端數據平台
    ("09 / 15", "08 / 14"),  # 洲際棒球場
    ("10 / 15", "09 / 14"),  # 龍潭三重
    ("11 / 15", "10 / 14"),  # 基層棒球
    ("12 / 15", "11 / 14"),  # 技術躍進
    ("13 / 15", "12 / 14"),  # 完整情境
    ("14 / 15", "13 / 14"),  # 未來方向
    ("15 / 15", "14 / 14"),  # 合作邀請
    # Fix the remaining /15 → /14 for slides 01-05
    ("01 / 15", "01 / 14"),
    ("02 / 15", "02 / 14"),
    ("03 / 15", "03 / 14"),
    ("04 / 15", "04 / 14"),
    ("05 / 15", "05 / 14"),
]

for old, new in replacements:
    count = content.count(old)
    content = content.replace(old, new)
    print(f"  {old} → {new}: {count} replacements")

with open("C:/Users/User/createslide/presentation.html", "w", encoding="utf-8") as f:
    f.write(content)

# Verify no stale numbers remain
remaining = []
for i, line in enumerate(content.split('\n'), 1):
    if '/ 15' in line or '/ 16' in line:
        remaining.append(f"Line {i}: {line.strip()}")

if remaining:
    print("\nWARNING - stale numbers found:")
    for r in remaining: print(r)
else:
    print("\nAll page numbers updated cleanly. Total: 14 slides.")
