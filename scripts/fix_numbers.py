with open("C:/Users/User/createslide/presentation.html", "r", encoding="utf-8") as f:
    content = f.read()

replacements = [
    ("01 / 16", "01 / 15"),
    ("02 / 16", "02 / 15"),
    # 03/16 already updated to 03/15
    ("05 / 16", "04 / 15"),
    ("06 / 16", "05 / 15"),
    ("07 / 16", "06 / 15"),
    ("08 / 16", "07 / 15"),
    ("09 / 16", "08 / 15"),
    ("10 / 16", "09 / 15"),
    ("11 / 16", "10 / 15"),
    ("12 / 16", "11 / 15"),
    ("13 / 16", "12 / 15"),
    ("14 / 16", "13 / 15"),
    ("15 / 16", "14 / 15"),
    ("16 / 16", "15 / 15"),
]

for old, new in replacements:
    content = content.replace(old, new)

with open("C:/Users/User/createslide/presentation.html", "w", encoding="utf-8") as f:
    f.write(content)

print("Done. Checking remaining '/ 16':")
for i, line in enumerate(content.split('\n'), 1):
    if '/ 16' in line:
        print(f"  Line {i}: {line.strip()}")
print("All done!")
