import win32com.client
import json
import csv
import os
from collections import defaultdict

# ==============================
# CONFIG
# ==============================
JSON_FILE = "assembly_dump.json"
BOM_FILE = "bom.csv"
OUTPUT_ASSEMBLY = "REBUILT_FROM_JSON.iam"
WORKING_DIR = os.getcwd()

# ==============================
# LOAD JSON
# ==============================
with open(JSON_FILE, "r", encoding="utf-8") as f:
    data = json.load(f)

occurrences = data["occurrences"]

# ==============================
# LOAD BOM
# ==============================
bom_parts = defaultdict(int)

with open(BOM_FILE, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        part_title = row["Part Title"].strip()
        qty = int(row["Quantity"])
        bom_parts[part_title] += qty

# ==============================
# CONNECT TO INVENTOR
# ==============================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = True

docs = inv.Documents
asm_doc = docs.Add(12291, "", True)  # kAssemblyDocumentObject
asm_doc.SaveAs(os.path.join(WORKING_DIR, OUTPUT_ASSEMBLY), False)

asm_def = asm_doc.ComponentDefinition
tg = inv.TransientGeometry

# ==============================
# MATRIX HELPER
# ==============================
def matrix_from_list(m):
    mat = tg.CreateMatrix()
    for r in range(3):
        for c in range(4):
            mat.Cell[r+1, c+1] = m[r][c]
    mat.Cell[4,4] = 1
    return mat

# ==============================
# INSERT COMPONENTS
# ==============================
json_count = defaultdict(int)
missing_files = set()

for occ in occurrences:
    part_name = occ["definition"].replace(".ipt", "")
    filename = f"{part_name}.ipt"
    part_path = os.path.join(WORKING_DIR, filename)

    json_count[part_name] += 1

    if not os.path.exists(part_path):
        missing_files.add(filename)
        continue

    transform = matrix_from_list(occ["transform"])
    asm_def.Occurrences.Add(part_path, transform)

# ==============================
# VALIDATION REPORT
# ==============================
print("\n🔎 BOM vs JSON CHECK\n")

for part, bom_qty in bom_parts.items():
    json_qty = json_count.get(part, 0)
    status = "OK" if bom_qty == json_qty else "MISMATCH"
    print(f"{part:20} BOM={bom_qty:3} | JSON={json_qty:3} → {status}")

if missing_files:
    print("\n❌ Missing IPT files:")
    for f in missing_files:
        print("  ", f)

# ==============================
# FINALIZE
# ==============================
asm_doc.Update()
asm_doc.Save()

print("\n✅ Assembly rebuilt successfully")
print(f"📁 Output: {OUTPUT_ASSEMBLY}")