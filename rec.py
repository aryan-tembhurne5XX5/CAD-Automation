import win32com.client
import csv
import os
import time

# ==============================
# CONFIG
# ==============================
BOM_FILE = "bom.csv"
OUTPUT_ASM = "reconstructed.iam"
SPACING_MM = 30  # visual spacing between parts

# ==============================
# CONNECT INVENTOR
# ==============================
inv = win32com.client.DispatchEx("Inventor.Application")
inv.Visible = True
time.sleep(2)

asm_doc = inv.Documents.Add(12291)  # kAssemblyDocumentObject
asm_def = asm_doc.ComponentDefinition

base_path = os.getcwd()

# ==============================
# LOAD BOM
# ==============================
bom_parts = []

with open(BOM_FILE, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        part = row["part title"].strip()
        qty = int(row["quantity"])
        bom_parts.append((part, qty))

# ==============================
# INSERT PARTS
# ==============================
x_offset = 0
is_first = True

for part_name, qty in bom_parts:
    part_file = os.path.join(base_path, f"{part_name}.ipt")

    if not os.path.exists(part_file):
        print(f"❌ Missing file: {part_file}")
        continue

    for i in range(qty):
        trans = inv.TransientGeometry.CreateMatrix()
        trans.Cell(1, 4) = x_offset

        occ = asm_def.Occurrences.Add(part_file, trans)

        if is_first:
            occ.Grounded = True
            is_first = False

        x_offset += SPACING_MM

# ==============================
# SAVE
# ==============================
out_path = os.path.join(base_path, OUTPUT_ASM)
asm_doc.SaveAs(out_path, False)

print("✅ Assembly created:", out_path)