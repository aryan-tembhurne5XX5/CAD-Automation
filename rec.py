import win32com.client
from pathlib import Path
import csv

# =========================
# CONFIG
# =========================
IPT_FOLDER = Path(r"E:\Phase 1\Assembly 1")
BOM_CSV    = Path(r"E:\Phase 1\BOM_1625891052.csv")
OUTPUT_ASM = Path(r"E:\Phase 1\Reconstructed.iam")

# Inventor document type (RAW ENUM)
kAssemblyDocumentObject = 12291

def run():
    inv = win32com.client.Dispatch("Inventor.Application")
    inv.Visible = True

    # Create new Assembly document
    asm_doc = inv.Documents.Add(
        kAssemblyDocumentObject,
        inv.FileManager.GetTemplateFile(
            kAssemblyDocumentObject,
            inv.Language
        ),
        True
    )

    asm_def = asm_doc.ComponentDefinition
    tg = inv.TransientGeometry

    # Read BOM
    with open(BOM_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    for row in rows:
        part = row["Part Number"].strip()
        qty  = int(row.get("Quantity", 1))

        part_path = IPT_FOLDER / f"{part}.ipt"
        if not part_path.exists():
            print(f"❌ Missing: {part_path}")
            continue

        for _ in range(qty):
            matrix = tg.CreateMatrix()
            asm_def.Occurrences.Add(str(part_path), matrix)
            print(f"✅ Inserted {part}")

    asm_doc.SaveAs(str(OUTPUT_ASM), False)
    print("\n🎉 Assembly reconstructed successfully")
    print(f"📦 Saved to: {OUTPUT_ASM}")

if __name__ == "__main__":
    run()