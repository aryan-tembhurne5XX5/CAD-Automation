import win32com.client
from pathlib import Path
import csv

# =========================
# CONFIG
# =========================
BASE_DIR = Path(__file__).parent
PARTS_DIR = BASE_DIR / "parts"
JSON_FILE = BASE_DIR / "assembly_data.json"
OUTPUT_IAM = BASE_DIR / "Reconstructed.iam"

# =========================
# HELPERS
# =========================
def make_matrix(inv, rotation, translation):
    """
    Build Inventor Matrix from rotation + translation
    """
    tg = inv.TransientGeometry
    m = tg.CreateMatrix()

    # rotation (3x3)
    m.Cell[1,1], m.Cell[1,2], m.Cell[1,3] = rotation[0]
    m.Cell[2,1], m.Cell[2,2], m.Cell[2,3] = rotation[1]
    m.Cell[3,1], m.Cell[3,2], m.Cell[3,3] = rotation[2]

    # translation
    m.Cell[1,4] = translation[0]
    m.Cell[2,4] = translation[1]
    m.Cell[3,4] = translation[2]

    return m

# =========================
# MAIN
# =========================
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