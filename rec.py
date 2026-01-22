import json
import win32com.client
from pathlib import Path

# =========================
# CONFIG
# =========================
BASE_DIR = Path(__file__).parent
PARTS_DIR =  r"E:\Phase 1\Assembly 1"
JSON_FILE =  r"E:\Phase 1\extractions\final_phase1_to_5.json"
OUTPUT_IAM =  r"E:\Phase 1\Assembly 1\Reconstructed.iam"

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
    # Load JSON
    with open(JSON_FILE, "r") as f:
        data = json.load(f)

    occurrences = data["occurrences"]

    # Start Inventor
    inv = win32com.client.Dispatch("Inventor.Application")
    inv.Visible = True

    # Create new assembly
    asm_doc = inv.Documents.Add(
        inv.DocumentTypeEnum.kAssemblyDocumentObject,
        inv.FileManager.GetTemplateFile(
            inv.DocumentTypeEnum.kAssemblyDocumentObject
        )
    )

    asm_def = asm_doc.ComponentDefinition

    print("🔧 Rebuilding assembly...")

    for occ_name, occ_data in occurrences.items():
        part_name = occ_data["definition"]
        part_path = PARTS_DIR / part_name

        if not part_path.exists():
            print(f"❌ Missing part: {part_name}")
            continue

        rotation = occ_data["transform"]["rotation"]
        translation = occ_data["transform"]["translation"]

        matrix = make_matrix(inv, rotation, translation)

        asm_def.Occurrences.Add(str(part_path), matrix)
        print(f"✅ Placed {occ_name}")

    # Save assembly
    asm_doc.SaveAs(str(OUTPUT_IAM), False)
    print("\n🎉 Assembly reconstruction complete")
    print(f"📦 Saved as: {OUTPUT_IAM}")

if __name__ == "__main__":
    run()