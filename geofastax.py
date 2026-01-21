import csv
import json
import time
import win32com.client
import pythoncom
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = Path(r"E:\Phase 1\Assembly 1\1093144795-M1.iam")
BOM_CSV_PATH  = ASSEMBLY_PATH.parent / "BOM_1093144795-M1.csv"
OUTPUT_JSON   = ASSEMBLY_PATH.parent / "geometry_fastener_axes.json"

FASTENER_KEYWORDS = ["RIVET", "NUT", "BOLT", "SCREW"]

# =====================================================
# INVENTOR CONNECTION
# =====================================================
def connect_inventor():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
        return inv

# =====================================================
# BOM FASTENER PARSING
# =====================================================
def read_fastener_part_numbers(csv_path):
    fasteners = set()

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            desc = (row.get("Part Title") or row.get("Description") or "").upper()
            title = (row.get("Title") or "").strip()

            if any(k in desc for k in FASTENER_KEYWORDS):
                fasteners.add(title)

    return fasteners

# =====================================================
# TRANSFORM → AXIS EXTRACTION (CORE)
# =====================================================
def extract_axis_from_transform(occ):
    """
    Fastener axis = local Z axis of transform
    """
    m = occ.Transformation

    direction = [
        m.Cell(1, 3),
        m.Cell(2, 3),
        m.Cell(3, 3)
    ]

    origin = [
        m.Cell(1, 4),
        m.Cell(2, 4),
        m.Cell(3, 4)
    ]

    return origin, direction

# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()

    if not ASSEMBLY_PATH.exists():
        raise FileNotFoundError("Assembly not found")

    if not BOM_CSV_PATH.exists():
        raise FileNotFoundError("BOM CSV not found")

    fastener_part_numbers = read_fastener_part_numbers(BOM_CSV_PATH)

    print(f"🔍 Fastener types detected from BOM: {fastener_part_numbers}")

    inv = connect_inventor()
    doc = inv.Documents.Open(str(ASSEMBLY_PATH), True)
    asm_def = doc.ComponentDefinition

    output = []

    for occ in asm_def.Occurrences:
        try:
            part_doc = occ.Definition.Document
            part_number = None

            try:
                props = part_doc.PropertySets.Item("Design Tracking Properties")
                part_number = props.Item("Part Number").Value
            except:
                continue

            if part_number not in fastener_part_numbers:
                continue

            origin, direction = extract_axis_from_transform(occ)

            output.append({
                "occurrence": occ.Name,
                "part_number": part_number,
                "origin": origin,
                "direction": direction,
                "source": "OccurrenceTransform",
                "confidence": 0.95
            })

        except:
            continue

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4)

    print("✅ Phase-4.3 complete")
    print(f"   → Fastener axes extracted: {len(output)}")
    print(f"   → Output: {OUTPUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()