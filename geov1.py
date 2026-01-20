import json
import time
import win32com.client
import pythoncom
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = Path(r"E:\Phase 1\extractions\geometry_hooks.json")

# =====================================================
# ENUMS
# =====================================================
kInsertConstraint = 100665344
kCylinderFace     = 67119536

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
# TRANSFORM EXTRACTION
# =====================================================
def extract_transform(occ):
    m = occ.Transformation
    return {
        "translation": [m.Cell(1,4), m.Cell(2,4), m.Cell(3,4)],
        "rotation": [
            [m.Cell(1,1), m.Cell(1,2), m.Cell(1,3)],
            [m.Cell(2,1), m.Cell(2,2), m.Cell(2,3)],
            [m.Cell(3,1), m.Cell(3,2), m.Cell(3,3)]
        ]
    }

# =====================================================
# AXIS EXTRACTION (INSERT ONLY)
# =====================================================
def extract_axis_direction(entity):
    try:
        if entity.Type == kCylinderFace:
            axis = entity.Geometry.Axis
            v = axis.Direction
            return [v.X, v.Y, v.Z]
    except:
        pass
    return None

# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()

    inv = connect_inventor()
    doc = inv.Documents.Open(ASSEMBLY_PATH, True)
    asm_def = doc.ComponentDefinition

    geometry = {
        "occurrence_transforms": {},
        "fastener_axes": {}
    }

    # --- Occurrence transforms ---
    for occ in asm_def.Occurrences:
        try:
            geometry["occurrence_transforms"][occ.Name] = extract_transform(occ)
        except:
            continue

    # --- Fastener axes from Insert constraints ---
    for c in asm_def.Constraints:
        if c.Type != kInsertConstraint:
            continue

        axis = extract_axis_direction(c.EntityOne) or extract_axis_direction(c.EntityTwo)
        if not axis:
            continue

        try:
            geometry["fastener_axes"][c.OccurrenceOne.Name] = axis
        except:
            pass

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(geometry, f, indent=4)

    print("✅ Phase-4.2 geometry hooks extracted")
    print(f"   → {OUTPUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()