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
# INVENTOR ENUMS
# =====================================================
kInsertConstraint = 100665344
kAxisEntity       = 67120288   # Axis
kCylinderFace    = 67119536   # Cylindrical Face

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
# TRANSFORM HELPERS
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

def transform_point(m, p):
    return [
        m.Cell(1,1)*p[0] + m.Cell(1,2)*p[1] + m.Cell(1,3)*p[2] + m.Cell(1,4),
        m.Cell(2,1)*p[0] + m.Cell(2,2)*p[1] + m.Cell(2,3)*p[2] + m.Cell(2,4),
        m.Cell(3,1)*p[0] + m.Cell(3,2)*p[1] + m.Cell(3,3)*p[2] + m.Cell(3,4),
    ]

def transform_vector(m, v):
    return [
        m.Cell(1,1)*v[0] + m.Cell(1,2)*v[1] + m.Cell(1,3)*v[2],
        m.Cell(2,1)*v[0] + m.Cell(2,2)*v[1] + m.Cell(2,3)*v[2],
        m.Cell(3,1)*v[0] + m.Cell(3,2)*v[1] + m.Cell(3,3)*v[2],
    ]

# =====================================================
# AXIS EXTRACTION (SAFE)
# =====================================================
def extract_axis_from_entity(ent):
    try:
        geom = ent.Geometry
        p = geom.PointOnLine
        d = geom.Direction
        return (
            [p.X, p.Y, p.Z],
            [d.X, d.Y, d.Z]
        )
    except:
        return None, None

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

    # -------------------------------------------------
    # OCCURRENCE TRANSFORMS
    # -------------------------------------------------
    for occ in asm_def.Occurrences:
        try:
            geometry["occurrence_transforms"][occ.Name] = extract_transform(occ)
        except:
            continue

    # -------------------------------------------------
    # FASTENER AXES FROM INSERT CONSTRAINTS
    # -------------------------------------------------
    for c in asm_def.Constraints:
        if c.Type != kInsertConstraint:
            continue

        candidates = [
            (c.OccurrenceOne, c.EntityOne),
            (c.OccurrenceOne, c.EntityTwo),
            (c.OccurrenceTwo, c.EntityOne),
            (c.OccurrenceTwo, c.EntityTwo),
        ]

        for occ, ent in candidates:
            if ent.Type not in (kAxisEntity, kCylinderFace):
                continue

            origin_local, dir_local = extract_axis_from_entity(ent)
            if not origin_local:
                continue

            m = occ.Transformation

            geometry["fastener_axes"][occ.Name] = {
                "origin": transform_point(m, origin_local),
                "direction": transform_vector(m, dir_local),
                "source": "InsertConstraint",
                "confidence": 1.0
            }

    # -------------------------------------------------
    # SAVE OUTPUT
    # -------------------------------------------------
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(geometry, f, indent=4)

    print("✅ Phase-4.2 geometry hooks extracted")
    print(f"   → {OUTPUT_JSON}")
    print(f"   → Fastener axes found: {len(geometry['fastener_axes'])}")

    doc.Close(True)

if __name__ == "__main__":
    run()