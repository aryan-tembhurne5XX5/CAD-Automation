import json
import time
import win32com.client
import pythoncom
import math
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
kAxisEntity       = 67120288

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
# VECTOR / TRANSFORM HELPERS
# =====================================================
def normalize(v):
    mag = math.sqrt(sum(x*x for x in v))
    return [x/mag for x in v] if mag else v

def transform_point(m, p):
    return [
        m.Cell(1,1)*p[0] + m.Cell(1,2)*p[1] + m.Cell(1,3)*p[2] + m.Cell(1,4),
        m.Cell(2,1)*p[0] + m.Cell(2,2)*p[1] + m.Cell(2,3)*p[2] + m.Cell(2,4),
        m.Cell(3,1)*p[0] + m.Cell(3,2)*p[1] + m.Cell(3,3)*p[2] + m.Cell(3,4),
    ]

def transform_vector(m, v):
    return normalize([
        m.Cell(1,1)*v[0] + m.Cell(1,2)*v[1] + m.Cell(1,3)*v[2],
        m.Cell(2,1)*v[0] + m.Cell(2,2)*v[1] + m.Cell(2,3)*v[2],
        m.Cell(3,1)*v[0] + m.Cell(3,2)*v[1] + m.Cell(3,3)*v[2],
    ])

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
        m = occ.Transformation
        geometry["occurrence_transforms"][occ.Name] = {
            "translation": [m.Cell(1,4), m.Cell(2,4), m.Cell(3,4)],
            "rotation": [
                [m.Cell(1,1), m.Cell(1,2), m.Cell(1,3)],
                [m.Cell(2,1), m.Cell(2,2), m.Cell(2,3)],
                [m.Cell(3,1), m.Cell(3,2), m.Cell(3,3)],
            ]
        }

    # --- Fastener axes from Insert constraints ---
    for c in asm_def.Constraints:
        if c.Type != kInsertConstraint:
            continue

        pairs = [
            (c.OccurrenceOne, c.EntityOne),
            (c.OccurrenceOne, c.EntityTwo),
            (c.OccurrenceTwo, c.EntityOne),
            (c.OccurrenceTwo, c.EntityTwo),
        ]

        for occ, ent in pairs:
            if ent.Type != kAxisEntity:
                continue

            axis = ent.Geometry
            origin_local = [
                axis.RootPoint.X,
                axis.RootPoint.Y,
                axis.RootPoint.Z
            ]
            dir_local = [
                axis.Direction.X,
                axis.Direction.Y,
                axis.Direction.Z
            ]

            m = occ.Transformation
            geometry["fastener_axes"][occ.Name] = {
                "origin": transform_point(m, origin_local),
                "direction": transform_vector(m, dir_local),
                "source": "InsertConstraint",
                "confidence": 1.0
            }

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(geometry, f, indent=4)

    print("✅ Phase-4.2 geometry hooks extracted")
    print(f"   Fastener axes: {len(geometry['fastener_axes'])}")
    print(f"   → {OUTPUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()