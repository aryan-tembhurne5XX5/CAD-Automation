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

FASTENER_KEYWORDS = ("RIVET", "BOLT", "SCREW", "NUT")

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
# TRANSFORMS
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
# FASTENER AXIS FROM PART (CORRECT)
# =====================================================
def extract_fastener_axis(occ):
    try:
        doc = occ.Definition.Document
        props = doc.PropertySets.Item("Design Tracking Properties")
        desc = (props.Item("Description").Value or "").upper()

        if not any(k in desc for k in FASTENER_KEYWORDS):
            return None

        comp_def = doc.ComponentDefinition
        axis = comp_def.WorkAxes.Item(1)

        p = axis.Line.PointOnLine
        d = axis.Line.Direction

        m = occ.Transformation

        return {
            "origin": transform_point(m, [p.X, p.Y, p.Z]),
            "direction": transform_vector(m, [d.X, d.Y, d.Z]),
            "source": "FastenerWorkAxis",
            "confidence": 1.0
        }
    except:
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

    for occ in asm_def.Occurrences:
        geometry["occurrence_transforms"][occ.Name] = extract_transform(occ)

        axis = extract_fastener_axis(occ)
        if axis:
            geometry["fastener_axes"][occ.Name] = axis

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(geometry, f, indent=4)

    print("✅ Phase-4.2 FIXED")
    print(f"   → Fastener axes extracted: {len(geometry['fastener_axes'])}")

    doc.Close(True)

if __name__ == "__main__":
    run()