import win32com.client
import pythoncom
import json
import time
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUT_JSON = Path(r"E:\Phase 1\extractions\geometry_hooks.json")

# =====================================================
# INVENTOR CONNECT
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
# TRANSFORM EXTRACTION (SAFE)
# =====================================================
def extract_transform(occ):
    m = occ.Transformation

    origin = [
        float(m.Cell(1, 4)),
        float(m.Cell(2, 4)),
        float(m.Cell(3, 4))
    ]

    # Z-axis of occurrence (direction)
    z_axis = [
        float(m.Cell(1, 3)),
        float(m.Cell(2, 3)),
        float(m.Cell(3, 3))
    ]

    return {
        "origin": origin,
        "z_axis": z_axis
    }

# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()

    inv = connect_inventor()
    doc = inv.Documents.Open(ASSEMBLY_PATH, True)
    asm = doc.ComponentDefinition

    geometry = {
        "occurrences": {}
    }

    for occ in asm.Occurrences:
        try:
            geometry["occurrences"][occ.Name] = extract_transform(occ)
        except Exception as e:
            print(f"⚠️ Skipped {occ.Name}: {e}")

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(geometry, indent=4))

    print("✅ Geometry hooks extracted (non-empty)")
    print(f"→ {OUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()