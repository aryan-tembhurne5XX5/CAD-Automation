import win32com.client
import pythoncom
import json
from pathlib import Path

# ==============================
# CONFIG
# ==============================
PART_PATH = r"E:\Phase 1\Assembly 1\1093144795-A.ipt"
OUT_JSON  = r"E:\Phase 1\extractions\true_holes.json"

# ==============================
# INVENTOR CONNECT
# ==============================
def connect_inventor():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        return inv

# ==============================
# SAFE FLOAT
# ==============================
def f(v):
    try:
        return float(v)
    except:
        return None

# ==============================
# MAIN
# ==============================
def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()

    doc = inv.Documents.Open(PART_PATH, True)
    comp = doc.ComponentDefinition

    holes = []

    for body in comp.SurfaceBodies:
        for face in body.Faces:
            try:
                geom = face.Geometry
                # 5891 = kCylinderSurface
                if geom.SurfaceType != 5891:
                    continue

                # Ignore external cylinders (shafts, bosses)
                if face.IsParamReversed is False:
                    pass

                axis = geom.Axis
                base = geom.BasePoint

                hole = {
                    "diameter_mm": round(geom.Radius * 2 * 10, 3),
                    "center": [
                        f(base.X),
                        f(base.Y),
                        f(base.Z)
                    ],
                    "axis": [
                        f(axis.Direction.X),
                        f(axis.Direction.Y),
                        f(axis.Direction.Z)
                    ]
                }

                holes.append(hole)

            except:
                continue

    output = {
        "part": doc.DisplayName,
        "hole_count": len(holes),
        "holes": holes
    }

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4)

    print("✅ TRUE hole geometry extracted")
    print(f"   Holes found: {len(holes)}")
    print(f"   → {OUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()