import win32com.client
import pythoncom
import json
import time
from pathlib import Path
from math import sqrt

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUT_JSON      = Path(r"E:\Phase 1\extractions\final_phase1_to_5.json")

# =====================================================
# CONSTANTS
# =====================================================
kCylinderFace = 67119536
MM_PER_CM = 10.0

# =====================================================
# HELPERS
# =====================================================
def vec_len(v):
    return sqrt(v[0]**2 + v[1]**2 + v[2]**2)

def vec_dot(a, b):
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]

def normalize(v):
    l = vec_len(v)
    if l == 0:
        return v
    return [v[0]/l, v[1]/l, v[2]/l]

def dist(a, b):
    return sqrt((a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2)

# =====================================================
# CONNECT INVENTOR
# =====================================================
def connect():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
        return inv

# =====================================================
# FORCE SOLID GENERATION (CRITICAL)
# =====================================================
def force_rebuild(doc):
    try:
        doc.ComponentDefinition.Rebuild()
        doc.Update()
        doc.Save()
    except:
        pass

# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()

    inv = connect()
    asm = inv.Documents.Open(ASSEMBLY_PATH, True)
    asm_def = asm.ComponentDefinition

    output = {
        "occurrences": {},
        "holes": [],
        "fastener_axes": [],
        "rivet_stacks": []
    }

    # =================================================
    # OCCURRENCES + TRANSFORMS
    # =================================================
    for occ in asm_def.Occurrences:
        m = occ.Transformation
        output["occurrences"][occ.Name] = {
            "definition": occ.Definition.Document.DisplayName,
            "transform": {
                "translation": [m.Cell(1,4), m.Cell(2,4), m.Cell(3,4)],
                "rotation": [
                    [m.Cell(1,1), m.Cell(1,2), m.Cell(1,3)],
                    [m.Cell(2,1), m.Cell(2,2), m.Cell(2,3)],
                    [m.Cell(3,1), m.Cell(3,2), m.Cell(3,3)]
                ]
            }
        }

    # =================================================
    # PART-LEVEL HOLE EXTRACTION (REAL GEOMETRY)
    # =================================================
    for occ in asm_def.Occurrences:
        doc = occ.Definition.Document
        if not doc.DisplayName.lower().endswith(".ipt"):
            continue

        force_rebuild(doc)
        comp = doc.ComponentDefinition

        for body in comp.SurfaceBodies:
            for face in body.Faces:
                if face.SurfaceType != kCylinderFace:
                    continue

                cyl = face.Geometry
                axis = cyl.Axis

                hole = {
                    "part": doc.DisplayName,
                    "occurrence": occ.Name,
                    "center": [
                        axis.RootPoint.X,
                        axis.RootPoint.Y,
                        axis.RootPoint.Z
                    ],
                    "direction": normalize([
                        axis.Direction.X,
                        axis.Direction.Y,
                        axis.Direction.Z
                    ]),
                    "diameter_mm": cyl.Radius * 2 * MM_PER_CM
                }

                output["holes"].append(hole)

    # =================================================
    # FASTENER AXIS (FROM SAME CYLINDER LOGIC)
    # =================================================
    for h in output["holes"]:
        if any(k in h["part"].upper() for k in ["RIVET", "FASTENER", "PIN"]):
            output["fastener_axes"].append(h)

    # =================================================
    # PHASE-5: BLIND RIVET STACK INFERENCE
    # =================================================
    for f in output["fastener_axes"]:
        stack = []

        for h in output["holes"]:
            if h["occurrence"] == f["occurrence"]:
                continue

            # axis alignment
            if abs(vec_dot(f["direction"], h["direction"])) < 0.95:
                continue

            # center proximity
            if dist(f["center"], h["center"]) > 1.5:
                continue

            # diameter compatibility (±0.3 mm)
            if abs(f["diameter_mm"] - h["diameter_mm"]) > 0.3:
                continue

            stack.append(h["occurrence"])

        if stack:
            output["rivet_stacks"].append({
                "fastener": f["occurrence"],
                "plates": sorted(set(stack)),
                "stack_size": len(set(stack)),
                "type": "blind_rivet",
                "confidence": 0.98
            })

    # =================================================
    # SAVE
    # =================================================
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4)

    print("✅ FINAL extraction complete")
    print(f"→ {OUT_JSON}")

    asm.Close(True)

# =====================================================
if __name__ == "__main__":
    run()