import win32com.client
import pythoncom
import json
from pathlib import Path

# ==============================
# CONFIG
# ==============================
PART_PATH = r"E:\Phase 1\Assembly 1\1093144795-A.ipt"
OUT_JSON  = r"E:\Phase 1\extractions\part_hole_debug.json"

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
# SAFE VALUE
# ==============================
def safe(v):
    try:
        return float(v)
    except:
        return None

# ==============================
# MAIN EXTRACTION
# ==============================
def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()

    doc = inv.Documents.Open(PART_PATH, True)
    comp = doc.ComponentDefinition

    output = {
        "part": doc.DisplayName,
        "holes": []
    }

    # ------------------------------
    # 1. Hole FEATURES
    # ------------------------------
    for h in comp.Features.HoleFeatures:
        hole_entry = {
            "feature_name": h.Name,
            "hole_type": h.HoleType,
            "suppressed": h.Suppressed,
            "diameter_mm": None,
            "centers": [],
            "pattern_count": 1
        }

        # ---- Diameter (robust) ----
        try:
            hole_entry["diameter_mm"] = safe(h.HoleDiameter.Value * 10)
        except:
            pass

        # ---- Center points ----
        try:
            for p in h.HoleCenterPoints:
                hole_entry["centers"].append([
                    safe(p.X), safe(p.Y), safe(p.Z)
                ])
        except:
            pass

        # ---- Pattern participation ----
        try:
            parents = h.ParentFeatures
            if parents and parents.Count > 0:
                hole_entry["pattern_count"] = parents.Count
        except:
            pass

        output["holes"].append(hole_entry)

    # ------------------------------
    # 2. CYLINDRICAL FACES (truth)
    # ------------------------------
    output["cylindrical_faces"] = []

    for body in comp.SurfaceBodies:
        for face in body.Faces:
            try:
                geom = face.Geometry
                if geom.SurfaceType == 5891:  # Cylinder
                    output["cylindrical_faces"].append({
                        "radius_mm": safe(geom.Radius * 10),
                        "axis_direction": [
                            safe(geom.AxisVector.X),
                            safe(geom.AxisVector.Y),
                            safe(geom.AxisVector.Z)
                        ],
                        "base_point": [
                            safe(geom.BasePoint.X),
                            safe(geom.BasePoint.Y),
                            safe(geom.BasePoint.Z)
                        ]
                    })
            except:
                continue

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4)

    print("✅ Diagnostic hole extraction complete")
    print(f"→ {OUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()