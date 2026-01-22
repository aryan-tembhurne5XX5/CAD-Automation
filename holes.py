import win32com.client
import json
import math
from pathlib import Path

# ==============================
# CONFIG
# ==============================
PART_PATH = r"E:\Phase 1\Assembly 1"      # folder with IPTs
OUTPUT_JSON = r"E:\Phase 1\extractions\part_holes.json"

# ==============================
# HELPERS
# ==============================
def mm(val_cm):
    return round(val_cm * 10, 4)

def pt_mm(pt):
    return [mm(pt.X), mm(pt.Y), mm(pt.Z)]

def vec(v):
    return [round(v.X, 4), round(v.Y, 4), round(v.Z, 4)]

# ==============================
# CONNECT INVENTOR
# ==============================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = True

results = []

# ==============================
# CORE EXTRACTION
# ==============================
def extract_holes_from_part(part_doc):
    cd = part_doc.ComponentDefinition
    holes_out = []

    # ---------- 1. Direct Hole Features ----------
    for hole in cd.Features.HoleFeatures:
        if hole.Suppressed:
            continue

        hdef = hole.Definition
        pdef = hole.PlacementDefinition

        # --- Diameter ---
        dia = None
        try:
            dia = hdef.Diameter.Value * 10
        except:
            try:
                dia = hdef.TapInfo.MajorDiameter * 10
            except:
                dia = None

        # --- Axis + Center ---
        if pdef.Type == 0:  # kSketchPlacementDefinition
            plane = pdef.Sketch.PlanarEntityGeometry
            axis_vec = vec(plane.Normal.AsVector())

            for pt in pdef.SketchPoints:
                center = pt_mm(pt.Geometry3d)
                holes_out.append({
                    "feature": hole.Name,
                    "diameter_mm": dia,
                    "axis": axis_vec,
                    "center_mm": center,
                    "patterned": False
                })

    # ---------- 2. Rectangular Patterns ----------
    for pat in cd.Features.RectangularPatternFeatures:
        if pat.Suppressed:
            continue

        parent = None
        for i in range(1, pat.ParentFeatures.Count + 1):
            pf = pat.ParentFeatures.Item(i)
            if pf.Type == 83886912:  # kHoleFeatureObject
                parent = pf
                break

        if not parent:
            continue

        hdef = parent.Definition
        pdef = parent.PlacementDefinition
        if pdef.Type != 0:
            continue

        plane = pdef.Sketch.PlanarEntityGeometry
        axis_vec = vec(plane.Normal.AsVector())

        try:
            dia = hdef.Diameter.Value * 10
        except:
            dia = None

        base_pt = pdef.SketchPoints.Item(1).Geometry3d

        for occ in pat.PatternElements:
            if occ.Suppressed:
                continue

            pt = base_pt.Copy()
            pt.TransformBy(occ.Transformation)

            holes_out.append({
                "feature": f"{pat.Name}:{occ.Index}",
                "diameter_mm": dia,
                "axis": axis_vec,
                "center_mm": pt_mm(pt),
                "patterned": True,
                "pattern_parent": pat.Name
            })

    return holes_out

# ==============================
# RUN FOR ALL PARTS
# ==============================
for ipt in Path(PART_PATH).glob("*.ipt"):
    print(f"🔍 {ipt.name}")
    doc = inv.Documents.Open(str(ipt), True)

    holes = extract_holes_from_part(doc)

    results.append({
        "part": ipt.name,
        "hole_count": len(holes),
        "holes": holes
    })

    doc.Close(True)

# ==============================
# SAVE
# ==============================
with open(OUTPUT_JSON, "w") as f:
    json.dump(results, f, indent=4)

print(f"\n✅ Hole extraction complete → {OUTPUT_JSON}")
inv.Quit()