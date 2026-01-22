import win32com.client
import json
import os
import time
import math

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = r"E:\Phase 1\extractions\assembly_dump.json"

# =====================================================
# BASIC UTILITIES
# =====================================================
def mat4(m):
    return [
        [m.Cell(1,1), m.Cell(1,2), m.Cell(1,3), m.Cell(1,4)],
        [m.Cell(2,1), m.Cell(2,2), m.Cell(2,3), m.Cell(2,4)],
        [m.Cell(3,1), m.Cell(3,2), m.Cell(3,3), m.Cell(3,4)],
        [m.Cell(4,1), m.Cell(4,2), m.Cell(4,3), m.Cell(4,4)],
    ]

def transform_point(M, p):
    return [
        round(M[0][0]*p.X + M[0][1]*p.Y + M[0][2]*p.Z + M[0][3], 6),
        round(M[1][0]*p.X + M[1][1]*p.Y + M[1][2]*p.Z + M[1][3], 6),
        round(M[2][0]*p.X + M[2][1]*p.Y + M[2][2]*p.Z + M[2][3], 6),
    ]

def transform_vector(M, v):
    return [
        round(M[0][0]*v.X + M[0][1]*v.Y + M[0][2]*v.Z, 6),
        round(M[1][0]*v.X + M[1][1]*v.Y + M[1][2]*v.Z, 6),
        round(M[2][0]*v.X + M[2][1]*v.Y + M[2][2]*v.Z, 6),
    ]

# =====================================================
# CONNECT TO INVENTOR
# =====================================================
try:
    inv = win32com.client.GetActiveObject("Inventor.Application")
except:
    inv = win32com.client.DispatchEx("Inventor.Application")
    inv.Visible = True
    time.sleep(5)

doc = inv.Documents.Open(ASSEMBLY_PATH, True)
asm = doc.ComponentDefinition

# =====================================================
# DATA STRUCTURE
# =====================================================
data = {
    "assembly": doc.DisplayName,
    "occurrences": [],
    "constraints": [],
    "patterns": [],
    "holes": []
}

# =====================================================
# PASS 1 — OCCURRENCES
# =====================================================
for occ in asm.Occurrences:
    try:
        M = mat4(occ.Transformation)
        data["occurrences"].append({
            "name": occ.Name,
            "definition": occ.Definition.Document.DisplayName,
            "full_path": occ.Definition.Document.FullFileName,
            "suppressed": bool(occ.Suppressed),
            "grounded": bool(occ.Grounded),
            "transform": M,
            "pattern_parent": occ.PatternElement.Parent.Name if occ.PatternElement else None
        })
    except:
        continue

# =====================================================
# PASS 2 — CONSTRAINTS
# =====================================================
for c in asm.Constraints:
    try:
        data["constraints"].append({
            "name": c.Name,
            "type": c.Type,
            "occurrence_1": c.OccurrenceOne.Name if hasattr(c, "OccurrenceOne") else None,
            "occurrence_2": c.OccurrenceTwo.Name if hasattr(c, "OccurrenceTwo") else None,
            "entity_1": c.EntityOne.Type if hasattr(c, "EntityOne") else None,
            "entity_2": c.EntityTwo.Type if hasattr(c, "EntityTwo") else None,
            "suppressed": bool(c.Suppressed)
        })
    except:
        continue

# =====================================================
# PASS 3 — COMPONENT PATTERNS (CORRECT API)
# =====================================================
features = asm.Features

for pat in features.RectangularPatternFeatures:
    try:
        data["patterns"].append({
            "name": pat.Name,
            "type": "Rectangular",
            "count": pat.PatternElements.Count,
            "elements": [
                {
                    "index": e.Index,
                    "suppressed": bool(e.Suppressed),
                    "transform": mat4(e.Transformation)
                }
                for e in pat.PatternElements
            ]
        })
    except:
        continue

for pat in features.CircularPatternFeatures:
    try:
        data["patterns"].append({
            "name": pat.Name,
            "type": "Circular",
            "count": pat.PatternElements.Count,
            "elements": [
                {
                    "index": e.Index,
                    "suppressed": bool(e.Suppressed),
                    "transform": mat4(e.Transformation)
                }
                for e in pat.PatternElements
            ]
        })
    except:
        continue

# =====================================================
# PASS 4 — HOLE GEOMETRY (ONLY SAFE METHOD)
# =====================================================
for occ in asm.Occurrences:
    try:
        part_doc = occ.Definition.Document
        if not part_doc.DisplayName.lower().endswith(".ipt"):
            continue

        cd = part_doc.ComponentDefinition
        M = mat4(occ.Transformation)

        for hole in cd.Features.HoleFeatures:
            if hole.Suppressed:
                continue

            pd = hole.PlacementDefinition
            if pd.Type != 0:  # NOT sketch-based → skip (unstable)
                continue

            sketch = pd.Sketch
            normal = sketch.PlanarEntityGeometry.Normal.AsVector()

            for pt in pd.SketchPoints:
                p3d = pt.Geometry3d
                data["holes"].append({
                    "occurrence": occ.Name,
                    "part": part_doc.DisplayName,
                    "hole": hole.Name,
                    "diameter_mm": round(hole.HoleDefinition.Diameter.Value * 10, 4),
                    "center_mm": transform_point(M, p3d),
                    "axis": transform_vector(M, normal),
                    "threaded": bool(hole.HoleDefinition.Tapped)
                })
    except:
        continue

# =====================================================
# SAVE JSON
# =====================================================
os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)

print("✅ Extraction complete")
print("📄 Output:", OUTPUT_JSON)

# =====================================================
# CLEANUP
# =====================================================
doc.Close(True)