import win32com.client
import json
import os

# ============================================================
# CONFIG
# ============================================================
ASSEMBLY_PATH = r"C:\InventorModels\MainAssembly.iam"

# ============================================================
# CONNECT INVENTOR
# ============================================================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = False
doc = inv.Documents.Open(ASSEMBLY_PATH, True)

if doc.DocumentType != 12291:
    raise Exception("Not an Assembly document")

asm = doc.ComponentDefinition
tg = inv.TransientGeometry

# ============================================================
# HELPERS
# ============================================================
def matrix(m):
    return [[m.Cell(i, j) for j in range(1,5)] for i in range(1,5)]

def vec(v): return [v.X, v.Y, v.Z]

def point(p): return [p.X, p.Y, p.Z]

# ============================================================
# EXPLODED VIEW (LOSSLESS)
# ============================================================
exploded = None
if doc.ActiveView.RepresentationType == 12294:  # kExplodedViewRepresentation
    exploded = {
        "name": doc.ActiveView.Name,
        "camera": {
            "eye": point(doc.ActiveView.Camera.Eye),
            "target": point(doc.ActiveView.Camera.Target)
        }
    }

# ============================================================
# OCCURRENCES (FULL)
# ============================================================
occurrences = []

for occ in asm.Occurrences:
    occurrences.append({
        "name": occ.Name,
        "path": occ.Definition.Document.FullFileName,
        "grounded": occ.Grounded,
        "visible": occ.Visible,
        "suppressed": occ.Suppressed,
        "transform": matrix(occ.Transformation),
        "pattern": {
            "element": occ.PatternElement.Name if occ.PatternElement else None,
            "parent": occ.PatternElement.Parent.Name if occ.PatternElement else None
        }
    })

# ============================================================
# CONSTRAINTS (ALL TYPES)
# ============================================================
constraints = []

for c in asm.Constraints:
    constraints.append({
        "name": c.Name,
        "type": c.Type,
        "suppressed": c.Suppressed,
        "occ1": getattr(c, "OccurrenceOne", None).Name if hasattr(c, "OccurrenceOne") else None,
        "occ2": getattr(c, "OccurrenceTwo", None).Name if hasattr(c, "OccurrenceTwo") else None,
        "offset": getattr(c, "Offset", None),
        "angle": getattr(c, "Angle", None)
    })

# ============================================================
# COMPONENT PATTERNS
# ============================================================
patterns = []

for p in asm.ComponentPatterns:
    patterns.append({
        "name": p.Name,
        "count": p.Count,
        "elements": [
            {
                "index": e.Index,
                "suppressed": e.Suppressed,
                "transform": matrix(e.Transformation)
            } for e in p.PatternElements
        ]
    })

# ============================================================
# PART DATA EXTRACTION
# ============================================================
parts = {}

for occ in asm.Occurrences:
    if occ.DefinitionDocumentType != 12290:
        continue

    part = occ.Definition.Document
    cd = part.ComponentDefinition

    pdata = {
        "features": [],
        "sketches": []
    }

    # ---------------- HOLES ----------------
    for h in cd.Features.HoleFeatures:
        axis = h.Axis
        pdata["features"].append({
            "type": "Hole",
            "name": h.Name,
            "diameter": h.HoleDiameter.Value,
            "depth": getattr(h, "HoleDepth", None),
            "tapped": h.Tapped,
            "axis": {
                "origin": point(axis.RootPoint),
                "direction": vec(axis.Direction)
            }
        })

    # ---------------- CUTS ----------------
    for c in cd.Features.ExtrudeFeatures:
        pdata["features"].append({
            "type": "Extrude",
            "name": c.Name,
            "operation": c.Operation
        })

    # ---------------- SKETCHES ----------------
    for sk in cd.Sketches:
        sketch = {
            "name": sk.Name,
            "points": [],
            "dimensions": []
        }

        for p in sk.SketchPoints:
            sketch["points"].append(point(p.Geometry))

        for d in sk.DimensionConstraints:
            sketch["dimensions"].append({
                "value": d.Parameter.Value,
                "type": d.Type
            })

        pdata["sketches"].append(sketch)

    parts[occ.Name] = pdata

# ============================================================
# FINAL JSON
# ============================================================
dump = {
    "assembly": doc.DisplayName,
    "file": ASSEMBLY_PATH,
    "exploded_view": exploded,
    "occurrences": occurrences,
    "constraints": constraints,
    "patterns": patterns,
    "parts": parts
}

out = os.path.join(os.path.dirname(ASSEMBLY_PATH), "FULL_ASSEMBLY_DUMP.json")
with open(out, "w") as f:
    json.dump(dump, f, indent=2)

print("✅ FULL LOSSLESS EXPORT COMPLETE")
print(out)