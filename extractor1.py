import win32com.client
import json
import os

# =================================================
# USER INPUT
# =================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"

# =================================================
# CONNECT TO INVENTOR
# =================================================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = False  # set True if you want UI

doc = inv.Documents.Open(ASSEMBLY_PATH, True)

if doc.DocumentType != 12291:  # kAssemblyDocumentObject
    raise Exception("Provided file is not an Assembly (.iam)")

asm = doc.ComponentDefinition

# =================================================
# HELPERS
# =================================================
def matrix4x4(m):
    return [
        [m.Cell(1,1), m.Cell(1,2), m.Cell(1,3), m.Cell(1,4)],
        [m.Cell(2,1), m.Cell(2,2), m.Cell(2,3), m.Cell(2,4)],
        [m.Cell(3,1), m.Cell(3,2), m.Cell(3,3), m.Cell(3,4)],
        [m.Cell(4,1), m.Cell(4,2), m.Cell(4,3), m.Cell(4,4)]
    ]

def transform_point(M, p):
    return [
        M[0][0]*p.X + M[0][1]*p.Y + M[0][2]*p.Z + M[0][3],
        M[1][0]*p.X + M[1][1]*p.Y + M[1][2]*p.Z + M[1][3],
        M[2][0]*p.X + M[2][1]*p.Y + M[2][2]*p.Z + M[2][3]
    ]

def transform_vector(M, v):
    return [
        M[0][0]*v.X + M[0][1]*v.Y + M[0][2]*v.Z,
        M[1][0]*v.X + M[1][1]*v.Y + M[1][2]*v.Z,
        M[2][0]*v.X + M[2][1]*v.Y + M[2][2]*v.Z
    ]

# =================================================
# PASS 1 — OCCURRENCES
# =================================================
occurrences = []

for occ in asm.Occurrences:
    occurrences.append({
        "name": occ.Name,
        "definition": occ.Definition.Document.FullFileName,
        "grounded": occ.Grounded,
        "suppressed": occ.Suppressed,
        "transform": matrix4x4(occ.Transformation),
        "pattern_element": occ.PatternElement.Name if occ.PatternElement else None,
        "pattern_parent": occ.PatternElement.Parent.Name if occ.PatternElement else None
    })

# =================================================
# PASS 2 — CONSTRAINTS
# =================================================
constraints = []

for c in asm.Constraints:
    constraints.append({
        "name": c.Name,
        "type": c.Type,
        "occurrence_one": c.OccurrenceOne.Name if hasattr(c, "OccurrenceOne") else None,
        "occurrence_two": c.OccurrenceTwo.Name if hasattr(c, "OccurrenceTwo") else None,
        "entity_one": c.EntityOne.Type if hasattr(c, "EntityOne") else None,
        "entity_two": c.EntityTwo.Type if hasattr(c, "EntityTwo") else None,
        "offset": getattr(c, "Offset", None),
        "angle": getattr(c, "Angle", None),
        "suppressed": c.Suppressed
    })

# =================================================
# PASS 3 — COMPONENT PATTERNS
# =================================================
patterns = []

for p in asm.ComponentPatterns:
    patterns.append({
        "name": p.Name,
        "count": p.Count,
        "seed_occurrence": p.Occurrences(1).Name,
        "elements": [
            {
                "index": e.Index,
                "suppressed": e.Suppressed,
                "transform": matrix4x4(e.Transformation)
            } for e in p.PatternElements
        ]
    })

# =================================================
# PASS 4 — DERIVED HOLE GEOMETRY
# =================================================
holes = []

for occ in asm.Occurrences:
    if occ.DefinitionDocumentType != 12290:  # kPartDocumentObject
        continue

    part = occ.Definition.Document
    cd = part.ComponentDefinition
    M = matrix4x4(occ.Transformation)

    for hole in cd.Features.HoleFeatures:
        axis = hole.Axis
        holes.append({
            "occurrence": occ.Name,
            "part": part.DisplayName,
            "hole_name": hole.Name,
            "diameter": hole.HoleDiameter.Value,
            "center": transform_point(M, axis.RootPoint),
            "axis": transform_vector(M, axis.Direction),
            "threaded": hole.Tapped
        })

# =================================================
# WRITE JSON
# =================================================
output = {
    "assembly": doc.DisplayName,
    "file": ASSEMBLY_PATH,
    "occurrences": occurrences,
    "constraints": constraints,
    "patterns": patterns,
    "derived_geometry": {
        "holes": holes
    }
}

out_path = os.path.join(os.path.dirname(ASSEMBLY_PATH), "assembly_dump.json")

with open(out_path, "w", encoding="utf-8") as f:
    json.dump(output, f, indent=2)

print("✅ Assembly exported successfully:")
print(out_path)