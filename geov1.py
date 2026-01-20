import win32com.client
import json
import os
import time
import math

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = r"E:\Phase 1\extractions\phase4_fastener_axes.json"

# =====================================================
# INVENTOR CONNECTION
# =====================================================
inv = win32com.client.DispatchEx("Inventor.Application")
inv.Visible = True
time.sleep(2)

doc = inv.Documents.Open(ASSEMBLY_PATH, True)
asm_def = doc.ComponentDefinition

# =====================================================
# VECTOR HELPERS
# =====================================================
def normalize(v):
    mag = math.sqrt(sum(x*x for x in v))
    return [x/mag for x in v] if mag else v

def transform_point(t, p):
    return [
        t.Cell(1,1)*p[0] + t.Cell(1,2)*p[1] + t.Cell(1,3)*p[2] + t.Cell(1,4),
        t.Cell(2,1)*p[0] + t.Cell(2,2)*p[1] + t.Cell(2,3)*p[2] + t.Cell(2,4),
        t.Cell(3,1)*p[0] + t.Cell(3,2)*p[1] + t.Cell(3,3)*p[2] + t.Cell(3,4),
    ]

def transform_vector(t, v):
    return normalize([
        t.Cell(1,1)*v[0] + t.Cell(1,2)*v[1] + t.Cell(1,3)*v[2],
        t.Cell(2,1)*v[0] + t.Cell(2,2)*v[1] + t.Cell(2,3)*v[2],
        t.Cell(3,1)*v[0] + t.Cell(3,2)*v[1] + t.Cell(3,3)*v[2],
    ])

# =====================================================
# PHASE-4.2 — AXIS EXTRACTION (CORRECT)
# =====================================================
fastener_axes = {}

for c in asm_def.Constraints:
    try:
        if c.Type != 100665344:  # Insert
            continue

        pairs = [
            (c.OccurrenceOne, c.EntityOne),
            (c.OccurrenceOne, c.EntityTwo),
            (c.OccurrenceTwo, c.EntityOne),
            (c.OccurrenceTwo, c.EntityTwo),
        ]

        for occ, ent in pairs:
            if ent.Type != 67120288:  # Axis
                continue

            geom = ent.Geometry

            origin_local = [
                geom.RootPoint.X,
                geom.RootPoint.Y,
                geom.RootPoint.Z
            ]

            dir_local = [
                geom.Direction.X,
                geom.Direction.Y,
                geom.Direction.Z
            ]

            t = occ.Transformation
            origin_world = transform_point(t, origin_local)
            dir_world = transform_vector(t, dir_local)

            fastener_axes[occ.Name] = {
                "origin": origin_world,
                "direction": dir_world,
                "confidence": 1.0,
                "source": "InsertConstraint"
            }

    except:
        continue

# =====================================================
# SAVE
# =====================================================
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(fastener_axes, f, indent=4)

print(f"✅ Phase-4.2 complete")
print(f"   Fastener axes extracted: {len(fastener_axes)}")
print(f"   Output → {OUTPUT_JSON}")

# =====================================================
# CLEANUP
# =====================================================
doc.Close(True)
inv.Quit()