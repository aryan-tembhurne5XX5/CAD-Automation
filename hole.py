import json
import math
from pathlib import Path
import win32com.client

# ===============================
# CONFIG
# ===============================
INPUT_JSON = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\assembly_with_holes.json")

# ===============================
# INVENTOR CONNECT
# ===============================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = False

# ===============================
# GEOMETRY HELPERS
# ===============================
def vec_to_tuple(v):
    return (round(v.X, 5), round(v.Y, 5), round(v.Z, 5))

def is_cylindrical(face):
    try:
        return face.SurfaceType == 103  # kCylinderSurface
    except:
        return False

def face_diameter(face):
    try:
        return round(face.Geometry.Radius * 2, 4)
    except:
        return None

# ===============================
# HOLE EXTRACTION CORE
# ===============================
def extract_holes_from_part(part_doc):
    holes = []

    comp_def = part_doc.ComponentDefinition
    features = comp_def.Features

    # ----------------------------------
    # 1️⃣ HoleFeatures (explicit)
    # ----------------------------------
    for hf in features.HoleFeatures:
        try:
            hole = {
                "source": "HoleFeature",
                "diameter_mm": round(hf.HoleDiameter * 10, 3),
                "depth_mm": None if hf.ExtentType == 0 else round(hf.Extent.Distance * 10, 3),
                "through": hf.ExtentType == 0,
                "axis_vector": vec_to_tuple(hf.PlacementDefinition.AxisVector),
                "patterned": False
            }
            holes.append(hole)
        except:
            continue

    # ----------------------------------
    # 2️⃣ Cut-Extrude Cylindrical Holes
    # ----------------------------------
    for ext in features.ExtrudeFeatures:
        try:
            if ext.Operation != 20481:  # kCutOperation
                continue

            for face in ext.Faces:
                if not is_cylindrical(face):
                    continue

                hole = {
                    "source": "CutExtrude",
                    "diameter_mm": face_diameter(face),
                    "depth_mm": round(ext.Extent.Distance * 10, 3)
                        if hasattr(ext.Extent, "Distance") else None,
                    "through": ext.ExtentType == 0,
                    "axis_vector": vec_to_tuple(face.Geometry.AxisVector),
                    "patterned": False
                }
                holes.append(hole)
        except:
            continue

    # ----------------------------------
    # 3️⃣ Pattern Multiplication
    # ----------------------------------
    for pattern in features.RectangularPatternFeatures:
        try:
            count = pattern.CountX * pattern.CountY
            for _ in range(count - 1):
                if holes:
                    dup = holes[-1].copy()
                    dup["patterned"] = True
                    holes.append(dup)
        except:
            pass

    for pattern in features.CircularPatternFeatures:
        try:
            count = pattern.Count
            for _ in range(count - 1):
                if holes:
                    dup = holes[-1].copy()
                    dup["patterned"] = True
                    holes.append(dup)
        except:
            pass

    return holes

# ===============================
# PART TYPE INFERENCE
# ===============================
def infer_part_type(occ, holes):
    name = occ["part_number"].lower()

    if "rivet" in name or "bolt" in name or "nut" in name:
        return "Fastener"

    if holes and all(h["through"] for h in holes):
        return "Plate"

    return "Structural"

# ===============================
# MAIN PIPELINE
# ===============================
def run():
    with open(INPUT_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    for occ in data["occurrences"]:
        try:
            part_path = occ["definition"]
            doc = inv.Documents.Open(part_path, False)

            if doc.DocumentType != 12291:  # PartDocument
                occ["holes"] = []
                occ["hole_count"] = 0
                occ["inferred_part_type"] = "Structural"
                doc.Close()
                continue

            holes = extract_holes_from_part(doc)

            occ["holes"] = holes
            occ["hole_count"] = len(holes)
            occ["inferred_part_type"] = infer_part_type(occ, holes)

            doc.Close()

        except Exception as e:
            occ["holes"] = []
            occ["hole_count"] = 0
            occ["inferred_part_type"] = "Unknown"
            occ["hole_error"] = str(e)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    print("✅ Phase-3 Hole Inference Completed")

# ===============================
if __name__ == "__main__":
    run()