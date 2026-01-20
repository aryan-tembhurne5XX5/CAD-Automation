import json
from pathlib import Path
import win32com.client

# ==========================================================
# CONFIG
# ==========================================================
INPUT_JSON = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\assembly_with_holes.json")

# ==========================================================
# INVENTOR CONNECT
# ==========================================================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = True

# ==========================================================
# INVENTOR ENUMS (SAFE HARDCODE)
# ==========================================================
kPartDocument = 12291
kAssemblyDocument = 12290
kCutOperation = 20481
kCylinderSurface = 103

# ==========================================================
# GEOMETRY HELPERS
# ==========================================================
def vec_tuple(v):
    return [round(v.X, 5), round(v.Y, 5), round(v.Z, 5)]

def is_cylindrical(face):
    try:
        return face.SurfaceType == kCylinderSurface
    except:
        return False

# ==========================================================
# HOLE EXTRACTION FROM PART
# ==========================================================
def extract_holes_from_part(part_doc):
    holes = []
    comp_def = part_doc.ComponentDefinition
    features = comp_def.Features

    # ------------------------------------------------------
    # 1️⃣ Explicit HoleFeatures
    # ------------------------------------------------------
    for hf in features.HoleFeatures:
        try:
            hole = {
                "source": "HoleFeature",
                "diameter_mm": round(hf.HoleDiameter * 10, 3),
                "depth_mm": None,
                "through": hf.ExtentType == 0,
                "axis_vector": vec_tuple(hf.PlacementDefinition.AxisVector),
                "patterned": False
            }
            holes.append(hole)
        except:
            pass

    # ------------------------------------------------------
    # 2️⃣ Cylindrical Cut-Extrude inference
    # ------------------------------------------------------
    for ext in features.ExtrudeFeatures:
        try:
            if ext.Operation != kCutOperation:
                continue

            for face in ext.Faces:
                if not is_cylindrical(face):
                    continue

                hole = {
                    "source": "CutExtrude",
                    "diameter_mm": round(face.Geometry.Radius * 2, 3),
                    "depth_mm": None,
                    "through": ext.ExtentType == 0,
                    "axis_vector": vec_tuple(face.Geometry.AxisVector),
                    "patterned": False
                }
                holes.append(hole)
        except:
            pass

    return holes

# ==========================================================
# PART TYPE INFERENCE
# ==========================================================
def infer_part_type(part_number, description, holes):
    name = (part_number + " " + description).lower()

    if any(x in name for x in ["rivet", "bolt", "nut", "screw"]):
        return "Fastener"

    if holes:
        return "Plate"

    return "Structural"

# ==========================================================
# MAIN PIPELINE
# ==========================================================
def run():
    with open(INPUT_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    asm_doc = inv.ActiveDocument
    asm_def = asm_doc.ComponentDefinition

    for occ in data["occurrences"]:
        try:
            # ---------------------------------------------
            # Resolve occurrence safely
            # ---------------------------------------------
            occ_obj = asm_def.Occurrences.ItemByName(occ["name"])

            if occ_obj.Suppressed:
                occ["holes"] = []
                occ["hole_count"] = 0
                occ["inferred_part_type"] = "Suppressed"
                continue

            doc = occ_obj.Definition.Document

            # ---------------------------------------------
            # Skip sub-assemblies
            # ---------------------------------------------
            if doc.DocumentType != kPartDocument:
                occ["holes"] = []
                occ["hole_count"] = 0
                occ["inferred_part_type"] = "SubAssembly"
                continue

            # ---------------------------------------------
            # Extract holes
            # ---------------------------------------------
            holes = extract_holes_from_part(doc)

            occ["holes"] = holes
            occ["hole_count"] = len(holes)
            occ["inferred_part_type"] = infer_part_type(
                occ.get("part_number", ""),
                occ.get("description", ""),
                holes
            )

        except Exception as e:
            occ["holes"] = []
            occ["hole_count"] = 0
            occ["inferred_part_type"] = "Error"
            occ["hole_error"] = str(e)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    print("✅ Phase-3 Hole Inference COMPLETED")

# ==========================================================
if __name__ == "__main__":
    run()