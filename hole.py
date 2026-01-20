import json
import pythoncom
import win32com.client
from pathlib import Path


# ==========================
# CONFIG
# ==========================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"   # CHANGE THIS
OUTPUT_JSON = r"E:\Phase 1\extractions\phase3_holes.json"


# Inventor constants
kAssemblyDocument = 12290
kPartDocument = 12289


# ==========================
# UTILS
# ==========================
def safe_str(value):
    try:
        return str(value)
    except:
        return ""


def infer_part_type(description, hole_count):
    desc = (description or "").upper()

    if "RIVET" in desc or "BOLT" in desc or "SCREW" in desc:
        return "Fastener"

    if hole_count > 0:
        return "Plate"

    return "Unknown"


# ==========================
# HOLE EXTRACTION (PART)
# ==========================
def extract_holes_from_part(part_doc):
    holes = []

    try:
        comp_def = part_doc.ComponentDefinition
        features = comp_def.Features

        if not hasattr(features, "HoleFeatures"):
            return holes, None

        for hf in features.HoleFeatures:
            hole_data = {
                "diameter": safe_str(getattr(hf, "Diameter", None)),
                "hole_type": safe_str(hf.HoleType),
                "termination": safe_str(hf.TerminationType)
            }
            holes.append(hole_data)

        return holes, None

    except Exception as e:
        return [], safe_str(e)


# ==========================
# MAIN
# ==========================
def run():
    pythoncom.CoInitialize()

    inv = win32com.client.Dispatch("Inventor.Application")
    inv.Visible = True

    # --------------------------
    # OPEN / GET ASSEMBLY
    # --------------------------
    asm_doc = inv.ActiveDocument

    if asm_doc is None or asm_doc.DocumentType != kAssemblyDocument:
        asm_doc = inv.Documents.Open(ASSEMBLY_PATH, True)

    asm_def = asm_doc.ComponentDefinition

    result = {
        "assembly_name": Path(ASSEMBLY_PATH).name,
        "occurrences": [],
        "constraints": []
    }

    # --------------------------
    # OCCURRENCES
    # --------------------------
    for occ in asm_def.Occurrences:
        occ_data = {
            "name": occ.Name,
            "definition": safe_str(occ.Definition.Document.DisplayName),
            "document_type": "Assembly" if occ.Definition.Document.DocumentType == kAssemblyDocument else "Part",
            "parent": safe_str(occ.Parent.Name) if occ.Parent else None,
            "suppressed": occ.Suppressed,
            "visible": occ.Visible,
            "part_number": "",
            "description": "",
            "hole_count": 0,
            "holes": [],
            "inferred_part_type": "Unknown"
        }

        doc = occ.Definition.Document

        # --------------------------
        # PART ONLY → HOLES
        # --------------------------
        if doc.DocumentType == kPartDocument:
            try:
                prop_sets = doc.PropertySets
                design_props = prop_sets.Item("Design Tracking Properties")

                occ_data["part_number"] = safe_str(
                    design_props.Item("Part Number").Value
                )
                occ_data["description"] = safe_str(
                    design_props.Item("Description").Value
                )

            except:
                pass

            holes, hole_error = extract_holes_from_part(doc)
            occ_data["holes"] = holes
            occ_data["hole_count"] = len(holes)

            occ_data["inferred_part_type"] = infer_part_type(
                occ_data["description"],
                occ_data["hole_count"]
            )

            if hole_error:
                occ_data["hole_error"] = hole_error

        result["occurrences"].append(occ_data)

    # --------------------------
    # CONSTRAINTS
    # --------------------------
    try:
        for c in asm_def.Constraints:
            c_data = {
                "constraint_type": safe_str(c.Type),
                "health": safe_str(c.HealthStatus),
                "occurrence_1": safe_str(c.OccurrenceOne.Name) if hasattr(c, "OccurrenceOne") else None,
                "occurrence_2": safe_str(c.OccurrenceTwo.Name) if hasattr(c, "OccurrenceTwo") else None,
                "entity_1_type": safe_str(c.EntityOne.Type) if hasattr(c, "EntityOne") else None,
                "entity_2_type": safe_str(c.EntityTwo.Type) if hasattr(c, "EntityTwo") else None,
            }
            result["constraints"].append(c_data)

    except Exception as e:
        result["constraint_error"] = safe_str(e)

    # --------------------------
    # WRITE OUTPUT
    # --------------------------
    Path(OUTPUT_JSON).parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=4)

    print("✅ Phase-3 hole extraction complete")
    print("📄 Output:", OUTPUT_JSON)


if __name__ == "__main__":
    run()