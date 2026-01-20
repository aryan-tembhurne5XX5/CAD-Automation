import json
from pathlib import Path
import win32com.client
import pythoncom

# =========================
# CONFIG
# ==========================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\Assembly1.iam"   # CHANGE THIS
OUTPUT_JSON = r"E:\Phase 1\extractions\phase3_holes.json"


# Inventor constants
kPartDocument = 12290
kAssemblyDocument = 12291
kHoleFeatureObject = 8388


# =========================
# UTILS
# =========================
def safe_str(val):
    try:
        return str(val)
    except:
        return ""


def connect_inventor():
    pythoncom.CoInitialize()
    inv = win32com.client.Dispatch("Inventor.Application")
    inv.Visible = True
    return inv


# =========================
# HOLE EXTRACTION
# =========================
def extract_holes_from_part(part_doc):
    holes = []

    try:
        comp_def = part_doc.ComponentDefinition
        features = comp_def.Features

        for feat in features:
            try:
                if feat.Type == kHoleFeatureObject:
                    hole = {
                        "diameter": None,
                        "depth": None,
                        "threaded": False,
                        "hole_type": safe_str(feat.HoleType)
                    }

                    try:
                        hole["diameter"] = feat.Diameter.Value
                    except:
                        pass

                    try:
                        hole["depth"] = feat.Depth.Value
                    except:
                        pass

                    try:
                        hole["threaded"] = feat.ThreadInfo is not None
                    except:
                        pass

                    holes.append(hole)

            except:
                continue

    except Exception as e:
        return [], str(e)

    return holes, None


# =========================
# MAIN RUN
# =========================
def run():
    inv = connect_inventor()

    # Open assembly safely
    try:
        asm_doc = inv.Documents.Open(ASSEMBLY_PATH, True)
    except Exception as e:
        print("❌ Failed to open assembly:", e)
        return

    if asm_doc.DocumentType != kAssemblyDocument:
        print("❌ Not an assembly document")
        return

    asm_def = asm_doc.ComponentDefinition

    result = {
        "assembly_name": Path(ASSEMBLY_PATH).name,
        "occurrences": [],
        "constraints": []
    }

    # =========================
    # OCCURRENCES
    # =========================
    for occ in asm_def.Occurrences:
        parent_name = None
        try:
            if hasattr(occ, "ParentOccurrence") and occ.ParentOccurrence:
                parent_name = occ.ParentOccurrence.Name
        except:
            parent_name = None

        occ_data = {
            "name": occ.Name,
            "definition": safe_str(occ.Definition.Document.DisplayName),
            "document_type": "Assembly"
            if occ.Definition.Document.DocumentType == kAssemblyDocument
            else "Part",
            "parent": parent_name,
            "suppressed": occ.Suppressed,
            "visible": occ.Visible,
            "part_number": "",
            "description": "",
            "hole_count": 0,
            "holes": [],
            "inferred_part_type": "Unknown"
        }

        # -------- Part metadata --------
        try:
            props = occ.Definition.Document.PropertySets
            design = props.Item("Design Tracking Properties")
            occ_data["part_number"] = safe_str(design.Item("Part Number").Value)
            occ_data["description"] = safe_str(design.Item("Description").Value)
        except:
            pass

        # -------- Hole extraction --------
        if occ.Definition.Document.DocumentType == kPartDocument:
            try:
                holes, err = extract_holes_from_part(occ.Definition.Document)
                occ_data["holes"] = holes
                occ_data["hole_count"] = len(holes)
                if err:
                    occ_data["hole_error"] = err
            except Exception as e:
                occ_data["hole_error"] = str(e)

        result["occurrences"].append(occ_data)

    # =========================
    # CONSTRAINTS
    # =========================
    for c in asm_def.Constraints:
        try:
            c_data = {
                "constraint_type": c.Type,
                "health": c.HealthStatus,
                "occurrence_1": safe_str(c.OccurrenceOne.Name),
                "occurrence_2": safe_str(c.OccurrenceTwo.Name),
                "entity_1_type": safe_str(c.EntityOne.Type),
                "entity_2_type": safe_str(c.EntityTwo.Type)
            }
            result["constraints"].append(c_data)
        except:
            continue

    # =========================
    # WRITE OUTPUT
    # =========================
    Path(OUTPUT_JSON).parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=4)

    print("✅ Phase-3 extraction complete")
    print("📄 Output:", OUTPUT_JSON)


# =========================
# ENTRY
# =========================
if __name__ == "__main__":
    run()