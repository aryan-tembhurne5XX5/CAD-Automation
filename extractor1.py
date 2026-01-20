import win32com.client
import os
import json

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\CAD-Automation\1093144795-M1.iam"
OUTPUT_PATH = r"E:\Phase 1\CAD-Automation\assembly_extraction.json"

# =====================================================
# INVENTOR ENUM MAPS
# =====================================================
DOC_TYPE_MAP = {
    12291: "Part",
    12290: "Assembly"
}

CONSTRAINT_TYPE_MAP = {
    100665856: "Mate",
    100665088: "Flush",
    100665344: "Insert",
    100666368: "Angle"
}

# =====================================================
# HELPERS
# =====================================================
def get_document_type(doc):
    try:
        return DOC_TYPE_MAP.get(doc.DocumentType, "Unknown")
    except:
        return "Unknown"

def get_entity_type(entity):
    try:
        return entity.Type
    except:
        return "Unknown"

def extract_holes_from_part(part_doc):
    holes = []
    try:
        comp_def = part_doc.ComponentDefinition
        hole_features = comp_def.Features.HoleFeatures

        for h in hole_features:
            try:
                holes.append({
                    "diameter": float(h.Diameter.Value),
                    "hole_type": h.HoleType,
                    "threaded": h.Tapped,
                    "suppressed": h.Suppressed
                })
            except:
                continue
    except:
        pass

    return holes

# =====================================================
# CONNECT TO INVENTOR
# =====================================================
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = False

doc = inv.Documents.Open(ASSEMBLY_PATH)
asm_def = doc.ComponentDefinition

# =====================================================
# DATA CONTAINERS
# =====================================================
data = {
    "assembly_name": doc.DisplayName,
    "occurrences": [],
    "constraints": []
}

# =====================================================
# EXTRACT OCCURRENCES + HOLES
# =====================================================
def extract_occurrences(occurrences, parent=None):
    for occ in occurrences:
        part_doc = occ.Definition.Document
        doc_type = get_document_type(part_doc)

        part_info = {
            "name": occ.Name,
            "definition": part_doc.DisplayName,
            "document_type": doc_type,
            "parent": parent,
            "suppressed": occ.Suppressed,
            "visible": occ.Visible
        }

        # iProperties
        try:
            props = part_doc.PropertySets.Item("Design Tracking Properties")
            part_info["part_number"] = props.Item("Part Number").Value
            part_info["description"] = props.Item("Description").Value
        except:
            part_info["part_number"] = None
            part_info["description"] = None

        # Hole intelligence (ONLY FOR PARTS)
        if doc_type == "Part":
            holes = extract_holes_from_part(part_doc)
            part_info["hole_count"] = len(holes)
            part_info["holes"] = holes

        data["occurrences"].append(part_info)

        # Recursive subassemblies
        try:
            if occ.SubOccurrences.Count > 0:
                extract_occurrences(occ.SubOccurrences, occ.Name)
        except:
            pass

extract_occurrences(asm_def.Occurrences)

# =====================================================
# EXTRACT SEMANTIC CONSTRAINTS
# =====================================================
for c in asm_def.Constraints:
    try:
        constraint = {
            "constraint_type": CONSTRAINT_TYPE_MAP.get(c.Type, "Unknown"),
            "health": c.HealthStatus,
            "occurrence_1": c.OccurrenceOne.Name if hasattr(c, "OccurrenceOne") else None,
            "occurrence_2": c.OccurrenceTwo.Name if hasattr(c, "OccurrenceTwo") else None,
            "entity_1_type": get_entity_type(c.EntityOne),
            "entity_2_type": get_entity_type(c.EntityTwo)
        }
        data["constraints"].append(constraint)
    except:
        continue

# =====================================================
# SAVE OUTPUT
# =====================================================
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=4)

print(f"Extraction complete → {OUTPUT_PATH}")

# =====================================================
# CLEANUP
# =====================================================
doc.Close(True)
inv.Quit()