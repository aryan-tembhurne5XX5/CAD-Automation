import win32com.client
import os
import json
import time

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\CAD-Automation\Assembly 1\1093144795-M1.iam"
OUTPUT_PATH = r"E:\Phase 1\extractions\assembly_extraction.json"

# =====================================================
# ENUM MAPS (Inventor)
# =====================================================
CONSTRAINT_TYPE_MAP = {
    100665856: "Mate",
    100665088: "Flush",
    100665344: "Insert",
    100666368: "Angle"
}

ENTITY_TYPE_MAP = {
    67119520: "Face",
    67120288: "Axis",
    83887776: "Edge"
}

# =====================================================
# SAFE HELPERS (PATCHED)
# =====================================================
def get_document_type(doc):
    """
    PATCH #1:
    Do NOT trust Inventor DocumentType blindly.
    Use file extension as authoritative fallback.
    """
    try:
        name = doc.DisplayName.lower()
        if name.endswith(".ipt"):
            return "Part"
        if name.endswith(".iam"):
            return "Assembly"
        return "Unknown"
    except:
        return "Unknown"


def get_entity_type(entity):
    """
    PATCH #2:
    Convert raw Inventor enums → semantic labels.
    """
    try:
        return ENTITY_TYPE_MAP.get(entity.Type, "Unknown")
    except:
        return "Unknown"


def extract_holes_from_part(part_doc):
    """
    PATCH #3:
    Guaranteed hole extraction for rivet intelligence.
    """
    holes = []
    try:
        comp_def = part_doc.ComponentDefinition
        hole_features = comp_def.Features.HoleFeatures

        for h in hole_features:
            try:
                holes.append({
                    "diameter": float(h.Diameter.Value),
                    "threaded": bool(h.Tapped),
                    "suppressed": bool(h.Suppressed)
                })
            except:
                continue
    except:
        pass

    return holes


# =====================================================
# INVENTOR CONNECTION (SAFE)
# =====================================================
if not os.path.isfile(ASSEMBLY_PATH):
    raise FileNotFoundError(f"Assembly not found: {ASSEMBLY_PATH}")

inv = win32com.client.DispatchEx("Inventor.Application")
inv.Visible = True
time.sleep(2)

doc = inv.Documents.Open(ASSEMBLY_PATH, True)
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
# OCCURRENCE EXTRACTION (PATCHED)
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
            "suppressed": bool(occ.Suppressed),
            "visible": bool(occ.Visible)
        }

        # iProperties
        try:
            props = part_doc.PropertySets.Item("Design Tracking Properties")
            part_info["part_number"] = props.Item("Part Number").Value
            part_info["description"] = props.Item("Description").Value
        except:
            part_info["part_number"] = None
            part_info["description"] = None

        # PATCH #3 APPLIED HERE
        if doc_type == "Part":
            holes = extract_holes_from_part(part_doc)
            part_info["hole_count"] = len(holes)
            part_info["holes"] = holes

        data["occurrences"].append(part_info)

        # Recursive for subassemblies
        try:
            if occ.SubOccurrences.Count > 0:
                extract_occurrences(occ.SubOccurrences, occ.Name)
        except:
            pass


extract_occurrences(asm_def.Occurrences)

# =====================================================
# CONSTRAINT EXTRACTION (PATCHED)
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

print(f"✅ Extraction complete → {OUTPUT_PATH}")

# =====================================================
# CLEANUP
# =====================================================
doc.Close(True)
inv.Quit()