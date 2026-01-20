import win32com.client
import os
import json

# ----------------------------
# CONFIG
# ----------------------------
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_DIR = r"E:\Phase 1\extractions"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ----------------------------
# CONNECT TO INVENTOR
# ----------------------------
inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = False

doc = inv.Documents.Open(ASSEMBLY_PATH)
asm_def = doc.ComponentDefinition

# ----------------------------
# DATA CONTAINERS
# ----------------------------
assembly_data = {
    "assembly_name": os.path.basename(ASSEMBLY_PATH),
    "occurrences": [],
    "constraints": [],
    "patterns": []
}

# ----------------------------
# EXTRACT OCCURRENCES (Hierarchy + Metadata)
# ----------------------------
def extract_occurrences(occurrences, parent=None):
    for occ in occurrences:
        part_info = {
            "name": occ.Name,
            "definition": occ.Definition.Document.DisplayName,
            "document_type": occ.Definition.DocumentType,
            "parent": parent,
            "suppressed": occ.Suppressed,
            "visible": occ.Visible
        }

        # iProperties (important for ML)
        try:
            props = occ.Definition.Document.PropertySets
            design_props = props.Item("Design Tracking Properties")
            part_info["part_number"] = design_props.Item("Part Number").Value
            part_info["description"] = design_props.Item("Description").Value
        except:
            part_info["part_number"] = None
            part_info["description"] = None

        assembly_data["occurrences"].append(part_info)

        # Recursive for subassemblies
        if occ.SubOccurrences.Count > 0:
            extract_occurrences(occ.SubOccurrences, occ.Name)

# Run extraction
extract_occurrences(asm_def.Occurrences)

# ----------------------------
# EXTRACT CONSTRAINTS (RULE GOLD)
# ----------------------------
for c in asm_def.Constraints:
    constraint_info = {
        "type": c.Type,
        "health": c.HealthStatus,
        "entity_one": str(c.EntityOne),
        "entity_two": str(c.EntityTwo)
    }
    assembly_data["constraints"].append(constraint_info)

# ----------------------------
# EXTRACT PATTERNS (IMPORTANT FOR QUANTITY RULES)
# ----------------------------
for p in asm_def.OccurrencePatterns:
    pattern_info = {
        "name": p.Name,
        "count": p.OccurrenceCount,
        "suppressed": p.Suppressed
    }
    assembly_data["patterns"].append(pattern_info)

# ----------------------------
# SAVE OUTPUT
# ----------------------------
output_file = os.path.join(OUTPUT_DIR, "assembly_extraction.json")
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(assembly_data, f, indent=4)

print(f"Extraction complete: {output_file}")

# ----------------------------
# CLEANUP
# ----------------------------
doc.Close(True)
inv.Quit()