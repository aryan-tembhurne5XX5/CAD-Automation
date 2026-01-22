import win32com.client
import json
import os
import time

ASM_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUT_JSON = r"E:\Phase 1\extractions\constraints_full.json"

inv = win32com.client.DispatchEx("Inventor.Application")
inv.Visible = True
time.sleep(2)

doc = inv.Documents.Open(ASM_PATH, True)
asm = doc.ComponentDefinition

def get_entity_data(entity):
    if entity is None:
        return None

    data = {
        "type": entity.Type,
        "surface": None,
        "reference_key": None
    }

    try:
        data["reference_key"] = entity.ReferenceKey
    except:
        pass

    try:
        geom = entity.Geometry
        if geom:
            data["surface"] = geom.Type
    except:
        pass

    return data

constraints = []

for c in asm.Constraints:
    try:
        constraints.append({
            "name": c.Name,
            "type": c.Type,
            "occurrence_1": c.OccurrenceOne.Name if hasattr(c, "OccurrenceOne") else None,
            "occurrence_2": c.OccurrenceTwo.Name if hasattr(c, "OccurrenceTwo") else None,
            "entity_1": get_entity_data(c.EntityOne),
            "entity_2": get_entity_data(c.EntityTwo),
            "suppressed": c.Suppressed
        })
    except:
        continue

with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(constraints, f, indent=2)

print("✅ Full constraint geometry extracted:", OUT_JSON)

doc.Close(True)
inv.Quit()