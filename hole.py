import json
import pythoncom
import win32com.client
from collections import defaultdict

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = r"E:\Phase 1\extractions\inferred_holes.json"

# =====================================================
# ENUMS
# =====================================================
kAssemblyDocument = 12291
kInsertConstraint = 100665344

# =====================================================
# INVENTOR CONNECTION
# =====================================================
def connect():
    try:
        inv = win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
    return inv


def open_assembly(inv):
    for d in inv.Documents:
        try:
            if d.FullFileName.lower() == ASSEMBLY_PATH.lower():
                return d
        except:
            pass
    return inv.Documents.Open(ASSEMBLY_PATH, True)


# =====================================================
# FASTENER DETECTION
# =====================================================
def is_fastener(occ):
    try:
        props = occ.Definition.Document.PropertySets.Item(
            "Design Tracking Properties"
        )
        desc = (props.Item("Description").Value or "").upper()
        return any(k in desc for k in ["RIVET", "BOLT", "SCREW", "PIN", "NUT"])
    except:
        return False


# =====================================================
# PHASE-3: CONSTRAINT-BASED HOLE INFERENCE
# =====================================================
def infer_holes_from_constraints(asm_def):
    hole_map = defaultdict(set)

    for c in asm_def.Constraints:
        if c.Type != kInsertConstraint:
            continue

        try:
            o1 = c.OccurrenceOne.Name
            o2 = c.OccurrenceTwo.Name
        except:
            continue

        hole_map[o1].add(o2)
        hole_map[o2].add(o1)

    results = []

    for occ in asm_def.Occurrences:
        if not is_fastener(occ):
            continue

        connected = list(hole_map.get(occ.Name, []))

        if len(connected) >= 2:
            results.append({
                "fastener": occ.Name,
                "hole_stack": connected,
                "confidence": round(0.8 + 0.05 * len(connected), 2)
            })

    return results


# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()

    inv = connect()
    doc = open_assembly(inv)

    if doc.DocumentType != kAssemblyDocument:
        raise RuntimeError("Not an assembly")

    asm_def = doc.ComponentDefinition
    holes = infer_holes_from_constraints(asm_def)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(holes, f, indent=4)

    print(f"✅ Inferred hole stacks: {len(holes)}")
    print(f"📁 Output written to: {OUTPUT_JSON}")


if __name__ == "__main__":
    run()