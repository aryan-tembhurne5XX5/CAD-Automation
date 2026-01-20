import json
import pythoncom
import win32com.client

# ================= USER CONFIG =================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_FILE   = r"E:\Phase 1\extractions\final_inferred_holes.json"
# ===============================================

# ---- ENUMS ----
kAssemblyDocumentObject = 12291
kInsertConstraint = 100665344
kCylinderFace = 67119536


def connect_inventor():
    try:
        inv = win32com.client.GetActiveObject("Inventor.Application")
    except Exception:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
    return inv


def open_assembly(inv):
    for doc in inv.Documents:
        try:
            if doc.FullFileName.lower() == ASSEMBLY_PATH.lower():
                return doc
        except Exception:
            pass
    return inv.Documents.Open(ASSEMBLY_PATH, True)


def extract_axis(entity):
    """
    Correct way to extract axis from InsertConstraint geometry
    """
    try:
        if entity.Type == kCylinderFace:
            return entity.Geometry.Axis
    except Exception:
        pass
    return None


def infer_holes(asm_def):
    tg = asm_def.Application.TransientGeometry
    results = []

    occurrences = list(asm_def.Occurrences)

    for c in asm_def.Constraints:
        if c.Type != kInsertConstraint:
            continue

        axis = extract_axis(c.EntityOne) or extract_axis(c.EntityTwo)
        if not axis:
            continue

        fastener = c.OccurrenceOne
        hole_stack = []

        for occ in occurrences:
            if occ.Suppressed or occ == fastener:
                continue

            try:
                bodies = occ.Definition.SurfaceBodies
            except Exception:
                continue

            for body in bodies:
                try:
                    hits = tg.CreateObjectCollection()
                    body.IntersectWithCurve(axis, hits)
                    if hits.Count > 0:
                        hole_stack.append(occ.Name)
                        break
                except Exception:
                    pass

        if len(hole_stack) >= 2:
            results.append({
                "fastener_occurrence": fastener.Name,
                "fastener_part": fastener.Definition.Document.PropertySets
                    .Item("Design Tracking Properties")
                    .Item("Part Number").Value,
                "hole_stack": hole_stack,
                "confidence": round(min(1.0, 0.75 + 0.05 * len(hole_stack)), 2)
            })

    return results


def run():
    pythoncom.CoInitialize()

    inv = connect_inventor()
    asm_doc = open_assembly(inv)

    if asm_doc.DocumentType != kAssemblyDocumentObject:
        raise RuntimeError("Opened document is not an Assembly")

    asm_def = asm_doc.ComponentDefinition
    holes = infer_holes(asm_def)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(holes, f, indent=4)

    print(f"✅ Assembly: {asm_doc.FullFileName}")
    print(f"✅ Inferred holes: {len(holes)}")
    print(f"📁 Output written to: {OUTPUT_FILE}")


if __name__ == "__main__":
    run()