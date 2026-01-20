import json
import pythoncom
import win32com.client
from pathlib import Path

OUTPUT_FILE = Path("final_inferred_holes.json")

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


def get_active_assembly(inv):
    doc = inv.ActiveDocument
    if not doc or doc.DocumentType != kAssemblyDocumentObject:
        raise RuntimeError("Active document must be an Assembly")
    return doc.ComponentDefinition


def extract_axis_from_entity(entity):
    """
    Robustly extract an axis from InsertConstraint entity
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

        axis = None
        axis = extract_axis_from_entity(c.EntityOne) or extract_axis_from_entity(c.EntityTwo)
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
    asm_def = get_active_assembly(inv)

    holes = infer_holes(asm_def)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(holes, f, indent=4)

    print(f"✅ FINAL inferred holes: {len(holes)}")
    print(f"📁 Output: {OUTPUT_FILE}")


if __name__ == "__main__":
    run()