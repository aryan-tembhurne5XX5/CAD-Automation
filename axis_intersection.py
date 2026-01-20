import json
import pythoncom
import win32com.client
from pathlib import Path

OUTPUT_FILE = Path("holes.json")


def safe(val, default=None):
    try:
        return val
    except Exception:
        return default


def connect_inventor():
    try:
        inv = win32com.client.GetActiveObject("Inventor.Application")
    except Exception:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
    return inv


def get_active_assembly(inv):
    doc = inv.ActiveDocument
    if not doc or doc.DocumentType != 12291:  # kAssemblyDocumentObject
        raise RuntimeError("Active document is not an Assembly")
    return doc.ComponentDefinition


def get_insert_constraints(asm_def):
    inserts = []
    for c in asm_def.Constraints:
        if c.Type == 100665344:  # InsertConstraintEnum
            inserts.append(c)
    return inserts


def get_axis_from_constraint(constraint):
    try:
        return constraint.EntityOne
    except Exception:
        return None


def axis_intersects_body(axis, body, tg):
    """
    Returns True if axis line intersects solid body
    """
    try:
        axis_line = axis.Geometry
        points = tg.CreateObjectCollection()
        body.IntersectWithCurve(axis_line, points)
        return points.Count > 0
    except Exception:
        return False


def infer_holes(asm_def):
    tg = asm_def.Application.TransientGeometry
    holes = []

    occurrences = list(asm_def.Occurrences)
    inserts = get_insert_constraints(asm_def)

    for ins in inserts:
        axis = get_axis_from_constraint(ins)
        if not axis:
            continue

        fastener_occ = ins.OccurrenceOne
        axis_id = safe(fastener_occ.Name)

        stack = []

        for occ in occurrences:
            if occ.Suppressed:
                continue
            if occ == fastener_occ:
                continue

            try:
                bodies = occ.Definition.SurfaceBodies
            except Exception:
                continue

            for body in bodies:
                if axis_intersects_body(axis, body, tg):
                    stack.append(occ.Name)
                    break

        if len(stack) >= 2:
            holes.append({
                "axis_id": axis_id,
                "fastener_part": fastener_occ.Definition.Document.PropertySets
                    .Item("Design Tracking Properties")
                    .Item("Part Number").Value,
                "hole_stack": stack,
                "confidence": round(min(1.0, 0.7 + 0.05 * len(stack)), 2)
            })

    return holes


def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()
    asm_def = get_active_assembly(inv)

    holes = infer_holes(asm_def)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(holes, f, indent=4)

    print(f"✔ Inferred {len(holes)} holes")
    print(f"✔ Output written to {OUTPUT_FILE}")


if __name__ == "__main__":
    run()