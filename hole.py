import json
import pythoncom
import win32com.client
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = r"E:\Phase 1\extractions\inferred_holes.json"

# =====================================================
# ENUMS
# =====================================================
kAssemblyDocument = 12291

# =====================================================
# INVENTOR CONNECTION
# =====================================================
def get_inventor():
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
        except:
            pass
    return inv.Documents.Open(ASSEMBLY_PATH, True)


# =====================================================
# FASTENER DETECTION (ROBUST)
# =====================================================
def is_fastener(occ):
    try:
        props = occ.Definition.Document.PropertySets.Item(
            "Design Tracking Properties"
        )
        desc = (props.Item("Description").Value or "").upper()
        if any(k in desc for k in ["RIVET", "BOLT", "SCREW", "PIN", "NUT"]):
            return True
    except:
        pass
    return False


# =====================================================
# AXIS EXTRACTION (GENERIC)
# =====================================================
def extract_axis_from_occurrence(occ):
    """
    Finds ANY usable axis from the part geometry.
    Works even if no InsertConstraint exists.
    """
    try:
        bodies = occ.Definition.SurfaceBodies
    except:
        return None

    for body in bodies:
        for face in body.Faces:
            try:
                geo = face.Geometry
                if geo.SurfaceType == 2:  # Cylinder
                    return geo.Axis
            except:
                continue
    return None


# =====================================================
# HOLE INFERENCE (GEOMETRY TRUTH)
# =====================================================
def infer_holes(asm_def):
    tg = asm_def.Application.TransientGeometry
    occurrences = list(asm_def.Occurrences)
    results = []

    fasteners = [o for o in occurrences if is_fastener(o)]

    for fastener in fasteners:
        axis = extract_axis_from_occurrence(fastener)
        if not axis:
            continue

        stack = []

        for occ in occurrences:
            if occ == fastener or occ.Suppressed:
                continue

            try:
                bodies = occ.Definition.SurfaceBodies
            except:
                continue

            for body in bodies:
                try:
                    hits = tg.CreateObjectCollection()
                    body.IntersectWithCurve(axis, hits)
                    if hits.Count > 0:
                        stack.append(occ.Name)
                        break
                except:
                    continue

        if len(stack) >= 2:
            results.append({
                "fastener": fastener.Name,
                "fastener_part": fastener.Definition.Document.DisplayName,
                "hole_stack": stack,
                "confidence": round(0.7 + 0.05 * len(stack), 2)
            })

    return results


# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()

    inv = get_inventor()
    doc = open_assembly(inv)

    if doc.DocumentType != kAssemblyDocument:
        raise RuntimeError("Not an assembly")

    asm_def = doc.ComponentDefinition
    holes = infer_holes(asm_def)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(holes, f, indent=4)

    print(f"✅ Inferred hole stacks: {len(holes)}")
    print(f"📁 Output: {OUTPUT_JSON}")


if __name__ == "__main__":
    run()