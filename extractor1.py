import win32com.client
import os
import json
import time

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_PATH   = r"E:\Phase 1\extractions\assembly_extraction2.json"

# =====================================================
# ENUM MAPS (SAFE)
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
# SAFE HELPERS
# =====================================================
def get_document_type(doc):
    """
    DO NOT trust DocumentType alone.
    Python + Inventor COM is unreliable here.
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
    try:
        if not entity:
            return "None"
        return ENTITY_TYPE_MAP.get(entity.Type, f"Unknown ({entity.Type})")
    except:
        return "Unknown"


def point_mm(p):
    # Inventor internal units = cm
    return [
        round(p.X * 10, 4),
        round(p.Y * 10, 4),
        round(p.Z * 10, 4)
    ]


def vec(v):
    return [
        round(v.X, 6),
        round(v.Y, 6),
        round(v.Z, 6)
    ]

# =====================================================
# HOLE EXTRACTION (FIXED FOR PYTHON)
# =====================================================
def extract_holes_from_part(part_doc):
    """
    ✔ Uses hole.Definition (NOT HoleDefinition)
    ✔ Uses Geometry3d (true 3D)
    ✔ Safe on all hole subtypes
    """
    holes = []

    try:
        cd = part_doc.ComponentDefinition
        hole_features = cd.Features.HoleFeatures
    except:
        return holes

    for hole in hole_features:
        if hole.Suppressed:
            continue

        try:
            hdef = hole.Definition                # ✅ CRITICAL FIX
            pdef = hole.PlacementDefinition
        except:
            continue

        # ---- Diameter (robust) ----
        diameter_mm = None
        try:
            if hasattr(hdef, "Diameter"):
                diameter_mm = hdef.Diameter.Value * 10
            elif hasattr(hdef, "TapInfo"):
                diameter_mm = hdef.TapInfo.MajorDiameter * 10
        except:
            diameter_mm = None

        # ---- Sketch-based holes (99%) ----
        if pdef.Type == 0:  # kSketchPlacementDefinition
            try:
                sketch_plane = pdef.Sketch.PlanarEntityGeometry
                axis = vec(sketch_plane.Normal.AsVector())
            except:
                axis = None

            for pt in pdef.SketchPoints:
                try:
                    center = pt.Geometry3d      # ✅ TRUE 3D CENTER
                except:
                    continue

                holes.append({
                    "feature_name": hole.Name,
                    "diameter_mm": round(diameter_mm, 4) if diameter_mm else None,
                    "center_mm": point_mm(center),
                    "axis_vector": axis,
                    "threaded": bool(hdef.Tapped) if hasattr(hdef, "Tapped") else False
                })

    return holes

# =====================================================
# RECURSIVE OCCURRENCE EXTRACTION
# =====================================================
def extract_occurrences(occurrences, parent=None):
    result = []

    for occ in occurrences:
        try:
            is_suppressed = bool(occ.Suppressed)

            occ_data = {
                "name": occ.Name,
                "parent": parent,
                "suppressed": is_suppressed,
                "visible": bool(occ.Visible),
                "document_type": "Unknown",
                "definition": None,
                "part_number": None,
                "description": None,
                "holes": []
            }

            if is_suppressed:
                occ_data["note"] = "Suppressed"
                result.append(occ_data)
                continue

            part_doc = occ.Definition.Document
            occ_data["document_type"] = get_document_type(part_doc)
            occ_data["definition"] = part_doc.DisplayName

            # iProperties
            try:
                props = part_doc.PropertySets.Item("Design Tracking Properties")
                occ_data["part_number"] = props.Item("Part Number").Value
                occ_data["description"] = props.Item("Description").Value
            except:
                pass

            # Holes (Parts only)
            if occ_data["document_type"] == "Part":
                occ_data["holes"] = extract_holes_from_part(part_doc)

            result.append(occ_data)

            # Recurse into subassemblies
            try:
                if occ.SubOccurrences.Count > 0:
                    result.extend(
                        extract_occurrences(occ.SubOccurrences, occ.Name)
                    )
            except:
                pass

        except Exception as e:
            print(f"⚠️ Occurrence error [{occ.Name}]: {e}")

    return result

# =====================================================
# MAIN
# =====================================================
def main():
    if not os.path.isfile(ASSEMBLY_PATH):
        print(f"❌ Assembly not found: {ASSEMBLY_PATH}")
        return

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    # Connect to Inventor
    try:
        inv = win32com.client.GetActiveObject("Inventor.Application")
        created = False
    except:
        inv = win32com.client.DispatchEx("Inventor.Application")
        inv.Visible = True
        created = True
        time.sleep(5)

    doc = None

    try:
        doc = inv.Documents.Open(ASSEMBLY_PATH, True)
        asm_def = doc.ComponentDefinition

        data = {
            "assembly_name": doc.DisplayName,
            "timestamp": time.ctime(),
            "occurrences": [],
            "constraints": []
        }

        print("⚙️ Extracting occurrences...")
        data["occurrences"] = extract_occurrences(asm_def.Occurrences)

        print("🔗 Extracting constraints...")
        for c in asm_def.Constraints:
            try:
                data["constraints"].append({
                    "name": c.Name,
                    "constraint_type": CONSTRAINT_TYPE_MAP.get(c.Type, f"Other ({c.Type})"),
                    "health": c.HealthStatus,
                    "occurrence_1": c.OccurrenceOne.Name if c.OccurrenceOne else None,
                    "occurrence_2": c.OccurrenceTwo.Name if c.OccurrenceTwo else None,
                    "entity_1_type": get_entity_type(c.EntityOne),
                    "entity_2_type": get_entity_type(c.EntityTwo)
                })
            except:
                continue

        with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

        print(f"✅ Extraction complete → {OUTPUT_PATH}")

    finally:
        if doc:
            doc.Close(True)
        if created:
            inv.Quit()

if __name__ == "__main__":
    main()