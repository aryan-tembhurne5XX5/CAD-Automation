import os
import json
import pythoncom
import win32com.client


# ------------------------------------------------------------
# Safe JSON loader
# ------------------------------------------------------------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------------
# Find occurrence by name
# ------------------------------------------------------------
def find_occurrence(asm_def, occ_name):
    for occ in asm_def.Occurrences:
        if occ.Name == occ_name:
            return occ
    return None


# ------------------------------------------------------------
# Assembly builder (components + constraints)
# ------------------------------------------------------------
def build_assembly(json_path, output_iam_path):
    pythoncom.CoInitialize()

    data = load_json(json_path)

    components = data["components"]
    constraints = data.get("constraints", [])

    base_dir = os.path.dirname(json_path)

    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = True

    tg = inventor.TransientGeometry

    # Create new assembly
    asm_doc = inventor.Documents.Add(12291)
    asm_def = asm_doc.ComponentDefinition

    print(f"Creating assembly with {len(components)} components...\n")

    # ------------------------------------------------------------
    # ADD COMPONENTS
    # ------------------------------------------------------------
    for comp in components:
        part_path = os.path.join(base_dir, comp["file_name"])

        if not os.path.exists(part_path):
            print(f"❌ Missing IPT: {part_path}")
            continue

        m = tg.CreateMatrix()

        r = comp["transform"]["rotation_matrix"]
        t = comp["transform"]["translation_mm"]

        # Rotation
        for i in range(3):
            for j in range(3):
                m.SetCell(i + 1, j + 1, r[i][j])

        # Translation (mm → cm)
        m.SetCell(1, 4, t["x"] / 10.0)
        m.SetCell(2, 4, t["y"] / 10.0)
        m.SetCell(3, 4, t["z"] / 10.0)

        occ = asm_def.Occurrences.Add(part_path, m)

        occ.Grounded = bool(comp.get("grounded", False))

        if comp.get("suppressed", False):
            occ.Suppress()
        else:
            occ.Unsuppress()

        print(f"✅ Added: {occ.Name}")

    # ------------------------------------------------------------
    # APPLY CONSTRAINTS (INTENT-LEVEL, SAFE)
    # ------------------------------------------------------------
    print(f"\nApplying {len(constraints)} constraints...\n")

    for c in constraints:
        occ1 = find_occurrence(asm_def, c["component_1"]["occurrence"])
        occ2 = find_occurrence(asm_def, c["component_2"]["occurrence"])

        if not occ1 or not occ2:
            print(f"⚠️ Skipped constraint {c['constraint_id']} (missing occurrence)")
            continue

        ctype = c["constraint_type"]
        params = c.get("parameters", {})

        try:
            # Mate / Flush using origin planes (robust default)
            if ctype in ("kMateConstraintObject", "kFlushConstraintObject"):
                p1 = occ1.Definition.WorkPlanes.Item(3)  # XY plane
                p2 = occ2.Definition.WorkPlanes.Item(3)

                offset_cm = (params.get("offset_mm") or 0) / 10.0

                asm_def.Constraints.AddMateConstraint(
                    p1, p2, offset_cm
                )

            # Angle constraint (Z axes)
            elif ctype == "kAngleConstraintObject":
                a1 = occ1.Definition.WorkAxes.Item(3)
                a2 = occ2.Definition.WorkAxes.Item(3)

                angle_rad = (params.get("angle_deg") or 0) * 3.141592653589793 / 180.0

                asm_def.Constraints.AddAngleConstraint(
                    a1, a2, angle_rad
                )

            # Insert (axis + plane heuristic)
            elif ctype == "kInsertConstraintObject":
                axis1 = occ1.Definition.WorkAxes.Item(3)
                axis2 = occ2.Definition.WorkAxes.Item(3)
                plane1 = occ1.Definition.WorkPlanes.Item(3)
                plane2 = occ2.Definition.WorkPlanes.Item(3)

                asm_def.Constraints.AddInsertConstraint(
                    axis1, axis2, plane1, plane2, 0
                )

            else:
                print(f"⚠️ Unsupported constraint type: {ctype}")
                continue

            print(f"🔗 Applied constraint: {c['constraint_id']}")

        except Exception as e:
            print(f"❌ Failed constraint {c['constraint_id']}: {e}")

    # ------------------------------------------------------------
    # SAVE
    # ------------------------------------------------------------
    asm_doc.SaveAs(output_iam_path, False)

    print("\n🎉 Assembly + constraints generated")
    print(f"📁 Saved as: {output_iam_path}")


# ------------------------------------------------------------
# Entry point
# ------------------------------------------------------------
if __name__ == "__main__":
    json_path = r"G:/Shubhangi college/Assembly 1 new/1093144795-M1.json"
    output_iam_path = r"G:/Shubhangi college/Assembly 1 new/Generated_Assembly.iam"

    build_assembly(json_path, output_iam_path)