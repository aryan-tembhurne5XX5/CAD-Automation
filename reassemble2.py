import os
import json
import math
import pythoncom
import win32com.client


# ------------------------------------------------------------
# Load JSON
# ------------------------------------------------------------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------------
# Find occurrence by name
# ------------------------------------------------------------
def find_occurrence(asm_def, name):
    for occ in asm_def.Occurrences:
        if occ.Name == name:
            return occ
    return None


# ------------------------------------------------------------
# Bind ReferenceKey → actual geometry
# ------------------------------------------------------------
def bind_refkey(asm_doc, refkey_string):
    ref_mgr = asm_doc.ReferenceKeyManager
    key_bytes = ref_mgr.StringToKey(refkey_string)
    return ref_mgr.BindKeyToObject(key_bytes)


# ------------------------------------------------------------
# Build Exact Assembly
# ------------------------------------------------------------
def build_exact_assembly(json_path, output_path):

    pythoncom.CoInitialize()

    data = load_json(json_path)
    components = data["components"]
    constraints = data.get("constraints", [])

    base_dir = os.path.dirname(json_path)

    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = True

    tg = inventor.TransientGeometry

    # Create new assembly
    asm_doc = inventor.Documents.Add(12291)  # Assembly doc
    asm_def = asm_doc.ComponentDefinition

    print(f"\nCreating assembly with {len(components)} components...\n")

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

        print(f"✅ Added: {occ.Name}")

    # ------------------------------------------------------------
    # APPLY CONSTRAINTS USING REFERENCEKEYS
    # ------------------------------------------------------------
    print(f"\nApplying {len(constraints)} constraints...\n")

    for c in constraints:

        try:
            ctype = c["constraint_type"]

            entity1 = bind_refkey(asm_doc, c["entity_one_refkey"])
            entity2 = bind_refkey(asm_doc, c["entity_two_refkey"])

            params = c.get("parameters", {})

            offset_cm = (params.get("offset_mm") or 0) / 10.0
            angle_rad = (params.get("angle_deg") or 0) * math.pi / 180.0

            # -----------------------------
            # Apply correct constraint
            # -----------------------------
            if ctype == "kMateConstraintObject":
                asm_def.Constraints.AddMateConstraint(
                    entity1, entity2, offset_cm
                )

            elif ctype == "kFlushConstraintObject":
                asm_def.Constraints.AddFlushConstraint(
                    entity1, entity2, offset_cm
                )

            elif ctype == "kAngleConstraintObject":
                asm_def.Constraints.AddAngleConstraint(
                    entity1, entity2, angle_rad
                )

            elif ctype == "kInsertConstraintObject":
                asm_def.Constraints.AddInsertConstraint(
                    entity1, entity2, offset_cm
                )

            elif ctype == "kTangentConstraintObject":
                asm_def.Constraints.AddTangentConstraint(
                    entity1, entity2
                )

            else:
                print(f"⚠️ Unsupported constraint type: {ctype}")
                continue

            print(f"🔗 Applied {ctype}: {c['constraint_id']}")

        except Exception as e:
            print(f"❌ Failed {c['constraint_id']}: {e}")

    # ------------------------------------------------------------
    # SAVE
    # ------------------------------------------------------------
    asm_doc.SaveAs(output_path, False)

    print("\n🎉 EXACT Assembly Reconstruction Complete")
    print(f"📁 Saved at: {output_path}")


# ------------------------------------------------------------
# Entry Point
# ------------------------------------------------------------
if __name__ == "__main__":

    json_path = r"G:/Shubhangi college/Assembly 1 new/1093144795-M1.json"
    output_path = r"G:/Shubhangi college/Assembly 1 new/Exact_Reconstructed_Assembly.iam"

    build_exact_assembly(json_path, output_path)