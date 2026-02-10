import os
import json
import pythoncom
import win32com.client


# ------------------------------------------------------------
# Safe JSON loader (no path assumptions, no mutation)
# ------------------------------------------------------------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------------
# Assembly builder
# ------------------------------------------------------------
def build_assembly(json_path, output_iam_path):
    pythoncom.CoInitialize()

    data = load_json(json_path)

    components = data["components"]

    base_dir = os.path.dirname(json_path)

    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = True

    # Create new assembly
    asm_doc = inventor.Documents.Add(12291)

    asm_def = asm_doc.ComponentDefinition
    tg = inventor.TransientGeometry

    print(f"Creating assembly with {len(components)} components...\n")

    for comp in components:
        part_name = comp["file_name"]
        part_path = os.path.join(base_dir, part_name)

        if not os.path.exists(part_path):
            print(f"❌ Missing IPT: {part_path}")
            continue

        # Create transform matrix
        m = tg.CreateMatrix()

        r = comp["transform"]["rotation_matrix"]
        t = comp["transform"]["translation_mm"]

        # Rotation
        m.SetCell(1, 1, r[0][0])
        m.SetCell(1, 2, r[0][1])
        m.SetCell(1, 3, r[0][2])

        m.SetCell(2, 1, r[1][0])
        m.SetCell(2, 2, r[1][1])
        m.SetCell(2, 3, r[1][2])

        m.SetCell(3, 1, r[2][0])
        m.SetCell(3, 2, r[2][1])
        m.SetCell(3, 3, r[2][2])

# Translation (Inventor uses cm internally)
        m.SetCell(1, 4, t["x"] / 10.0)
        m.SetCell(2, 4, t["y"] / 10.0)
        m.SetCell(3, 4, t["z"] / 10.0)


        occ = asm_def.Occurrences.Add(part_path, m)

        occ.Grounded = bool(comp.get("grounded", False))
        if comp.get("suppressed", False):
         occ.Suppress()
        else:
         occ.Unsuppress()


        print(f"✅ Added: {comp['occurrence_name']}")

    # Save assembly
    asm_doc.SaveAs(output_iam_path, False)

    print("\n🎉 Assembly automation completed")
    print(f"📁 Saved as: {output_iam_path}")


# ------------------------------------------------------------
# Entry point
# ------------------------------------------------------------
if __name__ == "__main__":
    json_path = r"G:/Shubhangi college/Assembly 1 new/1093144795-M1.json"
    output_iam_path = r"G:/Shubhangi college/Assembly 1 new/Generated_Assembly.iam"

    build_assembly(json_path, output_iam_path)