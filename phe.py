import win32com.client
import pythoncom
import json
import time
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = Path(r"E:\Phase 1\extractions\hole_features.json")

# =====================================================
# INVENTOR CONNECT
# =====================================================
def connect_inventor():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
        return inv

# =====================================================
# HOLE FEATURE EXTRACTION (CORRECT)
# =====================================================
def extract_holes_from_part(part_doc):
    holes = []

    # 🔥 CRITICAL: force Inventor to load features
    part_doc.Activate()
    part_doc.Update()

    comp_def = part_doc.ComponentDefinition
    hole_feats = comp_def.Features.HoleFeatures

    for h in hole_feats:
        try:
            hole_data = {
                "name": h.Name,
                "diameter_mm": float(h.Diameter.Value),
                "suppressed": bool(h.Suppressed),
                "patterned": False,
                "pattern_count": 1
            }

            # -------------------------------
            # Detect rectangular patterns
            # -------------------------------
            for pat in comp_def.Features.RectangularPatternFeatures:
                if h in list(pat.ParentFeatures):
                    hole_data["patterned"] = True
                    hole_data["pattern_count"] = pat.CountX * pat.CountY

            # -------------------------------
            # Detect circular patterns
            # -------------------------------
            for pat in comp_def.Features.CircularPatternFeatures:
                if h in list(pat.ParentFeatures):
                    hole_data["patterned"] = True
                    hole_data["pattern_count"] = pat.Count

            holes.append(hole_data)

        except:
            continue

    return holes

# =====================================================
# MAIN
# =====================================================
def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()
    asm = inv.Documents.Open(ASSEMBLY_PATH, True)

    seen_parts = set()
    output = []

    for occ in asm.ComponentDefinition.Occurrences:
        try:
            part_doc = occ.Definition.Document
            name = part_doc.DisplayName.lower()

            if not name.endswith(".ipt"):
                continue

            if name in seen_parts:
                continue
            seen_parts.add(name)

            holes = extract_holes_from_part(part_doc)

            output.append({
                "part": part_doc.DisplayName,
                "holes": holes
            })

        except:
            continue

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4)

    print("✅ HoleFeature extraction complete")
    print(f"   → {OUTPUT_JSON}")

    asm.Close(True)

# =====================================================
if __name__ == "__main__":
    run()