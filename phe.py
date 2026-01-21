import win32com.client
import pythoncom
import json
import time
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_PATH = r"E:\Phase 1\Assembly 1\1093144795-M1.iam"
OUTPUT_JSON   = Path(r"E:\Phase 1\extractions\part_holes.json")

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
# ENUM HELPERS
# =====================================================
def hole_type_name(hole):
    try:
        if hole.HoleType == 0:
            return "simple"
        if hole.HoleType == 1:
            return "counterbore"
        if hole.HoleType == 2:
            return "countersink"
    except:
        pass
    return "unknown"

def pattern_type_name(p):
    name = p.TypeName.lower()
    if "rectangular" in name:
        return "Rectangular"
    if "circular" in name:
        return "Circular"
    return "Sketch"

# =====================================================
# MAIN EXTRACTION
# =====================================================
def run():
    pythoncom.CoInitialize()

    inv = connect_inventor()
    asm = inv.Documents.Open(ASSEMBLY_PATH, True)

    extracted_parts = {}
    results = []

    for occ in asm.ComponentDefinition.Occurrences:
        try:
            part_doc = occ.Definition.Document
            part_name = part_doc.DisplayName

            if not part_name.lower().endswith(".ipt"):
                continue

            if part_name in extracted_parts:
                continue

            extracted_parts[part_name] = True
            part_def = part_doc.ComponentDefinition

            holes_out = []

            # -------------------------------
            # HOLE FEATURES
            # -------------------------------
            for hole in part_def.Features.HoleFeatures:
                hole_data = {
                    "hole_feature": hole.Name,
                    "diameter_mm": float(hole.Diameter.Value),
                    "hole_type": hole_type_name(hole),
                    "suppressed": bool(hole.Suppressed),
                    "patterned": False,
                    "pattern": None
                }

                # -------------------------------
                # PATTERN DETECTION
                # -------------------------------
                for pattern in part_def.Features.RectangularPatternFeatures:
                    if hole in list(pattern.ParentFeatures):
                        hole_data["patterned"] = True
                        hole_data["pattern"] = {
                            "pattern_name": pattern.Name,
                            "pattern_type": "Rectangular",
                            "instance_count": pattern.CountX * pattern.CountY
                        }

                for pattern in part_def.Features.CircularPatternFeatures:
                    if hole in list(pattern.ParentFeatures):
                        hole_data["patterned"] = True
                        hole_data["pattern"] = {
                            "pattern_name": pattern.Name,
                            "pattern_type": "Circular",
                            "instance_count": pattern.Count
                        }

                holes_out.append(hole_data)

            results.append({
                "part": part_name,
                "holes": holes_out
            })

        except:
            continue

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=4)

    print("✅ Phase-3 complete (correct hole extraction)")
    print(f"   → {OUTPUT_JSON}")

    asm.Close(True)

# =====================================================
# ENTRY
# =====================================================
if __name__ == "__main__":
    run()