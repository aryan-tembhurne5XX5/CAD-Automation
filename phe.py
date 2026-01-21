import win32com.client
import pythoncom
import json
import time
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_EXTRACTION = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
PART_ROOT_FOLDER   = Path(r"E:\Phase 1\Assembly 1")
OUTPUT_JSON        = Path(r"E:\Phase 1\extractions\hole_features.json")

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
def extract_holes_from_part(inv, part_path):
    holes = []

    part_doc = inv.Documents.Open(str(part_path), True)
    comp_def = part_doc.ComponentDefinition

    for h in comp_def.Features.HoleFeatures:
        try:
            hole_info = {
                "name": h.Name,
                "suppressed": bool(h.Suppressed),
                "hole_type": str(h.HoleType),
                "diameter_mm": None,
                "patterned": False,
                "pattern_count": 1
            }

            pd = h.PlacementDefinition

            # ---- Diameter extraction (SAFE) ----
            try:
                if hasattr(pd, "HoleDiameter"):
                    hole_info["diameter_mm"] = float(pd.HoleDiameter.Value)
                elif hasattr(pd, "TapDrillDiameter"):
                    hole_info["diameter_mm"] = float(pd.TapDrillDiameter.Value)
                elif hasattr(pd, "ClearanceDiameter"):
                    hole_info["diameter_mm"] = float(pd.ClearanceDiameter.Value)
            except:
                hole_info["diameter_mm"] = None

            # ---- Pattern detection ----
            for pat in comp_def.Features.RectangularPatternFeatures:
                try:
                    if h in list(pat.ParentFeatures):
                        hole_info["patterned"] = True
                        hole_info["pattern_count"] = pat.CountX * pat.CountY
                except:
                    pass

            for pat in comp_def.Features.CircularPatternFeatures:
                try:
                    if h in list(pat.ParentFeatures):
                        hole_info["patterned"] = True
                        hole_info["pattern_count"] = pat.Count
                except:
                    pass

            holes.append(hole_info)

        except Exception as e:
            print(f"⚠️ Skipped hole in {part_path.name}: {e}")

    part_doc.Close(True)
    return holes

# =====================================================
def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()

    with open(ASSEMBLY_EXTRACTION, "r", encoding="utf-8") as f:
        data = json.load(f)

    unique_parts = sorted(set(o["definition"] for o in data["occurrences"]))

    output = []

    for part_name in unique_parts:
        part_path = PART_ROOT_FOLDER / part_name
        if not part_path.exists():
            continue

        print(f"🔍 Extracting holes from {part_name}")
        holes = extract_holes_from_part(inv, part_path)

        output.append({
            "part": part_name,
            "holes": holes
        })

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4)

    print("\n✅ Phase-3 Hole Extraction COMPLETE")
    print(f"→ {OUTPUT_JSON}")

# =====================================================
if __name__ == "__main__":
    run()