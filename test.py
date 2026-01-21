import win32com.client
import pythoncom
import json
import time
from pathlib import Path

# ================= CONFIG =================
PART_PATH = Path(r"E:\Phase 1\Assembly 1\1093144795-A.ipt")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\_probe_holes.json")

# ================= INVENTOR =================
def connect_inventor():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
        return inv

def safe(val):
    try:
        return val
    except:
        return None

def safe_float(val):
    try:
        return float(val)
    except:
        return None

# ================= PROBE ====================
def probe_part(part_path):
    pythoncom.CoInitialize()
    inv = connect_inventor()

    doc = inv.Documents.Open(str(part_path), True)
    comp = doc.ComponentDefinition

    probe = {
        "part": part_path.name,
        "hole_features": []
    }

    for h in comp.Features.HoleFeatures:
        entry = {
            "name": h.Name,
            "suppressed": safe(h.Suppressed),
            "feature_type": safe(h.Type),
            "properties": {},
            "definition": {},
            "faces": []
        }

        # -------- Feature-level properties --------
        for attr in [
            "Diameter", "Tapped", "Threaded", "CounterboreDiameter",
            "CountersinkAngle", "Extent"
        ]:
            try:
                entry["properties"][attr] = safe_float(getattr(h, attr).Value)
            except:
                entry["properties"][attr] = None

        # -------- HoleDefinition inspection --------
        try:
            hd = h.HoleDefinition
            entry["definition"]["Type"] = str(safe(hd.Type))

            for attr in dir(hd):
                if "Diameter" in attr or "Thread" in attr:
                    try:
                        v = getattr(hd, attr)
                        entry["definition"][attr] = safe_float(v.Value) if hasattr(v, "Value") else str(v)
                    except:
                        pass
        except:
            entry["definition"] = None

        # -------- Face-based geometry --------
        try:
            for face in h.Faces:
                if face.SurfaceType == 16:  # Cylinder
                    cyl = face.Geometry
                    axis = cyl.Axis
                    entry["faces"].append({
                        "axis_dir": [
                            axis.Direction.X,
                            axis.Direction.Y,
                            axis.Direction.Z
                        ],
                        "axis_origin": [
                            axis.RootPoint.X,
                            axis.RootPoint.Y,
                            axis.RootPoint.Z
                        ],
                        "radius_mm": safe_float(cyl.Radius)
                    })
        except:
            pass

        probe["hole_features"].append(entry)

    doc.Close(True)
    return probe

# ================= MAIN =====================
if __name__ == "__main__":
    data = probe_part(PART_PATH)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    print("✅ Probe complete")
    print(f"→ {OUTPUT_JSON}")