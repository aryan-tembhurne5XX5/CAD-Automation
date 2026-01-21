import win32com.client
import pythoncom
import json
from pathlib import Path
import time

# ===============================
# CONFIG
# ===============================
PART_PATH = r"E:\Phase 1\Assembly 1\1093144795-A.ipt"
OUT_JSON  = Path(r"E:\Phase 1\extractions\part_full_dump.json")

# ===============================
# SAFE HELPERS
# ===============================
def safe(val):
    try:
        return float(val)
    except:
        return None

def try_get(obj, attr):
    try:
        return getattr(obj, attr)
    except:
        return None

# ===============================
# INVENTOR CONNECT
# ===============================
def connect_inventor():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
        return inv

# ===============================
# MAIN DUMP
# ===============================
def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()

    doc = inv.Documents.Open(PART_PATH, True)
    comp = doc.ComponentDefinition

    dump = {
        "part": doc.DisplayName,
        "document_type": str(doc.DocumentType),
        "surface_bodies": []
    }

    for b_idx, body in enumerate(comp.SurfaceBodies):
        body_data = {
            "body_index": b_idx,
            "faces": []
        }

        for f_idx, face in enumerate(body.Faces):
            face_data = {
                "face_index": f_idx,
                "surface_type": try_get(face.Geometry, "SurfaceType"),
                "is_param_reversed": try_get(face, "IsParamReversed"),
                "geometry_class": face.Geometry.__class__.__name__,
                "geometry": {}
            }

            geom = face.Geometry

            # Try extracting everything safely
            for attr in [
                "Radius",
                "Diameter",
                "Axis",
                "BasePoint",
                "Center",
                "Normal",
                "Direction"
            ]:
                val = try_get(geom, attr)

                if val is None:
                    face_data["geometry"][attr] = None
                else:
                    try:
                        face_data["geometry"][attr] = {
                            "X": safe(getattr(val, "X", None)),
                            "Y": safe(getattr(val, "Y", None)),
                            "Z": safe(getattr(val, "Z", None))
                        }
                    except:
                        face_data["geometry"][attr] = str(val)

            # Edge preview (count only)
            try:
                face_data["edge_count"] = face.Edges.Count
            except:
                face_data["edge_count"] = None

            body_data["faces"].append(face_data)

        dump["surface_bodies"].append(body_data)

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(dump, f, indent=4)

    print("✅ FULL PART GEOMETRY DUMP COMPLETE")
    print(f"   → {OUT_JSON}")

    doc.Close(True)

if __name__ == "__main__":
    run()