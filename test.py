import win32com.client
import pythoncom
import json
import time
from pathlib import Path

# ================= CONFIG =================
PART_PATH = Path(r"E:\Phase 1\Assembly 1\1093144795-A.ipt")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\_probe_everything.json")

MAX_ATTRS = 300  # safety limit

# ================= INVENTOR =================
def connect_inventor():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
        return inv

def safe_get(obj, attr):
    try:
        v = getattr(obj, attr)
        if callable(v):
            return "<callable>"
        return str(v)
    except:
        return None

def list_attrs(obj):
    out = {}
    for a in dir(obj):
        if a.startswith("_"):
            continue
        try:
            val = getattr(obj, a)
            if callable(val):
                out[a] = "<callable>"
            else:
                out[a] = str(val)
        except:
            out[a] = None
        if len(out) > MAX_ATTRS:
            break
    return out

# ================= MAIN PROBE =================
def run():
    pythoncom.CoInitialize()
    inv = connect_inventor()

    doc = inv.Documents.Open(str(PART_PATH), True)
    comp = doc.ComponentDefinition

    dump = {
        "part": PART_PATH.name,
        "property_sets": {},
        "features": [],
        "faces": [],
        "sketches": []
    }

    # ---------- PropertySets ----------
    for ps in doc.PropertySets:
        props = {}
        for p in ps:
            try:
                props[p.Name] = str(p.Value)
            except:
                props[p.Name] = None
        dump["property_sets"][ps.Name] = props

    # ---------- Features ----------
    for f in comp.Features:
        f_entry = {
            "name": safe_get(f, "Name"),
            "type": safe_get(f, "Type"),
            "all_attributes": list_attrs(f)
        }
        dump["features"].append(f_entry)

    # ---------- Faces ----------
    try:
        for body in comp.SurfaceBodies:
            for face in body.Faces:
                dump["faces"].append({
                    "surface_type": safe_get(face, "SurfaceType"),
                    "geometry": list_attrs(face.Geometry)
                })
    except:
        pass

    # ---------- Sketches ----------
    try:
        for sk in comp.Sketches:
            dump["sketches"].append({
                "name": sk.Name,
                "entities": list_attrs(sk)
            })
    except:
        pass

    doc.Close(True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(dump, f, indent=2)

    print("✅ FULL PROBE COMPLETE")
    print(f"→ {OUTPUT_JSON}")

if __name__ == "__main__":
    run()