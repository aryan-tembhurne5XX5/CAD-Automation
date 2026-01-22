import win32com.client
import json
import os

IPT_PATH = r"C:\InventorModels\Part.ipt"

inv = win32com.client.Dispatch("Inventor.Application")
inv.Visible = False
doc = inv.Documents.Open(IPT_PATH, True)

if doc.DocumentType != 12290:
    raise Exception("Not a Part document")

cd = doc.ComponentDefinition

def safe(val):
    try:
        return str(val)
    except:
        return None

dump = {
    "document": {},
    "parameters": [],
    "work_features": {},
    "features": {},
    "sketches": [],
    "bodies": []
}

# ---------------- DOCUMENT ----------------
dump["document"] = {
    "display_name": doc.DisplayName,
    "full_path": doc.FullFileName,
    "units": safe(doc.UnitsOfMeasure.LengthUnits),
    "model_states": [s.Name for s in doc.ModelStates]
}

# ---------------- PARAMETERS ----------------
for p in doc.Parameters:
    dump["parameters"].append({
        "name": p.Name,
        "value": safe(p.Value),
        "units": safe(p.Units),
        "expression": safe(p.Expression),
        "exposed": p.ExposedAsProperty
    })

# ---------------- WORK FEATURES ----------------
dump["work_features"]["planes"] = [wp.Name for wp in cd.WorkPlanes]
dump["work_features"]["axes"] = [wa.Name for wa in cd.WorkAxes]
dump["work_features"]["points"] = [wp.Name for wp in cd.WorkPoints]

# ---------------- FEATURES (BLIND) ----------------
for collection_name in dir(cd.Features):
    if collection_name.endswith("Features"):
        try:
            col = getattr(cd.Features, collection_name)
            feats = []
            for f in col:
                feats.append({
                    "name": f.Name,
                    "type": f.Type,
                    "suppressed": getattr(f, "Suppressed", None)
                })
            dump["features"][collection_name] = feats
        except:
            pass

# ---------------- SKETCHES ----------------
for sk in cd.Sketches:
    s = {
        "name": sk.Name,
        "plane": safe(sk.PlanarEntity),
        "points": [],
        "entities": [],
        "dimensions": []
    }

    for p in sk.SketchPoints:
        s["points"].append({
            "x": p.Geometry.X,
            "y": p.Geometry.Y,
            "z": p.Geometry.Z
        })

    for e in sk.SketchEntities:
        s["entities"].append({
            "type": e.Type
        })

    for d in sk.DimensionConstraints:
        s["dimensions"].append({
            "value": safe(d.Parameter.Value),
            "type": d.Type
        })

    dump["sketches"].append(s)

# ---------------- BODIES ----------------
for body in cd.SurfaceBodies:
    b = {
        "faces": body.Faces.Count,
        "edges": body.Edges.Count,
        "vertices": body.Vertices.Count,
        "volume": safe(body.Volume)
    }
    dump["bodies"].append(b)

# ---------------- WRITE ----------------
out = os.path.join(os.path.dirname(IPT_PATH), "RAW_IPT_DUMP.json")
with open(out, "w") as f:
    json.dump(dump, f, indent=2)

print("✅ Raw IPT dump complete:", out)