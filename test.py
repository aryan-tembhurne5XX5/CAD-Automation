import win32com.client
import pythoncom
import json
from pathlib import Path
import time

# ================================
# CONFIG
# ================================
PART_PATH = r"E:\Phase 1\Assembly 1\1093144795-A.ipt"
OUT_JSON  = Path(r"E:\Phase 1\extractions\part_debug_dump.json")

# ================================
# CONNECT INVENTOR
# ================================
def connect():
    try:
        inv = win32com.client.GetActiveObject("Inventor.Application")
    except:
        inv = win32com.client.Dispatch("Inventor.Application")
        inv.Visible = True
        time.sleep(2)
    return inv

# ================================
# MAIN
# ================================
def run():
    pythoncom.CoInitialize()
    inv = connect()

    doc = inv.Documents.Open(PART_PATH, True)
    comp = doc.ComponentDefinition

    dump = {
        "part": Path(PART_PATH).name,
        "hole_features": [],
        "sketches": [],
        "work_axes": [],
        "work_points": [],
        "cylindrical_faces": []
    }

    # ----------------------------
    # HOLE FEATURES (RAW)
    # ----------------------------
    for h in comp.Features.HoleFeatures:
        hole_data = {
            "name": h.Name,
            "hole_type_enum": int(h.HoleType),
            "suppressed": bool(h.Suppressed),
            "placement_definition_attrs": []
        }

        pd = h.PlacementDefinition
        for attr in dir(pd):
            if not attr.startswith("_"):
                try:
                    val = getattr(pd, attr)
                    hole_data["placement_definition_attrs"].append(attr)
                except:
                    pass

        dump["hole_features"].append(hole_data)

    # ----------------------------
    # SKETCH GEOMETRY
    # ----------------------------
    for sk in comp.Sketches:
        sk_data = {
            "name": sk.Name,
            "points": [],
            "circles": []
        }

        for pt in sk.SketchPoints:
            sk_data["points"].append([pt.Geometry.X, pt.Geometry.Y])

        for c in sk.SketchCircles:
            center = c.CenterSketchPoint.Geometry
            sk_data["circles"].append({
                "center": [center.X, center.Y],
                "radius": c.Radius
            })

        dump["sketches"].append(sk_data)

    # ----------------------------
    # WORK FEATURES
    # ----------------------------
    for ax in comp.WorkAxes:
        try:
            geo = ax.Line
            dump["work_axes"].append({
                "name": ax.Name,
                "origin": [geo.RootPoint.X, geo.RootPoint.Y, geo.RootPoint.Z],
                "direction": [geo.Direction.X, geo.Direction.Y, geo.Direction.Z]
            })
        except:
            pass

    for wp in comp.WorkPoints:
        try:
            p = wp.Point
            dump["work_points"].append({
                "name": wp.Name,
                "point": [p.X, p.Y, p.Z]
            })
        except:
            pass

    # ----------------------------
    # CYLINDRICAL FACES (IMPORTANT)
    # ----------------------------
    for body in comp.SurfaceBodies:
        for face in body.Faces:
            try:
                if face.SurfaceType == 16:  # Cylinder
                    cyl = face.Geometry
                    dump["cylindrical_faces"].append({
                        "radius": cyl.Radius,
                        "axis_origin": [
                            cyl.Axis.RootPoint.X,
                            cyl.Axis.RootPoint.Y,
                            cyl.Axis.RootPoint.Z
                        ],
                        "axis_direction": [
                            cyl.Axis.Direction.X,
                            cyl.Axis.Direction.Y,
                            cyl.Axis.Direction.Z
                        ]
                    })
            except:
                pass

    # ----------------------------
    # SAVE
    # ----------------------------
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(dump, f, indent=4)

    doc.Close(True)
    print(f"✅ Debug dump created → {OUT_JSON}")

if __name__ == "__main__":
    run()