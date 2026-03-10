import os
import json
import hashlib

# ==========================================
# 1. HASHING HELPERS (IDENTICAL TO OLD LOGIC)
# ==========================================
def r(x):
    if x is None:
        return 0
    return round(float(x), 6)

def vec(v):
    if v is None:
        return [0, 0, 0]
    return [r(x) for x in v]

def hash_string(s):
    return hashlib.sha256(s.encode()).hexdigest()[:16]

def face_hash(face):
    surface = face.get("surface_type")
    area = r(face.get("area_cm2"))
    
    center = face.get("center") or face.get("face_point_at_center")
    center = vec(center)
    
    normal = face.get("normal") or face.get("face_normal_at_center")
    normal = vec(normal)
    
    radius = r(face.get("radius_cm"))
    
    sig = f"{surface}|{area}|{center}|{normal}|{radius}"
    return "face_" + hash_string(sig)

def edge_hash(edge):
    curve = edge.get("curve_type")
    length = r(edge.get("length_cm"))
    
    mid = edge.get("midpoint") or edge.get("edge_midpoint")
    mid = vec(mid)
    
    tan = edge.get("tangent") or edge.get("edge_tangent_at_mid")
    tan = vec(tan)
    
    sig = f"{curve}|{length}|{mid}|{tan}"
    return "edge_" + hash_string(sig)


# ==========================================
# 2. FILE PROCESSING LOGIC
# ==========================================
def process_part_json(path):
    with open(path, 'r') as f:
        data = json.load(f)

    changed = False
    ref_to_hash_map = {} # Dictionary to translate Inventor keys to Hashes

    # Process Faces
    if "faces" in data:
        for f in data["faces"]:
            # Generate new hash
            new_hash = face_hash(f)
            f["geometry_hash"] = new_hash
            
            # Map Inventor's reference key to the new hash (if it exists)
            ref_key = f.get("reference_key_string")
            if ref_key:
                ref_to_hash_map[ref_key] = new_hash
                
        changed = True

    # Process Edges and B-Rep Topology
    if "edges" in data:
        for e in data["edges"]:
            e["geometry_hash"] = edge_hash(e)
            
            # Swap Inventor reference keys in adjacent_faces with our custom hashes
            if "adjacent_faces" in e:
                updated_adjacent_faces = []
                for inv_key in e["adjacent_faces"]:
                    # If we found a translation for this face, swap it!
                    if inv_key in ref_to_hash_map:
                        updated_adjacent_faces.append(ref_to_hash_map[inv_key])
                    else:
                        # Fallback just in case (should rarely happen)
                        updated_adjacent_faces.append(inv_key)
                        
                e["adjacent_faces"] = updated_adjacent_faces
                
        changed = True

    if changed:
        with open(path, "w") as f:
            json.dump(data, f, indent=2)
        print(f"Updated Part: {os.path.basename(path)}")


def process_assembly_json(path):
    with open(path, 'r') as f:
        data = json.load(f)

    changed = False

    if "constraints" in data:
        for c in data["constraints"]:
            for key in ["entity_one", "entity_two"]:
                if key not in c or c[key] is None:
                    continue

                ent = c[key]
                if ent.get("entity_type") == "FaceProxy":
                    ent["geometry_hash"] = face_hash(ent)
                    changed = True
                elif ent.get("entity_type") == "EdgeProxy":
                    ent["geometry_hash"] = edge_hash(ent)
                    changed = True

    if changed:
        with open(path, "w") as f:
            json.dump(data, f, indent=2)
        print(f"Updated Assembly: {os.path.basename(path)}")


# ==========================================
# 3. FOLDER EXECUTION
# ==========================================
if __name__ == "__main__":
    # CHANGE THESE PATHS TO MATCH YOUR MAC FOLDERS
    PARTS_DIR = "/Users/aryantembhurne/Desktop/Phase 1/Phase 2/dataset/Parts combined"
    ASSEMBLIES_DIR = "/Users/aryantembhurne/Desktop/Phase 1/Phase 2/dataset/Assemblies combined"

    print("--- Processing Part Files ---")
    if os.path.exists(PARTS_DIR):
        for root, dirs, files in os.walk(PARTS_DIR):
            for f in files:
                if f.endswith(".json"):
                    process_part_json(os.path.join(root, f))
    else:
        print(f"Directory not found: {PARTS_DIR}")

    print("\n--- Processing Assembly Files ---")
    if os.path.exists(ASSEMBLIES_DIR):
        for root, dirs, files in os.walk(ASSEMBLIES_DIR):
            for f in files:
                if f.endswith(".json"):
                    process_assembly_json(os.path.join(root, f))
    else:
        print(f"Directory not found: {ASSEMBLIES_DIR}")

    print("\nDone! All files successfully hashed.")