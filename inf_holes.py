import json
from pathlib import Path
from collections import defaultdict

# -------------------------------
# CONFIG
# -------------------------------

INPUT_JSON = Path("assembly_extraction.json")
OUTPUT_JSON = Path("assembly_with_inferred_holes.json")

# Inventor constraint enums (raw)
INSERT_CONSTRAINT = 100665344

# -------------------------------
# HELPERS
# -------------------------------

def safe_float(val):
    try:
        return float(val)
    except:
        return 0.0


def extract_thickness(desc: str) -> float:
    """
    Extracts thickness from strings like:
    'Th=1,25mm', 'Th=2mm', 'Th=1.5mm'
    """
    if not desc:
        return 0.0

    desc = desc.replace(",", ".").lower()
    if "th=" in desc:
        try:
            part = desc.split("th=")[1]
            mm = part.replace("mm", "").strip()
            return float(mm)
        except:
            return 0.0
    return 0.0


def infer_part_type(part):
    desc = (part.get("description") or "").upper()
    name = (part.get("part_number") or "").upper()

    if "RIVET" in desc or "RIVET" in name:
        return "Fastener"
    if "NUT" in desc:
        return "Fastener"
    if extract_thickness(part.get("description", "")) > 0:
        return "Structural"
    return "Unknown"


# -------------------------------
# MAIN PIPELINE
# -------------------------------

def run():
    with open(INPUT_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    occurrences = data["occurrences"]
    constraints = data["constraints"]

    # -------------------------------
    # Index parts
    # -------------------------------
    part_map = {}
    for occ in occurrences:
        occ["inferred_part_type"] = infer_part_type(occ)
        occ["thickness_mm"] = extract_thickness(occ.get("description", ""))
        part_map[occ["name"]] = occ

    # -------------------------------
    # Extract rivet insert constraints
    # -------------------------------
    rivet_to_parts = defaultdict(set)

    for c in constraints:
        if c["constraint_type"] != INSERT_CONSTRAINT:
            continue

        a = c["occurrence_1"]
        b = c["occurrence_2"]

        if a not in part_map or b not in part_map:
            continue

        part_a = part_map[a]
        part_b = part_map[b]

        # Identify rivet
        if part_a["inferred_part_type"] == "Fastener":
            rivet = a
            target = b
        elif part_b["inferred_part_type"] == "Fastener":
            rivet = b
            target = a
        else:
            continue

        rivet_to_parts[rivet].add(target)

    # -------------------------------
    # Infer hole stacks
    # -------------------------------
    inferred_holes = []

    for rivet, connected_parts in rivet_to_parts.items():
        stack = []
        total_thickness = 0.0

        for p in connected_parts:
            part = part_map[p]
            if part["inferred_part_type"] == "Structural":
                stack.append(p)
                total_thickness += part["thickness_mm"]

        if not stack:
            continue

        inferred_holes.append({
            "hole_id": f"HOLE_{rivet}",
            "rivet": rivet,
            "stack_parts": stack,
            "stack_thickness_mm": round(total_thickness, 3),
            "inferred_diameter_mm": 8.0,  # heuristic (M8 rivet)
            "confidence": "High"
        })

    # -------------------------------
    # Attach results
    # -------------------------------
    data["inferred_holes"] = inferred_holes

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    print(f"[OK] Inferred {len(inferred_holes)} hole stacks")
    print(f"[OK] Output written to {OUTPUT_JSON}")


if __name__ == "__main__":
    run()