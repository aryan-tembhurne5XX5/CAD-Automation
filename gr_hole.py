import json
from collections import defaultdict
from pathlib import Path

INPUT_JSON  = Path(r"E:\Phase 1\extractions\inferred_holes.json")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\grouped_holes.json")

with open(INPUT_JSON, "r", encoding="utf-8") as f:
    holes = json.load(f)

grouped = defaultdict(list)

for h in holes:
    plate = h["hole_stack"][0]
    fastener_type = h["fastener"].split(":")[0]  # part number only

    key = (plate, fastener_type)
    grouped[key].append(h["fastener"])

result = []

for (plate, fastener), instances in grouped.items():
    result.append({
        "plate": plate,
        "fastener_type": fastener,
        "hole_count": len(instances),
        "instances": instances,
        "confidence": round(0.8 + 0.01 * len(instances), 2)
    })

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(result, f, indent=4)

print(f"✅ Grouped holes written → {OUTPUT_JSON}")