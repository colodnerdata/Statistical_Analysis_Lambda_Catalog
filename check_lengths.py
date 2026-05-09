import json
from pathlib import Path
data = json.loads(Path("lambda_functions.json").read_text(encoding="utf-8"))
for fn in data["functions"]:
    lines = []
    for a in fn["arguments"]:
        name = f"[{a['name']}]" if a.get("optional") else a["name"]
        lines.append(f"{name}: {a['description']}")
    text = "\n\n".join(lines)
    print(f"{fn['name']:20s}  {len(text):3d} chars")
