import argparse
import json
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "json_path",
        nargs="?",
        type=Path,
        default=Path(__file__).resolve().with_name("lambda_functions.json"),
        help="Path to lambda_functions.json (defaults to the file next to this script)",
    )
    args = parser.parse_args()

    data = json.loads(args.json_path.read_text(encoding="utf-8"))
    for fn in data["functions"]:
        lines = []
        for a in fn["arguments"]:
            name = f"[{a['name']}]" if a.get("optional") else a["name"]
            lines.append(f"{name}: {a['description']}")
        text = "\n\n".join(lines)
        print(f"{fn['name']:20s}  {len(text):3d} chars")


if __name__ == "__main__":
    main()
