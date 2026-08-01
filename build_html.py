from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"

OUTPUTS = [
    ("gov_notice_board_pc_template.html", "gov_notice_board.html", "%%NOTICE_JSON%%", "notice.json"),
    ("gov_notice_board_mobile_template.html", "gov_notice_board_mobile.html", "%%NOTICE_JSON%%", "notice.json"),
    ("gov_major_press_pc_template.html", "gov_major_press_board.html", "%%MAJOR_JSON%%", "major.json"),
    ("gov_major_press_mobile_template.html", "gov_major_press_board_mobile.html", "%%MAJOR_JSON%%", "major.json"),
]


def build_html(compact: bool = True) -> None:
    for template_name, output_name, token, json_name in OUTPUTS:
        template_path = ROOT / template_name
        json_path = DATA_DIR / json_name
        output_path = ROOT / output_name

        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        if not json_path.exists():
            raise FileNotFoundError(f"JSON not found: {json_path} (run collector.py first)")

        template = template_path.read_text(encoding="utf-8")
        payload = json.loads(json_path.read_text(encoding="utf-8"))
        json_text = json.dumps(payload, ensure_ascii=False, separators=(",", ":") if compact else (",", ": "))
        if token not in template:
            raise ValueError(f"Token {token} missing in {template_name}")

        output_path.write_text(template.replace(token, json_text), encoding="utf-8")
        print(f"Wrote: {output_path}")


def main() -> None:
    build_html()


if __name__ == "__main__":
    main()
