from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from config.departments import DEPARTMENTS, MAJOR_PRESS
from parsers.generic import parse_announcement_links
from parsers.http_client import get_html
from parsers.types import DeptResult, Post

MAX_POSTS = 5
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"


def fetch_department(name: str, url: str, include_hint: str = "", exclude_hint: str = "") -> DeptResult:
    if not url.lower().startswith("http"):
        return DeptResult(name=name, ok=False, error="Invalid list URL")

    html, err = get_html(url)
    if err:
        return DeptResult(name=name, ok=False, error=err)

    posts = parse_announcement_links(html, url, include_hint, exclude_hint, MAX_POSTS)
    if not posts and include_hint:
        posts = parse_announcement_links(html, url, "", exclude_hint, MAX_POSTS)

    if not posts:
        return DeptResult(name=name, ok=False, error="No post links found (check HTML / include hint)")

    return DeptResult(name=name, ok=True, posts=posts)


def posts_to_json(posts: list[Post]) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    for post in posts:
        item = {"title": post.title, "url": post.url}
        if post.date:
            item["date"] = post.date
        items.append(item)
    return items


def collect_notices() -> dict:
    departments: list[dict] = []
    for dept in DEPARTMENTS:
        if not dept.enabled:
            continue
        print(f"Collecting: {dept.name}")
        result = fetch_department(dept.name, dept.url, dept.include_hint, dept.exclude_hint)
        entry = {
            "name": result.name,
            "ok": result.ok,
            "posts": posts_to_json(result.posts) if result.ok else [],
        }
        if not result.ok:
            print(f"  FAIL: {result.error}")
        else:
            print(f"  OK: {len(result.posts)} posts")
        departments.append(entry)

    return {
        "collectedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "departments": departments,
    }


def collect_major_press(notice_data: dict) -> dict:
    notice_map = {d["name"]: d for d in notice_data.get("departments", [])}
    departments: list[dict] = []

    for item in MAJOR_PRESS:
        name = item["name"]
        notices = notice_map.get(name, {}).get("posts", []) if notice_map.get(name, {}).get("ok") else []

        press: list[dict] = []
        html, err = get_html(item["press_url"])
        if not err and html:
            posts = parse_announcement_links(html, item["press_url"], item["include_hint"], "", MAX_POSTS)
            if not posts and item["include_hint"]:
                posts = parse_announcement_links(html, item["press_url"], "", "", MAX_POSTS)
            press = posts_to_json(posts)

        departments.append(
            {
                "name": name,
                "ok": bool(notices),
                "notices": notices,
                "press": press,
            }
        )

    return {
        "collectedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "departments": departments,
    }


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    notice = collect_notices()
    major = collect_major_press(notice)

    notice_path = DATA_DIR / "notice.json"
    major_path = DATA_DIR / "major.json"
    notice_path.write_text(json.dumps(notice, ensure_ascii=False, indent=2), encoding="utf-8")
    major_path.write_text(json.dumps(major, ensure_ascii=False, indent=2), encoding="utf-8")

    ok_count = sum(1 for d in notice["departments"] if d["ok"])
    fail_count = len(notice["departments"]) - ok_count
    print(f"\nDone: OK {ok_count} / FAIL {fail_count} / TOTAL {len(notice['departments'])}")
    print(f"Saved: {notice_path}")
    print(f"Saved: {major_path}")


if __name__ == "__main__":
    main()
