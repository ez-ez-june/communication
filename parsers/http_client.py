from __future__ import annotations

import time

import requests

HTTP_TIMEOUT = 20
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}


def get_html(url: str) -> tuple[str, str]:
    errors: list[str] = []
    session = requests.Session()
    for attempt in range(1, 7):
        try:
            response = session.get(url, headers=HEADERS, timeout=HTTP_TIMEOUT)
            if response.status_code < 200 or response.status_code >= 400:
                errors.append(f"HTTP {response.status_code}")
            else:
                response.encoding = response.apparent_encoding or response.encoding or "utf-8"
                text = response.text or ""
                if text.strip():
                    return text, ""
                errors.append("Empty HTML")
        except requests.RequestException as exc:
            errors.append(str(exc))
        if attempt < 6:
            time.sleep(attempt)
    return "", " / ".join(errors) if errors else "Fetch failed"
