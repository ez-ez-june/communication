from __future__ import annotations

from .js_urls import convert_js_to_url
from .router import try_specialized_parse
from .types import Post
from .utils import (
    clean_text,
    clean_url,
    extract_attr_value,
    extract_link_text,
    get_tag_attr_value,
    guess_date_near_link,
    is_junk_title,
    is_likely_post_link,
    to_absolute_url,
)

MAX_POSTS = 5


def parse_announcement_links(
    html: str,
    base_url: str,
    include_hint: str = "",
    exclude_hint: str = "",
    max_count: int = MAX_POSTS,
) -> list[Post]:
    specialized = try_specialized_parse(html, base_url, max_count)
    if specialized:
        return specialized

    result: list[Post] = []
    seen: set[str] = set()
    lower_html = html.lower()
    pos = 1
    loop_count = 0

    while len(result) < max_count and loop_count < 8000:
        loop_count += 1
        a_open = lower_html.find("<a ", pos)
        if a_open < 0:
            a_open = lower_html.find("<a>", pos)
        if a_open < 0:
            break

        href_pos = lower_html.find("href=", a_open)
        end_tag = lower_html.find(">", a_open)
        if href_pos < 0 or end_tag < 0 or href_pos > end_tag:
            pos = a_open + 2
            continue

        href = extract_attr_value(html, href_pos + 5)
        onclick_value = get_tag_attr_value(html, a_open, end_tag, "onclick")
        title_attr = get_tag_attr_value(html, a_open, end_tag, "title")
        accepted = False

        if is_likely_post_link(href, include_hint, exclude_hint):
            accepted = True
        elif any(token in href.lower() for token in ("fn_egov_select", "fn_detail", "goview")):
            js_abs = convert_js_to_url(href, onclick_value, base_url, html)
            if js_abs:
                href = js_abs
                accepted = True
        elif len(onclick_value) > 8 and any(
            token in onclick_value.lower() for token in ("fn_", "goview", "fnview")
        ):
            js_abs = convert_js_to_url(href, onclick_value, base_url, html)
            if js_abs:
                href = js_abs
                accepted = True

        if not accepted:
            pos = end_tag + 1
            continue

        title = clean_text(extract_link_text(html, end_tag))
        if len(title) < 2:
            title = clean_text(title_attr)
            if title.startswith("[") and "]" in title[:30]:
                title = title.split("]", 1)[1].strip()

        if len(title) < 2 or len(title) > 300 or is_junk_title(title):
            pos = end_tag + 1
            continue

        abs_url = clean_url(to_absolute_url(base_url, href))
        if not abs_url.lower().startswith("http"):
            pos = end_tag + 1
            continue
        if abs_url.lower() in seen:
            pos = end_tag + 1
            continue

        date_text = guess_date_near_link(html, a_open)
        seen.add(abs_url.lower())
        result.append(Post(title=title, url=abs_url, date=date_text))
        pos = end_tag + 1

    return result
