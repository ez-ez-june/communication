from __future__ import annotations

import html as html_lib
import re
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse, urlunparse


STATIC_EXTENSIONS = (".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".ico", ".woff", ".ttf")


def clean_text(value: str) -> str:
    text = value.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    for entity in ("&nbsp;", "&amp;", "&lt;", "&gt;", "&quot;", "&#39;", "&#034;", "&apos;"):
        text = text.replace(entity, html_lib.unescape(entity))
    text = decode_numeric_entities(text)
    while "  " in text:
        text = text.replace("  ", " ")
    text = text.strip()
    if text.startswith("새글"):
        text = text[2:].strip()
    return text


def decode_numeric_entities(value: str) -> str:
    def repl(match: re.Match[str]) -> str:
        try:
            code = int(match.group(1))
        except ValueError:
            return match.group(0)
        if 0 < code < 65536:
            return chr(code)
        return match.group(0)

    return re.sub(r"&#(\d+);", repl, value)


def clean_url(url: str) -> str:
    text = url.strip()
    text = text.replace("&amp;", "&").replace("&quot;", '"').replace("&#034;", '"').replace("/./", "/")
    lower = text.lower()
    marker = ";jsessionid="
    pos = lower.find(marker)
    if pos >= 0:
        q = text.find("?", pos)
        if q >= 0:
            text = text[:pos] + text[q:]
        else:
            text = text[:pos]
    return text


def extract_query_value(url: str, key: str) -> str:
    parsed = urlparse(url)
    query = parsed.query
    if not query and "?" in url:
        query = url.split("?", 1)[1]
    for chunk in query.split("&"):
        if chunk.lower().startswith(f"{key.lower()}="):
            return chunk.split("=", 1)[1]
    return ""


def scheme_host(base_url: str) -> str:
    parsed = urlparse(base_url)
    return f"{parsed.scheme}://{parsed.netloc}"


def to_absolute_url(base_url: str, href: str) -> str:
    href = href.strip()
    if not href:
        return ""
    if href.startswith(("http://", "https://")):
        return href
    if href.startswith("//"):
        parsed = urlparse(base_url)
        return f"{parsed.scheme}:{href}"
    if href.startswith("?"):
        bare = base_url.split("?", 1)[0]
        return bare + href
    return urljoin(base_url, href)


def has_extension(lower_href: str, ext: str) -> bool:
    pos = 0
    while True:
        pos = lower_href.find(ext, pos)
        if pos < 0:
            return False
        if ext == ".js" and lower_href[pos : pos + 4] == ".jsp":
            pos += 1
            continue
        end = pos + len(ext)
        if end >= len(lower_href):
            return True
        nxt = lower_href[end]
        if nxt in "?#&/;":
            return True
        pos += 1


def is_static_asset(lower_href: str) -> bool:
    return any(has_extension(lower_href, ext) for ext in STATIC_EXTENSIONS)


def is_bbs_numeric_detail(lower_href: str) -> bool:
    chunk = lower_href.rsplit("/", 1)[-1]
    chunk = chunk.split("?", 1)[0]
    return len(chunk) >= 3 and chunk.isdigit()


def is_likely_post_link(href: str, include_hint: str, exclude_hint: str) -> bool:
    h = href.lower().strip()
    if not h:
        return False
    if h.startswith(("javascript:", "mailto:")):
        return False
    if h in ("#", "#none") or h.startswith("#"):
        return False
    if any(token in h for token in ("logout", "login", "filedown", "download")):
        return False
    if is_static_asset(h):
        return False
    if exclude_hint and exclude_hint.lower() in h:
        return False
    if include_hint:
        return include_hint.lower() in h

    checks = (
        "nttid=",
        "nttno=",
        "article",
        "/view.do",
        "view.do?",
        "view.jsp",
        "boardview",
        "read.do",
        "selectboardarticle",
        "commonselectboardarticle",
        "boardcnts/view",
        "dtl.jsp",
        "idx=",
        "mode=view",
        "act=view",
        "list_no=",
        "cntid=",
        "noticedetail",
        "noticedtl",
        "noticeview",
        "selectbbsnttview",
        "viewrenew.do",
        "selectdoc.do",
        "nw_ntc_s001d",
    )
    if "mainview.do" in h:
        return False
    if any(token in h for token in checks):
        return True
    if "subview.do" in h and "enc=" in h:
        return True
    if "bbs/" in h and ("view" in h or "cntid=" in h):
        return True
    if "brd/" in h and "view" in h:
        return True
    if "/bbs/" in h and is_bbs_numeric_detail(h):
        return True
    if "bause=true" in h and "/bbs/" in h:
        return True
    if any(token in h for token in ("list.do", "list.jsp", "listrenew", "main.do", "index.do", "index.jsp")):
        return False
    return False


def is_junk_title(title: str) -> bool:
    t = title.lower().replace(" ", "")
    return t in {"more", "list", "prev", "next", "home", "login", "print", "share"} or len(t) <= 1


def extract_attr_value(source: str, start_pos: int) -> str:
    i = start_pos
    while i < len(source) and source[i] in " \t\r\n":
        i += 1
    if i >= len(source):
        return ""
    quote = source[i]
    if quote in "\"'":
        i += 1
        buf: list[str] = []
        while i < len(source):
            ch = source[i]
            if ch == quote:
                break
            buf.append(ch)
            i += 1
        return "".join(buf)
    buf = []
    while i < len(source) and source[i] not in " >\t\r\n":
        buf.append(source[i])
        i += 1
    return "".join(buf)


def get_tag_attr_value(html: str, tag_start: int, tag_end: int, attr_name: str) -> str:
    tag_html = html[tag_start:tag_end]
    lower_tag = tag_html.lower()
    needle = f"{attr_name.lower()}="
    pos = lower_tag.find(needle)
    if pos < 0:
        return ""
    return extract_attr_value(tag_html, pos + len(needle))


class _LinkTextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() in {"script", "style"}:
            self._skip += 1

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"script", "style"} and self._skip:
            self._skip -= 1

    def handle_data(self, data: str) -> None:
        if not self._skip:
            self._parts.append(data)


def extract_link_text(html: str, after_open_tag: int) -> str:
    lower = html.lower()
    close_pos = lower.find("</a>", after_open_tag + 1)
    if close_pos < 0:
        return ""
    inner = html[after_open_tag + 1 : close_pos]
    parser = _LinkTextParser()
    try:
        parser.feed(inner)
        parser.close()
    except Exception:
        return clean_text(re.sub(r"<[^>]+>", "", inner))
    return clean_text("".join(parser._parts))


def guess_date_near_link(html: str, link_pos: int) -> str:
    start = max(0, link_pos - 120)
    chunk = html[start : start + 400]
    for match in re.finditer(r"\d{4}[-./]\d{2}[-./]\d{2}", chunk):
        value = match.group(0).replace(".", "-").replace("/", "-")
        return value
    return ""


def extract_hidden_value(html: str, field_name: str) -> str:
    pattern = rf'<input[^>]*name\s*=\s*["\']{re.escape(field_name)}["\'][^>]*value\s*=\s*["\']([^"\']*)["\']'
    match = re.search(pattern, html, re.I)
    return match.group(1) if match else ""
