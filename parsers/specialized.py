from __future__ import annotations

import re
from urllib.parse import unquote

from .js_urls import convert_fn_detail_to_url
from .types import Post
from .utils import clean_text, clean_url, extract_query_value, is_junk_title, to_absolute_url


def _add_unique(result: list[Post], seen: set[str], title: str, url: str, max_count: int) -> None:
    if len(result) >= max_count or len(title) < 2 or is_junk_title(title):
        return
    key = url.lower()
    if key in seen:
        return
    seen.add(key)
    result.append(Post(title=title, url=url))


def extract_by_regex(
    html: str,
    base_url: str,
    pattern: str,
    max_count: int,
    title_first: bool = False,
) -> list[Post]:
    result: list[Post] = []
    seen: set[str] = set()
    for match in re.finditer(pattern, html, re.I | re.S | re.M):
        if title_first:
            title = clean_text(match.group(1))
            href = match.group(2)
        else:
            href = match.group(1)
            title = clean_text(match.group(2))
        abs_url = clean_url(to_absolute_url(base_url, href))
        if len(title) >= 2 and abs_url:
            _add_unique(result, seen, title, abs_url, max_count)
    return result


def extract_by_regex_title_href(html: str, base_url: str, pattern: str, max_count: int) -> list[Post]:
    return extract_by_regex(html, base_url, pattern, max_count, title_first=True)


def extract_mois_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    seen: set[str] = set()
    patterns = [
        r"inqire_notice\('(\d+)',\s*'BBSMSTR_000000000006'\)[^>]*>\s*([^<]{2,})",
        r'commonSelectBoardArticle\.do(?:;jsessionid=[^?"\']+)?\?bbsId=BBSMSTR_000000000006(?:&amp;|&)nttId=(\d+)[^"\']*["\'][^>]*>\s*([^<]{2,})',
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, html, re.I):
            ntt_id, title = match.group(1), clean_text(match.group(2))
            if len(title) < 2:
                continue
            href = (
                "https://www.mois.go.kr/frt/bbs/type013/commonSelectBoardArticle.do"
                f"?bbsId=BBSMSTR_000000000006&nttId={ntt_id}"
            )
            _add_unique(result, seen, title, href, max_count)
    return result


def extract_moe_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    seen: set[str] = set()
    menu_val = extract_query_value(base_url, "m") or "020501"
    pattern = r"goView\('333',\s*'(\d+)'[^\"]*\"[^>]*title=\"([^\"]+)\""
    for match in re.finditer(pattern, html, re.I):
        board_seq = match.group(1)
        title = clean_text(match.group(2))
        href = (
            "https://www.moe.go.kr/boardCnts/viewRenew.do?boardID=333"
            f"&boardSeq={board_seq}&lev=0&searchType=null&statusYN=W&page=1"
            f"&s=moe&m={menu_val}&opType=N"
        )
        _add_unique(result, seen, title, href, max_count)
    return result


def extract_mpm_links(html: str, base_url: str, max_count: int) -> list[Post]:
    pattern = r'href="(\?boardId=[^"]*mode=view[^"]*cntId=\d+[^"]*)"\s*>\s*([^<]{2,})\s*<'
    return extract_by_regex(html, base_url, pattern, max_count)


def extract_mpva_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    key_val = extract_query_value(base_url, "key") or "76"
    bbs_no = extract_query_value(base_url, "bbsNo") or "15"
    pattern = r'href="[^"]*selectBbsNttView\.do[^"]*nttNo=(\d+)[^"]*"[^>]*>\s*([^<]{2,})</a>'
    for match in re.finditer(pattern, html, re.I):
        ntt_no = match.group(1)
        title = clean_text(match.group(2))
        href = f"https://www.mpva.go.kr/mpva/selectBbsNttView.do?key={key_val}&bbsNo={bbs_no}&nttNo={ntt_no}"
        if len(title) >= 2:
            result.append(Post(title=title, url=href))
            if len(result) >= max_count:
                break
    return result


def extract_msit_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    key_matches = list(re.finditer(r"fn_detail\s*\(\s*['\"]?(\d+)['\"]?\s*\)", html, re.I))
    title_matches = list(re.finditer(r"sHtml\+= unescape\('([^']+)'\);", html, re.I))
    for idx, key_match in enumerate(key_matches):
        if idx >= len(title_matches):
            break
        ntt_seq_no = key_match.group(1)
        title = clean_text(unquote(title_matches[idx].group(1)))
        href = convert_fn_detail_to_url(f"fn_detail({ntt_seq_no})", base_url, html)
        if title and href:
            result.append(Post(title=title, url=href))
            if len(result) >= max_count:
                break
    return result


def extract_motir_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    board_id = "ATCL6e90bb9de"
    lower = base_url.lower()
    marker = "/kor/article/"
    pos = lower.find(marker)
    if pos >= 0:
        chunk = base_url[pos + len(marker) :]
        board_id = chunk.split("?", 1)[0].split("/", 1)[0] or board_id
    pattern = r"article\.view\s*\(\s*['\"]?(\d+)['\"]?\s*\)[^>]*>\s*<i>([^<]+)</i>"
    for match in re.finditer(pattern, html, re.I):
        article_id = match.group(1)
        title = clean_text(match.group(2))
        href = (
            f"https://www.motir.go.kr/kor/article/{board_id}/{article_id}/view"
            "?mno=&pageIndex=1&rowPageC=0&displayAuthor=&searchCategory=0"
            "&schClear=on&startDtD=&endDtD=&searchCondition=1&searchKeyword="
        )
        result.append(Post(title=title, url=href))
        if len(result) >= max_count:
            break
    return result


def extract_mnd_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    href_matches = list(re.finditer(r'href="(/bbs/mnd/11066/[^"]+/artclView\.do)"', html, re.I))
    title_matches = list(re.finditer(r"<strong><span>([^<]+)</span></strong>", html, re.I))
    for idx, href_match in enumerate(href_matches):
        if idx >= len(title_matches):
            break
        href = href_match.group(1)
        title = clean_text(title_matches[idx].group(1))
        abs_url = clean_url(to_absolute_url(base_url, href))
        if title and abs_url:
            result.append(Post(title=title, url=abs_url))
            if len(result) >= max_count:
                break
    return result


def extract_mss_links(html: str, base_url: str, max_count: int) -> list[Post]:
    result: list[Post] = []
    pattern = r'doBbsFView\(\'81\',\'(\d+)\',\'[^\']*\',\'(\d+)\'\);""\s+title="([^"]+)"'
    for match in re.finditer(pattern, html, re.I):
        bc_idx = match.group(1)
        parent_seq = match.group(2)
        title = clean_text(match.group(3))
        href = f"https://www.mss.go.kr/site/smba/ex/bbs/View.do?cbIdx=81&bcIdx={bc_idx}&parentSeq={parent_seq}"
        result.append(Post(title=title, url=href))
        if len(result) >= max_count:
            break
    return result
