from __future__ import annotations

import re

from .specialized import (
    extract_by_regex,
    extract_by_regex_title_href,
    extract_mnd_links,
    extract_moe_links,
    extract_mois_links,
    extract_motir_links,
    extract_mpva_links,
    extract_mpm_links,
    extract_mss_links,
    extract_msit_links,
)
from .types import Post
from .utils import clean_text

MAX_POSTS = 5


def try_specialized_parse(html: str, base_url: str, max_count: int = MAX_POSTS) -> list[Post]:
    lower_url = base_url.lower()

    if "msit.go.kr" in lower_url and "/bbs/list.do" in lower_url:
        return extract_msit_links(html, base_url, max_count)
    if "mofa.go.kr" in lower_url and (
        "/www/brd/m_4075/list.do" in lower_url or "/www/brd/m_4080/list.do" in lower_url
    ):
        return extract_by_regex(
            html,
            base_url,
            r'href="(\./view\.do\?seq=\d+[^"]*)"[^>]*>\s*(?:<span[^>]*>[^<]*</span>\s*)?([^<]{2,})',
            max_count,
        )
    if "mois.go.kr" in lower_url and "bbsmstr_000000000006" in lower_url:
        return extract_mois_links(html, base_url, max_count)
    if "moe.go.kr" in lower_url and "boardcnts/listrenew.do" in lower_url:
        return extract_moe_links(html, base_url, max_count)
    if "mpm.go.kr" in lower_url and ("newsnoitice" in lower_url or "noticelist" in lower_url):
        return extract_mpm_links(html, base_url, max_count)
    if "moj.go.kr" in lower_url and "/moj/223/subview.do" in lower_url:
        return extract_by_regex_title_href(
            html,
            base_url,
            r'title="([^"]+)"[\s\S]{0,120}?<a href="(/bbs/moj/184/\d+/artclView\.do)"',
            max_count,
        )
    if "mnd.go.kr" in lower_url and "/mnd/154/subview.do" in lower_url:
        return extract_mnd_links(html, base_url, max_count)
    if "mafra.go.kr" in lower_url and "/home/5108/subview.do" in lower_url:
        return extract_by_regex(
            html,
            base_url,
            r'href="(/bbs/home/791/\d+/artclView\.do)"[\s\S]{0,140}?\r?\n\s*([^<\r\n]{2,})\s*\r?\n\s*<span class="new">',
            max_count,
        )
    if "motir.go.kr" in lower_url and "/kor/article/" in lower_url:
        return extract_motir_links(html, base_url, max_count)
    if "mss.go.kr" in lower_url and "cbidx=81" in lower_url:
        return extract_mss_links(html, base_url, max_count)
    if "mcst.go.kr" in lower_url and "noticelist.jsp" in lower_url:
        return extract_by_regex(
            html,
            base_url,
            r'href="(noticeView\.jsp\?pSeq=[^"]+)"[^>]*title="([^"]+)"',
            max_count,
        )
    if "unikorea.go.kr" in lower_url and "bbs_0000000000000001" in lower_url:
        return extract_by_regex(
            html,
            base_url,
            r'href="(/web/unikorea/bbs/bbs_0000000000000001/\d+[^"]*)"[^>]*>\s*([^<]+?)\s*</a>',
            max_count,
        )
    if "molit.go.kr" in lower_url and "/usr/bord0201/m_69/brd.jsp" in lower_url:
        posts = extract_by_regex(
            html,
            base_url,
            r'href="(\./DTL\.jsp[^"]+)"[^>]*>([\s\S]*?)</a>',
            max_count,
        )
        return [
            Post(title=clean_text(re.sub(r"<[^>]+>", " ", post.title)), url=post.url, date=post.date)
            for post in posts
        ]
    if "mpva.go.kr" in lower_url and "selectbbsnttlist.do" in lower_url:
        return extract_mpva_links(html, base_url, max_count)
    if "mfds.go.kr" in lower_url and "/brd/m_689/" in lower_url:
        return extract_by_regex(
            html,
            base_url,
            r'href="(\./view\.do\?seq=\d+[^"]*)"[^>]*class="title"[^>]*>\s*([^<]{2,})',
            max_count,
        )
    if "mohw.go.kr" in lower_url and "bid=0003" in lower_url:
        return extract_by_regex(
            html,
            base_url,
            r'href="(/board\.es\?[^"]*act=view[^"]*list_no=\d+[^"]*)"[^>]*class="txt_title"[^>]*>[\s\S]*?(?:</span>\s*)?([^<]{2,})</a>',
            max_count,
        )

    return []
