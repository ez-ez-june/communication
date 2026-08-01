from __future__ import annotations

import re

from .utils import extract_hidden_value, extract_query_value, scheme_host


def convert_fn_egov_select_to_url(href: str, base_url: str) -> str:
    match = re.search(r"fn_egov_select\s*\(\s*([^,]+)\s*,\s*([^)]+)\)", href, re.I)
    if not match:
        return ""
    ntt_id = match.group(1).strip("'\" ")
    bbs_id = match.group(2).strip("'\" ")
    if not ntt_id or not bbs_id:
        return ""
    menu_no = extract_query_value(base_url, "menuNo") or "4050100"
    return (
        f"{scheme_host(base_url)}/nw/nes/detailNesDtaView.do"
        f"?searchBbsId1={bbs_id}&searchNttId1={ntt_id}&menuNo={menu_no}"
    )


def convert_fn_detail_to_url(href: str, base_url: str, html: str) -> str:
    match = re.search(r"fn_detail\s*\(\s*([^)]+)\)", href, re.I)
    if not match:
        return ""
    ntt_seq_no = match.group(1).strip("'\" ")
    if not ntt_seq_no:
        return ""
    bbs_seq_no = extract_hidden_value(html, "bbsSeqNo") or extract_query_value(base_url, "bbsSeqNo")
    s_code = extract_hidden_value(html, "sCode") or extract_query_value(base_url, "sCode") or "user"
    menu_id = extract_hidden_value(html, "mId") or extract_query_value(base_url, "mId")
    menu_pid = extract_hidden_value(html, "mPid") or extract_query_value(base_url, "mPid")
    path = f"/bbs/view.do?sCode={s_code}"
    if menu_id:
        path += f"&mId={menu_id}"
    if menu_pid:
        path += f"&mPid={menu_pid}"
    path += (
        f"&pageIndex=&bbsSeqNo={bbs_seq_no}&nttSeqNo={ntt_seq_no}"
        "&searchOpt=ALL&searchTxt="
    )
    return scheme_host(base_url) + path


def convert_go_view_to_url(onclick: str, base_url: str) -> str:
    match = re.search(r"goView\s*\(([^)]+)\)", onclick, re.I)
    if not match:
        return ""
    parts = [p.strip().strip("'\"") for p in match.group(1).split(",")]
    if len(parts) < 2:
        return ""
    board_id, board_seq = parts[0], parts[1]
    status_yn = parts[4] if len(parts) >= 5 else "W"
    curr_page = parts[5] if len(parts) >= 6 else "1"
    menu_val = extract_query_value(base_url, "m") or "020501"
    return (
        f"{scheme_host(base_url)}/boardCnts/viewRenew.do?boardID={board_id}"
        f"&boardSeq={board_seq}&lev=0&searchType=null&statusYN={status_yn}"
        f"&page={curr_page}&s=moe&m={menu_val}&opType=N"
    )


def convert_mogef_to_url(onclick: str, base_url: str) -> str:
    if "fn_selectView".lower() not in onclick.lower():
        return ""
    match = re.search(r"(\d{5,})", onclick)
    if not match:
        return ""
    bbt_sn = match.group(1)
    mid_val = extract_query_value(base_url, "mid") or "news400"
    return f"{scheme_host(base_url)}/nw/ntc/nw_ntc_s001d.do?mid={mid_val}&bbtSn={bbt_sn}"


def convert_fn_select_doc_to_url(onclick: str, base_url: str) -> str:
    match = re.search(r"fn_selectDoc\s*\(\s*([^)]+)\)", onclick, re.I)
    if not match:
        return ""
    doc_seq = match.group(1).strip("'\" ")
    if not doc_seq:
        return ""
    menu_seq = extract_query_value(base_url, "menuSeq") or "375"
    bbs_seq = extract_query_value(base_url, "bbsSeq") or "9"
    return (
        f"{scheme_host(base_url)}/doc/ko/selectDoc.do"
        f"?docSeq={doc_seq}&menuSeq={menu_seq}&bbsSeq={bbs_seq}"
    )


def convert_js_to_url(href: str, onclick: str, base_url: str, html: str) -> str:
    for fn in (
        lambda: convert_fn_egov_select_to_url(href, base_url),
        lambda: convert_fn_detail_to_url(href, base_url, html),
        lambda: convert_fn_egov_select_to_url(onclick, base_url),
        lambda: convert_fn_detail_to_url(onclick, base_url, html),
        lambda: convert_go_view_to_url(onclick, base_url),
        lambda: convert_mogef_to_url(onclick, base_url),
        lambda: convert_fn_select_doc_to_url(onclick, base_url),
    ):
        result = fn()
        if result:
            return result
    return ""
