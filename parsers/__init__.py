"""Government notice collectors ported from modGovNotice.bas."""

from .generic import parse_announcement_links
from .http_client import get_html
from .router import try_specialized_parse
from .types import DeptConfig, DeptResult, Post

__all__ = [
    "DeptConfig",
    "DeptResult",
    "Post",
    "get_html",
    "parse_announcement_links",
    "try_specialized_parse",
]
