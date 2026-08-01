from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Post:
    title: str
    url: str
    date: str = ""


@dataclass
class DeptConfig:
    name: str
    url: str
    enabled: bool = True
    include_hint: str = ""
    exclude_hint: str = ""


@dataclass
class DeptResult:
    name: str
    ok: bool
    posts: list[Post] = field(default_factory=list)
    notices: list[Post] = field(default_factory=list)
    press: list[Post] = field(default_factory=list)
    error: str = ""
