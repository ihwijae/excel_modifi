"""HTML parsing helpers for the SMPP screens."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional

import re
from bs4 import BeautifulSoup


@dataclass
class CorpFeatures:
    small_exists: bool
    small_confirm_date: Optional[str]
    small_expire_date: Optional[str]
    women_exists: bool
    women_confirm_date: Optional[str]
    women_expire_date: Optional[str]


def build_move_form_payload_from_list_html(html: str, biz_no: str) -> Dict[str, str]:
    """
    Parse the 목록 HTML and reconstruct the hidden ``moveForm`` payload that
    the site uses when javascript submits the 상세 요청.
    """
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form", attrs={"name": "moveForm"})
    if not form:
        raise RuntimeError("moveForm 폼을 찾을 수 없습니다.")

    payload: Dict[str, str] = {}
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        payload[name] = inp.get("value", "") or ""

    digits = re.sub(r"\D", "", biz_no)
    if not digits:
        raise ValueError("사업자등록번호에서 숫자를 추출할 수 없습니다.")
    payload["bsnmNo"] = digits
    if payload.get("searchBsnmNo") in ("", None):
        payload["searchBsnmNo"] = digits
    return payload


def parse_corp_features(html: str) -> CorpFeatures:
    """Parse the '기업특징' tab and extract validity dates."""
    soup = BeautifulSoup(html, "html.parser")

    tab = soup.find("div", class_=re.compile(r"tabContent.*a2"))
    if not tab:
        return CorpFeatures(False, None, None, False, None, None)

    table = tab.find("table")
    if not table:
        return CorpFeatures(False, None, None, False, None, None)

    tbody = table.find("tbody") or table

    small_exists = False
    small_confirm: Optional[str] = None
    small_expire: Optional[str] = None
    women_exists = False
    women_confirm: Optional[str] = None
    women_expire: Optional[str] = None

    for tr in tbody.find_all("tr"):
        tds = tr.find_all("td")
        if not tds:
            continue

        kind = tds[0].get_text(strip=True)
        clean_cells = [td.get_text(strip=True) or None for td in tds]

        if _is_no_entry_row(clean_cells):
            continue

        if "중소기업" in kind and len(clean_cells) >= 4:
            small_exists = True
            small_confirm = clean_cells[2]
            small_expire = clean_cells[3]

        if "여성기업" in kind and len(clean_cells) >= 4:
            women_exists = True
            women_confirm = clean_cells[2]
            women_expire = clean_cells[3]

    return CorpFeatures(
        small_exists=small_exists,
        small_confirm_date=small_confirm,
        small_expire_date=small_expire,
        women_exists=women_exists,
        women_confirm_date=women_confirm,
        women_expire_date=women_expire,
    )


def _is_no_entry_row(cells: list[Optional[str]]) -> bool:
    joined = " ".join(filter(None, cells))
    return "해당사항 없음" in joined or "정보 없음" in joined
