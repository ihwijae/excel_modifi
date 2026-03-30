"""HTTP client helpers for the SMPP portal."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional

import requests
from bs4 import BeautifulSoup


LOGIN_PAGE_URL = "https://www.smpp.go.kr/uat/uia/egovLoginUsr.do"
LOGIN_URL = "https://www.smpp.go.kr/uat/uia/actionLogin.do"
LIST_URL = "https://www.smpp.go.kr/cop/registcorp/selectRegistCorpListVw.do"
SUMMARY_URL = "https://www.smpp.go.kr/cop/registcorp/selectRegistCorpSumryInfoVw.do"


@dataclass
class SmppCredentials:
    """Simple container for SMPP login information."""

    user_id: str
    password: str


class SmppClient:
    """Wrapper around ``requests.Session`` with SMPP-specific helpers."""

    def __init__(self, creds: SmppCredentials, timeout: float = 30.0) -> None:
        self.creds = creds
        self.timeout = timeout
        self.session = requests.Session()

    def _default_headers(self) -> Dict[str, str]:
        return {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Origin": "https://www.smpp.go.kr",
            "Referer": LOGIN_URL,
        }

    def login(self) -> None:
        """Perform a SMPP login and ensure redirect target contains ``loginSuccess``."""
        headers = self._default_headers()
        initial = self.session.get(LOGIN_PAGE_URL, headers=headers, timeout=self.timeout)
        initial.raise_for_status()
        payload = _build_login_form_payload(initial.text, self.creds.user_id, self.creds.password)

        resp = self.session.post(
            LOGIN_URL,
            headers=headers,
            data=payload,
            allow_redirects=False,
            timeout=self.timeout,
        )
        location = resp.headers.get("Location", "")
        body = resp.text if resp.content else ""
        if (resp.status_code in (301, 302) and "loginSuccess" in location) or "loginSuccess" in body:
            return
        raise RuntimeError(f"SMPP 로그인에 실패했습니다. status={resp.status_code}, location={location}")

    def fetch_list_html_by_biz_no(self, biz_no: str) -> str:
        """Fetch the corporation list HTML filtered by the provided biz-no."""
        data = self._build_list_payload(biz_no)
        resp = self.session.post(LIST_URL, data=data, timeout=self.timeout)
        resp.raise_for_status()
        return resp.text

    def fetch_summary_html(self, payload: Dict[str, str]) -> str:
        """Fetch the detailed summary HTML using the moveForm payload."""
        resp = self.session.post(SUMMARY_URL, data=payload, timeout=self.timeout)
        resp.raise_for_status()
        return resp.text

    def _build_list_payload(self, biz_no: str) -> Dict[str, Optional[str]]:
        payload: Dict[str, Optional[str]] = {
            "chks": "",
            "fileType": "",
            "pageIndex": "1",
            "ctprvnNm": "",
            "signguNm": "",
            "cntrctEsntlNo": "",
            "entrpsNm": "",
            "searchBsnmNo": biz_no,
            "chargerNm": "",
            "detailPrdnm": "",
            "detailPrdnmNo": "",
            "ksicNm": "",
            "ksic": "",
            "prductNm": "",
            "ctprvnCode": "",
            "signguCode": "",
            "smbizCode": "",
            "femtrbleCode": "",
            "hitechCode": "",
            "envqualCode": "",
            "entrpsNmMbl": "",
            "searchBsnmNoMbl": "",
            "chargerNmMbl": "",
            "pageUnit": "15",
        }
        return payload


def _build_login_form_payload(html: str, user_id: str, password: str) -> Dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form", attrs={"name": "loginForm"}) or soup.find("form", attrs={"id": "loginForm"})
    if not form:
        raise RuntimeError("loginForm을 찾을 수 없습니다.")

    payload: Dict[str, str] = {}
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        payload[name] = inp.get("value", "") or ""

    payload["id"] = user_id
    payload["password"] = password
    return payload
