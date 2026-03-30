"""Business logic that ties SMPP HTTP calls together."""

from __future__ import annotations

from dataclasses import dataclass
from time import sleep
from typing import Callable, Iterable, List, Optional

from smpp_client import SmppClient, SmppCredentials
from smpp_parser import (
    CorpFeatures,
    build_move_form_payload_from_list_html,
    parse_corp_features,
)

ProgressCallback = Callable[[int, int, str], None]


@dataclass
class CorpCheckResult:
    biz_no: str
    features: Optional[CorpFeatures]
    payload: Optional[dict] = None
    error: Optional[str] = None

    def validity_string(self, women_only: bool = False) -> Optional[str]:
        """Return `YYYY.MM.DD~YYYY.MM.DD` format for the selected corp type."""
        if not self.features:
            return None
        if women_only:
            return format_period(
                self.features.women_confirm_date,
                self.features.women_expire_date,
            )
        return format_period(
            self.features.small_confirm_date,
            self.features.small_expire_date,
        )


def check_corps(
    user_id: str,
    password: str,
    biz_nos: Iterable[str],
    *,
    delay_seconds: float = 0.0,
    progress_callback: Optional[ProgressCallback] = None,
) -> List[CorpCheckResult]:
    client = SmppClient(SmppCredentials(user_id=user_id, password=password))
    client.login()

    normalized_biz_nos = [biz_no.strip() for biz_no in biz_nos if biz_no.strip()]
    total = len(normalized_biz_nos)
    results: List[CorpCheckResult] = []

    for idx, biz_no in enumerate(normalized_biz_nos, start=1):
        if progress_callback:
            progress_callback(idx, total, biz_no)

        try:
            list_html = client.fetch_list_html_by_biz_no(biz_no)
            payload = build_move_form_payload_from_list_html(list_html, biz_no)
            summary_html = client.fetch_summary_html(payload)
            features = parse_corp_features(summary_html)
            results.append(
                CorpCheckResult(
                    biz_no=biz_no,
                    features=features,
                    payload=payload,
                )
            )
        except Exception as exc:  # pylint: disable=broad-except
            results.append(
                CorpCheckResult(
                    biz_no=biz_no,
                    features=None,
                    payload=None,
                    error=str(exc),
                )
            )

        if delay_seconds > 0:
            sleep(delay_seconds)

    return results


def format_period(
    confirm_date: Optional[str],
    expire_date: Optional[str],
) -> Optional[str]:
    if not confirm_date and not expire_date:
        return None
    if confirm_date and expire_date:
        return f"{confirm_date}~{expire_date}"
    return confirm_date or expire_date
