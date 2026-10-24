"""M4: join decoded IZ rows to an ERP open items extract on KSeF number.

Mirrored by RunMatch in excel/vba/modMatch.bas - keep statuses identical.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from collections.abc import Sequence
from dataclasses import dataclass

from ksefiz.io_csv import Row, parse_amount
from ksefiz.iz import normalise

TOLERANCE = 0.005

STATUS_MATCHED = "Matched"
STATUS_NO_AMOUNT = "Matched (no amount)"
STATUS_AMOUNT_DIFF = "Amount diff"
STATUS_CURRENCY_DIFF = "Currency diff"
STATUS_UNMATCHED = "Unmatched"
STATUS_DUPLICATE_ERP = "Duplicate in ERP"

MATCH_FIELDS = [
    "iz", "ksef_number", "paid_amount", "paid_currency", "doc_no", "customer",
    "open_amount", "open_currency", "diff", "times_referenced", "total_paid", "status",
]


@dataclass(frozen=True)
class MatchRow:
    iz: str
    ksef_number: str
    paid_amount: float | None
    paid_currency: str
    doc_no: str
    customer: str
    open_amount: float | None
    open_currency: str
    diff: float | None
    times_referenced: int
    total_paid: float | None
    status: str

    def as_dict(self) -> dict[str, str]:
        def f(x: float | None) -> str:
            return "" if x is None else f"{x:.2f}"

        return {
            "iz": self.iz, "ksef_number": self.ksef_number,
            "paid_amount": f(self.paid_amount), "paid_currency": self.paid_currency,
            "doc_no": self.doc_no, "customer": self.customer,
            "open_amount": f(self.open_amount), "open_currency": self.open_currency,
            "diff": f(self.diff), "times_referenced": str(self.times_referenced),
            "total_paid": f(self.total_paid), "status": self.status,
        }


def _key(ksef: str) -> str:
    return normalise(ksef) or ksef.strip().upper()


def match(decoded: Sequence[Row], open_items: Sequence[Row]) -> list[MatchRow]:
    """One output row per decoded row.

    An invoice referenced by several decoded rows (partial payments, several IZ) is
    judged on the total paid across those rows; `diff` is total_paid - open_amount.
    """
    erp: dict[str, list[Row]] = defaultdict(list)
    for oi in open_items:
        erp[_key(oi.get("ksef_number", ""))].append(oi)

    keys = [_key(d.get("ksef_number", "")) for d in decoded]
    refs = Counter(keys)
    totals: dict[str, float | None] = {}
    for k, d in zip(keys, decoded, strict=True):
        a = parse_amount(d.get("amount", ""))
        prev = totals.get(k, 0.0)
        totals[k] = None if a is None or prev is None else prev + a

    out: list[MatchRow] = []
    for k, d in zip(keys, decoded, strict=True):
        paid = parse_amount(d.get("amount", ""))
        cur = d.get("currency", "").strip().upper()
        hits = erp.get(k, [])
        oi = hits[0] if hits else {}
        open_amt = parse_amount(oi.get("amount", "")) if hits else None
        open_cur = oi.get("currency", "").strip().upper()
        total = totals[k]
        diff = None if total is None or open_amt is None else round(total - open_amt, 2)

        if not hits:
            status = STATUS_UNMATCHED
        elif len(hits) > 1:
            status = STATUS_DUPLICATE_ERP
        elif total is None or open_amt is None:
            status = STATUS_NO_AMOUNT
        elif cur and open_cur and cur != open_cur:
            status = STATUS_CURRENCY_DIFF
        elif diff is not None and abs(diff) > TOLERANCE:
            status = STATUS_AMOUNT_DIFF
        else:
            status = STATUS_MATCHED

        out.append(MatchRow(
            iz=d.get("iz", "").strip(), ksef_number=k, paid_amount=paid, paid_currency=cur,
            doc_no=oi.get("doc_no", ""), customer=oi.get("customer", ""),
            open_amount=open_amt, open_currency=open_cur, diff=diff,
            times_referenced=refs[k], total_paid=total, status=status,
        ))
    return out
