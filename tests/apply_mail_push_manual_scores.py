"""Apply manual sorting scores to a mail push report.

This is a post-processing step for `tests/run_mail_push_suite.py`.

Usage:
  python tests/apply_mail_push_manual_scores.py \
    --report tests/reports/mail_push_report_20260208_210813.json \
    --scores tests/reports/mail_push_sorting_scores_20260208_210813.json

If not provided, defaults to the stable latest report files.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict


def _load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_json(path: Path, data: Dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _compute_final_score_100(sc: Dict[str, Any]) -> float:
    # weights: json 0.2, structure 0.2, relevance 0.2, sorting 0.2, tag 0.1, summary 0.1
    def f(k: str) -> float:
        v = sc.get(k)
        try:
            return float(v)
        except Exception:
            return 0.0

    return (
        f("json") * 0.2
        + f("structure") * 0.2
        + f("relevance") * 0.2
        + f("sorting_manual") * 0.2
        + f("tag") * 0.1
        + f("summary") * 0.1
    ) * 10.0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--report",
        default=str(Path(__file__).resolve().parent / "reports" / "mail_push_report.json"),
        help="Input report JSON path",
    )
    ap.add_argument(
        "--scores",
        default="",
        help="Manual scores JSON path (contains scores mapping)",
    )
    args = ap.parse_args()

    report_path = Path(args.report)
    report = _load_json(report_path)

    if args.scores:
        scores_path = Path(args.scores)
    else:
        raise SystemExit("--scores is required (manual sorting scores mapping)")

    manual = _load_json(scores_path)
    mapping = manual.get("scores") or {}
    if not isinstance(mapping, dict) or not mapping:
        raise SystemExit("scores file missing non-empty 'scores' mapping")

    threshold = int((report.get("suite") or {}).get("pass_threshold") or 70)

    passed = 0
    failed = 0
    errors = 0
    pending = 0
    durations = []

    for r in report.get("results") or []:
        status = r.get("status")
        if status == "ERROR":
            errors += 1
            continue

        case_id = str(r.get("id") or "")
        sc = r.get("scores") or {}

        if case_id in mapping:
            sc["sorting_manual"] = int(mapping[case_id])
            r["scores"] = sc
            final = round(_compute_final_score_100(sc), 1)
            r["final_score_100"] = final
            ok = bool(final >= threshold and float(sc.get("json") or 0) > 0)
            r["passed"] = ok
            r["status"] = "PASS" if ok else "FAIL"
            r["status_cn"] = "✅通过" if ok else "❌失败"
        else:
            pending += 1
            r["passed"] = None
            r["final_score_100"] = None
            r["status"] = "PENDING"
            r["status_cn"] = "⏳待评(排序)"

        if r.get("status") == "PASS":
            passed += 1
        elif r.get("status") == "FAIL":
            failed += 1

        try:
            durations.append(float(r.get("duration_s") or 0.0))
        except Exception:
            pass

    total = len(report.get("results") or [])
    denom = max(passed + failed, 1)
    pass_rate = f"{round((passed / denom) * 100, 1)}%"
    avg_duration_s = round(sum(durations) / max(len(durations), 1), 3)

    tech_overlap = [
        r.get("tech_ai_overlap")
        for r in report.get("results")
        if isinstance(r.get("tech_ai_overlap"), float)
    ]
    tech_ai_overlap_avg = (
        round(sum(tech_overlap) / max(len(tech_overlap), 1), 3) if tech_overlap else None
    )

    report["summary"] = {
        "total": total,
        "passed": passed,
        "failed": failed,
        "errors": errors,
        "pending": pending,
        "pass_rate": pass_rate,
        "avg_duration_s": avg_duration_s,
        "tech_ai_overlap_avg": ("{:.0%}".format(tech_ai_overlap_avg) if isinstance(tech_ai_overlap_avg, float) else ""),
    }

    # Write a scored copy next to the input report.
    scored_json = report_path.with_name(report_path.stem + "_scored.json")
    _save_json(scored_json, report)

    # Also refresh stable latest artifacts.
    latest_json = report_path.parent / "mail_push_report.json"
    _save_json(latest_json, report)

    # Re-render MD/XLSX using the suite renderer.
    from run_mail_push_suite import _make_excel, _render_md

    md = _render_md(report)
    scored_md = report_path.with_name(report_path.stem + "_scored.md")
    scored_md.write_text(md, encoding="utf-8")
    (report_path.parent / "mail_push_report.md").write_text(md, encoding="utf-8")

    scored_xlsx = report_path.with_name(report_path.stem + "_scored.xlsx")
    _make_excel(report, xlsx_path=scored_xlsx)
    _make_excel(report, xlsx_path=(report_path.parent / "mail_push_test_data.xlsx"))

    print(f"Scored report JSON: {scored_json}")
    print(f"Scored report MD: {scored_md}")
    print(f"Scored report XLSX: {scored_xlsx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
