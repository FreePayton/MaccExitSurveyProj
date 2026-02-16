#!/usr/bin/env python3
"""Deterministic analysis of MAcc exit survey rankings without external dependencies."""

from __future__ import annotations

import argparse
import csv
import re
import statistics
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
import xml.etree.ElementTree as ET

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_to_index(col_ref: str) -> int:
    total = 0
    for char in col_ref:
        total = total * 26 + (ord(char) - 64)
    return total - 1


def parse_xlsx(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in shared_root.findall("a:si", NS):
                text = "".join((node.text or "") for node in si.findall(".//a:t", NS))
                shared_strings.append(text)

        sheet_root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

    rows: List[Dict[int, str]] = []
    for row in sheet_root.findall(".//a:sheetData/a:row", NS):
        values: Dict[int, str] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            if not match:
                continue
            col_idx = col_to_index(match.group(1))
            cell_type = cell.attrib.get("t")

            value = ""
            if cell_type == "inlineStr":
                node = cell.find("a:is/a:t", NS)
                if node is not None:
                    value = node.text or ""
            else:
                node = cell.find("a:v", NS)
                if node is not None and node.text is not None:
                    raw = node.text
                    value = shared_strings[int(raw)] if cell_type == "s" else raw

            values[col_idx] = value.strip()
        rows.append(values)

    if len(rows) < 4:
        raise ValueError("Unexpected worksheet structure. Expected metadata rows + data.")

    max_col = max(max(row.keys(), default=0) for row in rows)
    header_row = rows[0]
    question_row = rows[1]

    columns: List[str] = []
    questions: Dict[str, str] = {}
    for idx in range(max_col + 1):
        col_name = header_row.get(idx, f"COL_{idx}")
        columns.append(col_name)
        questions[col_name] = question_row.get(idx, "")

    records: List[Dict[str, str]] = []
    for row in rows[3:]:
        record = {col_name: row.get(idx, "") for idx, col_name in enumerate(columns)}
        records.append(record)

    return columns, records, questions


def parse_course_name(question_text: str) -> str:
    if " - " in question_text:
        return question_text.split(" - ")[-1].strip()
    return question_text.strip() or "Unknown Course"


def clean_numeric(value: str) -> int | None:
    value = value.strip()
    if not value:
        return None
    if re.fullmatch(r"\d+(\.0+)?", value):
        return int(float(value))
    return None


def create_svg(path: Path, ranking: List[Dict[str, str]]) -> None:
    width, row_h, margin = 1100, 42, 170
    bar_max = width - margin - 220
    height = 90 + row_h * len(ranking)

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
        '<style>text { font-family: Arial, sans-serif; fill: #1f2937; } .title { font-size: 20px; font-weight: 700; } .label { font-size: 13px; } .score { font-size: 12px; }</style>',
        '<rect x="0" y="0" width="100%" height="100%" fill="#ffffff"/>',
        '<text x="24" y="34" class="title">MAcc Exit Survey 2024: Course Ranking (Higher = Better)</text>',
    ]

    for i, row in enumerate(ranking):
        y = 62 + i * row_h
        score = float(row["overall_score"])
        bar_w = max(0, min(bar_max, int(bar_max * (score / 100.0))))
        rank = row["rank"]
        label = row["course"]
        lines.append(f'<text x="24" y="{y + 19}" class="label">#{rank}</text>')
        lines.append(f'<text x="56" y="{y + 19}" class="label">{label}</text>')
        lines.append(f'<rect x="{margin}" y="{y}" width="{bar_max}" height="20" fill="#e5e7eb" rx="3"/>')
        lines.append(f'<rect x="{margin}" y="{y}" width="{bar_w}" height="20" fill="#2563eb" rx="3"/>')
        lines.append(f'<text x="{margin + bar_max + 10}" y="{y + 15}" class="score">{score:.1f} (n={row["num_responses"]})</text>')

    lines.append("</svg>")
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze and rank MAcc exit survey courses.")
    parser.add_argument("--input", default="Grad Program Exit Survey Data 2024.xlsx", help="Input XLSX file path")
    parser.add_argument("--output-dir", default="outputs", help="Directory for generated outputs")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    _, rows, questions = parse_xlsx(input_path)

    core_fields = [name for name in questions if name.startswith("Q35_")]
    elective_fields = ["Q76_1", "Q77_2", "Q78_3", "Q83_4", "Q82_5", "Q80_6", "Q81_9", "Q79_7"]

    long_rows: List[Dict[str, str]] = []
    course_scores: Dict[str, Dict[str, List[float]]] = defaultdict(lambda: {"core": [], "elective": []})

    for row in rows:
        is_finished = row.get("Finished", "").strip() in {"1", "true", "TRUE", "True"}
        if not is_finished:
            continue
        respondent_id = row.get("ResponseId", "")

        for field in core_fields:
            raw = row.get(field, "")
            rank = clean_numeric(raw)
            if rank is None or rank < 1 or rank > 8:
                continue
            course = parse_course_name(questions.get(field, ""))
            norm_score = ((9 - rank) / 8.0) * 100.0
            course_scores[course]["core"].append(norm_score)
            long_rows.append(
                {
                    "response_id": respondent_id,
                    "course": course,
                    "source_type": "core_rank",
                    "response_value": str(rank),
                    "normalized_score": f"{norm_score:.6f}",
                }
            )

        for field in elective_fields:
            raw = row.get(field, "")
            rating = clean_numeric(raw)
            if rating is None or rating < 1 or rating > 5:
                continue
            course = parse_course_name(questions.get(field, ""))
            norm_score = ((rating - 1) / 4.0) * 100.0
            course_scores[course]["elective"].append(norm_score)
            long_rows.append(
                {
                    "response_id": respondent_id,
                    "course": course,
                    "source_type": "elective_rating",
                    "response_value": str(rating),
                    "normalized_score": f"{norm_score:.6f}",
                }
            )

    ranking_rows = []
    for course, sources in course_scores.items():
        core_vals = sources["core"]
        elec_vals = sources["elective"]
        all_vals = core_vals + elec_vals
        if not all_vals:
            continue
        ranking_rows.append(
            {
                "course": course,
                "overall_score": statistics.fmean(all_vals),
                "num_responses": len(all_vals),
                "core_pref_score": statistics.fmean(core_vals) if core_vals else None,
                "core_n": len(core_vals),
                "elective_rating_score": statistics.fmean(elec_vals) if elec_vals else None,
                "elective_n": len(elec_vals),
            }
        )

    ranking_rows.sort(key=lambda r: (-r["overall_score"], -r["num_responses"], r["course"]))

    for idx, row in enumerate(ranking_rows, start=1):
        row["rank"] = idx

    long_csv = output_dir / "cleaned_responses_long.csv"
    with long_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["response_id", "course", "source_type", "response_value", "normalized_score"],
        )
        writer.writeheader()
        writer.writerows(long_rows)

    ranking_csv = output_dir / "course_ranking.csv"
    with ranking_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "rank",
                "course",
                "overall_score",
                "num_responses",
                "core_pref_score",
                "core_n",
                "elective_rating_score",
                "elective_n",
            ]
        )
        for row in ranking_rows:
            writer.writerow(
                [
                    row["rank"],
                    row["course"],
                    f"{row['overall_score']:.6f}",
                    row["num_responses"],
                    "" if row["core_pref_score"] is None else f"{row['core_pref_score']:.6f}",
                    row["core_n"],
                    "" if row["elective_rating_score"] is None else f"{row['elective_rating_score']:.6f}",
                    row["elective_n"],
                ]
            )

    figure_path = output_dir / "course_ranking.svg"
    create_svg(figure_path, [
        {
            "rank": str(row["rank"]),
            "course": row["course"],
            "overall_score": f"{row['overall_score']:.6f}",
            "num_responses": str(row["num_responses"]),
        }
        for row in ranking_rows
    ])

    summary = output_dir / "summary.md"
    top_five = ranking_rows[:5]
    with summary.open("w", encoding="utf-8") as f:
        f.write("# MAcc Exit Survey 2024 Course Ranking\n\n")
        f.write("## Method\n")
        f.write("- Included only completed responses (`Finished = 1`).\n")
        f.write("- Reshaped wide survey columns to long format in `cleaned_responses_long.csv`.\n")
        f.write("- Normalized scores to a 0-100 scale for comparability:\n")
        f.write("  - Core ranked courses (`Q35_*`): `((9 - rank) / 8) * 100` (rank 1 is best).\n")
        f.write("  - Elective ratings (`Q76_1` etc.): `((rating - 1) / 4) * 100` (rating 5 is best).\n")
        f.write("- Overall course score is the mean of all normalized scores for that course.\n\n")
        f.write("## Top 5 Courses\n\n")
        f.write("| Rank | Course | Overall Score | N |\n")
        f.write("|---:|---|---:|---:|\n")
        for row in top_five:
            f.write(f"| {row['rank']} | {row['course']} | {row['overall_score']:.2f} | {row['num_responses']} |\n")


if __name__ == "__main__":
    main()
