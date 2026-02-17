# MAcc Exit Survey 2024 Course Ranking

## Method
- Included only completed responses (`Finished = 1`).
- Reshaped wide survey columns to long format in `cleaned_responses_long.csv`.
- Normalized scores to a 0-100 scale for comparability:
  - Core ranked courses (`Q35_*`): `((9 - rank) / 8) * 100` (rank 1 is best).
  - Elective ratings (`Q76_1` etc.): `((rating - 1) / 4) * 100` (rating 5 is best).
- Overall course score is the mean of all normalized scores for that course.

## Top 5 Courses

| Rank | Course | Overall Score | N |
|---:|---|---:|---:|
| 1 | ACC 6020 Advanced Financial Application | 100.00 | 5 |
| 2 | ACC 679R Taxation of Business Entities II | 89.42 | 26 |
| 3 | ACC 6400 Advanced Tax Business Entities | 82.50 | 55 |
| 4 | ACC 6600 Business Law for Accountants (if taken as an elective) | 81.82 | 11 |
| 5 | ACC 6410 Tax Research & Procedure | 81.25 | 4 |
