# MaccExitSurveyProj

This repository contains a deterministic workflow to analyze **`Grad Program Exit Survey Data 2024.xlsx`** and rank courses based on student preferences/ratings.

## What the workflow does

1. Reads the dataset from the repository.
2. Cleans responses to include only completed submissions (`Finished = 1`).
3. Reshapes ranking/rating columns into a long format.
4. Computes a normalized 0-100 score for each course.
5. Produces a ranked output and two figures (default + UVU-themed variant).
6. Uploads results as GitHub Actions artifacts for grading.

## Run locally

```bash
python scripts/analyze_exit_survey.py \
  --input "Grad Program Exit Survey Data 2024.xlsx" \
  --output-dir outputs
```

## Outputs

- `outputs/course_ranking.csv` — final rank order of courses/programs.
- `outputs/cleaned_responses_long.csv` — cleaned and reshaped analytic dataset.
- `outputs/course_ranking.svg` — original ranking figure.
- `outputs/course_ranking_uvu_theme.svg` — UVU-themed ranking figure with forest-green bars and wider label space.
- `outputs/summary.md` — short explanation + top 5 courses.

## GitHub Actions

Workflow file: `.github/workflows/exit-survey-ranking.yml`

- Triggers on changes to the dataset, analysis script, or workflow file.
- Can also be run manually via **workflow_dispatch**.
- Uploads the outputs above as a single artifact named `exit-survey-analysis`.
