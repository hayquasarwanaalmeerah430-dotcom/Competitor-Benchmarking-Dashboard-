# Competitor Benchmarking Dashboard — Virtual Internship Programs

For this task I benchmarked three competitor virtual internship programs against each other — what they charge, what they actually offer, and which one gives students the most for their money. Instead of reading the spreadsheet and writing opinions, I built a script that does the whole analysis and turns it into a dashboard.

The script is `competitor_analysis.py`. It reads the workbook that came with the task (`competitor_benchmarking.xlsx`, four sheets of competitor data), cleans it up, works out the comparison numbers, draws six charts, and builds `dashboard.html` — a single page you can open in any browser. Everything in the results folder came out of that one run: the cleaned dataset (`cleaned_data.csv`), the computed insights (`insights.json`), the six charts and the dashboard itself.

## The three competitors

| | SkillBridge Interns | CodeNest Virtual | TechPulse Academy |
|---|---|---|---|
| Total cost | **PKR 1,500** | PKR 3,000 | PKR 5,500 |
| How the price works | program is free, certificate costs 1,500 | one payment | one payment |
| Duration | 4 weeks | 6 weeks | 8 weeks |
| Cost per week | PKR 375 | PKR 500 | PKR 688 |
| Internship fields | 8 | 6 | 10 |
| Live sessions | ✗ | ✓ | ✓ |
| Rating | 4.2 | 4.5 | 4.6 |
| Social followers | 27,700 | 20,200 | 40,000 |

## What the numbers say

Prices run from PKR 1,500 to PKR 5,500 — a 3.7× spread, with the average program costing PKR 3,333. But the cheapest headline price turned out to be the worst value: when I combined price, duration, fields, features and rating into one value-for-money score, the order flipped. TechPulse — the most expensive — scored **100** because you get the longest program, the most fields, live sessions and a stipend for top performers. CodeNest scored 77, and SkillBridge only 37: its program is free, but it's the shortest, has no live sessions, and the certificate is what you're actually paying for.

Each chart shows one piece of this: `chart_pricing.png` compares total cost against cost per week, `chart_offerings.png` covers fields and duration, `chart_value.png` is the value-for-money ranking, `chart_unit_economics.png` breaks the weekly cost down, `chart_rating.png` the ratings, and `chart_engagement.png` the social following — where TechPulse leads again with 40,000 followers.

An honest note on the data: the figures come from the workbook provided with the task, and the workbook itself labels them as sample data. So this edition demonstrates the benchmarking method on that dataset. I didn't stop there, though — see the last section.

## Requirements

The libraries the script needs are listed in **`requirements.txt`**, so the whole setup installs in one command from the project folder:

```bash
python -m pip install -r requirements.txt
```

- **pandas** — reads the workbook and does the cleaning and comparisons
- **numpy** — the number work underneath pandas
- **matplotlib** — draws the six charts
- **openpyxl** — the piece that lets pandas open `.xlsx` files at all. I learned this one the practical way: without it, the script stops at the very first sheet with `ImportError: Import openpyxl failed`.

## Run it yourself

```bash
python competitor_analysis.py
```

It finds the workbook on its own (any Excel file with "competitor" in the name, next to the script), then prints what it's doing as it goes and writes `cleaned_data.csv`, `insights.json`, the six charts and `dashboard.html`.

## Files in this project

- `competitor_analysis.py` — the whole pipeline, one file
- `competitor_benchmarking.xlsx` — the source workbook the task provided
- `cleaned_data.csv` — the workbook's four sheets, cleaned and merged
- `insights.json` — every KPI the script computed
- `dashboard.html` — the visual dashboard, charts embedded, open it in a browser
- `chart_pricing.png`, `chart_offerings.png`, `chart_value.png`, `chart_unit_economics.png`, `chart_rating.png`, `chart_engagement.png` — the six charts
- `requirements.txt` — the libraries needed, installable in one command (see above)
- `REPORT.md` — the full written analysis

## Data

Competitor benchmarking workbook (Pricing + Engagement sheets),collected from competitor websites. The workbook itself labels the figures as fictional sample data for illustration.

