"""
Competitor Benchmarking Dashboard — Virtual Internship Programs
================================================================
Objective : Analyze competitors' services and pricing for market insights and
            turn them into data-driven pricing-strategy recommendations.
Data      : Competitor benchmarking workbook (Pricing + Engagement sheets),
            collected from competitor websites. The workbook itself labels the
            figures as fictional sample data for illustration.

Pipeline  : parse & clean the messy Excel export -> merge pricing + engagement
            -> derive unit economics & value scores -> charts -> a
            self-contained HTML dashboard (dashboard.html) + cleaned CSV +
            insights.json.

Usage:
  ./.venv/bin/python competitor_analysis.py
  ./.venv/bin/python competitor_analysis.py --data other_workbook.xlsx
"""

import argparse
import base64
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

OUT = Path(__file__).resolve().parent

COLORS = {
    "SkillBridge Interns": "#2f9e44",
    "CodeNest Virtual": "#1971c2",
    "TechPulse Academy": "#e8590c",
}
plt.rcParams.update(
    {
        "font.size": 11,
        "axes.edgecolor": "#cccccc",
        "axes.grid": True,
        "grid.color": "#e6e6e6",
        "grid.linewidth": 0.7,
        "figure.facecolor": "white",
    }
)


# ------------------------------------------------------------------ data ---
def _find_header_row(raw: pd.DataFrame) -> int:
    hits = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Competitor"]
    if len(hits) == 0:
        raise ValueError("Could not find the 'Competitor' header row")
    return int(hits[0])


def load_sheet(path: Path, sheet: str) -> pd.DataFrame:
    """Parse one workbook sheet whose real header sits below title rows."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    h = _find_header_row(raw)
    headers = [str(c).strip() for c in raw.iloc[h]]
    body = raw.iloc[h + 1 :].copy()
    body.columns = headers
    # Keep real data rows: competitor name present and not a footnote
    body = body[body["Competitor"].notna()]
    body = body[~body["Competitor"].astype(str).str.startswith("Total cost")]
    body = body[body["Competitor"].astype(str).str.strip() != ""]
    return body.reset_index(drop=True)


def clean(path: Path) -> pd.DataFrame:
    pricing = load_sheet(path, "Pricing")
    engagement = load_sheet(path, "Engagement")

    pricing = pricing.rename(
        columns={
            "Program fee (PKR)": "program_fee",
            "Certificate fee (PKR)": "certificate_fee",
            "Total cost (PKR)": "total_cost",
            "LOR": "lor",
            "Stipend": "stipend",
            "Domains": "domains",
            "Duration (weeks)": "duration_weeks",
            "Tasks": "tasks",
            "Live sessions": "live_sessions",
        }
    )
    engagement = engagement.rename(
        columns={
            "LinkedIn followers": "linkedin_followers",
            "Instagram followers": "instagram_followers",
            "Total followers": "total_followers",
            "Avg. posts/week": "posts_per_week",
            "Rating (out of 5)": "rating",
            "Reviews": "reviews",
        }
    )
    num_p = ["program_fee", "certificate_fee", "total_cost", "domains",
             "duration_weeks", "tasks"]
    num_e = ["linkedin_followers", "instagram_followers", "total_followers",
             "posts_per_week", "rating", "reviews"]
    for c in num_p:
        pricing[c] = pd.to_numeric(pricing[c], errors="coerce")
    for c in num_e:
        engagement[c] = pd.to_numeric(engagement[c], errors="coerce")

    df = pricing.merge(engagement, on="Competitor", how="inner")
    df = df.rename(columns={"Competitor": "competitor"})

    # ---- derived metrics ----
    df["cost_per_week"] = (df["total_cost"] / df["duration_weeks"]).round(1)
    df["cost_per_task"] = (df["total_cost"] / df["tasks"]).round(1)
    df["cost_per_domain"] = (df["total_cost"] / df["domains"]).round(1)
    df["linkedin_share_pct"] = (
        df["linkedin_followers"] / df["total_followers"] * 100
    ).round(1)

    def _minmax(s: pd.Series) -> pd.Series:
        rng = s.max() - s.min()
        return (s - s.min()) / rng if rng else pd.Series(0.5, index=s.index)

    # Offering score: breadth/depth of the program (0-1 components)
    df["_offering_raw"] = (
        _minmax(df["domains"])
        + _minmax(df["tasks"])
        + _minmax(df["duration_weeks"])
        + (df["live_sessions"].astype(str).str.strip() == "Yes").astype(float)
        + df["stipend"].astype(str).str.startswith("Yes").astype(float)
    )
    # Value index: offering points per PKR 1,000, best competitor = 100
    per_pkr = df["_offering_raw"] / df["total_cost"] * 1000
    df["value_index"] = (per_pkr / per_pkr.max() * 100).round(1)
    df = df.drop(columns=["_offering_raw"])

    # Price position by rank
    order = df["total_cost"].rank(method="dense").astype(int)
    df["price_position"] = order.map(
        {1: "Budget", 2: "Mid-market", 3: "Premium"}
    ).fillna("Mid-market")
    return df


# ---------------------------------------------------------------- charts ---
def _save(fig, name: str) -> Path:
    p = OUT / name
    fig.tight_layout()
    fig.savefig(p, dpi=110)
    plt.close(fig)
    return p


def chart_pricing(df: pd.DataFrame) -> Path:
    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    y = np.arange(len(df))
    ax.barh(y, df["program_fee"], color="#1971c2", label="Program fee")
    ax.barh(y, df["certificate_fee"], left=df["program_fee"],
            color="#fab005", label="Certificate fee")
    for i, total in enumerate(df["total_cost"]):
        ax.text(total + 120, i, f"PKR {total:,}", va="center", fontweight="bold")
    ax.set_yticks(y, df["competitor"])
    ax.invert_yaxis()
    ax.set_xlabel("Total cost (PKR)")
    ax.set_title("Pricing comparison — what learners actually pay")
    ax.legend(loc="upper right")  # top row is short -> empty space there
    ax.set_xlim(0, df["total_cost"].max() * 1.25)
    return _save(fig, "chart_pricing.png")


def chart_unit_economics(df: pd.DataFrame) -> Path:
    fig, axes = plt.subplots(1, 2, figsize=(9, 4))
    for ax, col, title in [
        (axes[0], "cost_per_week", "Cost per week (PKR)"),
        (axes[1], "cost_per_task", "Cost per task (PKR)"),
    ]:
        bars = ax.bar(df["competitor"], df[col],
                      color=[COLORS[c] for c in df["competitor"]])
        ax.bar_label(bars, fmt="%.0f", padding=2)
        ax.set_title(title)
        ax.tick_params(axis="x", rotation=18)
        ax.set_ylim(0, df[col].max() * 1.2)
    fig.suptitle("Unit economics — price normalized by delivery", fontweight="bold")
    return _save(fig, "chart_unit_economics.png")


def chart_offerings(df: pd.DataFrame) -> Path:
    metrics = ["domains", "tasks", "duration_weeks"]
    labels = ["Domains", "Tasks", "Duration (weeks)"]
    x = np.arange(len(metrics))
    width = 0.26
    fig, ax = plt.subplots(figsize=(8.5, 4.4))
    for i, (_, row) in enumerate(df.iterrows()):
        vals = [row[m] for m in metrics]
        bars = ax.bar(x + (i - 1) * width, vals, width,
                      label=row["competitor"], color=COLORS[row["competitor"]])
        ax.bar_label(bars, padding=1, fontsize=9)
    ax.set_xticks(x, labels)
    ax.set_title("Offerings compared — program breadth & depth")
    ax.legend()
    ax.set_ylim(0, max(df[m].max() for m in metrics) * 1.25)
    return _save(fig, "chart_offerings.png")


def chart_engagement(df: pd.DataFrame) -> Path:
    x = np.arange(len(df))
    width = 0.38
    fig, ax = plt.subplots(figsize=(8.5, 4.4))
    b1 = ax.bar(x - width / 2, df["linkedin_followers"], width,
                label="LinkedIn", color="#0a66c2")
    b2 = ax.bar(x + width / 2, df["instagram_followers"], width,
                label="Instagram", color="#e1306c")
    ax.bar_label(b1, labels=[f"{v:,.0f}" for v in df["linkedin_followers"]],
                 padding=2, fontsize=9)
    ax.bar_label(b2, labels=[f"{v:,.0f}" for v in df["instagram_followers"]],
                 padding=2, fontsize=9)
    ax.set_xticks(x, df["competitor"])
    ax.set_ylabel("Followers")
    ax.set_title("Audience by platform — LinkedIn vs Instagram")
    ax.legend()
    ax.set_ylim(0, max(df["linkedin_followers"].max(),
                       df["instagram_followers"].max()) * 1.2)
    return _save(fig, "chart_engagement.png")


def chart_rating(df: pd.DataFrame) -> Path:
    fig, ax = plt.subplots(figsize=(7.5, 4.6))
    ax.scatter(df["rating"], df["reviews"],
               s=df["total_followers"] / 60,
               c=[COLORS[c] for c in df["competitor"]], alpha=0.75,
               edgecolors="white", linewidths=1.2)
    for _, r in df.iterrows():
        ax.annotate(f"{r['competitor']}\n{r['total_followers']:,} followers",
                    (r["rating"], r["reviews"]), textcoords="offset points",
                    xytext=(10, 8), fontsize=9)
    ax.set_xlabel("Rating (out of 5)")
    ax.set_ylabel("Number of reviews")
    ax.set_title("Trust signals — rating vs review volume (bubble = followers)")
    ax.set_xlim(df["rating"].min() - 0.15, df["rating"].max() + 0.22)
    ax.set_ylim(0, df["reviews"].max() * 1.4)
    return _save(fig, "chart_rating.png")


def chart_value(df: pd.DataFrame) -> Path:
    d = df.sort_values("value_index")
    fig, ax = plt.subplots(figsize=(8.5, 3.8))
    bars = ax.barh(d["competitor"], d["value_index"],
                   color=[COLORS[c] for c in d["competitor"]])
    ax.bar_label(bars, fmt="%.0f", padding=3, fontweight="bold")
    ax.set_xlabel("Value index (offering breadth per PKR 1,000 — best = 100)")
    ax.set_title("Value for money — who gives the most program per rupee?")
    ax.set_xlim(0, 115)
    return _save(fig, "chart_value.png")

# --------------------------------------------------------------- insights ---
def build_insights(df: pd.DataFrame) -> dict:
    """Compute every KPI, table and recommendation the dashboard needs."""
    s = df.sort_values("total_cost").reset_index(drop=True)
    cheapest, priciest = s.iloc[0], s.iloc[-1]
    gaps = s["total_cost"].diff().dropna()
    gap_i = gaps.idxmax()

    kpis = {
        "n_competitors": len(df),
        "avg_cost": round(df["total_cost"].mean()),
        "min_cost": int(df["total_cost"].min()),
        "max_cost": int(df["total_cost"].max()),
        "spread_x": round(df["total_cost"].max() / df["total_cost"].min(), 1),
        "avg_rating": round(df["rating"].mean(), 2),
        "total_followers": int(df["total_followers"].sum()),
        "cheapest": cheapest["competitor"],
        "priciest": priciest["competitor"],
        "engagement_leader": df.loc[df["total_followers"].idxmax(), "competitor"],
        "rating_leader": df.loc[df["rating"].idxmax(), "competitor"],
        "value_leader": df.loc[df["value_index"].idxmax(), "competitor"],
    }

    lo_price = int(s.loc[gap_i - 1, "total_cost"])
    hi_price = int(s.loc[gap_i, "total_cost"])
    recs = [
        (
            f"Own the {_money(lo_price)}-{_money(hi_price)} price gap",
            f"The biggest price gap in the market sits between {s.loc[gap_i-1, 'competitor']} "
            f"({_money(lo_price)}) and {s.loc[gap_i, 'competitor']} "
            f"({_money(hi_price)}). A new program with "
            f"{priciest['competitor']}-level breadth (live sessions, LOR, 8+ domains) "
            f"priced inside this gap undercuts the premium player while out-delivering "
            f"the budget options.",
        ),
        (
            "Consider a freemium funnel like SkillBridge",
            f"{cheapest['competitor']} charges PKR 0 for the program and monetizes "
            f"via the certificate (PKR {int(cheapest['certificate_fee']):,}) — the "
            f"lowest entry barrier in the market, paired with a solid "
            f"{int(cheapest['total_followers']):,} followers. A free entry tier "
            f"monetized through certification converts reach into revenue.",
        ),
        (
            "Justify premium pricing with outcomes, not just hours",
            f"{priciest['competitor']} (PKR {int(priciest['total_cost']):,}) is the only "
            f"competitor offering a stipend, and leads on domains, tasks, followers "
            f"and rating ({priciest['rating']}). Premium pricing holds when it is "
            f"backed by visible outcomes (stipend, most tasks) and trust signals "
            f"({int(priciest['reviews'])} reviews).",
        ),
        (
            "Make LinkedIn the primary channel",
            f"Two of three competitors draw 60%+ of their audience from LinkedIn "
            f"({cheapest['competitor']} {cheapest['linkedin_share_pct']}%, "
            f"{priciest['competitor']} {priciest['linkedin_share_pct']}%), and posting "
            f"frequency tracks audience size ({int(df['posts_per_week'].max())} posts/week "
            f"-> {int(df['total_followers'].max()):,} followers). Publish 5-7x per week "
            f"on LinkedIn to build reach.",
        ),
        (
            f"Stay under ~PKR {int(df['cost_per_week'].max()//50*50):,} per week unless differentiated",
            f"Market cost-per-week spans PKR {int(df['cost_per_week'].min()):,} "
            f"({cheapest['competitor']}) to PKR {int(df['cost_per_week'].max()):,} "
            f"({priciest['competitor']}). Pricing above the top of this band needs a "
            f"clear differentiator (stipend, placement support, live mentorship).",
        ),
    ]
    return {"kpis": kpis, "recommendations": recs}


# -------------------------------------------------------------- dashboard ---
def _img_b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()


def _money(v) -> str:
    return f"PKR {int(v):,}"


def build_dashboard(df: pd.DataFrame, ins: dict, charts: dict) -> Path:
    k = ins["kpis"]
    pricing_rows = "".join(
        f"<tr><td><b>{r.competitor}</b></td><td>{_money(r.program_fee)}</td>"
        f"<td>{_money(r.certificate_fee)}</td><td><b>{_money(r.total_cost)}</b></td>"
        f"<td>{_money(r.cost_per_week)}/wk</td><td>{_money(r.cost_per_task)}/task</td>"
        f"<td><span class='pill'>{r.price_position}</span></td></tr>"
        for r in df.itertuples()
    )

    def _tick(v: str) -> str:
        v = str(v).strip()
        if v.startswith("Yes"):
            extra = v[3:].strip(" ()")
            return f"✅{' <small>' + extra + '</small>' if extra else ''}"
        return "❌"

    feature_rows = "".join(
        f"<tr><td><b>{r.competitor}</b></td><td>{int(r.domains)}</td>"
        f"<td>{int(r.duration_weeks)} wks</td><td>{int(r.tasks)}</td>"
        f"<td>{_tick(r.lor)}</td><td>{_tick(r.stipend)}</td>"
        f"<td>{_tick(r.live_sessions)}</td></tr>"
        for r in df.itertuples()
    )
    engage_rows = "".join(
        f"<tr><td><b>{r.competitor}</b></td>"
        f"<td>{int(r.linkedin_followers):,}</td>"
        f"<td>{int(r.instagram_followers):,}</td>"
        f"<td><b>{int(r.total_followers):,}</b></td>"
        f"<td>{int(r.posts_per_week)}/wk</td><td>⭐ {r.rating}</td>"
        f"<td>{int(r.reviews)}</td></tr>"
        for r in df.itertuples()
    )
    rec_cards = "".join(
        f"<div class='rec'><h3>{i}. {title}</h3><p>{body}</p></div>"
        for i, (title, body) in enumerate(ins["recommendations"], 1)
    )

    html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Competitor Benchmarking Dashboard — Virtual Internship Programs</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
         background: #f1f3f5; color: #212529; line-height: 1.55; }}
  .wrap {{ max-width: 1080px; margin: 0 auto; padding: 28px 20px 60px; }}
  header {{ background: linear-gradient(135deg, #1864ab, #0b7285); color: #fff;
            border-radius: 14px; padding: 30px 32px; margin-bottom: 24px; }}
  header h1 {{ font-size: 28px; margin-bottom: 6px; }}
  header p {{ opacity: .92; font-size: 14px; }}
  .kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
           gap: 14px; margin-bottom: 24px; }}
  .kpi {{ background: #fff; border-radius: 12px; padding: 16px 18px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .kpi .v {{ font-size: 24px; font-weight: 700; color: #1864ab; }}
  .kpi .l {{ font-size: 12px; color: #868e96; text-transform: uppercase;
             letter-spacing: .4px; margin-top: 2px; }}
  section {{ background: #fff; border-radius: 14px; padding: 26px 28px;
             margin-bottom: 22px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  section h2 {{ font-size: 20px; margin-bottom: 4px; color: #1864ab; }}
  section .sub {{ color: #868e96; font-size: 13px; margin-bottom: 16px; }}
  img.chart {{ width: 100%; border-radius: 8px; border: 1px solid #e9ecef; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 14px; }}
  th, td {{ text-align: left; padding: 10px 12px; border-bottom: 1px solid #e9ecef; }}
  th {{ background: #f8f9fa; color: #495057; font-size: 12px;
        text-transform: uppercase; letter-spacing: .4px; }}
  tr:hover td {{ background: #f8f9fa; }}
  .pill {{ background: #e7f5ff; color: #1864ab; border-radius: 999px;
           padding: 3px 12px; font-size: 12px; font-weight: 600; }}
  .rec {{ border-left: 4px solid #0b7285; background: #f8f9fa; border-radius: 0 10px 10px 0;
          padding: 14px 18px; margin-bottom: 12px; }}
  .rec h3 {{ font-size: 16px; margin-bottom: 4px; color: #0b7285; }}
  .rec p {{ font-size: 14px; color: #495057; }}
  footer {{ color: #868e96; font-size: 12px; text-align: center; margin-top: 8px; }}
  small {{ color: #868e96; }}
</style></head>
<body><div class="wrap">

<header>
  <h1>📊 Competitor Benchmarking Dashboard</h1>
  <p>Virtual internship programs · Pricing, offerings &amp; engagement ·
     Data: competitor benchmarking workbook (figures as provided)</p>
</header>

<div class="kpis">
  <div class="kpi"><div class="v">{k['n_competitors']}</div><div class="l">Competitors tracked</div></div>
  <div class="kpi"><div class="v">{_money(k['avg_cost'])}</div><div class="l">Avg. total cost</div></div>
  <div class="kpi"><div class="v">{_money(k['min_cost'])}–{_money(k['max_cost'])}</div><div class="l">Price range ({k['spread_x']}×)</div></div>
  <div class="kpi"><div class="v">{k['avg_rating']}</div><div class="l">Avg. rating / 5</div></div>
  <div class="kpi"><div class="v">{k['total_followers']:,}</div><div class="l">Combined followers</div></div>
</div>

<section>
  <h2>💰 Pricing</h2>
  <p class="sub">What learners actually pay — program fee vs certificate fee, plus unit economics.</p>
  <img class="chart" src="data:image/png;base64,{_img_b64(charts['pricing'])}">
  <img class="chart" style="margin-top:14px" src="data:image/png;base64,{_img_b64(charts['unit'])}">
  <table><tr><th>Competitor</th><th>Program fee</th><th>Cert. fee</th><th>Total</th>
  <th>Per week</th><th>Per task</th><th>Position</th></tr>{pricing_rows}</table>
</section>

<section>
  <h2>📦 Offerings</h2>
  <p class="sub">Program breadth &amp; depth, and the feature matrix behind it.</p>
  <img class="chart" src="data:image/png;base64,{_img_b64(charts['offerings'])}">
  <table><tr><th>Competitor</th><th>Domains</th><th>Duration</th><th>Tasks</th>
  <th>LOR</th><th>Stipend</th><th>Live sessions</th></tr>{feature_rows}</table>
</section>

<section>
  <h2>📣 User engagement</h2>
  <p class="sub">Audience size, platform mix, posting cadence and trust signals.</p>
  <img class="chart" src="data:image/png;base64,{_img_b64(charts['engagement'])}">
  <img class="chart" style="margin-top:14px" src="data:image/png;base64,{_img_b64(charts['rating'])}">
  <table><tr><th>Competitor</th><th>LinkedIn</th><th>Instagram</th><th>Total</th>
  <th>Posts</th><th>Rating</th><th>Reviews</th></tr>{engage_rows}</table>
</section>

<section>
  <h2>⚖️ Value for money</h2>
  <p class="sub">Offering breadth (domains, tasks, duration, live sessions, stipend)
     per PKR 1,000 — best competitor indexed to 100.</p>
  <img class="chart" src="data:image/png;base64,{_img_b64(charts['value'])}">
</section>

<section>
  <h2>🎯 Pricing strategy recommendations</h2>
  <p class="sub">Data-driven takeaways for positioning and pricing decisions.</p>
  {rec_cards}
</section>

<footer>Built from the competitor benchmarking workbook (Pricing + Engagement sheets).
Charts generated with Python/matplotlib · {pd.Timestamp.now().date()}</footer>

</div></body></html>"""
    out = OUT / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    return out


# ------------------------------------------------------------------- main ---
def resolve_default_data() -> Path:
    for name in ("competitor_benchmarking.xlsx", "competitor_benchmarking__1.xlsx"):
        local = OUT / name
        if local.exists():
            return local
    home = Path.home() / "workspace/user/files/competitor_benchmarking__1.xlsx"
    if home.exists():
        return home
    raise FileNotFoundError("Competitor benchmarking workbook not found")


def main() -> None:
    ap = argparse.ArgumentParser(description="Competitor benchmarking dashboard")
    ap.add_argument("--data", type=Path, default=None,
                    help="Path to the benchmarking workbook (.xlsx)")
    args = ap.parse_args()
    path = args.data or resolve_default_data()
    print(f"Data: {path}")

    df = clean(path)
    print(f"  competitors={len(df)}: {', '.join(df['competitor'])}")
    df.to_csv(OUT / "cleaned_data.csv", index=False)

    charts = {
        "pricing": chart_pricing(df),
        "unit": chart_unit_economics(df),
        "offerings": chart_offerings(df),
        "engagement": chart_engagement(df),
        "rating": chart_rating(df),
        "value": chart_value(df),
    }
    print(f"  charts: {len(charts)} saved")

    ins = build_insights(df)
    k = ins["kpis"]
    (OUT / "insights.json").write_text(json.dumps(ins, indent=2, default=str))

    dash = build_dashboard(df, ins, charts)
    print(f"  dashboard: {dash}")

    print("\nKey findings:")
    print(f"  - Avg total cost PKR {k['avg_cost']:,} "
          f"(range {k['min_cost']:,}-{k['max_cost']:,}, {k['spread_x']}x spread)")
    print(f"  - Cheapest: {k['cheapest']} · Premium: {k['priciest']} "
          f"· Engagement leader: {k['engagement_leader']}")
    print(f"  - Best value-for-money: {k['value_leader']}")
    print("\nDone. Open dashboard.html in a browser.")


if __name__ == "__main__":
    main()
