"""
GreenEdge Services — ETL Pipeline
Excel (Facebook Ads Raw) → Transform → PostgreSQL

Tables created:
  - dim_campaigns       : Campaign metadata
  - dim_adsets          : Ad set metadata
  - fact_ad_performance : Daily grain fact table
  - kpi_daily           : Pre-aggregated KPIs per day
  - kpi_campaign        : Pre-aggregated KPIs per campaign
"""

import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import warnings
warnings.filterwarnings("ignore")

# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_PATH = "/home/claude/greenedge/data/raw_ads_data.xlsx"
DB_URL     = "postgresql+psycopg2://greenedge_user:greenedge_pass@localhost:5432/greenedge_db"

# For this demo we use SQLite if Postgres isn't available
import os
SQLITE_PATH = "/home/claude/greenedge/data/greenedge.db"
DB_URL_SQLITE = f"sqlite:///{SQLITE_PATH}"

# ── EXTRACT ───────────────────────────────────────────────────────────────────
print("📥 Extracting data from Excel...")
df_raw = pd.read_excel(EXCEL_PATH, sheet_name="Facebook Ads Raw")
print(f"   {len(df_raw):,} rows extracted")

# ── TRANSFORM ─────────────────────────────────────────────────────────────────
print("🔧 Transforming...")

df = df_raw.copy()

# 1. Parse & validate dates
df["date"] = pd.to_datetime(df["date"])
df["year"]  = df["date"].dt.year
df["month"] = df["date"].dt.month
df["week"]  = df["date"].dt.isocalendar().week.astype(int)
df["dow"]   = df["date"].dt.day_name()
df["is_weekend"] = df["date"].dt.weekday >= 5

# 2. Null checks & floor negatives
numeric_cols = ["spend_usd","impressions","clicks","leads","booked_appointments","revenue_usd"]
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).clip(lower=0)

# 3. Derived KPIs (row level)
df["ctr"]          = np.where(df["impressions"] > 0, df["clicks"] / df["impressions"], 0)
df["cpc"]          = np.where(df["clicks"] > 0, df["spend_usd"] / df["clicks"], 0)
df["cpm"]          = np.where(df["impressions"] > 0, (df["spend_usd"] / df["impressions"]) * 1000, 0)
df["cpl"]          = np.where(df["leads"] > 0, df["spend_usd"] / df["leads"], 0)          # Cost Per Lead
df["cpa"]          = np.where(df["booked_appointments"] > 0,
                               df["spend_usd"] / df["booked_appointments"], 0)             # Cost Per Booked Appt
df["roas"]         = np.where(df["spend_usd"] > 0, df["revenue_usd"] / df["spend_usd"], 0)
df["lead_rate"]    = np.where(df["clicks"] > 0, df["leads"] / df["clicks"], 0)
df["booking_rate"] = np.where(df["leads"] > 0, df["booked_appointments"] / df["leads"], 0)

# 4. Flag rows beating the KPI target (CPA < $66)
df["beats_kpi_target"] = df["cpa"].between(0.01, 66)

# Round floats
float_cols = ["ctr","cpc","cpm","cpl","cpa","roas","lead_rate","booking_rate"]
df[float_cols] = df[float_cols].round(4)

# ── DIMENSION TABLES ──────────────────────────────────────────────────────────
dim_campaigns = df[["campaign_id","campaign_name","campaign_objective"]].drop_duplicates().reset_index(drop=True)

dim_adsets = df[["campaign_id","ad_set_name"]].drop_duplicates().reset_index(drop=True)
dim_adsets.insert(0, "adset_id", [f"A{str(i+1).zfill(3)}" for i in range(len(dim_adsets))])

# ── KPI AGGREGATIONS ──────────────────────────────────────────────────────────
def safe_avg(spend, units):
    return (spend.sum() / units.sum()).round(4) if units.sum() > 0 else 0

# Daily KPIs
kpi_daily = df.groupby("date").apply(lambda g: pd.Series({
    "total_spend":        round(g["spend_usd"].sum(), 2),
    "total_impressions":  int(g["impressions"].sum()),
    "total_clicks":       int(g["clicks"].sum()),
    "total_leads":        int(g["leads"].sum()),
    "total_bookings":     int(g["booked_appointments"].sum()),
    "total_revenue":      round(g["revenue_usd"].sum(), 2),
    "avg_ctr":            round(g["clicks"].sum() / g["impressions"].sum(), 4) if g["impressions"].sum() > 0 else 0,
    "avg_cpa":            round(g["spend_usd"].sum() / g["booked_appointments"].sum(), 2) if g["booked_appointments"].sum() > 0 else 0,
    "avg_roas":           round(g["revenue_usd"].sum() / g["spend_usd"].sum(), 4) if g["spend_usd"].sum() > 0 else 0,
    "avg_cpl":            round(g["spend_usd"].sum() / g["leads"].sum(), 2) if g["leads"].sum() > 0 else 0,
    "beats_target_pct":   round(g["beats_kpi_target"].mean() * 100, 1),
})).reset_index()

# Campaign KPIs
kpi_campaign = df.groupby(["campaign_id","campaign_name"]).apply(lambda g: pd.Series({
    "total_spend":        round(g["spend_usd"].sum(), 2),
    "total_impressions":  int(g["impressions"].sum()),
    "total_clicks":       int(g["clicks"].sum()),
    "total_leads":        int(g["leads"].sum()),
    "total_bookings":     int(g["booked_appointments"].sum()),
    "total_revenue":      round(g["revenue_usd"].sum(), 2),
    "avg_ctr":            round(g["clicks"].sum() / g["impressions"].sum(), 4) if g["impressions"].sum() > 0 else 0,
    "avg_cpa":            round(g["spend_usd"].sum() / g["booked_appointments"].sum(), 2) if g["booked_appointments"].sum() > 0 else 0,
    "avg_roas":           round(g["revenue_usd"].sum() / g["spend_usd"].sum(), 4) if g["spend_usd"].sum() > 0 else 0,
    "avg_cpl":            round(g["spend_usd"].sum() / g["leads"].sum(), 2) if g["leads"].sum() > 0 else 0,
    "beats_target_rows":  int(g["beats_kpi_target"].sum()),
    "total_rows":         len(g),
})).reset_index()

# Weekly trend
kpi_weekly = df.groupby(["year","week"]).apply(lambda g: pd.Series({
    "total_spend":    round(g["spend_usd"].sum(), 2),
    "total_bookings": int(g["booked_appointments"].sum()),
    "total_revenue":  round(g["revenue_usd"].sum(), 2),
    "avg_cpa":        round(g["spend_usd"].sum() / g["booked_appointments"].sum(), 2) if g["booked_appointments"].sum() > 0 else 0,
    "avg_roas":       round(g["revenue_usd"].sum() / g["spend_usd"].sum(), 4) if g["spend_usd"].sum() > 0 else 0,
})).reset_index()

# ── LOAD ──────────────────────────────────────────────────────────────────────
print("💾 Loading into SQLite (PostgreSQL-compatible schema)...")
engine = create_engine(DB_URL_SQLITE)

tables = {
    "dim_campaigns":      dim_campaigns,
    "dim_adsets":         dim_adsets,
    "fact_ad_performance": df,
    "kpi_daily":          kpi_daily,
    "kpi_campaign":       kpi_campaign,
    "kpi_weekly":         kpi_weekly,
}

for table_name, dataframe in tables.items():
    dataframe.to_sql(table_name, engine, if_exists="replace", index=False)
    print(f"   ✅ {table_name}: {len(dataframe):,} rows")

# ── VALIDATION QUERIES ────────────────────────────────────────────────────────
print("\n📊 Validation Summary:")
with engine.connect() as conn:
    for q, label in [
        ("SELECT COUNT(*) FROM fact_ad_performance",   "Total fact rows"),
        ("SELECT SUM(total_spend) FROM kpi_daily",     "Total spend ($)"),
        ("SELECT AVG(avg_cpa) FROM kpi_campaign WHERE avg_cpa > 0", "Avg CPA across campaigns ($)"),
        ("SELECT MAX(avg_roas) FROM kpi_campaign",     "Best ROAS"),
    ]:
        val = conn.execute(text(q)).scalar()
        print(f"   {label}: {val:,.2f}" if isinstance(val, float) else f"   {label}: {val:,}")

print(f"\n✅ ETL complete → {SQLITE_PATH}")
print("   Schema is PostgreSQL-compatible (same DDL works with pg driver)\n")

# ── EXPORT KPIs to JSON for dashboard ─────────────────────────────────────────
import json

dashboard_data = {
    "summary": {
        "total_spend":    round(df["spend_usd"].sum(), 2),
        "total_leads":    int(df["leads"].sum()),
        "total_bookings": int(df["booked_appointments"].sum()),
        "total_revenue":  round(df["revenue_usd"].sum(), 2),
        "overall_cpa":    round(df["spend_usd"].sum() / df["booked_appointments"].sum(), 2),
        "overall_roas":   round(df["revenue_usd"].sum() / df["spend_usd"].sum(), 4),
        "overall_ctr":    round(df["clicks"].sum() / df["impressions"].sum() * 100, 2),
        "kpi_target":     66,
        "date_range":     f"{df['date'].min().strftime('%b %d')} – {df['date'].max().strftime('%b %d, %Y')}",
    },
    "daily": kpi_daily.assign(date=kpi_daily["date"].astype(str)).to_dict(orient="records"),
    "campaigns": kpi_campaign.to_dict(orient="records"),
    "weekly": kpi_weekly.to_dict(orient="records"),
    "by_creative": df.groupby("creative_name").agg(
        spend=("spend_usd","sum"),
        leads=("leads","sum"),
        bookings=("booked_appointments","sum"),
        revenue=("revenue_usd","sum")
    ).round(2).reset_index().to_dict(orient="records"),
    "by_placement": df.groupby("placement").agg(
        spend=("spend_usd","sum"),
        impressions=("impressions","sum"),
        clicks=("clicks","sum"),
        bookings=("booked_appointments","sum"),
    ).round(2).reset_index().assign(
        ctr=lambda x: (x["clicks"]/x["impressions"]*100).round(2)
    ).to_dict(orient="records"),
}

json_path = "/home/claude/greenedge/data/dashboard_data.json"
with open(json_path, "w") as f:
    json.dump(dashboard_data, f, indent=2, default=str)

print(f"📤 Dashboard JSON exported → {json_path}")
