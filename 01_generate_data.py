"""
GreenEdge Services — Synthetic Facebook Ads Data Generator
Simulates 90 days of ad campaign data for a landscaping/cleaning business in Miami, FL
Output: raw_ads_data.xlsx (mimics a real Facebook Ads Manager export)
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

np.random.seed(42)
random.seed(42)

# ── CONFIG ──────────────────────────────────────────────────────────────────
START_DATE = datetime(2024, 10, 1)
END_DATE   = datetime(2024, 12, 31)
DAYS       = (END_DATE - START_DATE).days + 1

CAMPAIGNS = [
    {"id": "C001", "name": "Lawn Care – Lead Gen",       "objective": "LEAD_GENERATION", "budget": 35},
    {"id": "C002", "name": "House Cleaning – Retarget",  "objective": "LEAD_GENERATION", "budget": 25},
    {"id": "C003", "name": "Holiday Deep Clean – Promo", "objective": "CONVERSIONS",     "budget": 50},
    {"id": "C004", "name": "Landscaping – Awareness",    "objective": "REACH",           "budget": 20},
]

AD_SETS = {
    "C001": ["Homeowners 30-55", "High Income Zip Codes"],
    "C002": ["Website Visitors 30d", "Lookalike 1% – Buyers"],
    "C003": ["Holiday Intent", "Email List Upload"],
    "C004": ["Broad Miami-Dade", "Suburban Homeowners"],
}

CREATIVES = ["Video – Before/After", "Carousel – Services", "Static – Promo Offer", "Story – Testimonial"]

PLACEMENTS = ["Facebook Feed", "Instagram Feed", "Facebook Stories", "Audience Network"]

# ── GENERATE ROWS ────────────────────────────────────────────────────────────
rows = []
for day_offset in range(DAYS):
    date = START_DATE + timedelta(days=day_offset)
    is_weekend = date.weekday() >= 5
    is_holiday_season = date.month == 12  # Higher competition in December

    for camp in CAMPAIGNS:
        for adset in AD_SETS[camp["id"]]:
            for creative in random.sample(CREATIVES, k=random.randint(1, 2)):
                placement = random.choice(PLACEMENTS)

                # Base spend influenced by budget, season, weekend
                base_spend = camp["budget"] * random.uniform(0.6, 1.1)
                if is_weekend:
                    base_spend *= 0.85
                if is_holiday_season:
                    base_spend *= 1.25

                spend = round(base_spend * random.uniform(0.8, 1.2), 2)

                # CPM varies by placement and season
                cpm_base = {"Facebook Feed": 18, "Instagram Feed": 22,
                            "Facebook Stories": 14, "Audience Network": 9}[placement]
                cpm = cpm_base * (1.3 if is_holiday_season else 1.0) * random.uniform(0.85, 1.15)

                impressions = int((spend / cpm) * 1000)
                ctr = random.uniform(0.012, 0.045)  # 1.2% – 4.5%
                if "Retarget" in camp["name"] or "Lookalike" in adset:
                    ctr *= 1.35  # Retargeting has better CTR
                clicks = int(impressions * ctr)

                # Lead rate
                lead_rate = random.uniform(0.06, 0.18)
                if camp["objective"] == "LEAD_GENERATION":
                    lead_rate *= 1.2
                if "Holiday" in camp["name"]:
                    lead_rate *= 1.4
                leads = int(clicks * lead_rate)

                # Booking rate (leads → booked appointments)
                booking_rate = random.uniform(0.25, 0.55)
                if "Retarget" in camp["name"]:
                    booking_rate *= 1.3
                booked_appointments = int(leads * booking_rate)

                # Revenue per appointment
                revenue_per_appt = random.uniform(180, 420)
                revenue = round(booked_appointments * revenue_per_appt, 2)

                rows.append({
                    "date":                date.strftime("%Y-%m-%d"),
                    "campaign_id":         camp["id"],
                    "campaign_name":       camp["name"],
                    "campaign_objective":  camp["objective"],
                    "ad_set_name":         adset,
                    "creative_name":       creative,
                    "placement":           placement,
                    "spend_usd":           spend,
                    "impressions":         impressions,
                    "clicks":              clicks,
                    "leads":               leads,
                    "booked_appointments": booked_appointments,
                    "revenue_usd":         revenue,
                    "reach":               int(impressions * random.uniform(0.70, 0.92)),
                    "frequency":           round(random.uniform(1.1, 3.5), 2),
                    "video_views":         int(impressions * random.uniform(0.10, 0.35))
                    if "Video" in creative else 0,
                })

df = pd.DataFrame(rows)

# ── SAVE TO EXCEL ─────────────────────────────────────────────────────────────
out_path = "/home/claude/greenedge/data/raw_ads_data.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Facebook Ads Raw", index=False)

    # Summary tab
    summary = df.groupby("campaign_name").agg(
        total_spend=("spend_usd", "sum"),
        total_leads=("leads", "sum"),
        total_bookings=("booked_appointments", "sum"),
        total_revenue=("revenue_usd", "sum"),
    ).round(2).reset_index()
    summary.to_excel(writer, sheet_name="Campaign Summary", index=False)

print(f"✅ Generated {len(df):,} rows → {out_path}")
print(f"   Date range: {df['date'].min()} to {df['date'].max()}")
print(f"   Total spend: ${df['spend_usd'].sum():,.2f}")
print(f"   Total leads: {df['leads'].sum():,}")
print(f"   Total bookings: {df['booked_appointments'].sum():,}")
