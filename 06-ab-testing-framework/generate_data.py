"""
generate_data.py
────────────────────────────────────────────────────────
Generates three realistic A/B test datasets:

  Test 1 — Email Subject Line Test
           Control: "Your Weekly Summary"
           Variant: "🔥 Don't Miss This Week's Highlights"
           Metric : Email open rate

  Test 2 — Landing Page Button Colour Test
           Control: Blue CTA button
           Variant: Green CTA button
           Metric : Click-through rate

  Test 3 — Pricing Page Layout Test
           Control: Original layout
           Variant: Redesigned layout with social proof
           Metric : Conversion rate (purchase)
────────────────────────────────────────────────────────
"""

import pandas as pd
import numpy as np
import random

random.seed(42)
np.random.seed(42)

def generate_test(test_name, control_name, variant_name,
                  metric_col, n_control, n_variant,
                  control_rate, variant_rate):
    rows = []

    for i in range(n_control):
        converted = np.random.binomial(1, control_rate)
        rows.append({
            "user_id":    f"U-C-{i+1:05d}",
            "group":      "Control",
            "group_name": control_name,
            "test":       test_name,
            "day":        random.randint(1, 14),
            metric_col:   converted,
        })

    for i in range(n_variant):
        converted = np.random.binomial(1, variant_rate)
        rows.append({
            "user_id":    f"U-V-{i+1:05d}",
            "group":      "Variant",
            "group_name": variant_name,
            "test":       test_name,
            "day":        random.randint(1, 14),
            metric_col:   converted,
        })

    return pd.DataFrame(rows)

# Test 1: Email open rate (variant clearly wins)
df1 = generate_test(
    test_name     = "Email Subject Line Test",
    control_name  = "Your Weekly Summary",
    variant_name  = "Don't Miss This Week's Highlights",
    metric_col    = "opened",
    n_control     = 2500,
    n_variant     = 2500,
    control_rate  = 0.21,
    variant_rate  = 0.27,
)

# Test 2: Button colour CTR (marginal difference — inconclusive)
df2 = generate_test(
    test_name     = "Landing Page Button Colour Test",
    control_name  = "Blue CTA Button",
    variant_name  = "Green CTA Button",
    metric_col    = "clicked",
    n_control     = 1800,
    n_variant     = 1800,
    control_rate  = 0.14,
    variant_rate  = 0.155,
)

# Test 3: Pricing page (variant wins significantly)
df3 = generate_test(
    test_name     = "Pricing Page Layout Test",
    control_name  = "Original Layout",
    variant_name  = "Redesigned with Social Proof",
    metric_col    = "converted",
    n_control     = 3200,
    n_variant     = 3200,
    control_rate  = 0.035,
    variant_rate  = 0.052,
)

df1.to_csv("test1_email.csv",   index=False)
df2.to_csv("test2_button.csv",  index=False)
df3.to_csv("test3_pricing.csv", index=False)

print("✅  Generated:")
print(f"    test1_email.csv   — {len(df1):,} rows")
print(f"    test2_button.csv  — {len(df2):,} rows")
print(f"    test3_pricing.csv — {len(df3):,} rows")
