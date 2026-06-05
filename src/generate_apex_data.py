"""
Apex Consumer Goods - Synthetic Data Generator
Generates 4 dirty CSVs simulating SAP transactional exports for a DRP project.

Design principles:
- Realistic FMCG coherence (weight vs volume, shelf life, demand variability)
- Believable injected errors (the kind SAP actually produces)
- Seeded exception scenarios so the model's flags have something to find

Usage:
    pip install -r requirements.txt
    python src/generate_apex_data.py
    # writes the 4 CSVs into ../data relative to this script

The seed is fixed (SEED = 42), so the output is deterministic: the same
seeded scenarios (orphans, FEFO risk, shortages) are reproduced every run,
which lets the Excel model's flag counts be validated against known numbers.
"""

import numpy as np
import pandas as pd
import random
from datetime import datetime, timedelta

# Reproducible but realistic
SEED = 42
np.random.seed(SEED)
random.seed(SEED)

TODAY = datetime(2026, 6, 3)

# ----------------------------------------------------------------------
# NETWORK DEFINITION
# ----------------------------------------------------------------------
PLANTS = {
    "PL-NTE": {"name": "Planta Norte (Monterrey)", "region": "Norte"},
    "PL-CEN": {"name": "Planta Centro (CDMX)", "region": "Centro"},
    "PL-SUR": {"name": "Planta Sur (Guadalajara)", "region": "Sur"},
}

# 15 DCs with a "home plant" and a base lead time + variability.
# Cross-region assignments exist on purpose (realistic, not perfectly regional).
DCS = {
    "CD-MTY": {"name": "CD Monterrey",    "plant": "PL-NTE", "lt": 1, "lt_sd": 0.3},
    "CD-SAL": {"name": "CD Saltillo",     "plant": "PL-NTE", "lt": 2, "lt_sd": 0.5},
    "CD-CHI": {"name": "CD Chihuahua",    "plant": "PL-NTE", "lt": 3, "lt_sd": 1.0},
    "CD-HER": {"name": "CD Hermosillo",   "plant": "PL-NTE", "lt": 4, "lt_sd": 1.2},
    "CD-TIJ": {"name": "CD Tijuana",      "plant": "PL-NTE", "lt": 5, "lt_sd": 1.5},
    "CD-CDMX": {"name": "CD Ciudad de Mexico", "plant": "PL-CEN", "lt": 1, "lt_sd": 0.2},
    "CD-PUE": {"name": "CD Puebla",       "plant": "PL-CEN", "lt": 1, "lt_sd": 0.3},
    "CD-QRO": {"name": "CD Queretaro",    "plant": "PL-CEN", "lt": 2, "lt_sd": 0.4},
    "CD-TOL": {"name": "CD Toluca",       "plant": "PL-CEN", "lt": 2, "lt_sd": 0.5},
    "CD-VER": {"name": "CD Veracruz",     "plant": "PL-CEN", "lt": 3, "lt_sd": 0.9},
    "CD-GDL": {"name": "CD Guadalajara",  "plant": "PL-SUR", "lt": 1, "lt_sd": 0.2},
    "CD-MOR": {"name": "CD Morelia",      "plant": "PL-SUR", "lt": 2, "lt_sd": 0.5},
    "CD-LEO": {"name": "CD Leon",         "plant": "PL-SUR", "lt": 2, "lt_sd": 0.6},
    "CD-MER": {"name": "CD Merida",       "plant": "PL-SUR", "lt": 4, "lt_sd": 1.3},
    "CD-CAN": {"name": "CD Cancun",       "plant": "PL-SUR", "lt": 5, "lt_sd": 1.6},
}

# ----------------------------------------------------------------------
# CATEGORY PROFILES (the weight-vs-volume tension is intentional)
# ----------------------------------------------------------------------
CATEGORIES = {
    "Bebidas Pesadas": {
        "n": 22, "weight_kg": (12, 24), "pallet_vol": (0.8, 1.2),
        "shelf_life": (180, 365), "demand": (200, 600), "demand_cv": (0.10, 0.20),
        "prefix": "BEB",
        "names": [("Agua Mineral {v}L", [0.5, 1, 1.5, 2]),
                  ("Refresco Cola {v}L", [0.6, 1, 2, 2.5]),
                  ("Bebida Deportiva {v}ml", [355, 500, 600]),
                  ("Agua Natural {v}L", [0.5, 1, 1.5]),
                  ("Refresco Toronja {v}ml", [355, 500, 600]),
                  ("Te Helado {v}L", [1, 1.5, 2])],
    },
    "Botanas Alto Volumen": {
        "n": 24, "weight_kg": (3, 7), "pallet_vol": (1.5, 2.4),
        "shelf_life": (90, 120), "demand": (300, 900), "demand_cv": (0.15, 0.30),
        "prefix": "BOT",
        "names": [("Papas Fritas {v}g", [45, 90, 150, 240]),
                  ("Frituras Maiz {v}g", [50, 100, 170]),
                  ("Cacahuates {v}g", [60, 120, 210]),
                  ("Botana Mixta {v}g", [100, 200, 350]),
                  ("Totopos {v}g", [150, 300, 500]),
                  ("Chicharron {v}g", [40, 80, 120])],
    },
    "Lacteos Perecederos": {
        "n": 18, "weight_kg": (6, 12), "pallet_vol": (1.0, 1.4),
        "shelf_life": (14, 30), "demand": (150, 450), "demand_cv": (0.25, 0.45),
        "prefix": "LAC",
        "names": [("Leche Entera {v}L", [1, 1.5, 2]),
                  ("Yogurt Bebible {v}ml", [220, 250, 900]),
                  ("Crema {v}ml", [200, 450, 900]),
                  ("Queso Fresco {v}g", [200, 400, 500]),
                  ("Leche Deslactosada {v}L", [1, 1.5]),
                  ("Yogurt Griego {v}g", [150, 500, 1000])],
    },
    "Abarrotes Secos": {
        "n": 20, "weight_kg": (8, 16), "pallet_vol": (0.9, 1.3),
        "shelf_life": (270, 540), "demand": (180, 500), "demand_cv": (0.08, 0.15),
        "prefix": "ABA",
        "names": [("Avena {v}g", [200, 400, 800]),
                  ("Arroz {v}kg", [1, 2, 5]),
                  ("Pasta {v}g", [200, 500, 1000]),
                  ("Azucar {v}kg", [1, 2, 5]),
                  ("Harina {v}kg", [1, 2]),
                  ("Cereal {v}g", [200, 300, 500])],
    },
    "Estacional Promo": {
        "n": 16, "weight_kg": (4, 14), "pallet_vol": (1.0, 2.0),
        "shelf_life": (60, 150), "demand": (100, 700), "demand_cv": (0.40, 0.70),
        "prefix": "PRO",
        "names": [("Edicion Especial Pack {v}", [6, 12, 24]),
                  ("Pack Promocional {v}pz", [4, 8, 12]),
                  ("Temporada Pack {v}", [6, 12]),
                  ("Combo Fiesta {v}pz", [10, 20]),
                  ("Edicion Limitada {v}pz", [6, 12, 18])],
    },
}

print("Network and category profiles defined.")
print(f"Plants: {len(PLANTS)} | DCs: {len(DCS)} | Categories: {len(CATEGORIES)}")
print(f"Total SKUs planned: {sum(c['n'] for c in CATEGORIES.values())}")

# ======================================================================
# 1. MASTER DATA — SKUs
# ======================================================================
sku_rows = []
sku_index = {}  # item_code -> clean attributes (for later coherence)

for cat, cfg in CATEGORIES.items():
    for i in range(cfg["n"]):
        code = f"{cfg['prefix']}-{1000 + i}"
        name_template, sizes = random.choice(cfg["names"])
        size = random.choice(sizes)
        # format size cleanly (drop trailing .0 for whole numbers)
        size_str = str(int(size)) if float(size).is_integer() else str(size)
        name = name_template.format(v=size_str)
        weight = round(np.random.uniform(*cfg["weight_kg"]), 2)
        vol = round(np.random.uniform(*cfg["pallet_vol"]), 3)
        shelf = int(np.random.uniform(*cfg["shelf_life"]))
        sku_rows.append({
            "Item_Code": code,
            "Description": name,
            "Category": cat,
            "Weight_kg": weight,
            "Pallet_Volume": vol,
            "Shelf_Life_Days": shelf,
        })
        sku_index[code] = {
            "category": cat, "weight": weight, "vol": vol, "shelf": shelf,
            "demand_range": cfg["demand"], "cv_range": cfg["demand_cv"],
        }

df_sku = pd.DataFrame(sku_rows)
ALL_CODES = list(sku_index.keys())
print(f"Generated {len(df_sku)} clean SKUs.")

# ======================================================================
# 2. DAILY INVENTORY + FORECAST  (100 SKUs x 15 DCs = ~1500 rows)
# ======================================================================
inv_rows = []
inv_lookup = {}  # (dc, code) -> dict, for seeding transit coherence later

for dc_code, dc in DCS.items():
    for code in ALL_CODES:
        meta = sku_index[code]
        avg_demand = round(np.random.uniform(*meta["demand_range"]) *
                           np.random.uniform(0.7, 1.3), 1)
        cv = np.random.uniform(*meta["cv_range"])
        demand_sd = round(avg_demand * cv, 1)
        lt = max(0.5, round(np.random.normal(dc["lt"], dc["lt_sd"]), 1))
        lt_sd = round(dc["lt_sd"], 2)
        # base inventory: a few days of supply, varied
        days_cover = np.random.uniform(2, 18)
        inv = int(avg_demand * days_cover)
        inv_rows.append({
            "DC": dc_code,
            "Item_Code": code,
            "Current_Inventory": inv,
            "Avg_Daily_Demand": avg_demand,
            "Demand_StdDev": demand_sd,
            "Lead_Time_Days": lt,
            "Lead_Time_StdDev": lt_sd,
        })
        inv_lookup[(dc_code, code)] = inv_rows[-1]

df_inv = pd.DataFrame(inv_rows)
print(f"Generated {len(df_inv)} inventory/forecast rows.")

# ======================================================================
# 3. IN-TRANSIT ORDERS  (~1000 rows)
# ======================================================================
transit_rows = []
order_counter = 50000
N_ORDERS = 1000

status_clean = "In Transit"

for _ in range(N_ORDERS):
    dc_code = random.choice(list(DCS.keys()))
    origin = DCS[dc_code]["plant"]
    code = random.choice(ALL_CODES)
    meta = sku_index[code]
    qty = int(np.random.uniform(*meta["demand_range"]) * np.random.uniform(1, 5))
    lt = DCS[dc_code]["lt"]
    # most orders dispatched recently, within a normal lead time window
    days_ago = int(np.random.uniform(0, max(1, lt)))
    dispatch = TODAY - timedelta(days=days_ago)
    order_counter += 1
    transit_rows.append({
        "Order_ID": f"SO-{order_counter}",
        "Origin": origin,
        "Destination": dc_code,
        "Item_Code": code,
        "Quantity": qty,
        "Dispatch_Date": dispatch,
        "SAP_Status": status_clean,
    })

df_transit = pd.DataFrame(transit_rows)
print(f"Generated {len(df_transit)} clean in-transit orders (before seeding/dirtying).")

# ======================================================================
# 4. SEED EXCEPTION SCENARIOS (so the model's flags have real targets)
# ======================================================================

# --- 4a. STOCKOUT RISK: pick DCs/SKUs with very low inventory vs demand,
#         and make sure no transit order saves them in time.
stockout_targets = random.sample(list(inv_lookup.keys()), 60)
for (dc_code, code) in stockout_targets:
    row = inv_lookup[(dc_code, code)]
    # 0.5 to 2 days of cover -> will run out before typical replenishment
    row["Current_Inventory"] = int(row["Avg_Daily_Demand"] * np.random.uniform(0.5, 2.0))

# --- 4b. FEFO RISK: dairy/perishable in transit dispatched long ago,
#         so it arrives with little remaining shelf life.
lac_codes = [c for c in ALL_CODES if c.startswith("LAC")]
fefo_orders = []
for _ in range(40):
    dc_code = random.choice(list(DCS.keys()))
    code = random.choice(lac_codes)
    shelf = sku_index[code]["shelf"]
    # dispatched a big chunk of its shelf life ago
    days_ago = int(shelf * np.random.uniform(0.55, 0.80))
    order_counter += 1
    fefo_orders.append({
        "Order_ID": f"SO-{order_counter}",
        "Origin": DCS[dc_code]["plant"],
        "Destination": dc_code,
        "Item_Code": code,
        "Quantity": int(np.random.uniform(100, 400)),
        "Dispatch_Date": TODAY - timedelta(days=days_ago),
        "SAP_Status": "In Transit",
    })

# --- 4c. INEFFICIENT LOAD: clusters of small orders on the same route
#         that together (or alone) underfill a truck. We tag them via small qty.
ineff_orders = []
for _ in range(50):
    dc_code = random.choice(list(DCS.keys()))
    code = random.choice(ALL_CODES)
    order_counter += 1
    ineff_orders.append({
        "Order_ID": f"SO-{order_counter}",
        "Origin": DCS[dc_code]["plant"],
        "Destination": dc_code,
        "Item_Code": code,
        "Quantity": int(np.random.uniform(20, 90)),  # deliberately small
        "Dispatch_Date": TODAY - timedelta(days=int(np.random.uniform(0, 3))),
        "SAP_Status": "In Transit",
    })

# --- 4d. GHOST SHIPMENTS: dispatched 6-9 days ago on 1-day routes, still "In Transit"
ghost_orders = []
short_lt_dcs = [d for d, v in DCS.items() if v["lt"] <= 2]
for _ in range(15):
    dc_code = random.choice(short_lt_dcs)
    code = random.choice(ALL_CODES)
    order_counter += 1
    ghost_orders.append({
        "Order_ID": f"SO-{order_counter}",
        "Origin": DCS[dc_code]["plant"],
        "Destination": dc_code,
        "Item_Code": code,
        "Quantity": int(np.random.uniform(150, 500)),
        "Dispatch_Date": TODAY - timedelta(days=int(np.random.uniform(6, 9))),
        "SAP_Status": "In Transit",
    })

df_transit = pd.concat([df_transit, pd.DataFrame(fefo_orders + ineff_orders + ghost_orders)],
                       ignore_index=True)
df_inv = pd.DataFrame(inv_rows)  # rebuild from mutated inv_lookup refs
print(f"Seeded scenarios. Transit now {len(df_transit)} rows.")
print(f"  Stockout targets: {len(stockout_targets)} | FEFO: {len(fefo_orders)} | "
      f"Ineff: {len(ineff_orders)} | Ghost: {len(ghost_orders)}")

# ======================================================================
# 5. FAIR SHARE SCENARIO (plant shortage: total DC need > plant supply)
#    We create a separate plant-supply file so the allocation has a constraint.
# ======================================================================
# Pick a handful of SKUs and declare a constrained plant supply for them.
fair_share_skus = random.sample(ALL_CODES, 8)
plant_supply_rows = []
for code in ALL_CODES:
    # For most SKUs supply is ample; for the chosen few it's constrained.
    total_demand_est = df_inv[df_inv["Item_Code"] == code]["Avg_Daily_Demand"].sum()
    if code in fair_share_skus:
        supply = int(total_demand_est * np.random.uniform(0.4, 0.7))  # shortage
    else:
        supply = int(total_demand_est * np.random.uniform(1.5, 3.0))  # ample
    plant = sku_index_plant = random.choice(list(PLANTS.keys()))
    plant_supply_rows.append({
        "Plant": plant,
        "Item_Code": code,
        "Available_Supply": supply,
    })
df_supply = pd.DataFrame(plant_supply_rows)
print(f"Fair-share scenario: {len(fair_share_skus)} constrained SKUs out of {len(ALL_CODES)}.")

# ======================================================================
# 6. INJECT ERRORS (the SAP noise the analyst must clean)
# ======================================================================

# ---- 6a. MASTER DATA errors ----
df_sku = df_sku.copy()

# Whitespace / invisible chars in codes and descriptions
dirty_idx = df_sku.sample(frac=0.12, random_state=1).index
df_sku.loc[dirty_idx, "Item_Code"] = df_sku.loc[dirty_idx, "Item_Code"].apply(
    lambda x: random.choice([f"  {x}", f"{x} ", f" {x} ", f"{x}\t"]))
desc_idx = df_sku.sample(frac=0.15, random_state=2).index
df_sku.loc[desc_idx, "Description"] = df_sku.loc[desc_idx, "Description"].apply(
    lambda x: random.choice([x.upper(), x.lower(), f"  {x}", x.replace(" ", "  ")]))

# Duplicate SKUs (same code, second row) -> dedup target
dupes = df_sku.sample(5, random_state=3).copy()
df_sku = pd.concat([df_sku, dupes], ignore_index=True)

# A few null/blank physical attributes
null_idx = df_sku.sample(6, random_state=4).index
df_sku.loc[null_idx, "Weight_kg"] = np.nan

# ---- 6b. INVENTORY errors ----
df_inv = df_inv.copy()

# DC codes written inconsistently
dc_variants = {"CD-MTY": ["CD Monterrey", "CDMTY", "cd-mty"],
               "CD-CDMX": ["CD CDMX", "CDMX", "cd-cdmx "]}
for clean, variants in dc_variants.items():
    idx = df_inv[df_inv["DC"] == clean].sample(frac=0.25, random_state=5).index
    df_inv.loc[idx, "DC"] = [random.choice(variants) for _ in idx]

# Null demand / inventory values
ni = df_inv.sample(frac=0.04, random_state=6).index
df_inv.loc[ni, "Avg_Daily_Demand"] = np.nan
ni2 = df_inv.sample(frac=0.03, random_state=7).index
df_inv.loc[ni2, "Current_Inventory"] = np.nan

# Negative / absurd inventory (data entry errors)
neg = df_inv.sample(8, random_state=8).index
df_inv.loc[neg, "Current_Inventory"] = [random.choice([-50, -10, 999999]) for _ in neg]

# Numbers stored as text with stray chars
df_inv["Avg_Daily_Demand"] = df_inv["Avg_Daily_Demand"].astype(object)
txt = df_inv.sample(frac=0.03, random_state=9).index
df_inv.loc[txt, "Avg_Daily_Demand"] = df_inv.loc[txt, "Avg_Daily_Demand"].apply(
    lambda x: f"{x} " if pd.notna(x) else x)

# ---- 6c. TRANSIT errors ----
df_transit = df_transit.copy()

# Inconsistent SAP status text
sidx = df_transit.sample(frac=0.30, random_state=10).index
df_transit.loc[sidx, "SAP_Status"] = [
    random.choice(["IN_TRANSIT", "En Transito", "in transit", "EN TRANSITO", "Entregado"])
    for _ in sidx]

# Mixed date formats: convert some to strings in different formats
df_transit["Dispatch_Date"] = pd.to_datetime(df_transit["Dispatch_Date"])
def mess_date(d, mode):
    if mode == 0: return d.strftime("%Y-%m-%d")      # ISO, unambiguous
    if mode == 1: return d.strftime("%d-%b-%Y")      # text month, unambiguous
    return d.strftime("%Y-%m-%d")
# Only two clearly-distinguishable formats: ISO and text-month.
# Both are non-ambiguous, so a single Locale step parses them cleanly.
date_modes = np.random.choice([0, 1], size=len(df_transit))
df_transit["Dispatch_Date"] = [mess_date(d, m) for d, m in
                               zip(df_transit["Dispatch_Date"], date_modes)]

# Orphan Item_Codes (not in master) -> cross-reference will catch
orph = df_transit.sample(12, random_state=11).index
df_transit.loc[orph, "Item_Code"] = [f"XXX-{random.randint(8000,8999)}" for _ in orph]

# Whitespace in transit item codes too
wsi = df_transit.sample(frac=0.10, random_state=12).index
df_transit.loc[wsi, "Item_Code"] = df_transit.loc[wsi, "Item_Code"].apply(lambda x: f" {x} ")

# A few negative / zero quantities
nq = df_transit.sample(6, random_state=13).index
df_transit.loc[nq, "Quantity"] = [random.choice([-100, 0, -25]) for _ in nq]

print("Errors injected into all three datasets.")

# ======================================================================
# 7. SHUFFLE + EXPORT
# ======================================================================
df_sku = df_sku.sample(frac=1, random_state=20).reset_index(drop=True)
df_inv = df_inv.sample(frac=1, random_state=21).reset_index(drop=True)
df_transit = df_transit.sample(frac=1, random_state=22).reset_index(drop=True)

import os
# Write to ../data relative to this script, so it works from any clone of the repo.
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
os.makedirs(OUT_DIR, exist_ok=True)

df_sku.to_csv(os.path.join(OUT_DIR, "Master_Data_SKUs.csv"), index=False)
df_inv.to_csv(os.path.join(OUT_DIR, "Daily_Inventory_Forecast.csv"), index=False)
df_transit.to_csv(os.path.join(OUT_DIR, "In_Transit_Orders.csv"), index=False)
df_supply.to_csv(os.path.join(OUT_DIR, "Plant_Supply.csv"), index=False)

print("\n=== FILES WRITTEN ===")
for f in ["Master_Data_SKUs.csv", "Daily_Inventory_Forecast.csv",
          "In_Transit_Orders.csv", "Plant_Supply.csv"]:
    p = os.path.join(OUT_DIR, f)
    rows = sum(1 for _ in open(p)) - 1
    print(f"  {f}: {rows} rows")

# ======================================================================
# 8. SEEDED SCENARIO AUDIT
# Prints the ground-truth counts of every error/scenario seeded above, so the
# Excel model's flags can be validated against known numbers.
# ======================================================================
print("\n=== SEEDED SCENARIO AUDIT (ground truth for model validation) ===")
print(f"Duplicate SKUs injected: 5")
print(f"Null Weight_kg in master: 6")
print(f"Orphan item codes in transit: 12")
print(f"Negative/absurd inventory values: 8")
print(f"Ghost shipments (>5d on short routes): 15")
print(f"FEFO-risk dairy orders: 40")
print(f"Inefficient (small) loads: 50")
print(f"Stockout-risk DC/SKU pairs: 60")
print(f"Fair-share constrained SKUs: 8 -> {sorted(fair_share_skus)}")
