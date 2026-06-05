# Technical Notes

Formulas, design decisions, and the reasoning behind them. This complements the README with the "how" and "why" at the level a technical reviewer would want.

---

## 1. Data layer (Power Query)

### Why Power Query instead of manual cleaning

Cleaning must be **repeatable and auditable**. Power Query records the cleaning as a chain of steps that re-runs with one click (Refresh). Manual cleaning doesn't scale and can't be reviewed.

### The five queries

| Query | Rows | Load mode |
|---|---|---|
| `Master_Clean` | 100 | Table in sheet |
| `Inventory_Clean` | 1,500 (45 errors) | Table in sheet — the 45 errors are seeded nulls, expected |
| `Transit_Clean` | 1,105 | Table in sheet |
| `Supply_Clean` | 100 | Table in sheet (must be loaded as a table so formulas can reference it) |
| `Transit_Enriched` | 1,105 | Table in sheet — `Transit_Clean` + enrichment |

### Date conversion with locale

Dates arrived as text in day/month/year order, in an English-UI Excel with a Mexican regional configuration. Converting with the default locale misreads `31/05/2026` (it tries month 31). The fix is explicit-locale conversion:

```
= Table.TransformColumnTypes(#"previous step", {{"Dispatch_Date", type date}}, "es-MX")
```

> **Lesson:** dates are the #1 source of pain in data cleaning. Always convert with an explicit locale, and prefer unambiguous source formats (ISO `2026-05-31`).

---

## 2. Enrichment: `Transit_Enriched`

Built entirely in Power Query (no copy-paste). Applied steps, in order:

1. **Source** — reads `Transit_Clean`
2. **Expanded Master_Clean** — merge (join) bringing `Weight_kg`, `Units_Per_Pallet`, `Shelf_Life_Days` from the master
3. **Renamed Columns** — strip the `Master_Clean.` prefix
4. **Added Conditional Column** — `Match_Status` (HUERFANO / OK)
5. **Changed Type with Locale** — `Dispatch_Date` → Date (es-MX)
6. **Added Custom** — `Dias_En_Transito`

### Orphan detection — a subtle bug worth documenting

`Match_Status` flags orders whose SKU doesn't exist in the master. The **first version** keyed off `Weight_kg = null`, but the master had seeded null weights on products that *do* exist — so they were falsely flagged, inflating the orphan count from 12 to 57.

**Fix:** key off `Shelf_Life_Days`, which is only null when the product genuinely isn't in the master:

```
each if [Shelf_Life_Days] = null then "HUERFANO" else "OK"
```

> **General principle:** to detect whether a join found a match, test a right-side column that is *never* null for valid records. A column that can be null for other reasons gives false positives.

### Days in transit

```
Duration.Days(DateTime.Date(DateTime.LocalNow()) - [Dispatch_Date])
```

Requires `Dispatch_Date` to already be type Date — which is why the type-conversion step must come *before* this one. Step order in Power Query is a dependency chain.

---

## 3. DRP engine

### Lookup choice: VLOOKUP vs XLOOKUP vs INDEX/MATCH

All three are used on purpose, to show judgment about when each fits:

- **VLOOKUP** — simple left-to-right lookups; also the classic the job posting asks for by name.
- **XLOOKUP** — modern, direction-agnostic, clean not-found handling; used to bring master attributes with `IFERROR`.
- **INDEX/MATCH** — the legacy pattern still common in inherited files.

```
=IFERROR(XLOOKUP(C2, Master_Clean[Item_Code], Master_Clean[Weight_kg]), "")
```

### In-transit quantity (SUMIFS, not SUMIF)

Two simultaneous conditions (same SKU **and** same destination) require SUMIFS:

```
=SUMIFS(Transit[Quantity], Transit[Item_Code], [SKU], Transit[Destination], [CD])
```

SUMIF allows only one condition; filtering by SKU alone would sum what's headed to *all* DCs.

### Statistical safety stock

```
SS = Z × √( LT × σD²  +  D² × σLT² )
```

- `Z = 1.65` → 95% service level
- Considers variability of **both** demand (σD) and lead time (σLT) — the second term `D²·σLT²` captures the risk of a late replenishment, more realistic than the demand-only version.
- `Z` lives in a named cell (`Z_Score`) so the entire network's service level changes by editing one cell.

### Reorder point

```
Reorder Point = (D × LT) + Safety Stock
```

### Status (IFS)

`DATO INVALIDO` → `PEDIR YA` → `EN CAMINO OK` → `SANO`.

The `PEDIR YA` vs `EN CAMINO OK` distinction considers in-transit stock, so it doesn't raise false alarms for product that already has reinforcement on the way.

---

## 4. Exception engine: `Analisis_Transito`

Spills `Transit_Enriched` live (`=Transit_Enriched[#All]`) and adds calculation + flag columns. Truck capacity: 24,000 kg or 33 pallets, whichever fills first.

### Load calculations

| Column | Formula (row 2) |
|---|---|
| `Pallets_Used` | `=IF(L2="HUERFANO","n/a",IF(ISNUMBER(J2),E2/J2,"n/a"))` |
| `Fill_Weight_%` | `=IF(L2="HUERFANO","n/a",IFERROR(E2*I2/24000*100,""))` |
| `Fill_Pallet_%` | `=IF(L2="HUERFANO","n/a",IFERROR(O2/33*100,""))` |
| `Camiones` | `=IF(L2="HUERFANO","n/a",IFERROR(MAX(CEILING(P2/100,1),CEILING(Q2/100,1)),""))` |
| `Eficiencia_%` | `=IF(L2="HUERFANO","n/a",IFERROR(MAX(P2,Q2)/(R2*100)*100,""))` |

**Dual-limit logic:** a truck fills by weight *or* space, whichever runs out first. Beverages fill by weight; snacks by space. So both fills are computed and the `MAX` is taken — that's the real limit. The `CEILING` refinement converts "142% full" (impossible) into "trucks needed" + a real efficiency that never exceeds 100%.

The `IF(L2="HUERFANO","n/a",...)` wrapper prevents `#DIV/0!` on orphan rows that have no `Units_Per_Pallet`.

### The four flags

```
Inefficient:  =IF(L2="HUERFANO","n/a",IF(S2<75,"CARGA INEFICIENTE","OK"))
FEFO:         =IF(L2="HUERFANO","n/a",IF(AND(LEFT(D2,3)="LAC",(K2-M2)/K2*100<70),"FEFO RIESGO","OK"))
Ghost:        =IF(L2="HUERFANO","n/a",IF(AND(G2="In Transit",M2>5),"FANTASMA","OK"))
```

### Counts dashboard (`Excepciones` sheet)

```
=COUNTIF(Analisis_Transito!T:T, "CARGA INEFICIENTE")
=COUNTIF(Analisis_Transito!U:U, "FEFO RIESGO")
=COUNTIF(Analisis_Transito!V:V, "FANTASMA")
=COUNTIF(Analisis_Transito!L:L, "HUERFANO")
=SUMIF(Fair_Share!N:N, "ESCASEZ", Fair_Share!O:O)
```

Final counts: **361** inefficient, **45** FEFO, **51** ghost, **12** orphan, **~22,717** units unfulfilled.

---

## 5. Fair Share Allocation: `Fair_Share`

Runs over the entire network; auto-detects shortage. Columns added to a live spill of `Inventory_Clean`:

| Column | Formula (row 2) |
|---|---|
| `Demanda_Total_SKU` | `=SUMIF($B$2:$B$1501,B2,$D$2:$D$1501)` |
| `Suministro_SKU` | `=SUMIF(Supply_Clean[Item_Code],B2,Supply_Clean[Available_Supply])` |
| `Ratio_Cobertura` | `=IFERROR(MIN(1,K2/J2),"")` |
| `Asignacion_Fair_Share` | `=IFERROR(D2*L2,"")` |
| `Estatus_Suministro` | `=IF(L2<1,"ESCASEZ","SUFICIENTE")` |
| `Demanda_No_Cubierta` | `=IF(N2="ESCASEZ",D2-M2,0)` |

**Why cap the ratio at 1 (`MIN`):** with surplus supply the ratio would exceed 1, which would allocate a DC more than it asked for. `MIN(1, ratio)` means: in surplus, each DC gets exactly its request; in shortage, its proportional share.

---

## 6. Data-type and formatting notes

- **Dates as serial numbers:** when the enriched table spills, `Dispatch_Date` shows as a serial (e.g., `46173`). Fixed with Number Format → Short Date (display only, doesn't change the value).
- **Defensive `IFERROR` / `IF` everywhere:** the data has seeded nulls and edge cases; a professional model never leaks `#DIV/0!` or `#N/A`. Formulas return `"n/a"` or `""` instead.
- **Exact text for COUNTIF:** the flag formulas produce exact strings (`"CARGA INEFICIENTE"`) so the counting COUNTIFs match reliably.

---

## 7. Lessons learned (the real path)

| Problem | Root cause | Fix |
|---|---|---|
| Dates impossible to parse | 4 mixed formats, ambiguous | Regenerate with 2 unambiguous formats |
| `DataFormat.Error` on dates | Converted without locale | Use "Using Locale" with es-MX |
| `Dias_En_Transito` = Error | Date column was text | Convert type before the calc; step order matters |
| `Match_Status` = Error everywhere | A column it used got renamed | Update the formula to the new name |
| 57 orphans (should be 12) | Keyed off `Weight_kg` (seeded nulls) | Key off `Shelf_Life_Days` |
| Stale quantities | Data was copy-pasted, not live | Use live reference `[#All]`; refresh |
| Impossible fill % (142%) | Plain division, no truck rounding | `CEILING` trucks + real efficiency |
| `#DIV/0!` on orphans | Orphans lack `Units_Per_Pallet` | Wrap formulas in `IF(L="HUERFANO","n/a",...)` |
| `SUMIF` couldn't find supply | `Supply_Clean` was Connection-Only | Load the query as a table |

**Principles that emerged:**

1. Dates are the #1 cleaning hazard — demand unambiguous formats, always convert with explicit locale.
2. In Power Query, steps form a dependency chain — changing something upstream can break something downstream.
3. To detect a failed join, test a column that's never null for valid records.
4. Never copy-paste data between sheets if you can reference it live.
5. Changing a base data field forces a review of everything that depends on it.
6. A professional model never shows errors — wrap formulas defensively.
7. Always validate counts against a ground truth instead of trusting the formula "should" be right.
8. A good flag fires on a meaningful minority — calibrating thresholds matters as much as the formula.
