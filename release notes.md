# Release Notes - Calculation Summary

Generated: June 5, 2026

This document summarizes the proposal calculation rules currently implemented in `pcs_proposal_web.py` and mirrored into the Profit Summary workbook formulas in `profit_summary_formulas.py`.

## June 24, 2026 Updates

- Standardized the Gaco E5320 constant name to `GACO_E5320_PRICE` throughout the program.
- Updated new proposal unit price handling so default material prices come from program constants and workbook formulas instead of stale or hard-coded posted values.
- Preserved the Profit Summary workbook price-per-unit formulas for silicone, Gaco patch, Bleed Trap, Gaco E5320, SW 1-Flash, SW Bleed Block, Drainage Mat, foam, RFC labor, and PCS labor when the value has not been manually overridden.
- Synced Profit Summary `Data` sheet constants before saving, including `Data!K8 = GACO_S42_BASE_PRICE = 195`.
- Hardened display logic so formula-backed unit prices use recalculated program values instead of stale cached workbook values.
- Added explicit manual unit-price override tracking for blank/new proposal creation so a stale displayed price is not saved as an override unless the user actually edits the price field.
- Updated 10/15/20 commission calculations to exclude office fee, matching the revised Profit Summary workbook formulas.
- Rebuilt and deployed the packaged macOS application to `/Applications/PCS_Proposal.app`.

## Core Inputs

- Total squares are calculated as `flat_roof_squares + wall_squares`.
- If both flat roof and wall square inputs are blank, the app can fall back to the existing saved total squares.
- Roof types are evaluated in this order:
  - TPO/EPDM
  - Metal
  - Mod Bit
  - Ballasted 60 mil
  - Ballasted 45 mil
  - Rock/Foam/Coat
- Product options currently driving material calculations are Gaco and Uniflex.
- Pricing mode is selected by `pcs_or_roofer_ind`:
  - `PCS Direct` uses PCS Direct pricing.
  - Anything else uses Roofer pricing.

## Price Per Square

Roofer pricing:

| Roof Type | 10 Year | 15 Year | 20 Year |
| --- | ---: | ---: | ---: |
| TPO/EPDM | 320 | 360 | 400 |
| Metal | 325 | 365 | 405 |
| Mod Bit | 330 | 370 | 410 |
| Ballasted 60 mil | 470 | 510 | 550 |
| Ballasted 45 mil | 575 | 615 | 655 |
| Rock/Foam/Coat | 670 | 710 | 750 |

PCS Direct pricing:

| Roof Type | 10 Year | 15 Year | 20 Year |
| --- | ---: | ---: | ---: |
| TPO/EPDM | 350 | 390 | 430 |
| Metal | 355 | 395 | 435 |
| Mod Bit | 360 | 400 | 440 |
| Ballasted 60 mil | 500 | 540 | 580 |
| Ballasted 45 mil | 605 | 645 | 685 |
| Rock/Foam/Coat | 700 | 740 | 780 |

Rules:

- If roof type or pricing mode changes, 10/15/20 year prices reset to the base table.
- The user can override the 10-year price per square.
- When 10-year price is overridden, the same delta is applied to 15-year and 20-year pricing.

## Labor Days

Calculated labor days:

- Ballasted 60 mil or Ballasted 45 mil: `ceil(squares / 30)`
- Rock/Foam/Coat: `ceil(squares / 75)`
- All other roof types: `ceil(squares / 45)`

Rules:

- Labor days reset to the calculated baseline when roof type or squares change.
- If labor days are blank or zero, the calculated baseline is used.
- The app tracks whether labor days were manually overridden.

## Coverage Factors

Coverage factors are the same for Gaco and Uniflex:

| Roof Type | 10 Year | 15 Year | 20 Year |
| --- | ---: | ---: | ---: |
| TPO/EPDM | 1.25 | 1.75 | 2.25 |
| Metal | 1.35 | 1.85 | 2.35 |
| Mod Bit | 1.25 | 1.75 | 2.25 |
| Ballasted 60 mil | 2.50 | 3.25 | 3.75 |
| Ballasted 45 mil | 3.00 | 4.50 | 5.50 |
| Rock/Foam/Coat | 1.25 | 1.75 | 2.25 |

Rules:

- Adjusted coverage is added to the 10/15/20 coverage factor when provided.
- Silicone units are calculated as `ceil((squares / 5) * coverage_factor)`.
- If the user manually changes 10-year silicone units, adjusted coverage is reset to 0.
- When 10-year silicone units are manually overridden, 15-year and 20-year units are derived from the 10-year value using the 15/10 and 20/10 coverage ratios.

## Material Unit And Price Rules

All material units and prices are normalized before totals:

- If units are 0 or blank, unit price is forced to 0.
- If units are greater than 0 and price is blank or 0, the base price is restored.
- Line totals use whole-number rounded units and whole-number rounded prices.

Base unit prices:

| Item | Base Price |
| --- | ---: |
| Gaco S42 silicone | 195 |
| Uniflex silicone | 185 |
| Gaco patch | 125 |
| Bleed Trap | 168 |
| Gaco E5320 | 185 |
| SW 1-Flash | 110 |
| SW Bleed Block | 100 |
| Drainage Mat | 164 |
| Gaco foam | 2600 |
| Uniflex foam | 2400 |
| RFC labor | 250 |
| PCS labor per day | 3250 |

Material quantity calculations:

| Item | Applies When | Units |
| --- | --- | --- |
| Silicone | Gaco or Uniflex | `ceil((squares / 5) * coverage_factor)` |
| Gaco Patch | Gaco, non Rock/Foam/Coat | `ceil(squares / 10)` |
| Gaco Patch | Gaco, Rock/Foam/Coat | `ceil(squares * 0.03)` |
| Bleed Trap | Gaco and Mod Bit | `ceil(squares / 5)` |
| SW 1-Flash | Uniflex and TPO/EPDM, Mod Bit, or Rock/Foam/Coat | `ceil(squares / 20)` |
| SW 1-Flash | Uniflex and Metal or ballasted roof types | `ceil(squares / 10)` |
| SW Bleed Block | Uniflex and Mod Bit | `ceil(squares / 5)` |
| Drainage Mat | Ballasted 60 mil or Ballasted 45 mil | `ceil(squares / 18)` |
| Foam | Rock/Foam/Coat | `ceil(squares / 25)` |

Line totals:

- Silicone total = silicone units x silicone price
- Gaco patch total = Gaco patch units x Gaco patch price
- Bleed Trap total = Bleed Trap units x Bleed Trap price
- SW 1-Flash total = SW 1-Flash units x SW 1-Flash price
- SW Bleed Block total = SW Bleed Block units x SW Bleed Block price
- Drainage Mat total = Drainage Mat units x Drainage Mat price
- Foam total = foam units x foam price

## Labor Totals

- RFC labor price is 250 only for Rock/Foam/Coat; otherwise it is 0.
- RFC labor total = `rfc_labor_price * squares`.
- If RFC labor price or squares are blank/0, RFC labor total is 0.
- PCS labor price defaults to 3250 if blank or 0.
- PCS labor total = `pcs_labor_price * labor_days`.

## Travel

Travel is included only when `include_travel` is Yes.

Travel constants:

| Item | Amount |
| --- | ---: |
| Gas per job | 250 |
| Hotel per room per night | 175 |
| Rooms per night | 6 |
| Food per day | 700 |
| Misc when labor days <= 2 | 250 |
| Misc when labor days > 2 | 500 |

Calculated travel total:

`250 + (6 * 175 * max(labor_days - 1, 0)) + (700 * labor_days) + misc`

Rules:

- If travel is turned on and the current travel total does not appear manually overridden, the app uses the calculated travel total.
- If travel is turned off after previously being on, travel total is reset to 0.

## Warranty

- Warranty total is 500 for each term only when product is Gaco and warranty is included.
- Otherwise warranty total is 0.
- Warranty totals are tracked separately for 10-year, 15-year, and 20-year prices.

## Office Fee

Office fee percent:

- Mark: 5%
- All other submitters: 5%

Office fee is calculated on each term subtotal and rounded to a whole dollar:

- 10-year office fee = `round(subtotal_10 * office_fee_pct, 0)`
- 15-year office fee = `round(subtotal_15 * office_fee_pct, 0)`
- 20-year office fee = `round(subtotal_20 * office_fee_pct, 0)`

Subtotal by term:

- 10-year subtotal = `(squares * price_per_sq_10) + warranty_10_total + travel_total + repair_costs_total`
- 15-year subtotal = `(squares * price_per_sq_15) + warranty_15_total + travel_total + repair_costs_total`
- 20-year subtotal = `(squares * price_per_sq_20) + warranty_20_total + travel_total + repair_costs_total`

Total price by term:

- 10-year total = `subtotal_10 + office_fee_10`
- 15-year total = `subtotal_15 + office_fee_15`
- 20-year total = `subtotal_20 + office_fee_20`

## Commission

Commission percent:

- Mark: 0%
- Richard: 0%
- All other submitters: 10%

Commission amount:

`round(commission_pct * (total_price_10 - foam_total - rfc_labor_total - scarifying_total - travel_total - repair_costs_total - office_fee_total), 0)`

## Total Cost

Total cost is the sum of:

- Silicone total
- Gaco patch total
- Bleed Trap total
- SW 1-Flash total
- SW Bleed Block total
- Drainage Mat total
- Foam total
- RFC labor total
- PCS labor total
- Scarifying total
- Travel total
- Repair costs total
- 10-year warranty total
- Commission amount

## Profit

Profit share:

`round(10% * (total_price_10 - total_cost), 0)`

PCS profit:

`total_price_10 - total_cost - profit_share`

Profit percentage:

`round(pcs_profit / total_price_10, 2)`

Daily profit:

`round(pcs_profit / labor_days, 0)`

## Profit Summary Workbook

The app writes the Python-calculated values into the Profit Summary workbook and also maintains Excel formulas for the workbook.

Important workbook mirrors:

- `M3`, `M5`, `M7`: 10/15/20 year price per square.
- `P3`, `P5`, `P7`: 10/15/20 year total prices.
- `E7`: labor days.
- `C11`, `H11`, `N11`: silicone units by term.
- Rows 11-17: material unit, price, and total formulas.
- Row 20: PCS labor.
- Row 23: warranty.
- Row 24: office fee.
- Row 26: total cost.
- Row 28: PCS profit.
- Row 29: profit percentage.
- Row 30: daily profit.
- Row 31: profit share.
- Row 32: commission.

## Recalculation Behavior

The app intentionally recalculates some values when key drivers change:

- Roof type or squares changed: labor days reset to baseline.
- Roof type or pricing mode changed: price per square resets to the selected table.
- Product, roof type, squares, or adjusted coverage changed: silicone units recalculate unless the user manually changed units.
- Product, roof type, or squares changed: related material quantities and default prices refresh.
- Submitted By changed: office fee and commission percentages refresh.
- Travel toggled off after being on: travel total resets to 0.

Manual override behavior:

- 10-year price per square can be overridden; 15-year and 20-year prices follow the same delta.
- Labor days can be overridden when roof type and squares have not changed.
- 10-year silicone units can be overridden; adjusted coverage is reset and 15/20-year units follow coverage ratios.
- Travel total can be manually overridden when travel is included.
