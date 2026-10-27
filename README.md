# Cash transfer RCT, Kerala: survey data processing and analysis

Code and results for a randomized cash transfer study among rice farmers in Kerala. The study gives cash at a time
of cash shortage during the Virippu rice season. It then tests whether households spend more on the farm activity
that comes right after each transfer.

There are two survey rounds:

| Round | Interviews | Season covered | Transfers |
|---|---|---|---|
| Round 1 (baseline) | July 2026 | Virippu 2025 costs; household data for the last 30 days | Before any transfer |
| Round 2 | 17–30 September 2026 | Virippu 2026 costs; household data for the last 30 days | After transfers in the 1st week of August and the 1st week of September 2026 |

**Treatment group:** `payment = 1` if the farmer is marked `Payment = Yes` in `RCT_payment_round1`. All other
households are the control group (`payment = 0`).

## Files

### Code

| File | What it does |
|---|---|
| `RCT_round1_analysis.py` | **Round 1, step 1:** merges the repeat sheets of the survey into one row per household (long to wide), fixes two duplicate farmer IDs and adds `round` and `payment`. It then computes Section 3 cultivation costs, cost per acre and outlier flags. |
| | **Round 1, later steps:** the baseline balance table, the treatment vs control cost comparison, item-level dated costs and the weekly bubble plots. |
| `RCT_round1_analysis.do` | Stata version of the Round 1 steps (Stata 16 or newer). |
| `RCT_round2_analysis.py` | **Round 2, step 1:** the same merge and cost steps for Virippu 2026. |
| | **Round 2, later steps:** cost comparison, bubble plots with the payment weeks, pesticide vs 3rd fertiliser dose, and PM-KISAN / Rs 2,000 use pies. |
| | **Comparisons with Round 1:** spending in the weeks right after each transfer (Round 1 as the pre-payment check) and household spending per person. |

### Results tables (Excel)

| File | Contents |
|---|---|
| `Balance_table_payment_round1.xlsx` | Round 1 balance of socio-economic characteristics, treatment vs control, with a joint orthogonality test |
| `Cost_comparison_payment_round1.xlsx` | Round 1 cost per acre by activity, treatment vs control |
| `Round2_cost_comparison_payment.xlsx` | Round 2 cost per acre by activity and timing relative to the transfers, treatment vs control |
| `Round2_spending_after_transfer.xlsx` | Spending in the 3 weeks after each transfer, with Round 1 as placebo and difference-in-differences |

### Figures

| File | Contents |
|---|---|
| `Cost_per_acre_bubble_overlay.jpg`, `Cost_per_acre_bubble_panels.jpg` | Round 1: weekly cost per acre by activity, with cumulative cost per acre |
| `Round2_cost_per_acre_bubble_overlay.jpg`, `Round2_cost_per_acre_bubble_panels.jpg` | Round 2: the same, with the two payment weeks marked |
| `Round2_pesticide_vs_3rd_fertiliser.jpg` | Round 2: pesticide and 3rd fertiliser dose by payment group |
| `Pesticide_3rd_fertiliser_round1_vs_round2.jpg` | The same, with Round 1 as the pre-payment check |
| `Round2_pmkisan_survey_use_pies.jpg` | Round 2: reported use of the Rs 2,000, treatment and control |
| `Round1_pmkisan_use_pie.jpg` | Round 1: reported use of PM-KISAN money, all households |

## Data not in this repository

The repository is public, so it holds no household-level data.

**Survey files you need to run the code:**
- `Survey_Round_1.xlsx` and `Survey_Round_1_questionnarie.xlsx`
- `RCT_payment_round1.xls`
- `Survey_round_2.xlsx` and `Survey_round_2_questionnarie.xlsx`

**Outputs left out because they contain household-level records** (listed in `.gitignore`):
- `Survey_Round_1_merged_wide.xlsx` and `Survey_round_2_merged_wide.xlsx`
- `PMKISAN_use_pies_round1_round2.xlsx`
- `Household_expenditure_per_person_round1_round2.xlsx`

The scripts recreate all four from the survey files.

## How to run

Requirements: Python 3 with `pandas`, `numpy`, `scipy`, `openpyxl`, `matplotlib` and `xlrd`. You need `xlrd` to read
the `.xls` payment file; without it, save the file as `RCT_payment_round1.xlsx` in the same folder.

```bash
pip install pandas numpy scipy openpyxl matplotlib xlrd

# 1. Round 1 (run first: Round 2 uses its outputs)
python RCT_round1_analysis.py <data_dir> <out_dir>

# 2. Round 2 and the cross-round analyses
python RCT_round2_analysis.py <data_dir> <round1_out_dir> <out_dir>
```

Without arguments, both scripts read the survey files from the current folder and write to `./outputs`. Run from
the survey files, the two Python scripts reproduce the tables and figures in this repository exactly (checked on
27 October 2026). The Stata do-file mirrors the Round 1 Python steps but has not been run in Stata.

## Main methods and data notes

- **Activity costs:** quantity × price per unit. For "other items", the amount paid is the cost, unless amount ÷
  quantity is implausibly low, in which case the amount is read as a unit price.
- **Special codes:** 999 and 9999 ("don't know") are treated as missing; a blank quantity and price count as 0.
- **Cost per acre:** activity cost ÷ rice area of that season (`rice_acres_2025` or `rice_acres_2026`).
- **Outliers:** an entry is flagged when |modified z-score| > 3.5 on log cost per acre (Iglewicz–Hoaglin). Flagged
  values are excluded from the comparisons.
- **Round 1 dates:** most Round 1 activity dates were recorded with the year 2026, so they are placed on the 2025
  calendar.
- **Round 1 form versions:** in the older version, the 1st fertiliser dose was recorded as a single item in the main
  sheet; it is included in the costs.
- **Round 2 harvest costs:** harvest had not finished by the interviews, so harvest costs are partly planned amounts.
  They are left out of the after-transfer windows.
- **Round 2 duplicate IDs:** farmer IDs 136, 208, 244 and 257 each appear twice, as different households. Both rows
  are kept.
- **Spending per person:** Round 2 did not ask household size, so per-person spending uses the Round 1 household
  size.
- **Treatment contrast:** 141 of 158 control households report receiving Rs 2,000 from the survey team in Round 2.
  The payment schedule by group should be confirmed before writing up results as cash vs no cash.
