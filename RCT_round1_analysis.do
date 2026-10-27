*==============================================================================*
* Cash transfer RCT - Survey Round 1: data merging and analysis (Stata)
*==============================================================================*
* Stata version of RCT_round1_analysis.py. Steps:
*   1. Merge the 20 repeat sheets of Survey_Round_1.xlsx into the main sheet
*      (long -> wide; older and newer form versions stacked), correct two
*      farmer IDs, add round / payment / in_payment_list, label variables.
*   2. Section 3 costs: activity costs (quantity x price per unit), section
*      totals, cost per acre, outlier flags  -> Survey_Round_1_merged_wide.dta
*   3. Baseline balance of socio-economic characteristics by payment
*   4. Per-acre cost components, treatment vs control
*   5. Item-level costs with dates on the 2025 calendar, weekly aggregates
*   6. Bubble plot: weekly cost per acre by activity + cumulative cost per acre
*
* Requires Stata 16 or newer. Run the whole file from the top: Steps 4-6 use
* the activity definitions set before Step 1.
* Written to mirror the Python script; differences:
*   - p-values of the Mann-Whitney (ranksum) test can differ slightly from
*     scipy (no continuity correction in Stata), and Stata's percentiles use a
*     different interpolation than pandas, so P5/P95 may differ a little.
*   - Randomization inference (Step 3): the Python p-value is one-sided
*     (permuted F at least as large as observed); read the "upper" row of
*     Stata's permute output, not the two-sided one.
*   - The Python version also writes a "likely cause" text for each outlier and
*     a cost_review sheet; here the outlier list carries the numbers only.
*==============================================================================*

version 16
clear all
set more off
capture set maxvar 10000      // not needed in Stata/BE: the dataset has about 800 variables

*------------------------------------------------------------------------------*
* CONFIGURATION - edit these two paths
*------------------------------------------------------------------------------*
global DATA "."          // folder with Survey_Round_1.xlsx, Survey_Round_1_questionnarie.xlsx, RCT_payment_round1.xls
global OUT  "outputs"    // all outputs are written here
capture mkdir "$OUT"
capture log close
log using "$OUT/RCT_round1_analysis.log", replace text

global SURVEY "$DATA/Survey_Round_1.xlsx"
global QUEST  "$DATA/Survey_Round_1_questionnarie.xlsx"
global PAY    "$DATA/RCT_payment_round1.xls"
global OLDFORM "vhWzpqVLG8rKBa5kqM2f3d"     // __version__ of the older form version

*==============================================================================*
* HELPER PROGRAMS
*==============================================================================*

* numeric copy of a variable that may have been imported as text
capture program drop numcopy
program define numcopy
    args src dst
    capture confirm numeric variable `src'
    if _rc gen double `dst' = real(`src')
    else   gen double `dst' = `src'
end

* lower-case text copy of a variable that may have been imported as a number
capture program drop strcopy
program define strcopy
    args src dst
    capture confirm string variable `src'
    if _rc gen strL `dst' = cond(missing(`src'), "", string(`src'))
    else   gen strL `dst' = ustrlower(`src')
end

* cost = quantity x price per unit, with the rules used throughout:
*   999 / 9999 in quantity or price -> missing ("don't know")
*   quantity and price both blank   -> 0 (activity not done)
*   quantity or price = 0           -> 0
*   quantity blank, unit says contract -> price is a lump sum
*   only one of quantity / price filled -> missing (incomplete)
capture program drop upcost
program define upcost
    args q u p newvar incvar
    tempvar Q P U
    numcopy `q' `Q'
    numcopy `p' `P'
    strcopy `u' `U'
    gen double `newvar' = `Q' * `P' if !missing(`Q', `P')
    replace `newvar' = `P' if missing(`Q') & `P' > 0 & !missing(`P') & regexm(`U', "contract|contact|lump|total")
    replace `newvar' = 0 if (`Q' == 0 | `P' == 0) & !inlist(`Q', 999, 9999) & !inlist(`P', 999, 9999)
    replace `newvar' = 0 if missing(`Q') & missing(`P')
    replace `newvar' = . if inlist(`Q', 999, 9999) | inlist(`P', 999, 9999)
    gen byte `incvar' = missing(`newvar')
end

* "other items" ask the total amount paid; the amount is read as a unit price
* (cost = quantity x amount) only where amount / quantity is implausibly low:
*   < Rs 1 per unit for any item, or < Rs 3 for fertiliser / lime items
* (not for units in ml or grams). 999 / 9999 in quantity -> quantity unknown.
capture program drop othcost
program define othcost
    args q u a item newvar incvar
    tempvar Q A U I QQ R base r1 r2
    numcopy `q' `Q'
    numcopy `a' `A'
    strcopy `u' `U'
    strcopy `item' `I'
    gen double `QQ' = `Q' if !inlist(`Q', 999, 9999)
    gen double `newvar' = `A'
    replace `newvar' = 0 if missing(`A') & missing(`QQ')
    replace `newvar' = 0 if `A' == 0
    gen double `R' = `A' / `QQ'
    gen byte `base' = `QQ' > 1 & !missing(`QQ') & `A' > 0 & !missing(`A') & !inlist(`A', 999, 9999) ///
        & !ustrregexm(`U', "\bml\b|\bg\b|gm|gram")
    gen byte `r1' = `base' & `R' < 1
    gen byte `r2' = `base' & !`r1' & `R' < 3 & ///
        ustrregexm(`I', "urea|potash|\bmop\b|fact|dap|sulph|phos|npk|complex|ammon|micro|magnes|zinc|borax|lime|line|dolomite")
    replace `newvar' = `QQ' * `A' if `r1' | `r2'
    replace `newvar' = . if inlist(`A', 999, 9999)
    replace `newvar' = . if missing(`A') & `QQ' > 0 & !missing(`QQ')
    gen byte `incvar' = missing(`newvar')
end

* Iglewicz-Hoaglin modified z-score of log(v), positive values only
capture program drop modz
program define modz
    args v z
    tempvar l ad
    gen double `l' = ln(`v') if `v' > 0 & !missing(`v')
    quietly summarize `l', detail
    local med = r(p50)
    gen double `ad' = abs(`l' - `med')
    quietly summarize `ad', detail
    local mad = r(p50)
    gen double `z' = .
    if `mad' > 0 & `mad' < . replace `z' = 0.6745 * (`l' - `med') / `mad'
end

* a date variable as a daily date (%td), whether imported as %td or %tc
capture program drop asdate
program define asdate
    args src dst
    local f : format `src'
    capture confirm numeric variable `src'
    if _rc {
        gen double `dst' = date(`src', "YMD")
    }
    else if strpos("`f'", "%tc") | strpos("`f'", "%tC") {
        gen double `dst' = dofc(`src')
    }
    else {
        quietly summarize `src'
        if r(max) > 1e8 & r(max) < . gen double `dst' = dofc(`src')
        else gen double `dst' = `src'
    }
    format `dst' %td
end

*==============================================================================*
* ACTIVITY DEFINITIONS (Section 3) - used in Steps 2, 4, 5 and 6
*==============================================================================*

* fixed activities: key | quantity | unit | price per unit | activity date | payment date | section
local fixed a_seed a_land b_labour b_fert c_mainlab c_mach c_translab c_fert1lab c_gapfill d_weed d_weedlab d_fert2lab e_pest e_pestlab e_fert3lab f_mach f_tractor f_labour
local f_a_seed     a_seed_qty a_seed_unit a_seed_price a_seed_actdate a_seed_paydate
local f_a_land     a_land_qty a_land_unit a_land_price a_land_actdate a_land_paydate
local f_b_labour   b_labour_qty b_labour_unit b_labour_amt b_labour_actdate b_labour_paydate
local f_b_fert     b_fert_qty b_fert_unit b_fert_amt b_fert_actdate b_fert_paydate
local f_c_mainlab  c_labour_qty c_labour_unit c_labour_amt c_labour_actdate_land c_labour_paydate_land
local f_c_mach     c_mach_qty c_mach_unit c_mach_amt c_mach_actdate_land c_mach_paydate_land
local f_c_translab c_labour_qty_001 c_labour_unit_001 c_labour_amt_001 c_labour_actdate c_labour_paydate
local f_c_fert1lab c_labour_fert1_qty c_labour_fert1_unit c_labour_fert1_amt c_labour_fert1_actdate c_labour_fert1_paydate
local f_c_gapfill  c_labour_gapfilling_qty c_labour_gapfilling_unit c_labour_gapfilling_amt c_labour_gapfilling_actdate c_labour_gapfilling_paydate
local f_d_weed     d_weed_qty d_weed_unit d_weed_price d_weed_actdate d_weed_paydate
local f_d_weedlab  d_weedlab_qty d_weedlab_unit d_weedlab_price d_weedlab_actdate d_weedlab_paydate
local f_d_fert2lab d_labour_fert2_qty d_labour_fert2_unit d_labour_fert2_amt d_labour_fert2_actdate d_labour_fert2_paydate
local f_e_pest     e_pest_qty e_pest_unit e_pest_price e_pest_actdate e_pest_paydate
local f_e_pestlab  e_labour_pest_qty e_labour_pest_unit e_labour_pest_amt e_labour_pest_actdate e_labour_pest_paydate
local f_e_fert3lab e_labour_fert3_qty e_labour_fert3_unit e_labour_fert3_amt e_labour_fert3_actdate e_labour_fert3_paydate
local f_f_mach     f_mach_qty f_mach_unit f_mach_amt f_mach_actdate f_mach_paydate
local f_f_tractor  f_tractor_qty f_tractor_unit f_tractor_amt f_tractor_actdate f_tractor_paydate
local f_f_labour   f_labour_qty f_labour_unit f_labour_amt f_labour_actdate f_labour_paydate

* repeat blocks: fertiliser doses (price per unit) and other items (amount paid); number of items
local fert c_fert1 d_fert2 e_fert3
local n_c_fert1 2
local n_d_fert2 6
local n_e_fert3 5
local other a_other b_other c_other d_other e_other f_other
local n_a_other 4
local n_b_other 4
local n_c_other 5
local n_d_other 3
local n_e_other 6
local n_f_other 4

* labels and sections, in season order
local order a_seed a_land a_other b_labour b_fert b_other c_mainlab c_mach c_translab c_fert1 c_fert1lab c_gapfill c_other d_weed d_weedlab d_fert2 d_fert2lab d_other e_pest e_pestlab e_fert3 e_fert3lab e_other f_mach f_tractor f_labour f_other
local L_a_seed "Seed"
local L_a_land "Land preparation"
local L_a_other "Other items, pre-nursery"
local L_b_labour "Nursery labour"
local L_b_fert "Nursery fertiliser"
local L_b_other "Other items, nursery"
local L_c_mainlab "Main-plot preparation labour"
local L_c_mach "Main-plot preparation machinery"
local L_c_translab "Transplanting labour"
local L_c_fert1 "Fertiliser, 1st dose"
local L_c_fert1lab "Labour, 1st fertiliser dose"
local L_c_gapfill "Gap-filling labour"
local L_c_other "Other items, after transplanting"
local L_d_weed "Herbicide / weedicide"
local L_d_weedlab "Labour, weedicide application"
local L_d_fert2 "Fertiliser, 2nd dose"
local L_d_fert2lab "Labour, 2nd fertiliser dose"
local L_d_other "Other items, after 2nd dose"
local L_e_pest "Pesticide"
local L_e_pestlab "Labour, pesticide application"
local L_e_fert3 "Fertiliser, 3rd dose / pesticide"
local L_e_fert3lab "Labour, 3rd fertiliser dose"
local L_e_other "Other items, pesticide / 4th dose"
local L_f_mach "Harvesting machinery"
local L_f_tractor "Tractor, harvest transport"
local L_f_labour "Labour, drying and packing"
local L_f_other "Other items, after harvest"
local SEC_A "Pre-nursery preparation"
local SEC_B "Nursery preparation"
local SEC_C "Transplanting"
local SEC_D "Weedicide / fertiliser application"
local SEC_E "Pesticide / fertiliser application"
local SEC_F "Harvesting"

*==============================================================================*
* STEP 1 - MERGE REPEAT SHEETS (LONG -> WIDE) INTO THE MAIN SHEET
*==============================================================================*

* block | count variable it follows | position variable | sheet(s): older form first
local blocks offfarm crop a_other b_other d_fert1 c_other d_fert2 d_other e_fert3 e_other f_other welfare
local cnt_offfarm offfarm_count
local pos_offfarm mem_pos
local sh_offfarm  offfarm_rep
local cnt_crop    other_crops_count
local pos_crop    crop_pos
local sh_crop     crop_rep
local cnt_a_other a_other_count
local pos_a_other a_other_pos
local sh_a_other  begin_repeat_OOtl5mSOZ begin_repeat_0W8hGwBOT
local cnt_b_other b_other_count
local pos_b_other b_other_pos
local sh_b_other  begin_repeat_FXOFif2yY begin_repeat_oFXmOdKGR
local cnt_d_fert1 d_fert1_count
local pos_d_fert1 d_fert1_pos
local sh_d_fert1  begin_repeat_q7UBsOtMr
local cnt_c_other c_other_count
local pos_c_other c_other_pos
local sh_c_other  begin_repeat_FV3YMN63y begin_repeat_kb3NZEaME
local cnt_d_fert2 d_fert2_count
local pos_d_fert2 d_fert2_pos
local sh_d_fert2  begin_repeat_BgK7Aa7Uw begin_repeat_pHGxYqKsw
local cnt_d_other d_other_count
local pos_d_other d_other_pos
local sh_d_other  begin_repeat_bDKIjSEX3 begin_repeat_9aGJkPAfe
local cnt_e_fert3 e_fert3_count
local pos_e_fert3 e_fert3_pos
local sh_e_fert3  begin_repeat_ntYu79uDD begin_repeat_y8vUlnzPY
local cnt_e_other e_other_count
local pos_e_other e_other_pos
local sh_e_other  begin_repeat_z8dV2uORP begin_repeat_PS5svQ16g
local cnt_f_other f_other_count
local pos_f_other f_other_pos
local sh_f_other  begin_repeat_lM3hqngUR begin_repeat_tXgwButw2
local cnt_welfare welfare_count
local pos_welfare scheme_pos
local sh_welfare  welfare_rep

foreach b of local blocks {
    local pos `pos_`b''
    local parts
    local i = 0
    foreach s of local sh_`b' {
        local ++i
        import excel using "$SURVEY", sheet("`s'") firstrow case(preserve) clear
        quietly ds _*, not
        local datavars `r(varlist)'
        local datavars : list datavars - pos
        keep _parent_index `pos' `datavars'
        * the position variable is stored as text in the Excel file
        destring `pos', replace
        confirm numeric variable `pos'
        * text everywhere so the two form versions append without type clashes
        foreach v of local datavars {
            capture confirm numeric variable `v'
            if !_rc tostring `v', replace format(%20.10g) force
        }
        tempfile part`i'
        quietly save `"`part`i''"'
        local parts `parts' part`i'          // names of the tempfile macros (paths may contain spaces)
    }
    gettoken first rest : parts
    use `"``first''"', clear
    foreach p of local rest {
        append using `"``p''"'
    }
    * check: positions are 1, 2, 3 ... within each household
    bysort _parent_index (`pos'): assert `pos' == _n
    * back to numbers where every value is numeric; dates back to %td
    quietly destring `datavars', replace
    foreach v of local datavars {
        if strpos("`v'", "date") {
            capture confirm numeric variable `v'
            if !_rc {
                quietly summarize `v'
                if r(max) > 1e8 & r(max) < . replace `v' = dofc(`v')
                format `v' %td
            }
        }
    }
    * reshape wide: stub_ + position -> e.g. a_other_item_1, a_other_item_2
    local stubs
    foreach v of local datavars {
        rename `v' `v'_
        local stubs `stubs' `v'_
    }
    reshape wide `stubs', i(_parent_index) j(`pos')
    * order columns item by item (all fields of item 1, then item 2, ...)
    local ord
    forvalues k = 1/30 {
        foreach v of local datavars {
            capture confirm variable `v'_`k'
            if !_rc local ord `ord' `v'_`k'
        }
    }
    order _parent_index `ord'
    local wide_`b' `ord'
    rename _parent_index _index
    tempfile w_`b'
    quietly save `"`w_`b''"'
}

* main sheet
import excel using "$SURVEY", sheet("Farmer Survey - Round 1") firstrow case(preserve) clear
isid _index
foreach b of local blocks {
    merge 1:1 _index using `"`w_`b''"', nogenerate keep(master match)
    order `wide_`b'', after(`cnt_`b'')
}

* farmer IDs: keep the text ID (one is "502 new"); correct two duplicate submissions
capture confirm string variable farmers_id
if _rc tostring farmers_id, replace format(%12.0f)
* strip all leading / trailing white space: two IDs contain a line break ("\n325", "116\n")
replace farmers_id = ustrregexra(farmers_id, "^\s+|\s+$", "")
assert farmers_id == "165" if _index == 234
replace farmers_id = "166" if _index == 234      // second submission of 165
assert farmers_id == "201" if _index == 16
replace farmers_id = "200" if _index == 16       // first submission of 201
duplicates report farmers_id
gen double farmer_id_num = real(farmers_id)
label var farmer_id_num "Farmer ID as a number (missing for text IDs such as '502 new')"

* payment (RCT_payment_round1.xls: farmer_id, Payment Yes/No)
preserve
    import excel using "$PAY", cellrange(A2) clear
    rename (A B) (farmers_id Payment)
    capture confirm string variable farmers_id
    if _rc tostring farmers_id, replace format(%12.0f)
    replace farmers_id = ustrregexra(farmers_id, "^\s+|\s+$", "")
    replace Payment = ustrregexra(Payment, "^\s+|\s+$", "")
    drop if inlist(farmers_id, "", ".")
    isid farmers_id
    tempfile pay
    save `"`pay'"'
restore
merge m:1 farmers_id using `"`pay'"'
display as text "Payment-list IDs not found in the survey:"
list farmers_id Payment if _merge == 2, noobs clean
drop if _merge == 2
gen byte in_payment_list = _merge == 3
gen byte payment = Payment == "Yes"
gen byte round = 1
drop _merge Payment
order round payment in_payment_list, after(farmers_id)
label var round "Survey round"
label var payment "1 = Payment 'Yes' in RCT_payment_round1; 0 = 'No' or not in the payment list"
label var in_payment_list "1 = farmer appears in RCT_payment_round1 (Yes or No)"
tab payment in_payment_list

* variable labels from the questionnaire (sheet survey: column B = name, C = English label)
preserve
    import excel using "$QUEST", sheet("survey") cellrange(A2) clear allstring
    keep if B != ""
    replace C = ustrregexra(C, "[" + char(36) + "][{][^}]*[}]", "#")
    replace C = subinstr(C, `"""', "'", .)
    replace C = subinstr(C, char(96), "'", .)
    replace C = ustrregexra(C, "\s+", " ")
    local nq = _N
    forvalues i = 1/`nq' {
        local qn`i' = B[`i']
        local ql`i' = usubstr(C[`i'], 1, 70)        // leaves room for " [item k]" within 80 characters
    }
restore
forvalues i = 1/`nq' {
    capture label var `qn`i'' `"`ql`i''"'
    forvalues k = 1/6 {
        capture label var `qn`i''_`k' `"`ql`i'' [item `k']"'
    }
}

*==============================================================================*
* STEP 2 - SECTION 3 COSTS, COST PER ACRE, OUTLIERS
*==============================================================================*

* fixed activities
foreach k of local fixed {
    tokenize `f_`k''
    upcost `1' `2' `3' cost_`k' ix_`k'
}
* fertiliser doses: sum of items; item costs kept as ic_<block>_<item> for the timeline
* The older form version recorded 1st-dose fertiliser as one item in the main sheet
* (c_fert1_name / _qty / _unit / _price / _actdate / _paydate); the newer version as
* up to 2 repeat items (c_fert1_qty_1 ...). Both are added; the main-sheet item is
* kept as ic_c_fert1_0. For the other doses the main-sheet columns do not exist.
foreach k of local fert {
    gen double cost_`k' = 0
    gen byte ix_`k' = 0
    forvalues j = 1/`n_`k'' {
        upcost `k'_qty_`j' `k'_unit_`j' `k'_price_`j' ic_`k'_`j' tmpinc
        replace cost_`k' = cost_`k' + cond(missing(ic_`k'_`j'), 0, ic_`k'_`j')
        replace ix_`k' = 1 if tmpinc
        drop tmpinc
    }
    capture confirm variable `k'_qty `k'_unit `k'_price `k'_name, exact
    if !_rc {
        upcost `k'_qty `k'_unit `k'_price ic_`k'_0 tmpinc
        replace cost_`k' = cost_`k' + cond(missing(ic_`k'_0), 0, ic_`k'_0)
        replace ix_`k' = 1 if tmpinc
        drop tmpinc
    }
    replace cost_`k' = . if ix_`k' & !(cost_`k' > 0)
}
* other items
foreach k of local other {
    gen double cost_`k' = 0
    gen byte ix_`k' = 0
    forvalues j = 1/`n_`k'' {
        othcost `k'_qty_`j' `k'_unit_`j' `k'_amt_`j' `k'_item_`j' ic_`k'_`j' tmpinc
        replace cost_`k' = cost_`k' + cond(missing(ic_`k'_`j'), 0, ic_`k'_`j')
        replace ix_`k' = 1 if tmpinc
        drop tmpinc
    }
    replace cost_`k' = . if ix_`k' & !(cost_`k' > 0)
}

* section totals and total (missing only if every part is missing)
foreach s in A B C D E F {
    local lc = lower("`s'")
    local parts
    foreach k of local order {
        if substr("`k'", 1, 1) == "`lc'" local parts `parts' cost_`k'
    }
    egen double cost_sec`s' = rowtotal(`parts'), missing
    label var cost_sec`s' "Cost (Rs.): Section `s' total - `SEC_`s''"
}
egen double cost_total = rowtotal(cost_secA cost_secB cost_secC cost_secD cost_secE cost_secF), missing
label var cost_total "Cost (Rs.): total cost of rice cultivation, Virippu 2025"
egen byte cost_n_missing = rowtotal(ix_*)
label var cost_n_missing "Activities with an item cost that could not be computed (counted as 0)"
drop ix_*

* cost per acre
local costvars
foreach k of local order {
    local costvars `costvars' cost_`k'
    label var cost_`k' "Cost (Rs.): `L_`k''"
}
local costvars `costvars' cost_secA cost_secB cost_secC cost_secD cost_secE cost_secF cost_total
foreach v of local costvars {
    local s = substr("`v'", 6, .)
    gen double cpa_`s' = `v' / rice_acres_2025 if rice_acres_2025 > 0 & !missing(rice_acres_2025)
    local lab : variable label `v'
    local short = substr(`"`lab'"', 13, .)
    label var cpa_`s' `"Cost per acre (Rs./acre): `short'"'
}

* outliers: |modified z| > 3.5 on log(cost per acre); hi_ = high-side flags (entry errors)
local cpavars
foreach v of local costvars {
    local s = substr("`v'", 6, .)
    local cpavars `cpavars' `s'
    modz cpa_`s' z_`s'
    gen byte out_`s' = abs(z_`s') > 3.5 & !missing(z_`s')
    gen byte hi_`s'  = z_`s' > 3.5 & !missing(z_`s')
}
* households without a usable rice area: screen the raw activity cost instead
foreach k of local order {
    modz cost_`k' zc_`k'
    replace out_`k' = 1 if zc_`k' > 3.5 & !missing(zc_`k') & !(rice_acres_2025 > 0 & !missing(rice_acres_2025))
    replace hi_`k'  = 1 if zc_`k' > 3.5 & !missing(zc_`k') & !(rice_acres_2025 > 0 & !missing(rice_acres_2025))
}
gen byte cpa_total_outlier = out_total
label var cpa_total_outlier "1 = total cost per acre is an outlier (|modified z| > 3.5, log scale)"
local outacts
foreach k of local order {
    local outacts `outacts' out_`k'
}
egen byte cost_outlier_n = rowtotal(`outacts')
label var cost_outlier_n "Number of activities (of 27) with an outlier cost per acre"

* outlier list and cost summary (Excel)
gen long obsno = _n
tempname po ps
tempfile outl summ
postfile `po' str12 farmers_id str60 activity str32 variable str5 direction double(acres cost cpa median times_median z) using `"`outl'"', replace
postfile `ps' str60 activity str32 variable double(n_pos n_zero n_miss mean median p5 p95 min max lower upper) double(n_high n_low) using `"`summ'"', replace
foreach s of local cpavars {
    if inlist("`s'", "secA", "secB", "secC", "secD", "secE", "secF", "total") {
        local lab = cond("`s'" == "total", "TOTAL cost of cultivation", "Section " + substr("`s'", 4, 1) + " total")
    }
    else local lab "`L_`s''"
    quietly summarize cpa_`s' if cpa_`s' > 0, detail
    local med = r(p50)
    local np = r(N)
    local mean = r(mean)
    local p5 = r(p5)
    local p95 = r(p95)
    local mn = r(min)
    local mx = r(max)
    quietly count if cpa_`s' == 0
    local nz = r(N)
    quietly count if missing(cpa_`s')
    local nm = r(N)
    tempvar lg ad
    quietly gen double `lg' = ln(cpa_`s') if cpa_`s' > 0 & !missing(cpa_`s')
    quietly summarize `lg', detail
    local lmed = r(p50)
    quietly gen double `ad' = abs(`lg' - `lmed')
    quietly summarize `ad', detail
    local lmad = r(p50)
    quietly count if out_`s' == 1 & !(z_`s' < -3.5)     // includes raw-cost flags for households without rice area
    local nh = r(N)
    quietly count if z_`s' < -3.5
    local nl = r(N)
    post `ps' ("`lab'") ("cpa_`s'") (`np') (`nz') (`nm') (`mean') (`med') (`p5') (`p95') (`mn') (`mx') ///
        (exp(`lmed' - 3.5 * `lmad' / 0.6745)) (exp(`lmed' + 3.5 * `lmad' / 0.6745)) (`nh') (`nl')
    quietly levelsof obsno if out_`s' == 1, local(rows)
    foreach r of local rows {
        local zz = z_`s'[`r']
        local dir = cond(`zz' > 0 | missing(`zz'), "high", "low")
        post `po' (farmers_id[`r']) ("`lab'") ("cpa_`s'") ("`dir'") (rice_acres_2025[`r']) ///
            (cost_`s'[`r']) (cpa_`s'[`r']) (`med') (cpa_`s'[`r'] / `med') (`zz')
    }
    drop `lg' `ad'
}
postclose `po'
postclose `ps'
drop z_* zc_* obsno

order `costvars' cost_n_missing, after(f_other_paydate_4)
order cpa_*, after(cost_n_missing)
order cost_outlier_n, after(cpa_total_outlier)
order out_* hi_* ic_*, last
compress
save "$OUT/Survey_Round_1_merged_wide.dta", replace
export excel using "$OUT/Survey_Round_1_merged_wide_stata.xlsx", sheet("data") firstrow(variables) replace
preserve
    use `"`summ'"', clear
    export excel using "$OUT/Survey_Round_1_merged_wide_stata.xlsx", sheet("cost_summary", replace) firstrow(variables)
    use `"`outl'"', clear
    gsort direction variable -times_median
    export excel using "$OUT/Survey_Round_1_merged_wide_stata.xlsx", sheet("cost_outliers", replace) firstrow(variables)
restore

*==============================================================================*
* STEP 3 - BASELINE BALANCE OF SOCIO-ECONOMIC CHARACTERISTICS BY PAYMENT
*==============================================================================*
use "$OUT/Survey_Round_1_merged_wide.dta", clear

* analysis-only recodes (the saved dataset is not changed)
gen double age_c = age if age >= 18                       // ages 3 and 6 for a respondent -> missing
gen double hh_members_c = hh_members if hh_members <= 20  // 58 and 61, inconsistent with member counts -> missing
foreach v in exp_food_amt exp_util_amt exp_comm_amt exp_health_amt exp_edu_amt exp_personal_amt exp_misc_amt ///
             inc_rice_vir_amt inc_rice_mun_amt inc_agnet_amt inc_offfarm_amt inc_salary_amt inc_nonfarm_amt ///
             inc_pension_amt inc_other_amt inc_loan_amt bank_distance bank_visits {
    numcopy `v' c_`v'
    replace c_`v' = . if inlist(c_`v', 999, 9999)          // "don't know" codes
    local vl : variable label `v'
    label var c_`v' `"`vl'"'
}
egen byte exp_nmiss = rowmiss(c_exp_food_amt c_exp_util_amt c_exp_comm_amt c_exp_health_amt c_exp_edu_amt c_exp_personal_amt c_exp_misc_amt)
egen double exp_sum = rowtotal(c_exp_food_amt c_exp_util_amt c_exp_comm_amt c_exp_health_amt c_exp_edu_amt c_exp_personal_amt c_exp_misc_amt)
replace exp_sum = . if exp_nmiss > 0

* indicators (as in the Python version, a missing answer counts as 0 unless noted)
gen byte female      = gender == 1
gen byte hh_head     = rel_head == 1
gen byte edu_low     = inlist(education, 1, 2, 3)
gen byte edu_primmid = inlist(education, 4, 5)
gen byte edu_sec     = education == 6
gen byte edu_hsplus  = inlist(education, 10, 11, 12, 13)
gen byte emp_selfag  = employment == 1
gen byte ag_sole     = ag_decision == 1
gen byte hh_sole     = hh_decision == 1
gen byte lease_in    = land_leased_in > 0 if !missing(land_leased_in)
gen byte other_crops = other_crops_count > 0 & !missing(other_crops_count)
gen byte has_loan    = c_inc_loan_amt > 0 if !missing(c_inc_loan_amt)
* select-multiple dummies: Excel header "has_phone/2" becomes has_phone_2 (found by label if named differently)
foreach t in has_phone/2 pay_dealers/3 pay_labourers/3 pay_grocery/3 {
    local nm = strtoname("`t'")
    capture confirm variable `nm', exact
    if _rc {
        foreach v of varlist _all {
            local vl : variable label `v'
            if `"`vl'"' == `"`t'"' local nm `v'
        }
    }
    confirm variable `nm', exact
    local sm_`=strtoname("`t'")' `nm'
}
gen byte pmkisan     = pmkisan_received == 1
gen byte smartphone  = `sm_has_phone_2' == 1
gen byte sms_on      = sms_alert == 1
gen byte upi         = has_upi == 1
gen byte upi_dealers = `sm_pay_dealers_3' == 1
gen byte upi_labour  = `sm_pay_labourers_3' == 1
gen byte upi_grocery = `sm_pay_grocery_3' == 1
label var female "Female respondent (share)"
label var age_c "Age (years)"
label var hh_head "Respondent is household head (share)"
label var edu_low "Education: not literate / below primary (share)"
label var edu_primmid "Education: primary or middle (share)"
label var edu_sec "Education: secondary (share)"
label var edu_hsplus "Education: higher secondary or above (share)"
label var emp_selfag "Self-employed in agriculture (share)"
label var ag_sole "Sole decision-maker: farm investment (share)"
label var hh_sole "Sole decision-maker: household investment (share)"
label var hh_members_c "Household size (members)"
label var lease_in "Leases in land (share)"
label var other_crops "Grows crops other than rice (share)"
label var exp_sum "Total consumption, sum of 7 items (Rs., last 30 days)"
label var has_loan "Has an outstanding loan (share)"
label var pmkisan "Ever received PM-KISAN (share)"
label var smartphone "Owns a smartphone (share)"
label var sms_on "Bank SMS alerts enabled (share)"
label var upi "Has UPI / digital wallet (share)"
label var upi_dealers "Pays input dealers by UPI (share)"
label var upi_labour "Pays labourers by UPI (share)"
label var upi_grocery "Pays grocery shops by UPI (share)"

local balvars female age_c hh_head edu_low edu_primmid edu_sec edu_hsplus emp_selfag ag_sole hh_sole ///
    hh_members_c working_age children migrants offfarm_count ///
    total_area land_owned lease_in land_leased_in rice_acres_2025 rice_prod_2025 rice_acres_2026 other_crops ///
    c_exp_food_amt c_exp_util_amt c_exp_comm_amt c_exp_health_amt c_exp_edu_amt c_exp_personal_amt c_exp_misc_amt exp_sum ///
    c_inc_rice_vir_amt c_inc_rice_mun_amt c_inc_agnet_amt c_inc_offfarm_amt c_inc_salary_amt c_inc_nonfarm_amt ///
    c_inc_pension_amt c_inc_other_amt c_inc_loan_amt has_loan ///
    pmkisan smartphone sms_on upi welfare_count c_bank_distance c_bank_visits upi_dealers upi_labour upi_grocery

tempname pb
tempfile baltab
postfile `pb' str32 variable str80 label double(N_T mean_T sd_T N_C mean_C sd_C diff se p_t normdiff p_mw) using `"`baltab'"', replace
foreach v of local balvars {
    local lab : variable label `v'
    quietly ttest `v', by(payment) unequal          // group 1 = control (0), group 2 = treatment (1)
    local diff = r(mu_2) - r(mu_1)
    local nd = `diff' / sqrt((r(sd_1)^2 + r(sd_2)^2) / 2)
    local NT = r(N_2)
    local mT = r(mu_2)
    local sT = r(sd_2)
    local NC = r(N_1)
    local mC = r(mu_1)
    local sC = r(sd_1)
    local se = r(se)
    local pt = r(p)
    quietly ranksum `v', by(payment)
    local pmw = 2 * normal(-abs(r(z)))
    post `pb' ("`v'") (`"`lab'"') (`NT') (`mT') (`sT') (`NC') (`mC') (`sC') (`diff') (`se') (`pt') (`nd') (`pmw')
}
postclose `pb'

* categorical distributions: chi-square tests
foreach v in education employment rel_head ag_decision hh_decision {
    tab `v' payment, chi2
}

* joint orthogonality test: OLS of payment on 28 baseline covariates
local X female age_c hh_head edu_sec edu_hsplus emp_selfag ag_sole hh_members_c working_age children migrants ///
    offfarm_count total_area land_owned rice_acres_2025 rice_prod_2025 other_crops exp_sum c_inc_rice_vir_amt ///
    c_inc_rice_mun_amt c_inc_loan_amt pmkisan smartphone sms_on upi welfare_count c_bank_distance c_bank_visits
preserve
    regress payment `X'
    keep if e(sample)
    display as result "Conventional F = " e(F) "   p = " Ftail(e(df_m), e(df_r), e(F))
    regress payment `X', vce(hc3)
    display as result "Robust HC3 F = " e(F) "   p = " Ftail(e(df_m), e(df_r), e(F))
    regress payment `X', vce(robust)
    display as result "Robust HC1 F = " e(F) "   p = " Ftail(e(df_m), e(df_r), e(F))
    * randomization inference: re-assign payment at random, 5,000 times
    * (compare the "upper" one-sided p-value with the Python output)
    permute payment F = e(F), reps(5000) rseed(20261026) nodots: regress payment `X'
    permute payment F = e(F), reps(5000) rseed(20261026) nodots: regress payment `X', vce(hc3)
    permute payment F = e(F), reps(5000) rseed(20261026) nodots: regress payment `X', vce(robust)
restore

preserve
    use `"`baltab'"', clear
    gen stars = cond(p_t < 0.01, "***", cond(p_t < 0.05, "**", cond(p_t < 0.10, "*", "")))
    order variable label N_T mean_T sd_T N_C mean_C sd_C diff stars se p_t normdiff p_mw
    export excel using "$OUT/Balance_table_payment_round1_stata.xlsx", sheet("balance_table") firstrow(variables) replace
    count if p_t < 0.10
    display as result r(N) " of " _N " characteristics differ at the 10% level"
restore

*==============================================================================*
* STEP 4 - PER-ACRE COST COMPONENTS, TREATMENT VS CONTROL
*==============================================================================*
use "$OUT/Survey_Round_1_merged_wide.dta", clear

* the older form version recorded 1st-dose fertiliser as one main-sheet item, the newer as up
* to 2 repeat items; both are in cost_c_fert1, so totals are complete in both versions
gen byte older_form = __version__ == "$OLDFORM"
tab older_form payment

tempname pc pr
tempfile cmptab robtab
postfile `pc' str32 variable str80 label double(N_T mean_T sd_T med_T inc_T N_C mean_C sd_C med_C inc_C diff p_t p_mw normdiff p_inc n_excl) using `"`cmptab'"', replace
postfile `pr' str32 variable str80 label double(all_mean_T all_mean_C all_med_T all_med_C all_p_t all_p_mw new_N_T new_N_C new_diff new_p_t) using `"`robtab'"', replace
local cmpvars `order' secA secB secC secD secE secF total
local L_secA "Section A: Pre-nursery preparation"
local L_secB "Section B: Nursery preparation"
local L_secC "Section C: Transplanting"
local L_secD "Section D: Weedicide / fertiliser application"
local L_secE "Section E: Pesticide / fertiliser application"
local L_secF "Section F: Harvesting"
local L_total "TOTAL cost of cultivation"
foreach s of local cmpvars {
    local lab `"`L_`s''"'
    if `"`lab'"' == "" local lab "`s' (cost per acre)"
    quietly count if out_`s' == 1
    local nx = r(N)
    tempvar y inc
    * main comparison: outliers excluded
    quietly gen double `y' = cpa_`s' if out_`s' != 1
    quietly ttest `y', by(payment) unequal          // group 1 = control (0), group 2 = treatment (1)
    local diff = r(mu_2) - r(mu_1)
    local nd = `diff' / sqrt((r(sd_1)^2 + r(sd_2)^2) / 2)
    local NT = r(N_2)
    local mT = r(mu_2)
    local sT = r(sd_2)
    local NC = r(N_1)
    local mC = r(mu_1)
    local sC = r(sd_1)
    local pt = r(p)
    quietly ranksum `y', by(payment)
    local pmw = 2 * normal(-abs(r(z)))
    quietly summarize `y' if payment == 1, detail
    local medT = r(p50)
    quietly summarize `y' if payment == 0, detail
    local medC = r(p50)
    quietly gen byte `inc' = `y' > 0 if !missing(`y')
    quietly summarize `inc' if payment == 1
    local iT = r(mean)
    quietly summarize `inc' if payment == 0
    local iC = r(mean)
    * Fisher's exact test only when some households have no cost (otherwise no p-value)
    local pinc = .
    capture quietly tab `inc' payment, exact
    if !_rc & r(r) == 2 local pinc = r(p_exact)
    post `pc' ("cpa_`s'") (`"`lab'"') (`NT') (`mT') (`sT') (`medT') (`iT') (`NC') (`mC') (`sC') (`medC') (`iC') ///
        (`diff') (`pt') (`pmw') (`nd') (`pinc') (`nx')
    * robustness: (a) all observations, outliers included; (b) newer form version only, outliers excluded
    quietly ttest cpa_`s', by(payment) unequal
    local amT = r(mu_2)
    local amC = r(mu_1)
    local apt = r(p)
    quietly ranksum cpa_`s', by(payment)
    local apmw = 2 * normal(-abs(r(z)))
    quietly summarize cpa_`s' if payment == 1, detail
    local admT = r(p50)
    quietly summarize cpa_`s' if payment == 0, detail
    local admC = r(p50)
    quietly ttest `y' if !older_form, by(payment) unequal
    post `pr' ("cpa_`s'") (`"`lab'"') (`amT') (`amC') (`admT') (`admC') (`apt') (`apmw') ///
        (r(N_2)) (r(N_1)) (r(mu_2) - r(mu_1)) (r(p))
    drop `y' `inc'
}
postclose `pc'
postclose `pr'

* composition: share of each section and input type in the household's total cost
* (households with no flagged total or section cost)
gen byte ok_share = !out_total & !out_secA & !out_secB & !out_secC & !out_secD & !out_secE & !out_secF & cost_total > 0 & !missing(cost_total)
foreach s in secA secB secC secD secE secF {
    gen double sh_`s' = cost_`s' / cost_total if ok_share
    ttest sh_`s', by(payment) unequal
}
egen double lab_cost = rowtotal(cost_b_labour cost_c_mainlab cost_c_translab cost_c_fert1lab cost_c_gapfill cost_d_weedlab cost_d_fert2lab cost_e_pestlab cost_e_fert3lab cost_f_labour)
egen double mach_cost = rowtotal(cost_a_land cost_c_mach cost_f_mach cost_f_tractor)
egen double input_cost = rowtotal(cost_a_seed cost_b_fert cost_d_weed cost_d_fert2 cost_e_pest cost_e_fert3)
egen double oth_cost = rowtotal(cost_a_other cost_b_other cost_c_other cost_d_other cost_e_other cost_f_other)
foreach g in lab mach input oth {
    gen double sh_`g' = `g'_cost / cost_total if ok_share
    ttest sh_`g', by(payment) unequal
}

preserve
    use `"`cmptab'"', clear
    gen stars = cond(p_t < 0.01, "***", cond(p_t < 0.05, "**", cond(p_t < 0.10, "*", "")))
    order variable label N_T mean_T sd_T med_T inc_T N_C mean_C sd_C med_C inc_C diff stars p_t p_mw normdiff p_inc n_excl
    export excel using "$OUT/Cost_comparison_payment_round1_stata.xlsx", sheet("cost_comparison") firstrow(variables) replace
    use `"`robtab'"', clear
    export excel using "$OUT/Cost_comparison_payment_round1_stata.xlsx", sheet("robustness", replace) firstrow(variables)
restore

*==============================================================================*
* STEP 5 - ITEM-LEVEL COSTS WITH DATES, WEEKLY AGGREGATES
*==============================================================================*
use "$OUT/Survey_Round_1_merged_wide.dta", clear
tempfile base items
save `"`base'"'
clear
gen long _index = .
save `"`items'"', emptyok replace

* fixed activities: one item per household
foreach k of local fixed {
    use `"`base'"', clear
    tokenize `f_`k''
    asdate `4' act
    asdate `5' pay
    keep _index payment rice_acres_2025 cost_`k' act pay hi_`k'
    rename (cost_`k' hi_`k') (cost bad)
    gen str12 key = "`k'"
    append using `"`items'"'
    save `"`items'"', replace
}
* repeat blocks: one item per position; item 0 = the older form's single 1st-dose
* fertiliser item in the main sheet (dates c_fert1_actdate / c_fert1_paydate)
foreach k in `fert' `other' {
    forvalues j = 0/`n_`k'' {
        use `"`base'"', clear
        if `j' == 0 {
            capture confirm variable ic_`k'_0, exact
            if _rc continue
            local sfx
        }
        else local sfx _`j'
        asdate `k'_actdate`sfx' act
        asdate `k'_paydate`sfx' pay
        keep _index payment rice_acres_2025 ic_`k'_`j' act pay hi_`k'
        rename (ic_`k'_`j' hi_`k') (cost bad)
        gen str12 key = "`k'"
        append using `"`items'"'
        save `"`items'"', replace
    }
}
use `"`items'"', clear
gen str1 sec = upper(substr(key, 1, 1))
keep if cost > 0 & !missing(cost)
count
display as text "Items flagged as entry errors and left out: " _continue
count if bad == 1
drop if bad == 1

* most dates were saved with year 2026 although they refer to Virippu 2025:
* put every date on the 2025 calendar; 1 January / January dates are placeholders
foreach d in act pay {
    gen int yr = year(`d')
    tab yr, missing
    drop yr
    gen double `d'_n = mdy(month(`d'), day(`d'), 2025)
    replace `d'_n = . if `d'_n < mdy(2, 1, 2025)
    gen double `d'_wk = `d'_n - mod(dow(`d'_n) + 6, 7)          // Monday of the week
    format `d'_n `d'_wk %td
}
save "$OUT/cost_items_dated.dta", replace

* weekly totals by activity (activity date): total spent, farmers, items
preserve
    keep if !missing(act_wk)
    bysort key act_wk _index: gen byte first = _n == 1
    gen byte one = 1
    collapse (sum) total = cost farmers = first items = one, by(key sec act_wk)
    export excel using "$OUT/Weekly_spending_by_activity_stata.xlsx", firstrow(variables) replace
restore

*==============================================================================*
* STEP 6 - BUBBLE PLOT: WEEKLY COST PER ACRE + CUMULATIVE COST PER ACRE
*==============================================================================*
use "$OUT/Survey_Round_1_merged_wide.dta", clear
quietly count if rice_acres_2025 > 0 & !missing(rice_acres_2025)
local NHH = r(N)
display as text "Households with rice area: `NHH'"

* bubbles: per farmer, activity and week, then the mean over the farmers who spent
use "$OUT/cost_items_dated.dta", clear
keep if rice_acres_2025 > 0 & !missing(rice_acres_2025) & !missing(act_wk)
gen double cpa = cost / rice_acres_2025
preserve
    collapse (sum) cpa, by(key sec _index act_wk)
    collapse (mean) cpa (count) farmers = cpa, by(key sec act_wk)
    gen double x = act_wk + 3.5
    tempfile bub
    save `"`bub'"'
restore
* cumulative line: weekly cost per acre summed over farmers / all households with rice area;
* every week filled in (weeks without spending add 0), starting at 0 one week before the first spending
collapse (sum) cpa, by(act_wk)
replace cpa = cpa / `NHH'
set obs `=_N + 1'
summarize act_wk, meanonly
replace act_wk = r(min) - 7 in L
replace cpa = 0 in L
tsset act_wk, delta(7)
tsfill
replace cpa = 0 if missing(cpa)
tsset, clear
sort act_wk
gen double cum = sum(cpa)
gen double xc = act_wk + 7                       // value reached by the end of the week
quietly summarize cum
local endval = r(max)
quietly summarize xc if cum >= `endval' / 2
local halfx = r(min)
display as result "Cost per acre by season end: Rs " %9.0fc `endval' "; half reached in the week ending " %td `halfx' - 1
keep xc cum
append using `"`bub'"'
format x xc %td

local m1 = mdy(2,1,2025)
local ticks
forvalues mo = 2/12 {
    local ticks `ticks' `=mdy(`mo', 1, 2025)'
}
local cA "42 120 214"
local cB "235 104 52"
local cC "27 175 122"
local cD "237 161 0"
local cE "232 123 164"
local cF "0 131 0"

* Bubble size: Stata scales [aw=] markers separately within each plot, so sizes would not be
* comparable across colours. Instead, 8 size classes on one scale for all bubbles
* (marker diameter proportional to the square root of the number of farmers).
quietly summarize farmers, meanonly
local maxf = r(max)
gen byte szc = max(1, round(8 * sqrt(farmers / `maxf'))) if !missing(farmers)
local scat
local p = 0
foreach s in A B C D E F {
    quietly levelsof szc if sec == "`s'", local(cls)
    local lg_`s'
    local best = 99
    foreach c of local cls {
        local ++p
        * legend key: the plot whose size class is closest to the middle of the scale
        if abs(`c' - 4) < `best' {
            local best = abs(`c' - 4)
            local lg_`s' `p'
        }
        local scat `scat' (scatter cpa x if sec == "`s'" & szc == `c', msymbol(O) msize(`=0.7 * `c'') mcolor("`c`s''%75") mlcolor(white) mlwidth(vthin))
    }
}
local pl = `p' + 1
twoway `scat' (line cum xc if cum > 0, sort lcolor(black) lwidth(medthick)), ///
    yscale(log) ylabel(10 50 100 500 1000 5000 10000 50000, format(%9.0fc) angle(0) labsize(small)) ///
    xscale(range(`m1' `=mdy(12,31,2025)')) xlabel(`ticks', format(%tdMon) labsize(small)) ///
    ytitle("Cost per acre (Rs., log scale)") xtitle("Week of activity, 2025") ///
    legend(order(`lg_A' "A. Pre-nursery preparation" `lg_B' "B. Nursery preparation" `lg_C' "C. Transplanting" ///
        `lg_D' "D. Weedicide / 2nd fertiliser dose" `lg_E' "E. Pesticide / 3rd fertiliser dose" `lg_F' "F. Harvesting" ///
        `pl' "Cumulative cost per acre, average farmer") pos(3) cols(1) size(small) region(lcolor(none))) ///
    title("Rice cultivation cost per acre by week and activity, Virippu 2025", size(medium) pos(11)) ///
    subtitle("Bubbles: mean cost per acre among farmers who spent on the activity that week (area ~ farmers, 8 size classes)." ///
        "Line: cumulative cost per acre of the average farmer (`NHH' households).", size(small) pos(11)) ///
    note("Dates placed on the 2025 calendar; entry errors and undated items excluded. Log scale compresses the late rise of the line.", size(vsmall)) ///
    graphregion(color(white)) plotregion(color("252 252 251")) xsize(13) ysize(7.5)
graph export "$OUT/Cost_per_acre_bubble_overlay_stata.png", width(2600) replace
capture graph export "$OUT/Cost_per_acre_bubble_overlay_stata.jpg", width(2600) replace
if _rc display as text "This Stata version cannot export .jpg; use the .png file."

* alternative: cumulative line in its own panel on a linear scale (panels 3 : 1.35 in height)
twoway `scat', yscale(log) ylabel(10 50 100 500 1000 5000 10000, format(%9.0fc) angle(0) labsize(small)) ///
    xscale(range(`m1' `=mdy(12,31,2025)')) xlabel(`ticks', format(%tdMon) labsize(small)) xtitle("") ///
    ytitle("Cost per acre (Rs., log)") ///
    legend(order(`lg_A' "A" `lg_B' "B" `lg_C' "C" `lg_D' "D" `lg_E' "E" `lg_F' "F") rows(1) size(small) region(lcolor(none))) ///
    graphregion(color(white)) plotregion(color("252 252 251")) name(gtop, replace) nodraw
twoway (area cum xc, sort color(gs14)) (line cum xc, sort lcolor(black) lwidth(medthick)), ///
    xscale(range(`m1' `=mdy(12,31,2025)')) xlabel(`ticks', format(%tdMon) labsize(small)) ///
    ylabel(0(10000)50000, format(%9.0fc) angle(0) labsize(small)) ytitle("Cumulative cost per acre (Rs.)") ///
    xtitle("Week of activity, 2025") legend(off) graphregion(color(white)) plotregion(color("252 252 251")) ///
    fysize(45) name(gbot, replace) nodraw
graph combine gtop gbot, cols(1) xcommon imargin(small) graphregion(color(white)) ///
    title("Rice cultivation cost per acre by week and activity, Virippu 2025", size(medium)) xsize(13) ysize(9.4)
graph export "$OUT/Cost_per_acre_bubble_panels_stata.png", width(2600) replace
capture graph export "$OUT/Cost_per_acre_bubble_panels_stata.jpg", width(2600) replace

log close
*==============================================================================* END
