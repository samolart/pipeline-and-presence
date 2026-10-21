# Methods appendix

**Data.** `analysis/gen_data.py` (seed 20261021) writes seven CSVs with deliberately injected quality issues: missing or zero amounts, wrong probabilities, no next step, stale activity, lost without reason, no owner, inactive owner, missing industry, and 12 duplicate accounts.

**Rules (open opportunities unless stated).** Missing or zero amount (all); no owner (all); owner inactive; close date before 19 Oct 2026; no activity for more than 60 days; probability differs from the stage default (10/25/50/75); no next step; lost without reason (closed lost). Accounts: missing industry; duplicate normalised name.

**Cleaning.** Probability reset to the stage default; duplicate accounts mapped to a master record by normalised name; industry taken from the master else "Unassigned". Amounts are never imputed.

**KPIs.** Weighted pipeline = amount x stage probability. FY27 coverage = weighted pipeline closing in FY27 / (FY27 target - won YTD). Win rate = won / (won + lost) over closed deals in the last 12 months. Cycle time = created to closed for won deals.

**Checks.** The Excel workbook recomputes every headline with formulas (COUNTIFS, SUMIFS, AVERAGEIFS, INDEX-MATCH) and the Checks sheet compares them to the Python results; all pass. The dashboard's KPIs were compared to Python as well.

**Reproduce.**
```
pip install -r requirements.txt
python analysis/gen_data.py
python analysis/audit.py
python analysis/build_workbook.py   # then recalculate in LibreOffice
python analysis/build_dashboard.py
python analysis/build_comms.py
python analysis/make_nudges.py
```
