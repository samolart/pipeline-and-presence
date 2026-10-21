# Pipeline and Presence: one-page summary

**What it is.** A working toolkit for a Clients & Markets team, built on simulated Salesforce-style data for a fictional professional-services firm (950 opportunities, 332 accounts, 28 owners, 180 feedback responses, 6 events). Nothing here is real client or firm data.

**The problem.** A BD meeting is only as good as the CRM behind it. The audit finds that only **45%** of the 373 open opportunities are free of data issues.

| Finding (simulated, as of 19 Oct 2026) | Value |
|---|---|
| Open pipeline / probability-weighted | KES 6,816m / KES 2,718m |
| FY27 coverage (weighted FY27 pipeline / (target - won YTD)) | 0.66x overall; Deals 5.3x, Consulting 0.77x, Assurance 0.36x, Tax & Legal 0.26x, Risk 0.24x |
| Win rate, last 12 months | 37.4% |
| Stale (60+ days no activity) / close date passed / owner left | 22% / 15% / 11% of open opportunities |
| Priority accounts with an open opportunity | 38 of 40 |
| Client feedback | 3.64 of 5; 50 responses awaiting follow-up |

**What I built.**
1. **Audit and cleaning** (Python): 8 opportunity and 2 account rules, safe fixes only (no invented amounts), and a 386-row owner follow-up list.
2. **Excel workbook** with live formulas: LoS and Industry dashboards, priority accounts, DQ scorecard, owner hygiene, and a Checks sheet that ties every number back to Python.
3. **Interactive dashboard** (single HTML file) with LoS, industry and owner filters.
4. **Marketing and communications pack:** Salesforce adoption and clinic kit, 4-week content calendar, Tax Alert mailer template, client-feedback report, daily news brief sample (real public news, linked), event pack, and Power Automate flow specs with a Python nudge prototype.

**What it shows I can do.** Turn messy CRM data into numbers a partner can trust, then turn the numbers into habits: training, clinics, nudges, calendars and mailers.

**Limits.** Data, targets and the firm are invented to protect confidentiality; the pipeline is reproducible (seeded) and the rules transfer to a real CRM. Probabilities use stage defaults, not a calibrated model.
