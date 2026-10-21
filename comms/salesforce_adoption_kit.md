# Salesforce adoption kit for Lines of Service and Industry Groups

*Built from the data-quality audit in this repository (simulated data). The aim: make the CRM trustworthy enough that the monthly BD meeting can run straight from the dashboard.*

## 1. Where we are today (simulated baseline, 2026-10-19)
- **45%** of 373 open opportunities have no data issue. Target: 85% within one quarter.
- Biggest issues: no activity for more than 60 days (22%), close date already passed (15%), owner has left (11%), no next step (8%).
- 17 accounts have no industry and 24 sit in duplicate pairs, which distorts the industry dashboards.
- 386 individual fixes are listed in `outputs/owner_followups.csv`.

## 2. Data-entry standards (one page to share with all owners)
| Field | Standard |
|---|---|
| Amount | Always an estimate in KES, never blank or zero. Update when the proposal is sent. |
| Stage and probability | Use the stage default. Do not overwrite probability by hand. |
| Close date | Always in the future for open opportunities. If the date passes, move it or close the opportunity the same week. |
| Next step | One concrete action with a date. Required from Proposal stage. |
| Owner | A current employee. When someone leaves, reassign within five working days. |
| Lost reason | Required on every Closed Lost opportunity. |
| Account | Search before creating. Industry is mandatory. Use the "Ltd" naming standard. |
| Activity | Log meetings and calls the same week. No activity for 60 days flags the opportunity as stale. |

## 3. Training session plan (bi-annual, 90 minutes)
| Time | Segment | Who leads | Output |
|---|---|---|---|
| 0:00 | Why the CRM matters: the BD meeting runs from the dashboard | Senior Manager | Shared purpose |
| 0:10 | Live walk-through of the LoS dashboard and what each number means | Associate | Everyone can read the dashboard |
| 0:30 | Hands-on: update five of your own opportunities | Everyone | Fixes made in the session |
| 1:00 | The five most common mistakes and how to avoid them | Senior Associate | Checklist |
| 1:15 | Quick wins: mobile app, reports, follow-up tasks | Associate | Habits |
| 1:25 | Questions and next clinic date | Manager | Calendar invite |
Materials: slide pack (10 slides), data-entry one-pager, before-and-after screenshots, short quiz, feedback form.
Internal communications: invitation two weeks before, reminder two days before, recording and one-pager within 24 hours, leaderboard of clean rate by Line of Service one month later.

## 4. Salesforce clinic (30 minutes, monthly, per Line of Service)
1. Open with the LoS data-quality number and the three most common issues.
2. Take live questions and fix records on screen.
3. Log every question in a shared tracker with owner, date, answer and whether it needs escalating to the CRM administrator.

**Common troubleshooting**
| Problem | Likely cause | Fix |
|---|---|---|
| "My opportunity is missing from the dashboard" | Wrong stage, close date outside the filter, or you are not the owner | Check the three fields; correct or reassign |
| Numbers on the dashboard do not match my list | Report filter or refresh time | Confirm the filter and the refresh timestamp |
| Cannot create an account | Possible duplicate | Search by name stem; ask the administrator to merge |
| Probability keeps resetting | Stage default applies | Change the stage, not the probability |
| Old opportunity will not close | Missing lost reason or amount | Fill required fields, then close |

## 5. Dashboard and report specifications
| Dashboard | Audience | Key components | Filters | Refresh |
|---|---|---|---|---|
| LoS pipeline | LoS leader, BD meeting | Open and weighted pipeline by stage; FY coverage vs target; win rate; cycle time; data-clean rate | LoS, owner, FY | Daily |
| Industry group | Industry leader | Weighted pipeline by industry; priority accounts with open opportunities; last touch | Industry, tier | Daily |
| Data quality | CMD team | Eight rule failures, by owner and LoS; trend | LoS, owner | Daily |
| Priority accounts | Industry business partner | Open opps, weighted value, days since touch; accounts needing outreach | Industry | Daily |
Reconstructed with live formulas in `model/pipeline_presence_model.xlsx` (sheets Dashboard_LoS, Dashboard_Industry, Priority_Accounts, DQ_Scorecard).

## 6. How adoption is measured
| Measure | Today (simulated) | Target in 90 days |
|---|---|---|
| Open opportunities with no issue | 45% | 85% |
| Open opportunities stale | 22% | under 10% |
| Close date passed | 15% | under 3% |
| Owner has left | 11% | 0% |
| Accounts missing industry | 17 | 0 |
| Clinic attendance | n/a | 80% of owners with open opportunities |

## 7. Follow-up routine for the Associate
1. Monday: refresh the data-quality scorecard and the owner list.
2. Tuesday: send each owner their personal list (see `automation/nudges`).
3. Thursday: clinic or one-to-one for owners with more than five issues.
4. Friday: update the clean-rate trend; report to the Senior Associate.
5. Before each monthly BD meeting: confirm numbers, prepare the dashboard slide, list open data issues.
