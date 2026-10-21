"""Step 5. Marketing and communications pack, generated from outputs/audit_results.json and the real news items fetched on 21 Oct 2026."""
import json, os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = json.load(open(os.path.join(ROOT, "outputs", "audit_results.json")))
OUT = os.path.join(ROOT, "comms"); os.makedirs(OUT, exist_ok=True); os.makedirs(os.path.join(ROOT, "automation"), exist_ok=True)
W = lambda name, s: open(os.path.join(OUT, name), "w").write(s)
p = lambda v, d=0: f"{v*100:.{d}f}%"
L = {b["los"]: b for b in A["by_los"]}; I = {b["industry"]: b for b in A["by_industry"]}
dq = {d["rule"]: d for d in A["dq"]}
ev = A["events"]; held = [e for e in ev if e["attended"]]
inv, reg, att, opp = (sum(e[k] for e in held) for k in ("invited", "registered", "attended", "opps_created"))
conf = [e for e in ev if "planned" in e["event_name"]][0]
fb = A["feedback"]
tot = A["open"]

# ---------------------------------------------------------------- 1. daily news brief (REAL news)
W("daily_news_brief_sample.md", f"""# Daily firmwide brief: Wednesday 21 October 2026 (sample)

*Format: five items, each with a one-line "why it matters" and the industry groups to forward it to. Built from public items published 7 to 19 October 2026. Check every figure against the linked source before sending a real brief.*

## 1. Central Bank holds its rate at 8.75%
The Monetary Policy Committee kept the Central Bank Rate at 8.75% at its 7 October 2026 meeting, the same level as in August. Banking-sector rates in August: lending 14.34%, deposit 6.91%, savings 3.54%.
**Why it matters:** borrowing costs are steady, which helps planning for clients with variable-rate debt and for capital-project decisions.
**Forward to:** Financial Services, Real Estate & Infrastructure, Consumer & Retail.
Sources: [Central Bank of Kenya statement](https://www.centralbank.go.ke/2026/10/07/mpc-retains-the-cbr-at-8-75-percent-3/) | [People Daily](https://peopledaily.digital/business/cbk-holds-interest-rate-at-8-75-as-kenya-growth-forecast-rises-to-5)

## 2. Inflation edges up to 6.8% in September
Annual inflation rose to 6.8% from 6.6% in August, inside the 2.5% to 7.5% target range. Core inflation rose to 4.0% from 3.4%; non-core fell to 14.0% from 14.7%.
**Why it matters:** the rise is led by core prices, so cost pressure is broadening beyond food and fuel. Useful context for pricing conversations.
**Forward to:** Consumer & Retail, Agribusiness, Public Sector & Development.
Source: [Nairobi Business Monthly, 16 Oct 2026](https://nairobibusinessmonthly.com/kenya-inflation-hits-6-8pc-as-growth-forecast-rises/)

## 3. Growth forecast raised to 5.0% for 2026
The CBK raised its 2026 growth forecast to 5.0% from 4.9%, led by industry and services, and projects 5.3% for 2027. Growth in the first quarter of 2026 was 5.3%.
**Why it matters:** a firmer outlook supports client investment plans and demand for advisory work.
**Forward to:** all industry groups.
Source: [Nairobi Business Monthly, 16 Oct 2026](https://nairobibusinessmonthly.com/kenya-inflation-hits-6-8pc-as-growth-forecast-rises/)

## 4. Risks the CBK is watching
El Nino weather could disrupt agriculture, food supply and prices. Geopolitical tension in the Middle East could push up oil prices and raise transport and production costs.
**Why it matters:** both are natural topics for client conversations on resilience and supply chains.
**Forward to:** Agribusiness, Energy & Resources, Consumer & Retail.
Source: [Nairobi Business Monthly, 16 Oct 2026](https://nairobibusinessmonthly.com/kenya-inflation-hits-6-8pc-as-growth-forecast-rises/)

## 5. Markets snapshot (19 October 2026)
USD 129.92, GBP 173.14, EUR 145.79 (Kenya shillings). 91-day Treasury bill 8.750%.
Source: [Central Bank of Kenya](https://www.centralbank.go.ke/2026/10/07/mpc-retains-the-cbr-at-8-75-percent-3/)

---
**From the CMD team:** the Budget Breakfast follow-up and the Annual Industry Conference (19 November) invitations are going out this week. Reply to this mail with questions.
""")

# ---------------------------------------------------------------- 2. content calendar (xlsx)
cal = [
 ("2026-10-26", "Mon", "LinkedIn", "Thought leadership", "All industries", "Rates held at 8.75%: what steady borrowing costs mean for 2027 planning", "TL-01", "Partner review", "Draft"),
 ("2026-10-27", "Tue", "X", "Thought leadership", "All industries", "Three charts: inflation, policy rate and growth outlook in one view", "TL-02", "Senior Associate", "Draft"),
 ("2026-10-28", "Wed", "LinkedIn", "Industry insight", "Consumer & Retail", "Core inflation is rising: pricing questions for retailers", "IN-01", "Industry partner", "Idea"),
 ("2026-10-29", "Thu", "LinkedIn", "Event promotion", "All industries", "Save the date: Annual Industry Conference, 19 November", "EV-01", "Events lead", "Approved"),
 ("2026-10-30", "Fri", "X", "Insight", "Agribusiness", "El Nino watch: what to plan for in supply chains", "IN-02", "Industry partner", "Idea"),
 ("2026-11-02", "Mon", "LinkedIn", "Thought leadership", "Financial Services", "Customer data and risk in digital lending: five questions for boards", "TL-03", "Partner review", "Idea"),
 ("2026-11-03", "Tue", "LinkedIn", "People and culture", "Firmwide", "Meet the team: how our graduates build client skills in year one", "PC-01", "HR and Creative", "Idea"),
 ("2026-11-04", "Wed", "X", "Event promotion", "All industries", "Speakers announced for the Annual Industry Conference", "EV-02", "Events lead", "Idea"),
 ("2026-11-05", "Thu", "LinkedIn", "Industry insight", "Energy & Resources", "Energy transition: lessons from our September forum", "IN-03", "Industry partner", "Idea"),
 ("2026-11-06", "Fri", "LinkedIn", "Tax alert", "All industries", "Tax alert: summary of the latest regulatory update [insert measure]", "TA-01", "Tax partner", "Idea"),
 ("2026-11-09", "Mon", "LinkedIn", "Thought leadership", "Public Sector & Development", "Public finance clinic: three takeaways for county finance teams", "TL-04", "Industry partner", "Idea"),
 ("2026-11-10", "Tue", "X", "Insight", "Health", "Health leaders dinner: what we heard", "IN-04", "Industry partner", "Idea"),
 ("2026-11-12", "Thu", "LinkedIn", "Event promotion", "All industries", "One week to go: register for the Annual Industry Conference", "EV-03", "Events lead", "Idea"),
 ("2026-11-16", "Mon", "LinkedIn", "Event promotion", "All industries", "Agenda and what to expect on the day", "EV-04", "Events lead", "Idea"),
 ("2026-11-19", "Thu", "LinkedIn", "Event live", "All industries", "Live from the Annual Industry Conference: key quotes and photos", "EV-05", "Events lead", "Idea"),
 ("2026-11-20", "Fri", "LinkedIn", "Event follow-up", "All industries", "Thank you and the three themes from the conference", "EV-06", "Partner review", "Idea")]
wb = Workbook(); ws = wb.active; ws.title = "Calendar"
hd = ["Date", "Day", "Channel", "Content pillar", "Industry or audience", "Headline / post idea", "Approved template ID", "Approver", "Status"]
F = "Arial"; ws.append(hd)
for c in ws[1]: c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="12202B"); c.alignment = Alignment(wrap_text=True, vertical="center")
for r in cal: ws.append([pd.Timestamp(r[0]).to_pydatetime()] + list(r[1:]))
for row in ws.iter_rows(min_row=2):
    row[0].number_format = "yyyy-mm-dd"
    for c in row: c.font = Font(name=F, size=10)
dv = DataValidation(type="list", formula1='"Idea,Draft,Partner review,Approved,Scheduled,Published"', allow_blank=False); ws.add_data_validation(dv); dv.add(f"I2:I{len(cal)+1}")
for k, v in {"A": 12, "B": 7, "C": 11, "D": 20, "E": 26, "F": 70, "G": 12, "H": 18, "I": 14}.items(): ws.column_dimensions[k].width = v
ws.freeze_panes = "A2"
s2 = wb.create_sheet("Status summary")
s2.append(["Status", "Posts"]); 
for i, st in enumerate(["Idea", "Draft", "Partner review", "Approved", "Scheduled", "Published"]): s2.append([st, f'=COUNTIF(Calendar!$I$2:$I${len(cal)+1},A{i+2})'])
s2.append(["Total", "=SUM(B2:B7)"]); s2.append([]); s2.append(["Pillar", "Posts"])
for i, pl in enumerate(sorted({r[3] for r in cal})): s2.append([pl, f'=COUNTIF(Calendar!$D$2:$D${len(cal)+1},A{10+i})'])
for row in s2.iter_rows():
    for c in row: c.font = Font(name=F, size=10, bold=(c.row in (1, 8, 9)))
s2.column_dimensions["A"].width = 24
ws3 = wb.create_sheet("How to use")
for t in ["Workflow: Idea > Draft > Partner review > Approved > Scheduled > Published.", "Every post uses an approved template ID from the Creative team so the brand stays consistent across channels.",
          "Thought-leadership and tax items need partner approval before they are shared or sent to clients.", "Add a UTM tag to every link (utm_source=channel, utm_campaign=pillar) so Salesforce and web analytics can attribute leads.",
          "This calendar is a sample. Dates, topics and approvers are illustrative; templates are placeholders."]: ws3.append([t])
ws3.column_dimensions["A"].width = 130
wb.save(os.path.join(OUT, "content_calendar.xlsx"))

# ---------------------------------------------------------------- 3. tax alert mailer (HTML email template)
W("tax_alert_mailer_template.html", """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Tax Alert template</title></head>
<body style="margin:0;background:#EEF1F0;font-family:Arial,Helvetica,sans-serif;color:#102A30">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:16px">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;background:#fff;border:1px solid #D5DDDB">
<tr><td style="background:#12202B;color:#fff;padding:14px 20px;font-size:12px;letter-spacing:1px;text-transform:uppercase">[Firm name] | Tax Alert | [Date]</td></tr>
<tr><td style="padding:22px 20px 8px"><h1 style="margin:0;font-size:22px;line-height:1.25">[Headline: what changed and who it affects]</h1></td></tr>
<tr><td style="padding:8px 20px;font-size:15px;line-height:1.5"><p style="margin:0 0 10px"><b>In short:</b> [One or two sentences: the change, the effective date, the main impact.]</p>
<p style="margin:0 0 10px"><b>What changes</b></p><ul style="margin:0 0 10px;padding-left:20px"><li>[Change 1]</li><li>[Change 2]</li><li>[Change 3]</li></ul>
<p style="margin:0 0 10px"><b>What you should do</b></p><ul style="margin:0 0 10px;padding-left:20px"><li>[Action 1 with date]</li><li>[Action 2]</li></ul></td></tr>
<tr><td style="padding:8px 20px 20px"><a href="[link with UTM tags]" style="display:inline-block;background:#E4572E;color:#fff;text-decoration:none;padding:10px 16px;font-size:14px">Read the full alert</a></td></tr>
<tr><td style="padding:14px 20px;border-top:1px solid #D5DDDB;font-size:13px">Contact: [Name, title, email, phone] &nbsp;|&nbsp; [Name, title, email, phone]</td></tr>
<tr><td style="padding:12px 20px;background:#EEF1F0;font-size:11px;color:#5B6E72;line-height:1.4">This alert is general information and not advice for your specific situation. [Approved by: partner name, date]. To stop receiving alerts, [unsubscribe link].</td></tr>
</table></td></tr></table></body></html>
""")

# ---------------------------------------------------------------- 4. Salesforce adoption kit
W("salesforce_adoption_kit.md", f"""# Salesforce adoption kit for Lines of Service and Industry Groups

*Built from the data-quality audit in this repository (simulated data). The aim: make the CRM trustworthy enough that the monthly BD meeting can run straight from the dashboard.*

## 1. Where we are today (simulated baseline, {A['asof']})
- **{p(A['open_clean_rate'])}** of {tot['open_n']} open opportunities have no data issue. Target: 85% within one quarter.
- Biggest issues: no activity for more than {A['stale_days']} days ({p(dq['f_stale']['rate'])}), close date already passed ({p(dq['f_past_close']['rate'])}), owner has left ({p(dq['f_inactive_owner']['rate'])}), no next step ({p(dq['f_no_next_step']['rate'])}).
- {A['acc_dq'][0]['failing']} accounts have no industry and {A['acc_dq'][1]['failing']} sit in duplicate pairs, which distorts the industry dashboards.
- {A['followups_n']} individual fixes are listed in `outputs/owner_followups.csv`.

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
| Open opportunities with no issue | {p(A['open_clean_rate'])} | 85% |
| Open opportunities stale | {p(dq['f_stale']['rate'])} | under 10% |
| Close date passed | {p(dq['f_past_close']['rate'])} | under 3% |
| Owner has left | {p(dq['f_inactive_owner']['rate'])} | 0% |
| Accounts missing industry | {A['acc_dq'][0]['failing']} | 0 |
| Clinic attendance | n/a | 80% of owners with open opportunities |

## 7. Follow-up routine for the Associate
1. Monday: refresh the data-quality scorecard and the owner list.
2. Tuesday: send each owner their personal list (see `automation/nudges`).
3. Thursday: clinic or one-to-one for owners with more than five issues.
4. Friday: update the clean-rate trend; report to the Senior Associate.
5. Before each monthly BD meeting: confirm numbers, prepare the dashboard slide, list open data issues.
""")

# ---------------------------------------------------------------- 5. event pack
rr, ar, cr = reg / inv, att / reg, opp / att
W("event_pack.md", f"""# Event pack: Annual Industry Conference, 19 November 2026 (planned, simulated)

## 1. Planning assumptions from the five events already held this FY
| Measure | Value | How it is used |
|---|---|---|
| Invited, registered, attended | {inv}, {reg}, {att} | Funnel |
| Registration rate | {p(rr,0)} | Invitations needed |
| Attendance rate | {p(ar,0)} | Seats and catering |
| Opportunities per 10 attendees | {cr*10:.1f} | Success target |

For {conf['invited']} invitations: about {round(conf['invited']*rr)} registrations, {round(conf['invited']*rr*ar)} attendees and {round(conf['invited']*rr*ar*cr)} new opportunities. Plan catering for {round(conf['invited']*rr*0.9)} and set the registration cap just above it.

## 2. Run-of-show (half day)
| Time | Item | Owner |
|---|---|---|
| 08:00 | Registration and coffee | Events team |
| 08:45 | Welcome and the year in numbers | Senior Partner |
| 09:00 | Keynote: outlook for Kenya's economy | Guest speaker |
| 09:45 | Panels by industry (three parallel rooms) | Industry leaders |
| 11:15 | Break and client meetings | All |
| 11:45 | Plenary: lessons from this year's forums | Partners |
| 12:30 | Close, feedback QR code | MC |
| 12:35 | Lunch and networking | All |

## 3. Countdown checklist
| When | Task | Owner | Done |
|---|---|---|---|
| T-8 weeks | Confirm venue, date, budget and speakers | Manager | [ ] |
| T-6 weeks | Invitation list from Salesforce (priority and key accounts first); remove duplicates and inactive contacts | Associate | [ ] |
| T-5 weeks | Save-the-date sent and posted | Associate / Creative | [ ] |
| T-4 weeks | Agenda and speaker biographies final | Senior Associate | [ ] |
| T-3 weeks | Formal invitations; registration page tested | Associate | [ ] |
| T-2 weeks | Reminder 1; review registrations vs {round(conf['invited']*rr)} target | Associate | [ ] |
| T-1 week | Reminder 2 to non-responders; speaker briefings; run slides and signage | All | [ ] |
| T-2 days | Final headcount to venue; name badges; partner client lists by industry | Associate | [ ] |
| Event day | Check-in; capture attendance in Salesforce; photos and quotes for social | Team | [ ] |
| T+1 day | Thank-you email to attendees, "sorry we missed you" to registrants | Associate | [ ] |
| T+3 days | Attendance uploaded to Salesforce; opportunities and follow-up tasks created | Associate | [ ] |
| T+7 days | Feedback summary and conversion report to partners | Associate | [ ] |

## 4. Success measures
| Measure | Target | Source |
|---|---|---|
| Registration rate | {p(rr,0)} or better | Registration tool |
| Attendance rate | {p(ar,0)} or better | Check-in list |
| Priority accounts attending | at least 30 of 40 | Salesforce |
| New opportunities within 60 days | {round(conf['invited']*rr*ar*cr)} or more | Salesforce |
| Satisfaction | 4.2 out of 5 or better | Feedback form |

## 5. Budget template (fill in with real quotes)
| Line | Quote | Notes |
|---|---|---|
| Venue and AV | | |
| Catering | | per head x {round(conf['invited']*rr*0.9)} |
| Speakers and travel | | |
| Printing and signage | | |
| Photography and video | | |
| Contingency (10%) | | |
""")

# ---------------------------------------------------------------- 6. client feedback report
bt = pd.DataFrame(fb["by_theme"]).set_index("theme")
lo = sorted(fb["by_los"], key=lambda r: r["mean"])
W("client_feedback_report.md", f"""# Client feedback consolidation (simulated survey)

**Headline.** Average rating {fb['avg']:.2f} out of 5 across {fb['n']} responses. Net sentiment is {fb['nps_like']:.0f} points (share rating 4 or 5 minus share rating 1 or 2). {fb['open_followups']} responses still need a follow-up.

## Ratings by Line of Service
| Line of service | Responses | Average rating |
|---|---|---|
""" + "\n".join(f"| {r['line_of_service']} | {r['count']} | {r['mean']:.2f} |" for r in lo) + f"""

## What clients talk about (themes tagged from comments)
| Theme | Positive | Neutral | Negative |
|---|---|---|---|
""" + "\n".join(f"| {t} | {int(r.get('Positive',0))} | {int(r.get('Neutral',0))} | {int(r.get('Negative',0))} |" for t, r in bt.iterrows()) + f"""

## What to do
1. **Close the loop:** contact the {fb['open_followups']} open responses within 10 working days, starting with ratings of 1 or 2. Report responses and outcomes to managers weekly.
2. **Share wins and gaps with partners:** the lowest-rated Line of Service is {lo[0]['line_of_service']} ({lo[0]['mean']:.2f}); the highest is {lo[-1]['line_of_service']} ({lo[-1]['mean']:.2f}).
3. **Feed themes into marketing:** use positive quotes (with permission) in thought leadership; review the negative comments under communication (the largest negative theme) and agree one fix per Line of Service.
4. **Track the same measures each quarter** so changes are visible.

*Method: keyword tagging of comments into five themes (analysis/audit.py). It is a first pass; a human review of the "negative" comments is part of the routine.*
""")

# ---------------------------------------------------------------- 7. Power Automate flow specs
open(os.path.join(ROOT, "automation", "power_automate_flows.md"), "w").write(f"""# Power Automate flow specifications

*Design documents for three flows that remove repetitive CMD work. They are specifications, not exported flows: build and test in the firm's environment with the right connectors and approvals.*

## Flow 1: Weekly stale-opportunity nudge
- **Trigger:** recurrence, Tuesday 08:00 (EAT).
- **Data:** Salesforce report "Open opportunities with an issue" (rules in `analysis/audit.py`).
- **Steps:** (1) Get report rows. (2) Group by owner. (3) For each owner build an HTML table: opportunity, issue, action. (4) Send an email from the CMD shared mailbox; copy the Senior Associate when an owner has more than five issues. (5) Append a row to a SharePoint list for the weekly clean-rate trend.
- **Controls:** do not email owners marked inactive; send those to the CRM administrator for reassignment.
- **Success:** clean rate rises from {p(A['open_clean_rate'])} toward 85%.
- Python prototype: `analysis/make_nudges.py`.

## Flow 2: Daily news digest
- **Trigger:** weekdays 07:00.
- **Steps:** (1) Pull items from approved RSS feeds and the news-monitoring tool. (2) Filter by keyword lists per industry group. (3) Post the candidate items to a Teams channel for the Associate to review and rank. (4) After approval (reaction or form), merge into the mailer template and send to the firmwide list.
- **Control:** a human approves before anything is sent.

## Flow 3: CMD request intake
- **Trigger:** new item in the CMD request portal (SharePoint form).
- **Steps:** (1) Classify by type (event, thought leadership, data request, social). (2) Assign to the on-duty Associate. (3) Acknowledge to the requester with the expected turnaround. (4) Chase at 80% of the turnaround time. (5) Close with a short satisfaction question.
- **Success:** requests answered within the agreed time; fewer follow-up emails.
""")
print("comms written")
