# Pipeline and Presence

A CRM analytics and marketing-operations toolkit for a Clients & Markets team, built on **simulated** Salesforce-style data for a fictional firm. No real client, firm or staff data is used.

**Start here:** `docs/executive_summary.md` (one page), then **[View Presentation](https://samolart.github.io/pipeline-and-presence/)** (open in a browser) and `model/pipeline_presence_model.xlsx`.

| Folder | Contents |
|---|---|
| `analysis/` | Data generator, audit, workbook and dashboard builders, comms and nudge scripts |
| `data/` | Simulated raw CSVs and cleaned versions |
| `outputs/` | Audit results (JSON), owner follow-up list, priority accounts |
| `model/` | Excel workbook with live formulas and a Checks sheet |
| `dashboard/` | Interactive single-file HTML dashboard |
| `comms/` | Salesforce adoption kit, content calendar, Tax Alert mailer, feedback report, daily news brief, event pack |
| `automation/` | Power Automate flow specs and generated owner nudges |
| `docs/` | Summary, methods, CV bullets, cover letter, interview Q&A, 30-60-90 plan |

Reproduce: see `docs/methods_appendix.md`. Seeded, so results are identical on every run.

**Limits:** simulated data and targets; stage-default probabilities; Power Automate flows are specifications, not exported flows. The daily news brief uses real public items from October 2026, each linked to its source.
