# Power Automate flow specifications

*Design documents for three flows that remove repetitive CMD work. They are specifications, not exported flows: build and test in the firm's environment with the right connectors and approvals.*

## Flow 1: Weekly stale-opportunity nudge
- **Trigger:** recurrence, Tuesday 08:00 (EAT).
- **Data:** Salesforce report "Open opportunities with an issue" (rules in `analysis/audit.py`).
- **Steps:** (1) Get report rows. (2) Group by owner. (3) For each owner build an HTML table: opportunity, issue, action. (4) Send an email from the CMD shared mailbox; copy the Senior Associate when an owner has more than five issues. (5) Append a row to a SharePoint list for the weekly clean-rate trend.
- **Controls:** do not email owners marked inactive; send those to the CRM administrator for reassignment.
- **Success:** clean rate rises from 45% toward 85%.
- Python prototype: `analysis/make_nudges.py`.

## Flow 2: Daily news digest
- **Trigger:** weekdays 07:00.
- **Steps:** (1) Pull items from approved RSS feeds and the news-monitoring tool. (2) Filter by keyword lists per industry group. (3) Post the candidate items to a Teams channel for the Associate to review and rank. (4) After approval (reaction or form), merge into the mailer template and send to the firmwide list.
- **Control:** a human approves before anything is sent.

## Flow 3: CMD request intake
- **Trigger:** new item in the CMD request portal (SharePoint form).
- **Steps:** (1) Classify by type (event, thought leadership, data request, social). (2) Assign to the on-duty Associate. (3) Acknowledge to the requester with the expected turnaround. (4) Chase at 80% of the turnaround time. (5) Close with a short satisfaction question.
- **Success:** requests answered within the agreed time; fewer follow-up emails.
