"""Prototype of the weekly nudge flow: one email per owner from outputs/owner_followups.csv. Drafts only, nothing is sent."""
import os
import pandas as pd
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
fu = pd.read_csv(os.path.join(ROOT, "outputs", "owner_followups.csv"), keep_default_na=False)
os.makedirs(os.path.join(ROOT, "automation", "nudges"), exist_ok=True)
out = ["# Draft owner nudges (generated; simulated data)\n"]
for owner, d in fu.groupby("owner"):
    d = d.sort_values(["priority", "amount_kes"], ascending=[True, False])
    out.append(f"## To: {owner}\n**Subject:** {len(d)} Salesforce updates needed before the BD meeting\n\nHello {owner.split()[0] if owner != '(no owner)' else 'team'},\n\nThese records need a quick update so the pipeline dashboard is accurate:\n\n| Opportunity | Issue | What to do | Priority |\n|---|---|---|---|")
    for r in d.head(8).itertuples():
        out.append(f"| {str(r.opp_name).replace("|","-")} ({r.opp_id}) | {r.issue} | {r.action} | {r.priority} |")
    if len(d) > 8: out.append(f"\n...and {len(d)-8} more in the attached list.")
    out.append("\nThank you. The next Salesforce clinic is on the first Thursday of the month.\n\nCMD team\n")
open(os.path.join(ROOT, "automation", "nudges", "owner_nudges.md"), "w").write("\n".join(out))
print(fu.owner.nunique(), "owners")
