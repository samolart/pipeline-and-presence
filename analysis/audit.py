"""Step 2. Data-quality audit, safe cleaning and pipeline KPIs (Python reference).
The Excel model recomputes the same flags and KPIs with live formulas; the Checks sheet compares the two."""
import json, os, re
import numpy as np, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW, PROC, OUT = (os.path.join(ROOT, d) for d in ("data/raw", "data/processed", "outputs"))
ASOF = pd.Timestamp("2026-10-19")
STALE_DAYS = 60
FY27 = (pd.Timestamp("2026-07-01"), pd.Timestamp("2027-06-30"))
STAGE_PROB = {"Prospecting": 10, "Qualification": 25, "Proposal": 50, "Negotiation": 75, "Closed Won": 100, "Closed Lost": 0}
rd = lambda n, dates=(): pd.read_csv(os.path.join(RAW, f"{n}.csv"), parse_dates=list(dates), keep_default_na=False)
owners = rd("owners"); acc = rd("accounts", ["created_date"]); op = rd("opportunities", ["created_date", "close_date", "last_modified", "last_activity_date"])
act = rd("activities", ["activity_date"]); fb = rd("feedback", ["survey_date"]); ev = rd("events"); tg = rd("targets")
owners["active"] = owners.active.astype(str).str.lower().eq("true")
op["amount_kes"] = pd.to_numeric(op.amount_kes.replace("", np.nan))
op["probability"] = pd.to_numeric(op.probability)
inactive = set(owners[~owners.active].owner_id)

# ---------------- accounts: duplicates and missing industry
def norm(n):
    n = n.lower()
    n = re.sub(r"\b(limited|ltd)\b", "", n)
    n = n.replace("county govt", "county government")
    return re.sub(r"[^a-z0-9]", "", n)
acc["norm"] = acc.account_name.map(norm)
acc["dup_group_size"] = acc.groupby("norm").account_id.transform("count")
acc["is_duplicate"] = acc.dup_group_size > 1
master = acc.sort_values("account_id").groupby("norm").account_id.first()
acc["master_id"] = acc.norm.map(master)
acc["missing_industry"] = acc.industry.eq("")
ind_of = acc.set_index("account_id").industry
m_ind = acc.master_id.map(ind_of)
acc["industry_clean"] = np.where(acc.missing_industry, np.where(m_ind.ne(""), m_ind, "Unassigned"), acc.industry)

# ---------------- opportunities: flags
op["is_open"] = ~op.stage.isin(["Closed Won", "Closed Lost"])
op["f_missing_amount"] = op.amount_kes.isna() | (op.amount_kes == 0)
op["f_no_owner"] = op.owner_id.eq("")
op["f_inactive_owner"] = op.is_open & op.owner_id.isin(inactive)
op["f_past_close"] = op.is_open & (op.close_date < ASOF)
op["days_since_activity"] = (ASOF - op.last_activity_date).dt.days
op["f_stale"] = op.is_open & (op.days_since_activity > STALE_DAYS)
op["f_prob_mismatch"] = op.is_open & (op.probability != op.stage.map(STAGE_PROB))
op["f_no_next_step"] = op.is_open & op.next_step.eq("")
op["f_lost_no_reason"] = (op.stage == "Closed Lost") & op.lost_reason.eq("")
FL = ["f_missing_amount", "f_no_owner", "f_inactive_owner", "f_past_close", "f_stale", "f_prob_mismatch", "f_no_next_step", "f_lost_no_reason"]
op["n_flags"] = op[FL].sum(axis=1)
op["clean"] = op.n_flags == 0
op = op.merge(acc[["account_id", "account_name", "industry_clean", "tier", "master_id", "is_duplicate"]], on="account_id", how="left")
op["prob_clean"] = op.stage.map(STAGE_PROB)
op["weighted_kes"] = op.amount_kes.fillna(0) * op.prob_clean / 100
op["cycle_days"] = (op.close_date - op.created_date).dt.days
own = owners.set_index("owner_id")
op["owner_name"] = op.owner_id.map(own.owner_name).fillna("(no owner)")

# ---------------- DQ scorecard
O = op[op.is_open]
rules = {
    "f_missing_amount": ("Amount missing or zero", "all opportunities", op),
    "f_no_owner": ("No owner assigned", "all opportunities", op),
    "f_inactive_owner": ("Owner has left or is inactive", "open", O),
    "f_past_close": ("Open with close date in the past", "open", O),
    "f_stale": (f"No activity for more than {STALE_DAYS} days", "open", O),
    "f_prob_mismatch": ("Probability does not match stage", "open", O),
    "f_no_next_step": ("No next step recorded", "open", O),
    "f_lost_no_reason": ("Lost without a reason", "closed lost", op[op.stage == "Closed Lost"]),
}
dq = [dict(rule=k, label=v[0], scope=v[1], failing=int(v[2][k].sum()), population=int(len(v[2])), rate=float(v[2][k].mean())) for k, v in rules.items()]
acc_dq = [dict(rule="missing_industry", label="Account missing industry", failing=int(acc.missing_industry.sum()), population=int(len(acc))),
          dict(rule="is_duplicate", label="Account is part of a duplicate pair", failing=int(acc.is_duplicate.sum()), population=int(len(acc)))]
open_clean = float(O.clean.mean())

# ---------------- KPIs
def stats(d):
    return dict(open_n=int(len(d)), open_kes_m=float(d.amount_kes.fillna(0).sum() / 1e6), weighted_kes_m=float(d.weighted_kes.sum() / 1e6))
fy_open = O[(O.close_date >= FY27[0]) & (O.close_date <= FY27[1])]
won_ytd = op[(op.stage == "Closed Won") & (op.close_date >= FY27[0]) & (op.close_date <= ASOF)]
last12 = op[(~op.is_open) & (op.close_date > ASOF - pd.Timedelta(days=365)) & (op.close_date <= ASOF)]
bylos = []
for l in tg.line_of_service:
    d = O[O.line_of_service == l]; f = fy_open[fy_open.line_of_service == l]; w = won_ytd[won_ytd.line_of_service == l]; c = last12[last12.line_of_service == l]
    tgt = float(tg.set_index("line_of_service").loc[l, "fy27_bd_target_kes_m"])
    won_m = float(w.amount_kes.fillna(0).sum() / 1e6)
    remaining = max(tgt - won_m, 0.0)
    bylos.append(dict(los=l, **stats(d), fy27_pipeline_weighted_kes_m=float(f.weighted_kes.sum() / 1e6), target_kes_m=tgt, won_ytd_kes_m=won_m, remaining_kes_m=remaining,
                      coverage_x=float(f.weighted_kes.sum() / 1e6 / remaining) if remaining else None, closed_12m=int(len(c)), win_rate=float((c.stage == "Closed Won").mean()),
                      avg_cycle_days_won=float(op[(op.stage == "Closed Won") & (op.line_of_service == l)].cycle_days.mean()), dq_clean_rate=float(d.clean.mean())))
byind = []
for i in sorted(O.industry_clean.unique()):
    d = O[O.industry_clean == i]; c = last12[last12.industry_clean == i]
    byind.append(dict(industry=i, **stats(d), closed_12m=int(len(c)), win_rate=float((c.stage == "Closed Won").mean()) if len(c) else None))
pri = acc[(acc.tier == "Priority")]
pri_open = O[O.account_id.isin(pri.account_id) | O.master_id.isin(pri.master_id)]
last_touch = act.groupby("account_id").activity_date.max()
pri_rows = []
for r in pri.itertuples():
    d = O[(O.account_id == r.account_id)]
    lt = last_touch.get(r.account_id)
    pri_rows.append(dict(account_id=r.account_id, account=r.account_name, industry=r.industry_clean, driver=own.owner_name[r.account_driver], open_opps=int(len(d)), weighted_kes_m=float(d.weighted_kes.sum() / 1e6),
                         days_since_touch=int((ASOF - lt).days) if lt is not None and not pd.isna(lt) else None))
pri_df = pd.DataFrame(pri_rows)
priority = dict(n=int(len(pri)), with_open_opp=int((pri_df.open_opps > 0).sum()), touched_90d=int((pri_df.days_since_touch <= 90).sum()), no_open_no_touch=int(((pri_df.open_opps == 0) & ((pri_df.days_since_touch.isna()) | (pri_df.days_since_touch > 90))).sum()))
source = []
for s_, d in O.groupby("lead_source"):
    c = last12[last12.lead_source == s_]
    source.append(dict(source=s_, open_n=int(len(d)), open_kes_m=float(d.amount_kes.fillna(0).sum() / 1e6), win_rate=float((c.stage == "Closed Won").mean()) if len(c) else None, closed_12m=int(len(c))))
stage = [dict(stage=s_, n=int((O.stage == s_).sum()), kes_m=float(O[O.stage == s_].amount_kes.fillna(0).sum() / 1e6)) for s_ in ["Prospecting", "Qualification", "Proposal", "Negotiation"]]
top10 = O.sort_values("weighted_kes", ascending=False).head(10)[["opp_id", "opp_name", "line_of_service", "stage", "amount_kes", "weighted_kes", "owner_name", "close_date"]].copy()
top10["close_date"] = top10.close_date.dt.strftime("%Y-%m-%d")

# ---------------- owner follow-ups (one row per flagged open opportunity and issue)
ISSUE_TXT = {"f_missing_amount": "Add an estimated amount", "f_no_owner": "Assign an owner", "f_inactive_owner": "Reassign: owner has left", "f_past_close": "Update the close date or close the opportunity",
             "f_stale": "Log a recent activity or confirm status", "f_prob_mismatch": "Correct probability to match stage", "f_no_next_step": "Record the next step", "f_lost_no_reason": "Add the lost reason"}
fu = []
for r in op.itertuples():
    for f in FL:
        if getattr(r, f):
            fu.append(dict(owner=r.owner_name, opp_id=r.opp_id, opp_name=r.opp_name, stage=r.stage, amount_kes=None if pd.isna(r.amount_kes) else float(r.amount_kes), issue=rules[f][0], action=ISSUE_TXT[f], priority="High" if (f in ("f_past_close", "f_inactive_owner", "f_no_owner") or (r.tier == "Priority")) else "Normal"))
fu = pd.DataFrame(fu)
owner_rows = []
for o, d in O.groupby("owner_name"):
    owner_rows.append(dict(owner=o, open_n=int(len(d)), flagged=int((~d.clean).sum()), clean_rate=float(d.clean.mean()), issues=int(d.n_flags.sum())))
owner_df = pd.DataFrame(owner_rows).sort_values(["flagged", "open_n"], ascending=False)

# ---------------- feedback and events
THEMES = {"responsiveness": ["respond", "come back", "turnaround", "available"], "insight": ["insight", "advice", "thought leadership"], "value": ["fees", "pricing", "scoping", "cost"],
          "communication": ["updates", "told", "point of contact", "timeline"], "digital": ["portal", "emails", "invitations"]}
def theme(c):
    c = c.lower()
    for t, ks in THEMES.items():
        if any(k in c for k in ks):
            return t
    return "other"
fb["theme"] = fb.comment.map(theme)
fb["sentiment"] = np.where(fb.rating >= 4, "Positive", np.where(fb.rating <= 2, "Negative", "Neutral"))
fb_los = fb.groupby("line_of_service").rating.agg(["mean", "count"]).reset_index()
fb_theme = fb.groupby(["theme", "sentiment"]).size().unstack(fill_value=0).reset_index()
nps_like = float(((fb.rating >= 4).mean() - (fb.rating <= 2).mean()) * 100)
ev["attend_rate"] = np.where(ev.registered > 0, ev.attended / ev.registered, np.nan)
ev["opps_per_10_attendees"] = np.where(ev.attended > 0, ev.opps_created / ev.attended * 10, np.nan)

res = dict(asof=str(ASOF.date()), stale_days=STALE_DAYS, n_opps=int(len(op)), n_open=int(len(O)), n_accounts=int(len(acc)), dq=dq, acc_dq=acc_dq, open_clean_rate=open_clean,
           open=stats(O), fy27=dict(pipeline_weighted_kes_m=float(fy_open.weighted_kes.sum() / 1e6), won_ytd_kes_m=float(won_ytd.amount_kes.fillna(0).sum() / 1e6), target_kes_m=float(tg.fy27_bd_target_kes_m.sum())),
           win_rate_12m=float((last12.stage == "Closed Won").mean()), by_los=bylos, by_industry=byind, priority=priority, by_source=source, by_stage=stage, top10=top10.replace({np.nan: None}).to_dict("records"),
           feedback=dict(n=int(len(fb)), avg=float(fb.rating.mean()), nps_like=nps_like, by_los=fb_los.to_dict("records"), by_theme=fb_theme.to_dict("records"), open_followups=int((fb.follow_up_status == "Open").sum())),
           events=ev.replace({np.nan: None}).to_dict("records"), followups_n=int(len(fu)), owner_hygiene=owner_df.head(10).to_dict("records"))
json.dump(res, open(os.path.join(OUT, "audit_results.json"), "w"), indent=1, default=float)
os.makedirs(PROC, exist_ok=True)
keep = ["opp_id", "account_id", "master_id", "account_name", "industry_clean", "tier", "opp_name", "line_of_service", "stage", "amount_kes", "probability", "prob_clean", "created_date", "close_date", "last_activity_date", "owner_id", "owner_name", "lead_source", "opp_type", "next_step", "lost_reason", "weighted_kes", "days_since_activity", "n_flags", "clean"] + FL
op[keep].to_csv(os.path.join(PROC, "opportunities_clean.csv"), index=False, date_format="%Y-%m-%d")
acc.drop(columns=["norm"]).to_csv(os.path.join(PROC, "accounts_clean.csv"), index=False, date_format="%Y-%m-%d")
fu.to_csv(os.path.join(OUT, "owner_followups.csv"), index=False)
fb.to_csv(os.path.join(PROC, "feedback_tagged.csv"), index=False, date_format="%Y-%m-%d")
pri_df.to_csv(os.path.join(OUT, "priority_accounts.csv"), index=False)
if __name__ == "__main__":
    print("open", res["open"], "clean rate", round(open_clean, 3), "win12m", round(res["win_rate_12m"], 3), "followups", len(fu))
    print(pd.DataFrame(dq)[["label", "failing", "population", "rate"]].round(3).to_string())
    print(pd.DataFrame(bylos).round(2).to_string())
    print(priority); print(pd.DataFrame(byind).round(2).to_string())
