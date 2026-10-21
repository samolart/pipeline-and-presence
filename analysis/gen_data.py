"""Step 1. Generate a SIMULATED Salesforce-style export for a fictional professional-services firm in Nairobi.
No real client, firm or employee data is used. Every name is invented. Seeded for reproducibility.
Deliberate data-quality problems are injected so the audit has something real to find."""
import os, random
import numpy as np, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
rng = np.random.default_rng(20261021)
random.seed(20261021)
ASOF = pd.Timestamp("2026-10-19")
FY27_START, FY27_END = pd.Timestamp("2026-07-01"), pd.Timestamp("2027-06-30")

LOS = ["Assurance", "Tax & Legal", "Deals", "Consulting", "Risk"]
LOS_W = [0.20, 0.27, 0.15, 0.28, 0.10]
LOS_MU = {"Assurance": 16.2, "Tax & Legal": 15.6, "Deals": 17.0, "Consulting": 16.6, "Risk": 15.9}  # ln KES
INDUSTRIES = ["Financial Services", "Energy & Resources", "Technology, Media & Telecom", "Public Sector & Development", "Health", "Consumer & Retail", "Agribusiness", "Real Estate & Infrastructure"]
IND_W = [0.22, 0.12, 0.14, 0.16, 0.08, 0.12, 0.08, 0.08]
STAGES = ["Prospecting", "Qualification", "Proposal", "Negotiation", "Closed Won", "Closed Lost"]
STAGE_PROB = {"Prospecting": 10, "Qualification": 25, "Proposal": 50, "Negotiation": 75, "Closed Won": 100, "Closed Lost": 0}
SOURCES = ["Existing client", "Referral", "Event", "Thought leadership", "Digital", "Cold outreach"]

first = ["Amina", "Brian", "Cynthia", "Daniel", "Esther", "Felix", "Grace", "Hassan", "Irene", "James", "Kevin", "Lucy", "Mark", "Nancy", "Omar", "Purity", "Quincy", "Rose", "Samuel", "Tabitha", "Victor", "Winnie", "Yusuf", "Zainab", "Peter", "Mercy", "Joseph", "Faith", "Collins", "Naomi"]
last = ["Otieno", "Wanjiru", "Mwangi", "Kamau", "Njoroge", "Achieng", "Abdi", "Chebet", "Kiprop", "Mutua", "Ochieng", "Wafula", "Karanja", "Maina", "Odhiambo", "Barasa", "Nyambura", "Kariuki", "Juma", "Langat"]
# ---- owners
owners = []
for i in range(28):
    grade = "Partner" if i < 7 else "Director" if i < 14 else "Senior Manager"
    owners.append(dict(owner_id=f"U{i+1:03d}", owner_name=f"{first[i % 30]} {last[(i*7) % 20]}", grade=grade, los=LOS[i % 5], active=not (i in (11, 23))))
owners = pd.DataFrame(owners)

# ---- accounts
stem = {"Financial Services": ["Capital", "Bank", "Insurance", "Microfinance", "Sacco", "Asset Managers"], "Energy & Resources": ["Power", "Energy", "Petroleum", "Geothermal", "Mining"],
        "Technology, Media & Telecom": ["Digital", "Telecom", "Media", "Networks", "Cloud"], "Public Sector & Development": ["County Government", "Authority", "Development Fund", "Foundation", "Commission"],
        "Health": ["Hospital", "Pharma", "Health Group", "Diagnostics"], "Consumer & Retail": ["Foods", "Retail", "Beverages", "Supermarkets", "Hotels"],
        "Agribusiness": ["Growers", "Agri", "Tea Estates", "Dairy", "Horticulture"], "Real Estate & Infrastructure": ["Properties", "Construction", "Infrastructure", "Housing", "Logistics"]}
pre = ["Savanna", "Rift", "Kilimanjaro", "Lakeview", "Mara", "Baobab", "Acacia", "Tembo", "Simba", "Jua", "Pwani", "Highland", "Nile", "Equator", "Safari", "Kibo", "Amani", "Jambo", "Umoja", "Harambee", "Twiga", "Nyota", "Mto", "Zuri"]
names, rows = set(), []
n_acc = 320
while len(rows) < n_acc:
    ind = rng.choice(INDUSTRIES, p=IND_W)
    nm = f"{random.choice(pre)} {random.choice(stem[ind])}"
    suffix = "" if ind == "Public Sector & Development" else " Limited"
    full = nm + suffix
    if full in names:
        continue
    names.add(full)
    tier = "Priority" if len(rows) < 40 else "Key" if len(rows) < 110 else "Standard"
    rows.append(dict(account_id=f"A{len(rows)+1:04d}", account_name=full, industry=ind, tier=tier, account_type=rng.choice(["Listed company", "Private company", "Public sector", "NGO / development partner", "Multinational subsidiary"], p=[.15, .38, .2, .1, .17]),
                     city=rng.choice(["Nairobi", "Mombasa", "Kisumu", "Nakuru", "Eldoret"], p=[.72, .12, .06, .06, .04]),
                     account_driver=owners.owner_id.sample(1, random_state=int(rng.integers(1e6))).iloc[0], created_date=ASOF - pd.Timedelta(days=int(rng.integers(120, 2200)))))
acc = pd.DataFrame(rows)
# priority accounts skew to bigger industries and partner drivers
partners = owners[owners.grade == "Partner"].owner_id.tolist()
acc.loc[acc.tier == "Priority", "account_driver"] = [partners[i % 7] for i in range((acc.tier == "Priority").sum())]

# ---- opportunities
opps = []
n_opp = 950
acc_w = acc.tier.map({"Priority": 5.0, "Key": 2.2, "Standard": 1.0}).values
acc_w = acc_w / acc_w.sum()
for i in range(n_opp):
    a = acc.iloc[int(rng.choice(len(acc), p=acc_w))]
    los = rng.choice(LOS, p=LOS_W)
    amt = float(np.round(np.exp(rng.normal(LOS_MU[los], 0.85)) / 1000) * 1000)
    amt = min(max(amt, 800_000), 220_000_000)
    created = ASOF - pd.Timedelta(days=int(rng.integers(5, 420)))
    cycle = int(rng.integers(60, 300))
    closed_due = created + pd.Timedelta(days=cycle)
    if closed_due < ASOF:
        stage = rng.choice(["Closed Won", "Closed Lost"], p=[0.38, 0.62])
        close = closed_due
        last_mod = close
    else:
        stage = rng.choice(["Prospecting", "Qualification", "Proposal", "Negotiation"], p=[.22, .3, .3, .18])
        close = closed_due
        last_mod = ASOF - pd.Timedelta(days=int(rng.integers(0, 40)))
    owner = a.account_driver if rng.random() < 0.6 else owners.owner_id.sample(1, random_state=int(rng.integers(1e6))).iloc[0]
    last_act = last_mod - pd.Timedelta(days=int(rng.integers(0, 14))) if stage not in ("Closed Won", "Closed Lost") else close
    opps.append(dict(opp_id=f"O{i+1:05d}", account_id=a.account_id, opp_name=f"{a.account_name.split(' Limited')[0]} | {los} | {rng.choice(['advisory', 'annual', 'review', 'transformation', 'compliance', 'readiness', 'due diligence', 'support'])}",
                     line_of_service=los, stage=stage, amount_kes=amt, probability=STAGE_PROB[stage], created_date=created, close_date=close, last_modified=last_mod, last_activity_date=last_act, owner_id=owner,
                     lead_source=rng.choice(SOURCES, p=[.3, .2, .12, .12, .1, .16]), opp_type=rng.choice(["New client", "Existing client"], p=[.3, .7]),
                     next_step=("" if stage in ("Closed Won", "Closed Lost") else rng.choice(["Send proposal", "Partner meeting", "Follow-up call", "Await tender outcome", "Scope workshop", "Pricing review"])),
                     lost_reason=(rng.choice(["Price", "Incumbent retained", "Scope change", "Independence conflict", "No decision", "Competitor"]) if stage == "Closed Lost" else "")))
op = pd.DataFrame(opps)
# push a share of open opportunities to a future close date inside FY27 (realistic forward pipeline)
openm = ~op.stage.isin(["Closed Won", "Closed Lost"])
op.loc[openm, "close_date"] = [ASOF + pd.Timedelta(days=int(d)) for d in rng.integers(-40, 260, openm.sum())]

# ---- inject data-quality problems (deliberate)
ix = lambda frac, mask=None: rng.choice(op.index[mask] if mask is not None else op.index, size=int(frac * (mask.sum() if mask is not None else len(op))), replace=False)
op.loc[ix(0.05), "amount_kes"] = np.nan                                    # missing amount
op.loc[ix(0.02), "amount_kes"] = 0                                         # zero amount
op.loc[ix(0.06, openm), "probability"] = rng.choice([0, 90, 100, 5], size=int(0.06 * openm.sum()))  # probability inconsistent with stage
op.loc[ix(0.08, openm), "next_step"] = ""                                  # open with no next step
stale_idx = ix(0.22, openm)
op.loc[stale_idx, "last_activity_date"] = [ASOF - pd.Timedelta(days=int(d)) for d in rng.integers(61, 200, len(stale_idx))]
lost = op.stage == "Closed Lost"
op.loc[ix(0.18, lost), "lost_reason"] = ""                                 # lost without reason
op.loc[ix(0.03), "owner_id"] = ""                                          # no owner
inactive = owners[~owners.active].owner_id.tolist()
op.loc[ix(0.04, openm), "owner_id"] = rng.choice(inactive, size=int(0.04 * openm.sum()))  # owned by someone who left
# accounts: missing industry, duplicate variants
acc.loc[rng.choice(acc.index, 14, replace=False), "industry"] = ""
dups = acc.sample(12, random_state=7).copy()
dups["account_id"] = [f"A{n_acc+k+1:04d}" for k in range(len(dups))]
dups["account_name"] = dups.account_name.str.replace(" Limited", " Ltd", regex=False).str.replace("County Government", "County Govt", regex=False)
dups["tier"] = "Standard"
acc = pd.concat([acc, dups], ignore_index=True)
# duplicates: re-point 1-2 opportunities of the original to the duplicate
orig_by_name = {r.account_name: r.account_id for r in acc.iloc[:n_acc].itertuples()}
for r in dups.itertuples():
    base = r.account_name.replace(" Ltd", " Limited").replace("County Govt", "County Government")
    oid = orig_by_name.get(base)
    cand = op.index[op.account_id == oid][:2]
    if len(cand):
        op.loc[cand[-1], "account_id"] = r.account_id

# ---- activities (touchpoints)
act = []
for r in acc.itertuples():
    k = int(rng.poisson({"Priority": 14, "Key": 6, "Standard": 2.2}[r.tier]))
    for _ in range(k):
        act.append(dict(account_id=r.account_id, activity_type=rng.choice(["Client meeting", "Call", "Event attended", "Proposal sent", "Thought-leadership download", "Email click"], p=[.22, .2, .12, .1, .16, .2]),
                        activity_date=ASOF - pd.Timedelta(days=int(rng.integers(0, 400))), owner_id=r.account_driver))
act = pd.DataFrame(act)
act.insert(0, "activity_id", [f"T{i+1:05d}" for i in range(len(act))])

# ---- client feedback (simulated survey comments)
themes = {"responsiveness": ["Quick to respond and always available when we need guidance.", "Slow to come back on urgent queries.", "Turnaround on our requests was faster than expected."],
          "insight": ["The team brought genuinely useful sector insight.", "Advice felt generic and not tailored to our industry.", "The thought leadership alerts help us stay ahead of regulation."],
          "value": ["Fees were fair for the quality delivered.", "Pricing was higher than we expected for the scope.", "Clear scoping avoided surprises on cost."],
          "communication": ["Updates were clear and regular.", "We were not told about changes to the timeline early enough.", "One point of contact made everything easier."],
          "digital": ["The client portal made sharing documents simple.", "Too many separate emails from different teams.", "Event invitations are relevant to our work."]}
fb = []
for i in range(180):
    th = rng.choice(list(themes))
    j = int(rng.choice(3, p=[.5, .3, .2]) if rng.random() < .8 else rng.integers(0, 3))
    rating = int(np.clip({0: rng.normal(4.3, .6), 1: rng.normal(2.4, .8), 2: rng.normal(4.0, .8)}[j], 1, 5).round())
    fb.append(dict(feedback_id=f"F{i+1:04d}", account_id=acc.account_id.iloc[int(rng.integers(0, n_acc))], line_of_service=rng.choice(LOS, p=LOS_W), survey_date=ASOF - pd.Timedelta(days=int(rng.integers(5, 330))),
                   rating=rating, comment=themes[th][j], follow_up_status=rng.choice(["Open", "Responded", "Closed"], p=[.3, .4, .3])))
fb = pd.DataFrame(fb)

# ---- events
ev = pd.DataFrame([
    dict(event_id="E01", event_name="Banking & Fintech Roundtable", industry="Financial Services", date="2026-07-23", invited=120, registered=64, attended=47, opps_created=6),
    dict(event_id="E02", event_name="Budget Breakfast: Tax Changes Briefing", industry="All industries", date="2026-08-06", invited=260, registered=148, attended=109, opps_created=11),
    dict(event_id="E03", event_name="Energy Transition Forum", industry="Energy & Resources", date="2026-09-10", invited=90, registered=51, attended=38, opps_created=4),
    dict(event_id="E04", event_name="Public Sector Finance Clinic", industry="Public Sector & Development", date="2026-09-24", invited=80, registered=42, attended=33, opps_created=3),
    dict(event_id="E05", event_name="Health Leaders Dinner", industry="Health", date="2026-10-08", invited=40, registered=27, attended=22, opps_created=3),
    dict(event_id="E06", event_name="Annual Industry Conference (planned)", industry="All industries", date="2026-11-19", invited=400, registered=0, attended=0, opps_created=0)])

# ---- FY27 targets by LoS (KES m) - simulated planning inputs
targets = pd.DataFrame({"line_of_service": LOS, "fy27_bd_target_kes_m": [1100, 1500, 900, 1900, 600]})

for nm, df in (("owners", owners), ("accounts", acc), ("opportunities", op), ("activities", act), ("feedback", fb), ("events", ev), ("targets", targets)):
    df.to_csv(os.path.join(RAW, f"{nm}.csv"), index=False, date_format="%Y-%m-%d")
    print(nm, len(df))
open(os.path.join(RAW, "README.txt"), "w").write("SIMULATED DATA. Generated by analysis/gen_data.py (seed 20261021). Fictional firm, clients and staff. Not real Salesforce data. As-of date 2026-10-19. Deliberate data-quality issues are injected.\n")
