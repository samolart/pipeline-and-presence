"""Step 4. Build the interactive dashboard (single HTML file, data embedded). Reads data/processed and outputs/audit_results.json."""
import json, os
import pandas as pd
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = json.load(open(os.path.join(ROOT, "outputs", "audit_results.json")))
op = pd.read_csv(os.path.join(ROOT, "data", "processed", "opportunities_clean.csv"), keep_default_na=False)
pri = pd.read_csv(os.path.join(ROOT, "outputs", "priority_accounts.csv"), keep_default_na=False)
fu = pd.read_csv(os.path.join(ROOT, "outputs", "owner_followups.csv"), keep_default_na=False)
tg = pd.read_csv(os.path.join(ROOT, "data", "raw", "targets.csv"))
FL = ["f_missing_amount", "f_no_owner", "f_inactive_owner", "f_past_close", "f_stale", "f_prob_mismatch", "f_no_next_step", "f_lost_no_reason"]
rows = []
for r in op.itertuples():
    rows.append([r.opp_id, r.opp_name, r.line_of_service, r.industry_clean, r.stage, None if r.amount_kes == "" else round(float(r.amount_kes) / 1e6, 4), int(r.prob_clean), r.owner_name, r.lead_source, r.close_date,
                 r.created_date, r.tier, "".join("1" if getattr(r, f) in (True, "True") else "0" for f in FL), int(r.days_since_activity), r.account_name])
data = dict(asof=A["asof"], fy=["2026-07-01", "2027-06-30"], targets=dict(zip(tg.line_of_service, tg.fy27_bd_target_kes_m)), rules=[d["label"] for d in A["dq"]], rule_keys=FL,
            opps=rows, priority=pri.to_dict("records"), fu=fu.head(400).to_dict("records"), audit=dict(open=A["open"], win=A["win_rate_12m"], clean=A["open_clean_rate"], by_los=A["by_los"]),
            feedback=A["feedback"], events=A["events"], acc_dq=A["acc_dq"])
J = json.dumps(data, separators=(",", ":"), default=float)
HTML = r"""<title>Pipeline and Presence</title><meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Familjen+Grotesk:wght@500;600;700&family=Hanken+Grotesk:wght@400;500;600&family=Martian+Mono:wght@400;600&display=swap">
<style>
/* Layout concept: an operations console. Summary strip first, then the panels a BD meeting walks through: pipeline, coverage, data quality, priority accounts, chasing list. */
:root{--bg:#EDF0EE;--panel:#FAFBFA;--ink:#102A30;--mute:#5B6E72;--rule:#D5DDDB;--ocean:#2A7FA8;--coral:#E4572E;--ochre:#D9A21B;--mint:#2E9C7A;--soft:#E3EAE8;--hot:#F8DCD3;
--disp:'Familjen Grotesk','Trebuchet MS',system-ui,sans-serif;--body:'Hanken Grotesk',system-ui,-apple-system,'Segoe UI',sans-serif;--mono:'Martian Mono',ui-monospace,Menlo,monospace}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#0C1B1F;--panel:#12272C;--ink:#E6EEEC;--mute:#98AEB2;--rule:#26414A;--ocean:#59B3DB;--coral:#FF7C57;--ochre:#F0BC3E;--mint:#52CBA3;--soft:#183840;--hot:#47262A;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#0C1B1F;--panel:#12272C;--ink:#E6EEEC;--mute:#98AEB2;--rule:#26414A;--ocean:#59B3DB;--coral:#FF7C57;--ochre:#F0BC3E;--mint:#52CBA3;--soft:#183840;--hot:#47262A;color-scheme:dark}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--body);font-size:15px;line-height:1.45;padding-block:18px 56px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.wrap{max-width:1240px;margin:0 auto;display:flex;flex-direction:column;gap:16px}
header{display:flex;flex-wrap:wrap;gap:10px 20px;align-items:flex-end;justify-content:space-between}
h1{font:700 clamp(26px,3.4vw,38px)/1.05 var(--disp);margin:0;letter-spacing:-.01em}h2{font:600 15px/1.2 var(--disp);margin:0 0 10px;letter-spacing:.01em}
.sub{color:var(--mute);font-size:13px;margin:4px 0 0}
.badge{font:600 11px/1 var(--mono);letter-spacing:.06em;padding:6px 9px;border:1.5px solid var(--coral);color:var(--coral);text-transform:uppercase}
.bar{display:flex;flex-wrap:wrap;gap:10px 16px;align-items:center;background:var(--panel);border:1px solid var(--rule);padding:10px 14px}
.bar label{font:600 11px/1 var(--mono);letter-spacing:.06em;color:var(--mute);text-transform:uppercase;display:flex;flex-direction:column;gap:5px}
select,button{font:500 14px var(--body);color:var(--ink);background:var(--bg);border:1px solid var(--rule);padding:7px 9px;min-width:150px}
button{cursor:pointer;min-width:0}button:focus-visible,select:focus-visible{outline:2px solid var(--ocean);outline-offset:1px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
.kpi{background:var(--panel);border:1px solid var(--rule);padding:12px 14px;display:flex;flex-direction:column;gap:2px;min-width:0}
.kpi b{font:700 clamp(24px,2.6vw,32px)/1.05 var(--disp);font-variant-numeric:tabular-nums}.kpi span{font-size:12.5px;color:var(--mute)}.kpi i{font:600 11px/1.3 var(--mono);font-style:normal}
.kpi.warn b{color:var(--coral)}.kpi.good b{color:var(--mint)}
.grid{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:16px}
.panel{background:var(--panel);border:1px solid var(--rule);padding:14px 16px;min-width:0}.c4{grid-column:span 4}.c5{grid-column:span 5}.c6{grid-column:span 6}.c7{grid-column:span 7}.c8{grid-column:span 8}.c12{grid-column:span 12}
@media (max-width:900px){.c4,.c5,.c6,.c7,.c8{grid-column:span 12}}
.row{display:grid;grid-template-columns:minmax(96px,190px) minmax(0,1fr) 104px;gap:8px;align-items:center;font-size:13px;margin:6px 0}
.row .lab{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row .val{font:500 12px var(--mono);text-align:right;font-variant-numeric:tabular-nums}
.track{position:relative;height:16px;background:var(--soft)}.fill{position:absolute;inset:0 auto 0 0;background:var(--ocean)}.fill.w{background:var(--ochre)}.fill.bad{background:var(--coral)}
.mark{position:absolute;top:-3px;bottom:-3px;width:2px;background:var(--ink)}
.scroll{overflow-x:auto;min-width:0}table{border-collapse:collapse;width:100%;font-size:13px}
th{font:600 10.5px/1.2 var(--mono);letter-spacing:.06em;text-transform:uppercase;color:var(--mute);text-align:left;padding:7px 8px;border-bottom:2px solid var(--ink);white-space:nowrap}
td{padding:7px 8px;border-bottom:1px solid var(--rule);vertical-align:top}td.n,th.n{text-align:right;font-family:var(--mono);font-size:12px;font-variant-numeric:tabular-nums}
.pill{display:inline-block;font:600 10.5px/1 var(--mono);padding:3px 6px;background:var(--hot);color:var(--coral);border:1px solid var(--coral)}.pill.ok{background:var(--soft);color:var(--mint);border-color:var(--mint)}.pill.hi{background:var(--coral);color:#fff}
.note{font-size:12.5px;color:var(--mute);margin:8px 0 0}.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--mute);margin-top:6px}.legend i{display:inline-block;width:10px;height:10px;margin-right:5px}
.theme{position:fixed;right:max(12px,env(safe-area-inset-right,0px));bottom:max(12px,env(safe-area-inset-bottom,0px))}
.two{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}
</style>
<div class="wrap">
<header><div><h1>Pipeline and Presence</h1><p class="sub" id="sub"></p></div><span class="badge">Simulated data. No real client or firm information.</span></header>
<div class="bar"><label>Line of service<select id="fLos"></select></label><label>Industry<select id="fInd"></select></label><label>Owner<select id="fOwn"></select></label><button id="reset" type="button">Reset filters</button><span class="sub" id="fnote"></span></div>
<div class="kpis" id="kpis"></div>
<div class="grid">
<div class="panel c5"><h2>Open pipeline by stage (KES m)</h2><div id="stage"></div><div class="legend"><span><i style="background:var(--ocean)"></i>Open value</span><span><i style="background:var(--ochre)"></i>Weighted</span></div></div>
<div class="panel c7"><h2>FY27 coverage by line of service</h2><div id="cov"></div><p class="note">Bar = weighted pipeline closing in FY27. Marker = target still to win (target less won year to date). A bar short of the marker means coverage below 1.0x. Targets are simulated.</p></div>
<div class="panel c6"><h2>Pipeline by industry (weighted, KES m)</h2><div id="ind"></div></div>
<div class="panel c6"><h2>Where deals come from (open pipeline and win rate)</h2><div id="src"></div></div>
<div class="panel c6"><h2>Data quality: share failing each rule</h2><div id="dq"></div><p class="note" id="dqnote"></p></div>
<div class="panel c6"><h2>Owner hygiene: open opportunities with an issue</h2><div id="own"></div></div>
<div class="panel c12"><h2>Priority accounts: who needs attention</h2><div class="scroll"><table id="pri"></table></div><p class="note" id="prinote"></p></div>
<div class="panel c12"><h2>Chasing list: top opportunities to fix (filters apply)</h2><div class="scroll"><table id="fix"></table></div><p class="note">Sorted by weighted value. The full list of 386 owner follow-ups is in outputs/owner_followups.csv.</p></div>
<div class="panel c6"><h2>Client feedback (180 simulated responses)</h2><div id="fb"></div></div>
<div class="panel c6"><h2>Events: funnel from invitation to opportunity</h2><div class="scroll"><table id="ev"></table></div></div>
</div>
<p class="note">All figures are generated from a simulated Salesforce-style export (analysis/gen_data.py). Coverage targets are simulated. Probability is reset to the default for each stage before weighting. Opportunities with no amount count as zero in totals and are listed for follow-up, never estimated.</p>
</div>
<button class="theme" id="thm" type="button">Light / dark</button>
<script>
var D=__DATA__;
var C={ocean:'var(--ocean)',coral:'var(--coral)',ochre:'var(--ochre)'};
var ASOF=new Date(D.asof+'T00:00:00Z'),FY0=new Date(D.fy[0]+'T00:00:00Z'),FY1=new Date(D.fy[1]+'T00:00:00Z');
var O=D.opps.map(function(r){var fl=r[12];return{id:r[0],name:r[1],los:r[2],ind:r[3],stage:r[4],amt:r[5],prob:r[6],owner:r[7],src:r[8],close:new Date(r[9]+'T00:00:00Z'),created:r[10],tier:r[11],fl:fl.split('').map(Number),days:r[13],acct:r[14],open:r[4].indexOf('Closed')!==0}});
O.forEach(function(o){o.w=(o.amt||0)*o.prob/100;o.nf=o.fl.reduce(function(a,b){return a+b},0);o.clean=o.nf===0;});
var $=function(i){return document.getElementById(i)};
function fm(v,d){d=d==null?0:d;return v.toLocaleString('en-GB',{minimumFractionDigits:d,maximumFractionDigits:d})}
function uniq(k){var s={};O.forEach(function(o){s[o[k]]=1});return Object.keys(s).sort()}
function fill(id,list){var e=$(id);e.innerHTML='<option value="">All</option>'+list.map(function(x){return'<option>'+x.replace(/&/g,'&amp;')+'</option>'}).join('')}
fill('fLos',Object.keys(D.targets));fill('fInd',uniq('ind'));fill('fOwn',uniq('owner'));
function sel(){return{los:$('fLos').value,ind:$('fInd').value,own:$('fOwn').value}}
function pick(o,s){return(!s.los||o.los===s.los)&&(!s.ind||o.ind===s.ind)&&(!s.own||o.owner===s.own)}
function bars(el,rows,max,fmt,cls){el.innerHTML=rows.map(function(r){var w=max?Math.min(100,r.v/max*100):0;var h='<div class="row"><span class="lab" title="'+r.l+'">'+r.l+'</span><span class="track"><span class="fill '+(cls||'')+'" style="width:'+w+'%"></span>'+(r.v2!=null?'<span class="fill w" style="width:'+Math.min(100,r.v2/max*100)+'%;height:6px;top:5px;bottom:auto"></span>':'')+(r.mark!=null?'<span class="mark" style="left:'+Math.min(99.5,r.mark/max*100)+'%"></span>':'')+'</span><span class="val">'+fmt(r)+'</span></div>';return h}).join('')}
function render(){
var s=sel(),all=O.filter(function(o){return pick(o,s)}),op=all.filter(function(o){return o.open});
var last12=all.filter(function(o){return!o.open&&o.close>new Date(ASOF-365*864e5)&&o.close<=ASOF});
var won=last12.filter(function(o){return o.stage==='Closed Won'});
var sum=function(a,f){return a.reduce(function(x,o){return x+f(o)},0)};
var openV=sum(op,function(o){return o.amt||0}),wV=sum(op,function(o){return o.w});
var fy=op.filter(function(o){return o.close>=FY0&&o.close<=FY1});
var wonYtd=all.filter(function(o){return o.stage==='Closed Won'&&o.close>=FY0&&o.close<=ASOF});
var los=s.los?[s.los]:Object.keys(D.targets);
var tgt=sum(los.map(function(l){return{v:D.targets[l]}}),function(x){return x.v});
var wonLos=sum(wonYtd,function(o){return o.amt||0});
var rem=Math.max(tgt-wonLos,0);var cov=(!s.ind&&!s.own&&rem>0)?sum(fy,function(o){return o.w})/rem:null;
var cl=op.length?op.filter(function(o){return o.clean}).length/op.length:null;
$('sub').textContent='As of '+D.asof+'. FY27 runs 1 Jul 2026 to 30 Jun 2027. KES millions unless stated.';
$('fnote').textContent=all.length+' opportunities in view';
var k=[['Open pipeline','KES '+fm(openV)+'m',op.length+' open opportunities',''],['Weighted pipeline','KES '+fm(wV)+'m','Stage-default probabilities',''],
['FY27 coverage',cov==null?'n/a':fm(cov,2)+'x',cov==null?'Targets exist by line of service only':'Weighted FY27 pipeline / target still to win',cov!=null&&cov<1?'warn':'good'],
['Win rate, 12 months',last12.length?fm(won.length/last12.length*100)+'%':'n/a',won.length+' won of '+last12.length+' closed',''],
['Data-clean open opps',cl==null?'n/a':fm(cl*100)+'%',op.filter(function(o){return!o.clean}).length+' need a fix',cl!=null&&cl<.6?'warn':'good'],
['Stale opportunities',op.filter(function(o){return o.fl[4]}).length+'','No activity for more than 60 days','']];
$('kpis').innerHTML=k.map(function(x){return'<div class="kpi '+x[3]+'"><span>'+x[0]+'</span><b>'+x[1]+'</b><span>'+x[2]+'</span></div>'}).join('');
var st=['Prospecting','Qualification','Proposal','Negotiation'];
var srow=st.map(function(g){var a=op.filter(function(o){return o.stage===g});return{l:g,v:sum(a,function(o){return o.amt||0}),v2:sum(a,function(o){return o.w}),n:a.length}});
bars($('stage'),srow,Math.max.apply(null,srow.map(function(r){return r.v}))||1,function(r){return fm(r.v)+' ('+r.n+')'});
var crow=Object.keys(D.targets).filter(function(l){return!s.los||l===s.los}).map(function(l){var f=fy.filter(function(o){return o.los===l});var wl=sum(O.filter(function(o){return o.los===l&&o.stage==='Closed Won'&&o.close>=FY0&&o.close<=ASOF}),function(o){return o.amt||0});var r=Math.max(D.targets[l]-wl,0);return{l:l,v:sum(f,function(o){return o.w}),mark:r,r:r}});
var cm=Math.max.apply(null,crow.map(function(r){return Math.max(r.v,r.r)}))||1;
bars($('cov'),crow,cm,function(r){return fm(r.v)+' / '+fm(r.r)},'');
if(s.ind||s.own){$('cov').insertAdjacentHTML('afterbegin','<p class="note">Industry and owner filters change the pipeline bar only. Targets stay at line-of-service level.</p>')}
var by=function(k,list){return list.map(function(x){var a=op.filter(function(o){return o[k]===x});return{l:x,v:sum(a,function(o){return o.w})}}).filter(function(r){return r.v>0}).sort(function(a,b){return b.v-a.v})};
var ir=by('ind',uniq('ind'));bars($('ind'),ir,ir.length?ir[0].v:1,function(r){return fm(r.v)});
var sr=uniq('src').map(function(x){var a=op.filter(function(o){return o.src===x}),c=last12.filter(function(o){return o.src===x});return{l:x,v:sum(a,function(o){return o.amt||0}),wr:c.length?c.filter(function(o){return o.stage==='Closed Won'}).length/c.length:null}}).sort(function(a,b){return b.v-a.v});
bars($('src'),sr,sr.length?sr[0].v:1,function(r){return fm(r.v)+' | '+(r.wr==null?'n/a':fm(r.wr*100)+'%')});
var rules=D.rules,dq=rules.map(function(l,i){var pop=i===7?all.filter(function(o){return o.stage==='Closed Lost'}):(i<2?all:op);var f=pop.filter(function(o){return o.fl[i]}).length;return{l:l,v:pop.length?f/pop.length:0,f:f,p:pop.length}});
bars($('dq'),dq,.4,function(r){return fm(r.v*100)+'% ('+r.f+')'},'bad');
$('dqnote').textContent='Account records: '+D.acc_dq[0].failing+' missing an industry and '+D.acc_dq[1].failing+' in duplicate pairs (whole database, not filtered).';
var owm={};op.forEach(function(o){var x=owm[o.owner]=owm[o.owner]||{n:0,b:0};x.n++;if(!o.clean)x.b++});
var orow=Object.keys(owm).map(function(k){return{l:k,v:owm[k].b,n:owm[k].n}}).sort(function(a,b){return b.v-a.v}).slice(0,10);
bars($('own'),orow,orow.length?orow[0].v:1,function(r){return r.v+' of '+r.n},'bad');
var pr=D.priority.map(function(p){var need=p.open_opps===0||p.days_since_touch==null||p.days_since_touch>90;return p.__n=need,p}).filter(function(p){return!s.ind||p.industry===s.ind}).sort(function(a,b){return(b.__n-a.__n)||(b.weighted_kes_m-a.weighted_kes_m)}).slice(0,12);
$('pri').innerHTML='<thead><tr><th>Account</th><th>Industry</th><th>Driver</th><th class="n">Open opps</th><th class="n">Weighted KES m</th><th class="n">Days since touch</th><th>Status</th></tr></thead><tbody>'+pr.map(function(p){return'<tr><td>'+p.account+'</td><td>'+p.industry+'</td><td>'+p.driver+'</td><td class="n">'+p.open_opps+'</td><td class="n">'+fm(p.weighted_kes_m,1)+'</td><td class="n">'+(p.days_since_touch==null?'none':p.days_since_touch)+'</td><td>'+(p.__n?'<span class="pill">Reach out</span>':'<span class="pill ok">On track</span>')+'</td></tr>'}).join('')+'</tbody>';
var pa=D.priority;$('prinote').textContent=pa.filter(function(p){return p.open_opps>0}).length+' of '+pa.length+' priority accounts have an open opportunity; '+pa.filter(function(p){return p.days_since_touch!=null&&p.days_since_touch<=90}).length+' were touched in the last 90 days. Showing accounts that need attention first.';
var fx=op.filter(function(o){return!o.clean}).sort(function(a,b){return b.w-a.w}).slice(0,10);
$('fix').innerHTML='<thead><tr><th>Opportunity</th><th>Owner</th><th>Stage</th><th class="n">KES m</th><th>Issues</th></tr></thead><tbody>'+fx.map(function(o){var is=o.fl.map(function(f,i){return f?'<span class="pill'+(i===3||i===2||i===1?' hi':'')+'">'+D.rules[i].replace('No activity for more than 60 days','Stale').replace('Open with close date in the past','Close date passed').replace('Owner has left or is inactive','Owner left').replace('Probability does not match stage','Probability off').replace('Amount missing or zero','No amount').replace('No owner assigned','No owner').replace('No next step recorded','No next step').replace('Lost without a reason','No lost reason')+'</span> ':''}).join('');return'<tr><td>'+o.name+'</td><td>'+o.owner+'</td><td>'+o.stage+'</td><td class="n">'+(o.amt==null?'-':fm(o.amt,1))+'</td><td>'+is+'</td></tr>'}).join('')+'</tbody>';
}
function static_(){var f=D.feedback;var mx=Math.max.apply(null,f.by_theme.map(function(t){return(t.Positive||0)+(t.Neutral||0)+(t.Negative||0)}));
$('fb').innerHTML='<div class="two"><div class="kpi"><span>Average rating</span><b>'+fm(f.avg,2)+' / 5</b></div><div class="kpi"><span>Net sentiment</span><b>'+fm(f.nps_like,0)+' pts</b></div><div class="kpi"><span>Open follow-ups</span><b>'+f.open_followups+'</b></div></div>'+f.by_theme.map(function(t){var n=(t.Positive||0),u=(t.Neutral||0),g=(t.Negative||0),T=n+u+g;return'<div class="row"><span class="lab">'+t.theme+'</span><span class="track"><span class="fill" style="width:'+n/mx*100+'%;background:var(--mint)"></span><span class="fill bad" style="left:'+n/mx*100+'%;width:'+g/mx*100+'%"></span></span><span class="val">+'+n+' / -'+g+'</span></div>'}).join('')+'<p class="note">Green = positive (rating 4-5), red = negative (1-2), by theme tagged from comment keywords.</p>';
$('ev').innerHTML='<thead><tr><th>Event</th><th class="n">Invited</th><th class="n">Registered</th><th class="n">Attended</th><th class="n">Opps</th><th class="n">Per 10 attendees</th></tr></thead><tbody>'+D.events.map(function(e){return'<tr><td>'+e.event_name+'</td><td class="n">'+e.invited+'</td><td class="n">'+(e.registered||'-')+'</td><td class="n">'+(e.attended||'-')+'</td><td class="n">'+(e.opps_created||'-')+'</td><td class="n">'+(e.opps_per_10_attendees==null?'planned':fm(e.opps_per_10_attendees,1))+'</td></tr>'}).join('')+'</tbody>'}
['fLos','fInd','fOwn'].forEach(function(i){$(i).addEventListener('change',render)});
$('reset').addEventListener('click',function(){['fLos','fInd','fOwn'].forEach(function(i){$(i).value=''});render()});
$('thm').addEventListener('click',function(){var r=document.documentElement,d=r.getAttribute('data-theme');var dark=d?d==='dark':window.matchMedia('(prefers-color-scheme: dark)').matches;r.setAttribute('data-theme',dark?'light':'dark')});
static_();render();
</script>"""
open(os.path.join(ROOT, "dashboard", "pipeline_presence_dashboard.html"), "w").write(HTML.replace("__DATA__", J))
print("bytes", len(HTML) + len(J))
