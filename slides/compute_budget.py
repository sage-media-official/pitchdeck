#!/usr/bin/env python3
"""Recompute the budget with agency fee carved in at 20% of the total (flagship
totals unchanged: Aggr 790 / Bal 399 / Cons 185 L). The 12 channel categories
(the working media budget = 80% of total) are rescaled to fit; agency is no
longer an itemised line."""

# sub-line items: name -> (category, aggr, bal, cons)   [Agency Fees removed]
LINES = [
 ("Print","ATL",35,18,6),("OOH","ATL",40,20,8),("Radio","ATL",15,6,0),("TVC","ATL",50,20,0),
 ("Social Media","Digital",30,20,10),("Online Influencer Marketing","Digital",35,18,8),
 ("3D Plugins on Website","Digital",8,5,3),("Visualizer","Digital",12,7,4),("WhatsApp","Digital",8,5,3),
 ("Emailers","Digital",4,2,1),("Interactive I-brochures","Digital",6,3,2),("Website","Digital",12,8,5),
 ("Blogs","Digital",5,3,2),("SEO","Digital",10,7,4),
 ("SEM + SMM + Display/GDN","Performance",75,44,22),
 ("Images / Adaptations","Content",10,6,3),("Videos — Product","Content",12,7,4),
 ("AID Engagements","Content",10,6,3),("Digital Film — DVC","Content",18,8,0),
 ("Walk-through","Content",8,4,2),("Corporate AV","Content",6,4,2),
 ("Exhibition for the year","Exhibition",35,20,10),
 ("Signage","Visibility",25,15,8),("Counter Branding","Visibility",20,12,7),
 ("Carpenter (Nukkad)","Meets",30,20,12),("Dealer / Contractor","Meets",20,13,8),
 ("Portal Set-up","Loyalty",10,7,5),("Operational Cost","Loyalty",15,9,5),("Contractor + AID","Loyalty",15,9,5),
 ("Collaterals / Merchandise / Gifts","Collaterals",18,10,6),
 ("Samples","Samples",25,15,9),
 ("Experience Stores / Centers","Experience",40,18,0),
 ("Brand Ambassador","Brand Ambassador",80,0,0),
]
CAT_ORDER = ["ATL","Digital","Performance","Content","Exhibition","Visibility","Meets","Loyalty","Collaterals","Samples","Experience","Brand Ambassador"]
TOTAL = {"aggr":790,"bal":399,"cons":185}
MEDIA = {k:round(0.8*v) for k,v in TOTAL.items()}          # 632 / 319 / 148
AGENCY = {k:TOTAL[k]-MEDIA[k] for k in TOTAL}              # 158 / 80 / 37

def largest_remainder(raw, target):
    """round list of floats to ints summing exactly to target"""
    floors=[int(x) for x in raw]; rem=target-sum(floors)
    fracs=sorted(range(len(raw)), key=lambda i: raw[i]-floors[i], reverse=True)
    for i in range(rem): floors[fracs[i]]+=1
    return floors

scaled={}  # scen -> list of ints aligned to LINES
for si,scen in enumerate(["aggr","bal","cons"]):
    cur=[l[2+si] for l in LINES]
    s=sum(cur); f=MEDIA[scen]/s
    raw=[c*f for c in cur]
    scaled[scen]=largest_remainder(raw, MEDIA[scen])

# category aggregation
def cat_totals(scen):
    d={c:0 for c in CAT_ORDER}
    for l,v in zip(LINES, scaled[scen]): d[l[1]]+=v
    return d
cats={s:cat_totals(s) for s in scaled}

print("=== MEDIA / AGENCY / TOTAL ===")
for s in ["aggr","bal","cons"]:
    print(f"{s}: media {MEDIA[s]}  agency {AGENCY[s]} ({AGENCY[s]/TOTAL[s]*100:.1f}%)  total {TOTAL[s]}")

print("\n=== CATEGORY TABLE (rescaled media) ===")
print(f"{'Category':22} {'Aggr':>5} {'Bal':>5} {'Cons':>5}")
for c in CAT_ORDER:
    print(f"{c:22} {cats['aggr'][c]:>5} {cats['bal'][c]:>5} {cats['cons'][c]:>5}")
print(f"{'MEDIA TOTAL':22} {MEDIA['aggr']:>5} {MEDIA['bal']:>5} {MEDIA['cons']:>5}")
print(f"{'Agency (20%)':22} {AGENCY['aggr']:>5} {AGENCY['bal']:>5} {AGENCY['cons']:>5}")
print(f"{'GRAND TOTAL':22} {TOTAL['aggr']:>5} {TOTAL['bal']:>5} {TOTAL['cons']:>5}")

# Balanced by-channel (sorted desc) for bar chart
print("\n=== BALANCED by-channel (desc) ===")
balcats=sorted([(c,cats['bal'][c]) for c in CAT_ORDER if cats['bal'][c]>0], key=lambda x:-x[1])
for c,v in balcats: print(f"{c:14} {v}")

# strategic buckets
def buckets(scen):
    c=cats[scen]
    return {
      "ATL / Brand": c["ATL"],
      "Demand Gen": c["Digital"]+c["Performance"],
      "Channel & Influencer": c["Exhibition"]+c["Visibility"]+c["Meets"]+c["Loyalty"]+c["Collaterals"]+c["Samples"]+c["Experience"]+c["Brand Ambassador"],
      "Content": c["Content"],
      "Agency (20%)": AGENCY[scen],
    }
print("\n=== STRATEGIC BUCKETS ===")
for s in ["aggr","bal","cons"]:
    b=buckets(s); print(s, b, "sum", sum(b.values()))

# donut % of TOTAL (balanced)
print("\n=== DONUT % of total (balanced) ===")
b=buckets("bal")
for k,v in b.items(): print(f"{k:24} {v:>4}  {v/TOTAL['bal']*100:5.1f}%")

# appendix sub-lines rescaled
print("\n=== APPENDIX sub-lines (rescaled) ===")
print(f"{'Line':30} {'Aggr':>5} {'Bal':>5} {'Cons':>5}")
for i,l in enumerate(LINES):
    print(f"{l[0]:30} {scaled['aggr'][i]:>5} {scaled['bal'][i]:>5} {scaled['cons'][i]:>5}")
