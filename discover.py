#!/usr/bin/env python3
"""Find which career-site platform each company uses, keep the ones with India openings,
and merge them into companies.yaml.

  python discover.py                          # probe every name in india_companies.txt
  python discover.py --names "Zoho, Freshworks, Razorpay"
  python discover.py --min-india 1            # keep boards with at least N India jobs (default 1)

For each name it tries slug variants on Greenhouse, Lever, Ashby, SmartRecruiters, Recruitee and
Workday. A slug can belong to a different company with the same name, so check the printed sample
job/location for anything surprising and delete wrong entries from companies.yaml.
"""
import argparse, re, sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import requests, yaml

BASE = Path(__file__).parent
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (personal job-search agent)"
INDIA = re.compile(r"india|hyderabad|bengaluru|bangalore|pune|chennai|mumbai|gurgaon|gurugram|noida|delhi|kolkata", re.I)


def variants(name):
    n = re.sub(r"[^a-z0-9 ]", "", name.lower().replace("&", " and ")).strip()
    parts = n.split()
    out = ["".join(parts), "-".join(parts)]
    if len(parts) > 1 and len(parts[0]) >= 4:
        out.append(parts[0])
    out.append("".join(parts) + "india")
    camel = "".join(p.capitalize() for p in parts)  # SmartRecruiters slugs are case-sensitive
    return list(dict.fromkeys(out)), camel


def jget(url, **kw):
    r = S.get(url, timeout=12, **kw)
    return r.json() if r.status_code == 200 else None


def probe_greenhouse(s):
    d = jget(f"https://boards-api.greenhouse.io/v1/boards/{s}/jobs")
    jobs = (d or {}).get("jobs") or []
    hits = [j for j in jobs if INDIA.search((j.get("location") or {}).get("name", ""))]
    return (len(hits), f"{hits[0]['title']} / {hits[0]['location']['name']}") if hits else (0, "")


def probe_lever(s):
    d = jget(f"https://api.lever.co/v0/postings/{s}?mode=json")
    hits = [j for j in (d if isinstance(d, list) else []) if INDIA.search((j.get("categories") or {}).get("location", "") or "")]
    return (len(hits), f"{hits[0]['text']} / {hits[0]['categories']['location']}") if hits else (0, "")


def probe_ashby(s):
    d = jget(f"https://api.ashbyhq.com/posting-api/job-board/{s}")
    hits = [j for j in (d or {}).get("jobs", []) if INDIA.search(j.get("location", "") or "")]
    return (len(hits), f"{hits[0]['title']} / {hits[0]['location']}") if hits else (0, "")


def probe_sr(s):
    d = jget(f"https://api.smartrecruiters.com/v1/companies/{s}/postings", params={"country": "in", "limit": 20})
    hits = [j for j in (d or {}).get("content", []) if ((j.get("location") or {}).get("country") or "").lower() == "in"]
    return (len(hits), hits[0]["name"]) if hits else (0, "")


def probe_recruitee(s):
    d = jget(f"https://{s}.recruitee.com/api/offers/")
    hits = [j for j in (d or {}).get("offers", []) if INDIA.search(f"{j.get('location', '')} {j.get('country', '')}")]
    return (len(hits), hits[0]["title"]) if hits else (0, "")


PROBES = [("greenhouse", probe_greenhouse), ("lever", probe_lever), ("ashby", probe_ashby),
          ("recruitee", probe_recruitee)]


def probe_workday(s):
    """Workday needs tenant + shard + site; find them by following the tenant's root URL redirect."""
    for shard in ("wd1", "wd3", "wd5", "wd12", "wd501", "wd103", "wd2", "wd4"):
        host = f"https://{s}.{shard}.myworkdayjobs.com"
        try:
            r = S.get(host + "/", timeout=6, allow_redirects=True)
        except Exception:
            continue
        if "myworkdayjobs.com" not in r.url or r.status_code != 200:
            continue
        parts = [p for p in r.url.split(".com/", 1)[-1].split("?")[0].split("/") if p]
        parts = [p for p in parts if not re.fullmatch(r"[a-z]{2}-[A-Z]{2}", p)]
        if not parts:
            continue
        site = parts[0]
        try:
            q = S.post(f"{host}/wday/cxs/{s}/{site}/jobs", timeout=12,
                       json={"appliedFacets": {}, "limit": 1, "offset": 0, "searchText": "India"})
            total = q.json().get("total", 0) if q.status_code == 200 else 0
        except Exception:
            continue
        if total:
            return dict(ats="workday", tenant=s, wd=shard, site=site), total, "(workday search 'India')"
    return None


def discover(name):
    slugs, camel = variants(name)
    for s in slugs:
        for ats, fn in PROBES:
            try:
                n, sample = fn(s)
            except Exception:
                continue
            if n:
                return name, dict(name=name, ats=ats, slug=s), n, sample
    for s in [camel] + slugs[:1]:
        try:
            n, sample = probe_sr(s)
        except Exception:
            n, sample = 0, ""
        if n:
            return name, dict(name=name, ats="smartrecruiters", slug=s), n, sample
    for s in slugs[:3]:
        w = probe_workday(s.replace("-", ""))
        if w:
            entry, n, sample = w
            return name, dict(name=name, **entry), n, sample
    return name, None, 0, ""


def key(e):
    return (e["ats"], e.get("slug") or e.get("tenant") or e.get("url"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--names", help="comma-separated company names (default: india_companies.txt)")
    ap.add_argument("--min-india", type=int, default=1)
    a = ap.parse_args()
    if a.names:
        names = [n.strip() for n in a.names.split(",") if n.strip()]
    else:
        names = [l.strip() for l in (BASE / "india_companies.txt").read_text().splitlines()
                 if l.strip() and not l.startswith("#")]

    cfg_path = BASE / "companies.yaml"
    cfg = yaml.safe_load(cfg_path.read_text()) if cfg_path.exists() else {}
    cfg.setdefault("companies", [])
    have = {key(e) for e in cfg["companies"]}
    added, missed = 0, []

    with ThreadPoolExecutor(10) as ex:
        for name, entry, n, sample in ex.map(discover, names):
            if entry and n >= a.min_india:
                new = key(entry) not in have
                print(f"{'+' if new else '='} {name:28} {entry['ats']:16} {n:4} India jobs   e.g. {sample}", flush=True)
                if new:
                    cfg["companies"].append(entry)
                    have.add(key(entry))
                    added += 1
            else:
                missed.append(name)

    cfg_path.write_text(yaml.safe_dump(cfg, sort_keys=False, width=200, default_flow_style=None))
    print(f"\nadded {added} companies; {len(cfg['companies'])} total in companies.yaml")
    if missed:
        print(f"no supported platform with India jobs found for {len(missed)}: {', '.join(missed)}")
        print("(custom career sites - add them as 'sitemap' entries, or find the Workday URL by hand)")


if __name__ == "__main__":
    sys.exit(main())
