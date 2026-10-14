#!/usr/bin/env python3
"""Lite career-portal job agent.

Polls company career-site APIs (Greenhouse, Lever, Ashby, SmartRecruiters, Workday),
keeps only India + entry-level/backend roles, stores everything in SQLite, and
regenerates jobs.xlsx. Your Status / Notes columns in the Excel are preserved.

  python agent.py                       # one run
  python agent.py --loop --every 180    # run forever, every 3 hours
"""
import argparse, html, json, re, sqlite3, time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import requests, yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).parent
DB_PATH, XLSX = BASE / "jobs.db", BASE / "jobs.xlsx"
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (personal job-search agent)"
KNOWN = set()  # URLs already in the DB (lets the sitemap adapter skip re-fetching pages)

DEFAULT_FILTERS = {
    "max_years": 1,
    "title_include": r"software|developer|engineer|backend|back-end|sde|python|java|graduate|trainee|associate|apprentice|mainframe|data|analyst|intern",
    "title_exclude": r"\b(senior|sr|lead|principal|staff|manager|architect|director|head|vp|iii|iv)\b",
    "location_include": r"india|hyderabad|bengaluru|bangalore|pune|chennai|mumbai|gurgaon|gurugram|noida|delhi|kolkata|ahmedabad|kochi|\d+ locations?",
}


def log(msg):
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def get(url, **kw):
    r = S.get(url, timeout=25, **kw)
    r.raise_for_status()
    return r.json()


def clean(t):
    t = html.unescape(t or "")
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t).strip()


# ---------------- ATS adapters: each yields dicts with title/location/url/posted + jd or jd_fn ----------------
def greenhouse(c):
    d = get(f"https://boards-api.greenhouse.io/v1/boards/{c['slug']}/jobs?content=true")
    for j in d["jobs"]:
        yield dict(title=j["title"], location=(j.get("location") or {}).get("name", ""),
                   url=j["absolute_url"], posted=(j.get("updated_at") or "")[:10], jd=clean(j.get("content")))


def lever(c):
    for j in get(f"https://api.lever.co/v0/postings/{c['slug']}?mode=json"):
        yield dict(title=j["text"], location=(j.get("categories") or {}).get("location", ""),
                   url=j["hostedUrl"], posted=datetime.fromtimestamp(j["createdAt"] / 1000).strftime("%Y-%m-%d"),
                   jd=clean(j.get("descriptionPlain") or j.get("description")))


def ashby(c):
    for j in get(f"https://api.ashbyhq.com/posting-api/job-board/{c['slug']}")["jobs"]:
        yield dict(title=j["title"], location=j.get("location", ""), url=j["jobUrl"],
                   posted=(j.get("publishedAt") or "")[:10],
                   jd=clean(j.get("descriptionHtml") or j.get("descriptionPlain")))


def _sr_jd(slug, pid):
    secs = get(f"https://api.smartrecruiters.com/v1/companies/{slug}/postings/{pid}")["jobAd"]["sections"]
    return clean(" ".join(v.get("text", "") for v in secs.values() if isinstance(v, dict)))


def smartrecruiters(c):
    off = 0
    while True:
        d = get(f"https://api.smartrecruiters.com/v1/companies/{c['slug']}/postings",
                params={"country": "in", "limit": 100, "offset": off})
        for j in d["content"]:
            loc = j.get("location") or {}
            yield dict(title=j["name"], location=f"{loc.get('city', '')}, India",
                       url=f"https://jobs.smartrecruiters.com/{c['slug']}/{j['id']}",
                       posted=(j.get("releasedDate") or "")[:10],
                       jd_fn=lambda pid=j["id"]: _sr_jd(c["slug"], pid))
        off += 100
        if off >= d.get("totalFound", 0):
            break


def workday(c):
    host = f"https://{c['tenant']}.{c['wd']}.myworkdayjobs.com"
    api = f"{host}/wday/cxs/{c['tenant']}/{c['site']}"
    seen = set()
    for term in c.get("search", ["software", "engineer", "graduate", "associate", "trainee"]):
        off = 0
        while off < c.get("max", 200):
            r = S.post(f"{api}/jobs", timeout=25,
                       json={"appliedFacets": {}, "limit": 20, "offset": off, "searchText": term})
            r.raise_for_status()
            d = r.json()
            for j in d.get("jobPostings", []):
                p = j["externalPath"]
                if p in seen:
                    continue
                seen.add(p)
                yield dict(title=j["title"], location=j.get("locationsText", ""),
                           url=f"{host}/en-US/{c['site']}{p}", posted=j.get("postedOn", ""),
                           jd_fn=lambda p=p: clean(get(f"{api}{p}")["jobPostingInfo"]["jobDescription"]))
            off += 20
            if off >= d.get("total", 0):
                break


def recruitee(c):
    for j in get(f"https://{c['slug']}.recruitee.com/api/offers/")["offers"]:
        yield dict(title=j["title"], location=j.get("location") or f"{j.get('city', '')}, {j.get('country', '')}",
                   url=j["careers_url"], posted=(j.get("published_at") or "")[:10], jd=clean(j.get("description")))


def _jsonld_location(d):
    locs = d.get("jobLocation") or []
    locs = locs if isinstance(locs, list) else [locs]
    out = []
    for l in locs:
        a = (l or {}).get("address", {}) if isinstance(l, dict) else {}
        ctry = a.get("addressCountry", "")
        ctry = ctry.get("name", "") if isinstance(ctry, dict) else ctry
        out.append(", ".join(filter(None, [a.get("addressLocality", ""), a.get("addressRegion", ""), ctry])))
    return " | ".join(out)


def sitemap(c):
    """Custom career sites: walk sitemap.xml, fetch each unseen job page, parse its schema.org JobPosting JSON-LD.
    Config: {name, ats: sitemap, url: https://example.com/sitemap.xml, match: 'regex for job URLs', max: 300}"""
    todo, urls = [c["url"]], []
    while todo:
        x = S.get(todo.pop(), timeout=25).text
        for u in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", x):
            if u.endswith(".xml"):
                todo.append(u)
            elif re.search(c.get("match", r"job|career|opening"), u, re.I):
                urls.append(u)
    for u in urls[: c.get("max", 300)]:
        if u in KNOWN:
            yield dict(url=u, known=True)
            continue
        try:
            page = S.get(u, timeout=25).text
        except Exception:
            continue
        time.sleep(0.3)
        for blob in re.findall(r'<script[^>]*ld\+json[^>]*>(.*?)</script>', page, re.S | re.I):
            try:
                data = json.loads(blob)
            except Exception:
                continue
            for d in (data if isinstance(data, list) else data.get("@graph", [data])):
                if isinstance(d, dict) and d.get("@type") == "JobPosting":
                    yield dict(title=d.get("title", ""), location=_jsonld_location(d), url=u,
                               posted=(d.get("datePosted") or "")[:10], jd=clean(d.get("description")))
                    break


ADAPTERS = dict(greenhouse=greenhouse, lever=lever, ashby=ashby,
                smartrecruiters=smartrecruiters, workday=workday, recruitee=recruitee, sitemap=sitemap)


def fetch_company(c):
    try:
        jobs = list(ADAPTERS[c["ats"]](c))
        log(f"{c['name']}: {len(jobs)} listings")
        return c, jobs
    except Exception as e:
        log(f"{c['name']}: FAILED ({e})")
        return c, None


# ---------------- filtering ----------------
def min_years(jd):
    """Smallest 'N years ... experience' requirement found in the JD; None if not stated."""
    found = []
    for m in re.finditer(r"(\d{1,2})\s*(?:\+|-|–|to)?\s*\d{0,2}\s*\+?\s*(?:years|yrs)", jd, re.I):
        ctx = jd[max(0, m.start() - 80): m.end() + 80].lower()
        if "experience" in ctx:
            found.append(int(m.group(1)))
    return min(found) if found else None


# ---------------- storage ----------------
def init_db():
    db = sqlite3.connect(DB_PATH)
    db.execute("""CREATE TABLE IF NOT EXISTS jobs(
        url TEXT PRIMARY KEY, company TEXT, title TEXT, location TEXT, posted TEXT, jd TEXT,
        exp INTEGER, first_seen TEXT, last_seen TEXT, active INTEGER DEFAULT 1, keep INTEGER DEFAULT 1,
        status TEXT DEFAULT 'New', notes TEXT DEFAULT '')""")
    return db


def sync_from_xlsx(db):
    """Pull your Status/Notes edits from the Excel back into the DB before regenerating it."""
    if not XLSX.exists():
        return
    try:
        wb = load_workbook(XLSX)
        for ws in wb:
            hdr = [c.value for c in ws[1]]
            if not {"Link", "Status", "Notes"} <= set(hdr):
                continue
            li, si, ni = hdr.index("Link"), hdr.index("Status"), hdr.index("Notes")
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r[li]:
                    st, nt = r[si] or "New", r[ni] or ""
                    db.execute("UPDATE jobs SET status=?, notes=? WHERE url=? AND (status IS NOT ? OR notes IS NOT ?)",
                               (st, nt, r[li], st, nt))
        db.commit()
    except Exception as e:
        log(f"could not read existing Excel (is it open?): {e}")


COLS = ["Company", "Title", "Location", "Posted", "Exp (yrs)", "Link", "First Seen", "Status", "Notes", "JD"]
WIDTHS = [20, 45, 22, 12, 9, 40, 12, 16, 30, 80]


def export(db):
    wb = Workbook()
    for name, active in (("Jobs", 1), ("Closed", 0)):
        ws = wb.active if name == "Jobs" else wb.create_sheet(name)
        ws.title = name
        ws.append(COLS)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF", name="Arial")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        rows = db.execute("""SELECT company,title,location,posted,exp,url,first_seen,status,notes,jd
                             FROM jobs WHERE keep=1 AND active=? ORDER BY first_seen DESC, company""", (active,))
        for r in rows:
            r = list(r)
            r[9] = (r[9] or "")[:32000]
            ws.append(r)
            link = ws.cell(ws.max_row, 6)
            link.hyperlink = r[5]
            link.font = Font(color="0563C1", underline="single", name="Arial")
        for i, w in enumerate(WIDTHS, 1):
            ws.column_dimensions[chr(64 + i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        dv = DataValidation(type="list", allow_blank=True,
                            formula1='"New,Interested,Applied,Referral asked,Interview,Rejected,Skip"')
        ws.add_data_validation(dv)
        dv.add(f"H2:H{max(ws.max_row, 2) + 500}")
    tmp = XLSX.with_suffix(".tmp.xlsx")
    try:
        wb.save(tmp)
        tmp.replace(XLSX)
        log("jobs.xlsx updated")
    except PermissionError:
        log("jobs.xlsx is open in Excel - will update on the next run")


# ---------------- one polling cycle ----------------
def run_once():
    cfg = yaml.safe_load((BASE / "companies.yaml").read_text())
    f = {**DEFAULT_FILTERS, **(cfg.get("filters") or {})}
    title_in, title_out = re.compile(f["title_include"], re.I), re.compile(f["title_exclude"], re.I)
    loc_in = re.compile(f["location_include"], re.I)

    db = init_db()
    t0 = db.total_changes
    sync_from_xlsx(db)
    synced = db.total_changes - t0
    snap = lambda: db.execute("SELECT COUNT(*), COALESCE(SUM(active),0) FROM jobs WHERE keep=1").fetchone()
    before = snap()
    KNOWN.clear()
    KNOWN.update(r[0] for r in db.execute("SELECT url FROM jobs"))
    start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with ThreadPoolExecutor(8) as ex:
        results = list(ex.map(fetch_company, cfg["companies"]))

    pending, ok = [], set()
    for c, jobs in results:
        if jobs is None:
            continue
        ok.add(c["name"])
        for j in jobs:
            if j.get("known"):
                db.execute("UPDATE jobs SET last_seen=?, active=1 WHERE url=?", (start, j["url"]))
                continue
            if not title_in.search(j["title"]) or title_out.search(j["title"]):
                continue
            if j["location"] and not loc_in.search(j["location"]):
                continue
            if db.execute("SELECT 1 FROM jobs WHERE url=?", (j["url"],)).fetchone():
                db.execute("UPDATE jobs SET last_seen=?, active=1 WHERE url=?", (start, j["url"]))
            else:
                pending.append((c["name"], j))

    def enrich(item):
        name, j = item
        try:
            jd = j.get("jd") or (j["jd_fn"]() if j.get("jd_fn") else "")
        except Exception:
            jd = ""
        return name, j, jd

    new = 0
    with ThreadPoolExecutor(6) as ex:
        for name, j, jd in ex.map(enrich, pending):
            exp = min_years(jd)
            keep = int(exp is None or exp <= f["max_years"])
            db.execute("INSERT OR IGNORE INTO jobs(url,company,title,location,posted,jd,exp,first_seen,last_seen,keep)"
                       " VALUES(?,?,?,?,?,?,?,?,?,?)",
                       (j["url"], name, j["title"], j["location"], j.get("posted", ""), jd, exp, start, start, keep))
            new += keep

    for name in ok:  # only close jobs for companies we fetched successfully this run
        db.execute("UPDATE jobs SET active=0 WHERE company=? AND last_seen<? AND keep=1", (name, start))
    changed = bool(synced or pending or snap() != before or not XLSX.exists())
    if not changed:
        db.rollback()  # nothing new: leave jobs.db / jobs.xlsx byte-identical (keeps git history small)
        log("no changes")
    else:
        db.commit()
        log(f"{new} new matching jobs")
        export(db)
    db.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--loop", action="store_true")
    ap.add_argument("--every", type=int, default=180, help="minutes between runs")
    a = ap.parse_args()
    while True:
        try:
            run_once()
        except Exception as e:
            log(f"run failed: {e}")
        if not a.loop:
            break
        time.sleep(a.every * 60)


if __name__ == "__main__":
    main()
