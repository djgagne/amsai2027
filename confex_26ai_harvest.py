#!/usr/bin/env python3
"""
Harvest sessions, presentations, authors and abstracts for one Confex "Program"
(default: AMS 107th Annual Meeting, 26th Conference on Artificial Intelligence for
Environmental Science, Program/1955) via the public Confex Meeting API, and write:

  26AI_program.xlsx   - Presentations / Sessions / Authors sheets (filterable)
  26AI_program.json   - same data, compact, for the 26AI Program Explorer page
  confex_cache/       - every raw API response, so re-runs are instant and nothing is lost

API reference (Confex Meeting API, 2020):
  GET https://{client}.confex.com/{client}/{meeting}/meetingapi.cgi/{Object}/{id}
  GET .../{Object}/{id}/{ChildList_X}   -> full child objects in one request
  Tree: Program -> [Symposium] -> Session -> Paper ; Entry -> Role -> Person

Usage:
  pip install requests openpyxl
  python confex_26ai_harvest.py                 # full harvest
  python confex_26ai_harvest.py --inspect       # print field names of each object type
  python confex_26ai_harvest.py --program 1955 --meeting 107ANNUAL --client ams
  python confex_26ai_harvest.py --refresh       # ignore cache, re-download
  python confex_26ai_harvest.py --repair 26AI_program.json --out 26AI_program_fixed
  python confex_26ai_harvest.py --repair new.json --slots-from first_run.json --out fixed
"""
import argparse
import html
import json
import os
import re
import sys
import time

try:
    import requests
except ImportError:
    sys.exit("Please `pip install requests openpyxl` first.")

# ----------------------------------------------------------------------------- HTTP

class Confex:
    def __init__(self, client, meeting, cache_dir, refresh=False, delay=0.25):
        self.api = f"https://{client}.confex.com/{client}/{meeting}/meetingapi.cgi"
        self.app = f"https://{client}.confex.com/{client}/{meeting}/meetingapp.cgi"
        self.cache_dir = cache_dir
        self.refresh = refresh
        self.delay = delay
        self.s = requests.Session()
        self.s.headers.update({
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "26AI-program-harvester (personal research use)",
        })
        os.makedirs(cache_dir, exist_ok=True)

    def _cache_path(self, path):
        return os.path.join(self.cache_dir, re.sub(r"[^A-Za-z0-9_.-]+", "_", path) + ".json")

    def get(self, path):
        """GET an API path like 'Program/1955' or 'Session/123/ChildList_Paper'."""
        cp = self._cache_path(path)
        if not self.refresh and os.path.exists(cp):
            with open(cp, encoding="utf-8") as f:
                return json.load(f)
        url = f"{self.api}/{path}"
        for attempt in range(5):
            try:
                r = self.s.get(url, timeout=60)
                if r.status_code == 404:
                    return None
                if r.status_code in (429, 500, 502, 503, 504):
                    raise requests.HTTPError(f"HTTP {r.status_code}")
                r.raise_for_status()
                data = r.json()
                with open(cp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False)
                time.sleep(self.delay)
                return data
            except (requests.RequestException, ValueError) as e:
                wait = 2 ** attempt
                print(f"  ! {path}: {e} - retrying in {wait}s", file=sys.stderr)
                time.sleep(wait)
        print(f"  ! giving up on {path}", file=sys.stderr)
        return None

    def get_children(self, obj, listname):
        """Full child objects for obj[listname]; batch request first, then one-by-one fallback."""
        refs = obj.get(listname) or []
        if not refs:
            return []
        batch = self.get(f"{obj['_url']}/{listname}")
        if isinstance(batch, list) and batch and isinstance(batch[0], dict):
            return batch
        out = []
        for ref in refs:
            ref = ref if isinstance(ref, str) else (ref or {}).get("_url")
            if ref:
                o = self.get(ref)
                if isinstance(o, dict):
                    out.append(o)
        return out

# ----------------------------------------------------------------------------- helpers

def clean(text):
    """Strip HTML tags/entities and normalise whitespace (keeps paragraph breaks)."""
    if text is None:
        return ""
    if isinstance(text, (list, dict)):
        text = json.dumps(text, ensure_ascii=False)
    t = str(text)
    t = re.sub(r"(?i)<\s*br\s*/?>|</\s*p\s*>|</\s*div\s*>|</\s*li\s*>", "\n", t)
    t = re.sub(r"<[^>]+>", "", t)
    t = html.unescape(t).replace("\r", "").replace("_x000D_", "")
    # entity-escaped or truncated HTML (e.g. '&lt;span style=' or '<span style=') left after unescaping
    t = re.sub(r"</?(?:span|div|p|font|b|i|u|em|strong|sup|sub|br|a|li|ul|ol)\b[^>]*>?", "", t, flags=re.I)
    t = re.sub(r"[ \t\xa0]+", " ", t)
    t = re.sub(r"\n\s*\n\s*(\n\s*)+", "\n\n", t)
    return t.strip()

def pick(obj, *keys):
    """First non-empty value among candidate field names (case-insensitive)."""
    if not isinstance(obj, dict):
        return ""
    lower = {k.lower(): k for k in obj}
    for k in keys:
        real = lower.get(k.lower())
        if real is not None and obj[real] not in (None, "", [], {}):
            return obj[real]
    return ""

def list_keys(obj, needle):
    return [k for k in obj if k.startswith(("ChildList_", "ParentList_")) and needle.lower() in k.lower()]

def truthy(v):
    return str(v).strip().lower() in ("1", "y", "yes", "true", "t", "x")

def person_name(role):
    first = clean(pick(role, "FirstName", "Person_FirstName", "First"))
    mid = clean(pick(role, "MiddleName", "Person_MiddleName", "Middle"))
    last = clean(pick(role, "LastName", "Person_LastName", "Last"))
    name = " ".join(p for p in (first, mid, last) if p)
    return name or clean(pick(role, "Name", "FullName", "DisplayName", "label"))

def affiliation(role):
    aff = clean(pick(role, "Affiliation", "Institution", "Organization", "Company", "Address_Affiliation"))
    if not aff:
        a = pick(role, "Address")
        if isinstance(a, dict):
            aff = clean(pick(a, "Affiliation", "Institution", "Organization"))
    if not aff:  # any scalar field whose name looks like an affiliation
        for k, v in role.items():
            if (isinstance(v, str) and v.strip() and not k.startswith(("_", "ChildList_", "ParentList_"))
                    and re.search(r"affil|institut|organi[sz]ation|employer|company", k, re.I)
                    and not re.search(r"id$|url", k, re.I)):
                aff = clean(v)
                break
    return aff

def role_kind(role):
    return clean(pick(role, "Role", "RoleName", "RoleType", "Type", "label"))

def is_presenter(role):
    if any(truthy(pick(role, k)) for k in ("Presenter", "isPresenter", "Presenting", "IsPresenting", "Speaker")):
        return True
    return "present" in role_kind(role).lower() or "speaker" in role_kind(role).lower()

def position(role):
    try:
        return int(pick(role, "Position", "Order", "AuthorOrder", "Sort", "Rank") or 999)
    except (TypeError, ValueError):
        return 999

def norm_date(v):
    v = clean(v)
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", v)
    if m:
        return m.group(0)
    m = re.search(r"(\d{4})(\d{2})(\d{2})", v)
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else v

def norm_time(v):
    """Return 24h 'HH:MM' from '15:00:00', '3:00 PM', '2027-01-12T15:00:00' etc."""
    v = clean(v)
    m = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp]\.?[Mm]\.?)?", v.split("T")[-1])
    if not m:
        return ""
    h, mi, ap = int(m.group(1)), int(m.group(2)), (m.group(3) or "").lower().replace(".", "")
    if ap == "pm" and h != 12:
        h += 12
    if ap == "am" and h == 12:
        h = 0
    return f"{h:02d}:{mi:02d}"

def slot_fields(obj):
    """Date / start / end / room from any slot-like object, by exact names first, then by key pattern."""
    out = {
        "date": norm_date(pick(obj, "Date", "SlotDate", "SessionDate", "StartDate", "Day", "Slot_Date")),
        "start": norm_time(pick(obj, "StartTime", "SlotStartTime", "TimeStart", "Start", "Slot_StartTime", "Time")),
        "end": norm_time(pick(obj, "EndTime", "SlotEndTime", "TimeEnd", "End", "Slot_EndTime")),
        "room": clean(pick(obj, "Location", "Room", "RoomName", "SlotLocation", "Slot_Location",
                           "LocationName", "Place", "Venue")),
    }
    if not isinstance(obj, dict):
        return out
    for k, v in obj.items():
        if k.startswith(("_", "ChildList_", "ParentList_")) or isinstance(v, (list, dict)) or v in (None, ""):
            continue
        kl = k.lower()
        if not out["date"] and "date" in kl and re.search(r"\d{4}-?\d{2}-?\d{2}", str(v)):
            out["date"] = norm_date(v)
        if not out["end"] and "end" in kl and norm_time(v) and "date" not in kl:
            out["end"] = norm_time(v)
        if not out["start"] and "start" in kl and norm_time(v) and "date" not in kl:
            out["start"] = norm_time(v)
        if not out["room"] and re.search(r"room|location|venue", kl) and not re.search(r"id$|url|map", kl):
            out["room"] = clean(v)
    return out

def is_slot(o):
    return isinstance(o, dict) and (o.get("_name") == "SlotData" or str(o.get("_url", "")).startswith("SlotData/"))

def slot_paper_id(o):
    """SlotData/Paper491770_Slot31888 -> 'Paper/491770'."""
    m = re.search(r"Paper(\d+)", str(o.get("_url", "")))
    if m:
        return f"Paper/{m.group(1)}"
    for k in list_keys(o, "Paper"):
        refs = o.get(k) or []
        if refs:
            return refs[0] if isinstance(refs[0], str) else refs[0].get("_url")
    ref = pick(o, "Paper", "Entry", "EntryURL")
    return ref if isinstance(ref, str) and ref.startswith("Paper/") else ""

def session_slot(cx, s):
    """The session's own SlotData record, if the API exposes one. Only SlotData records are read for
    schedule info: Paper and Session records carry other dates (submission, creation) that look similar."""
    info = {"date": "", "start": "", "end": "", "room": ""}
    for key in list_keys(s, "SlotData"):
        for ref in (s.get(key) or [])[:3]:
            ref = ref if isinstance(ref, str) else (ref or {}).get("_url")
            if not ref or "Paper" in ref:
                continue
            slot = cx.get(ref)
            if is_slot(slot):
                for k, v in slot_fields(slot).items():
                    info[k] = info[k] or v
    return info

def fmt12(hhmm):
    if not hhmm:
        return ""
    h, m = map(int, hhmm.split(":"))
    return f"{(h % 12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}"

def add_minutes(hhmm, n):
    h, m = map(int, hhmm.split(":"))
    t = h * 60 + m + n
    return f"{t // 60:02d}:{t % 60:02d}"

def valid_time(t):
    return bool(t) and t != "00:00"

def valid_room(r):
    r = (r or "").strip()
    return bool(r) and r not in ("0", "-", "TBD", "TBA")

def finalize_times(sessions, presentations, slots_by_paper, slots_by_session=None):
    """Schedule rules:
    - A talk's date/start/end/room come only from its own SlotData record.
    - A session's date and start come from its talks (or, with no talks, from its SlotData).
    - A session's own SlotData can supply its end and room, only if they are plausible.
    - Ends not in the data are inferred and flagged in *_estimated fields.
    - A 'room' used by two different sessions at the same date and start is not a real room; it is dropped."""
    slots_by_session = slots_by_session or {}
    for p in presentations:
        sl = slots_by_paper.get(p["paper_id"]) or {}
        p["date"] = norm_date(sl.get("date", ""))
        p["time"] = norm_time(sl.get("start", "")) if valid_time(norm_time(sl.get("start", ""))) else ""
        e = norm_time(sl.get("end", ""))
        p["end"] = e if valid_time(e) and p["time"] and e > p["time"] else ""
        p["room"] = sl.get("room", "") if valid_room(sl.get("room", "")) else ""
        p["end_estimated"] = ""

    by_sess = {}
    for p in presentations:
        by_sess.setdefault(p["session_id"], []).append(p)

    # typical talk length (minutes between consecutive oral talks)
    gaps = []
    for ps in by_sess.values():
        if any("poster" in (x["session_type"] or "").lower() for x in ps):
            continue
        starts = sorted({to_minutes(x["time"]) for x in ps if x["time"]})
        gaps += [b - a for a, b in zip(starts, starts[1:]) if 0 < b - a <= 60]
    talk = max(set(gaps), key=gaps.count) if gaps else 15

    for s in sessions:
        ps = by_sess.get(s["session_id"], [])
        timed = [x for x in ps if x["time"] and x["date"]]
        own = s.get("_slot") or {}
        extra = [sl for sl in slots_by_session.get(s["session_id"], []) if valid_time(norm_time(sl.get("start")))]
        poster = "poster" in (s.get("type") or "").lower()
        if timed:
            s["date"] = min(x["date"] for x in timed)
            s["start"] = min(x["time"] for x in timed if x["date"] == s["date"])
        elif valid_time(norm_time(own.get("start"))) and own.get("date"):
            s["date"], s["start"] = norm_date(own["date"]), norm_time(own["start"])
        elif extra:
            first = min(extra, key=lambda sl: (norm_date(sl.get("date")), norm_time(sl.get("start"))))
            s["date"], s["start"] = norm_date(first.get("date")), norm_time(first.get("start"))
        else:
            s["date"], s["start"] = "", ""
        own_end = norm_time(own.get("end"))
        s["end"] = own_end if valid_time(own_end) and s["start"] and own_end > s["start"] else ""
        s["end_estimated"] = ""
        rooms = sorted({x["room"] for x in ps if x["room"]})
        s["room"] = own.get("room") if valid_room(own.get("room")) else "; ".join(rooms)
        if not s["end"] and timed and not poster:
            s["end"] = add_minutes(max(x["time"] for x in timed), talk)
            s["end_estimated"] = "Yes"

    # Parallel oral sessions sharing a date and start share a block: snap inferred ends to the block's end.
    block_end = {}
    for s in sessions:
        if s["end"] and s["start"] and "poster" not in (s.get("type") or "").lower():
            k = (s["date"], s["start"])
            block_end[k] = max(block_end.get(k, ""), s["end"])
    for s in sessions:
        if s["end_estimated"]:
            s["end"] = block_end.get((s["date"], s["start"]), s["end"])

    # A room can't host two sessions at once: drop room values that collide.
    seen, bad = {}, set()
    for s in sessions:
        for r in filter(None, (x.strip() for x in s["room"].split(";"))):
            k = (r, s["date"], s["start"])
            if s["start"] and k in seen and seen[k] != s["session_id"]:
                bad.add(r)
            seen.setdefault(k, s["session_id"])
    if bad:
        print(f"  ! ignoring room values used by parallel sessions (not real rooms): {', '.join(sorted(bad))}")
        for s in sessions:
            s["room"] = "; ".join(r for r in (x.strip() for x in s["room"].split(";")) if r and r not in bad)
        for p in presentations:
            if p["room"] in bad:
                p["room"] = ""

    for s in sessions:
        ps = by_sess.get(s["session_id"], [])
        poster = "poster" in (s.get("type") or "").lower()
        timed = [x for x in ps if x["time"]]
        if not poster:  # each oral talk ends when the next one starts (or at session end)
            starts = sorted({x["time"] for x in timed})
            for x in timed:
                if x["end"]:
                    continue
                later = [t for t in starts if t > x["time"]]
                x["end"] = later[0] if later else s["end"]
                x["end_estimated"] = "" if later else s["end_estimated"]
        for x in ps:
            x["session_start"], x["session_end"] = s["start"], s["end"]
            x["date"] = x["date"] or s["date"]
            x["room"] = x["room"] or s["room"]
            if not x["time"] and poster:
                x["time"] = s["start"]
        s.pop("_slot", None)
    return talk

def roles_for(cx, entry):
    roles = []
    for key in list_keys(entry, "Role"):
        if key.startswith("ChildList_"):
            roles.extend(r for r in cx.get_children(entry, key) if isinstance(r, dict))
    # Some meetings only store a Person link on the Role; pull the person if names are missing.
    for r in roles:
        if not person_name(r):
            for pk in list_keys(r, "Person"):
                refs = r.get(pk) or []
                if refs:
                    p = cx.get(refs[0] if isinstance(refs[0], str) else refs[0].get("_url"))
                    if isinstance(p, dict):
                        for k, v in p.items():
                            r.setdefault(k, v)
                    break
    return roles

# ----------------------------------------------------------------------------- walk

def collect_sessions(cx, program):
    """Program -> sessions (directly, via Symposium, or via SlotData)."""
    sessions, seen = [], set()

    def add(objs):
        for s in objs:
            if isinstance(s, dict) and s.get("_name", "Session") == "Session" and s.get("_url") not in seen:
                seen.add(s.get("_url"))
                sessions.append(s)

    for key in list_keys(program, "Session"):
        if key.startswith("ChildList_"):
            add(cx.get_children(program, key))
    for key in list_keys(program, "Symposium"):
        if key.startswith("ChildList_"):
            for sym in cx.get_children(program, key):
                for k2 in list_keys(sym, "Session"):
                    if k2.startswith("ChildList_"):
                        add(cx.get_children(sym, k2))
    if not sessions:
        for key in list_keys(program, "SlotData"):
            for slot in cx.get_children(program, key):
                for k2 in list_keys(slot, "Session"):
                    add(cx.get_children(slot, k2))
                for k in ("Session", "Entry", "EntryURL"):
                    ref = slot.get(k)
                    if isinstance(ref, str) and ref.startswith("Session/"):
                        add([cx.get(ref)])
    return sessions

def sort_all(sessions, presentations, authors):
    sessions.sort(key=lambda s: (s["n_papers"] == 0, s["date"] or "9999", to_minutes(s["start"]),
                                 s["number"], s["title"]))
    order = {s["session_id"]: k for k, s in enumerate(sessions)}
    presentations.sort(key=lambda p: (order.get(p["session_id"], 9999), p["date"] or "9999",
                                      to_minutes(p["time"]), p["paper_number"], p["title"]))
    when = {p["paper_id"]: p for p in presentations}
    for a in authors:
        p = when.get(a.get("paper_id"))
        if p:
            a["date"], a["time"] = p["date"], p["time"]
    authors.sort(key=lambda a: (a["author"].split()[-1].lower() if a["author"] else "", a["author"]))

def to_minutes(t):
    m = re.match(r"\s*(\d{1,2}):(\d{2})\s*([AaPp][Mm])?", t or "")
    if not m:
        return 9999
    h, mi, ap = int(m.group(1)), int(m.group(2)), (m.group(3) or "").lower()
    if ap == "pm" and h != 12:
        h += 12
    if ap == "am" and h == 12:
        h = 0
    return h * 60 + mi

def harvest(cx, program_id):
    program = cx.get(f"Program/{program_id}")
    if not isinstance(program, dict):
        sys.exit(f"Could not load Program/{program_id} from {cx.api}")
    conf_title = clean(pick(program, "Title", "Name", "label")) or f"Program {program_id}"
    print(f"Program: {conf_title}")

    raw_sessions = collect_sessions(cx, program)
    print(f"Found {len(raw_sessions)} sessions")

    sessions, presentations, authors = [], [], []
    slots_by_paper, slots_by_session = {}, {}
    for i, s in enumerate(raw_sessions, 1):
        sid = s.get("_url", "")
        title = clean(pick(s, "Title", "SessionTitle", "Name", "label"))
        number = clean(pick(s, "FinalSessionNumber", "SessionNumber", "SessionNum", "Number", "SessionCode"))
        stype = clean(pick(s, "SessionType", "Type", "Format", "GoodType"))
        is_poster = "poster" in (title + " " + stype).lower() or truthy(pick(s, "isPoster", "Poster"))
        slot = session_slot(cx, s)
        s_roles = roles_for(cx, s)
        chairs = "; ".join(
            f"{person_name(r)} ({affiliation(r)})" if affiliation(r) else person_name(r)
            for r in sorted(s_roles, key=position) if person_name(r)
        )
        hosts = clean(pick(s, "Hosts", "Host", "JointWith", "Sponsors", "HostedBy"))
        summary = clean(pick(s, "Description", "Summary", "Abstract", "Synopsis"))

        children = []
        for key in list_keys(s, "Paper") + list_keys(s, "SlotData"):
            if key.startswith("ChildList_"):
                children.extend(cx.get_children(s, key))
        papers, seen_p = [], set()
        for c in children:
            if is_slot(c):
                pid_ = slot_paper_id(c)
                if pid_:
                    slots_by_paper[pid_] = slot_fields(c)
                slots_by_session.setdefault(sid, []).append(slot_fields(c))
            elif isinstance(c, dict) and c.get("_url", "").startswith("Paper/") and c["_url"] not in seen_p:
                seen_p.add(c["_url"])
                papers.append(c)
        print(f"  [{i}/{len(raw_sessions)}] {number or sid}: {title[:70]}  ({len(papers)} papers)")

        sess = {
            "session_id": sid, "number": number, "title": title,
            "type": "Poster" if is_poster else (stype or "Oral"),
            "date": "", "start": "", "end": "", "room": "", "_slot": slot,
            "chairs": chairs, "hosts": hosts, "summary": summary,
            "n_papers": 0, "url": f"{cx.app}/{sid}" if sid else "",
        }

        for p in papers:
            if not isinstance(p, dict):
                continue
            if clean(pick(p, "GoodType")).lower() == "break":
                continue
            pid = p.get("_url", "")
            withdrawn = str(pick(p, "Withdrawn")).lower() in ("w", "1", "y", "yes", "true")
            p_roles = sorted(roles_for(cx, p), key=position)
            names = [person_name(r) for r in p_roles if person_name(r)]
            presenters = [person_name(r) for r in p_roles if is_presenter(r) and person_name(r)]
            affs = []
            for r in p_roles:
                a = affiliation(r)
                if a and a not in affs:
                    affs.append(a)
            presentations.append({
                "paper_id": pid,
                "paper_number": clean(pick(p, "FinalPaperNumber", "PaperNumber", "Number", "AbstractNumber")),
                "title": clean(pick(p, "Title", "label")),
                "presenter": "; ".join(presenters) or (names[0] if names else ""),
                "authors": "; ".join(names),
                "affiliations": "; ".join(affs),
                "abstract": clean(pick(p, "Abstract", "AbstractText", "Body", "Summary")),
                "session_id": sid, "session_number": number, "session_title": title,
                "session_type": sess["type"],
                "date": "", "time": "", "end": "", "room": "",
                "withdrawn": "Yes" if withdrawn else "",
                "url": f"{cx.app}/{pid}" if pid else "",
            })
            sess["n_papers"] += 1
            for order, r in enumerate(p_roles, 1):
                if person_name(r):
                    authors.append({
                        "author": person_name(r), "affiliation": affiliation(r),
                        "role": role_kind(r), "presenter": "Yes" if is_presenter(r) else "",
                        "order": order, "paper_number": presentations[-1]["paper_number"],
                        "paper_title": presentations[-1]["title"],
                        "session_number": number, "paper_url": presentations[-1]["url"],
                        "paper_id": pid,
                    })
        sessions.append(sess)

    talk = finalize_times(sessions, presentations, slots_by_paper, slots_by_session)
    sort_all(sessions, presentations, authors)
    return conf_title, sessions, presentations, authors, talk

# ----------------------------------------------------------------------------- inspect

def inspect(cx, program_id):
    program = cx.get(f"Program/{program_id}")
    def show(label, o):
        if isinstance(o, dict):
            print(f"\n== {label} ({o.get('_url')})")
            for k, v in o.items():
                v = json.dumps(v, ensure_ascii=False)
                print(f"   {k:32} {v[:110]}")
    show("Program", program)
    sessions = collect_sessions(cx, program)
    if sessions:
        s = sessions[0]; show("Session", s)
        for k in list_keys(s, "SlotData"):
            refs = s.get(k) or []
            if refs: show("SlotData", cx.get(refs[0])); break
        papers = []
        for k in list_keys(s, "Paper"):
            if k.startswith("ChildList_"): papers += cx.get_children(s, k)
        if papers:
            show("Paper", papers[0])
            rs = roles_for(cx, papers[0])
            if rs: show("Role", rs[0])

# ----------------------------------------------------------------------------- output

def write_xlsx(path, conf_title, sessions, presentations, authors, source, talk_minutes=15):
    import datetime as dt
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_font = Font(name="Arial", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F4E5F")
    body_font = Font(name="Arial", size=10)
    est_font = Font(name="Arial", size=10, italic=True, color="7F7F7F")
    link_font = Font(name="Arial", size=10, color="0563C1", underline="single")

    def as_date(v):
        try:
            return dt.date.fromisoformat(v) if v else None
        except ValueError:
            return v
    def as_time(v):
        try:
            h, m = map(int, v.split(":")[:2]); return dt.time(h, m)
        except (ValueError, AttributeError):
            return v or None

    # (header, key, width, kind) kind: date | time | time_est:<flagkey> | text
    def sheet(ws, cols, rows):
        ws.append([c[0] for c in cols])
        for c_idx, col in enumerate(cols, 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.font, cell.fill = head_font, head_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c_idx)].width = col[2]
        for r_idx, r in enumerate(rows, 2):
            for c_idx, (_, key, _, kind) in enumerate(cols, 1):
                v = r.get(key, "")
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = body_font
                cell.alignment = Alignment(vertical="top", wrap_text=kind == "text")
                if kind == "date":
                    cell.value = as_date(v); cell.number_format = "ddd d mmm yyyy"
                elif kind.startswith("time"):
                    cell.value = as_time(v); cell.number_format = "h:mm AM/PM"
                    if kind.startswith("time_est:") and r.get(kind.split(":", 1)[1]):
                        cell.font = est_font
                else:
                    cell.value = v
                    if isinstance(v, str) and v.startswith("https://"):
                        cell.hyperlink, cell.font = v, link_font
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    ws = wb.active
    ws.title = "Presentations"
    sheet(ws, [
        ("Date", "date", 15, "date"), ("Start", "time", 10, "time"), ("End", "end", 10, "time_est:end_estimated"),
        ("Session #", "session_number", 10, "text"), ("Session", "session_title", 36, "text"),
        ("Type", "session_type", 13, "text"), ("Session start", "session_start", 11, "time"),
        ("Room", "room", 14, "text"), ("Paper #", "paper_number", 10, "text"), ("Title", "title", 50, "text"),
        ("Presenter", "presenter", 22, "text"), ("Authors", "authors", 36, "text"),
        ("Affiliations", "affiliations", 30, "text"), ("Abstract", "abstract", 90, "text"),
        ("Withdrawn", "withdrawn", 10, "text"), ("Link", "url", 30, "text"),
    ], presentations)

    sheet(wb.create_sheet("Sessions"), [
        ("Date", "date", 15, "date"), ("Start", "start", 10, "time"), ("End", "end", 10, "time_est:end_estimated"),
        ("Session #", "number", 10, "text"), ("Title", "title", 50, "text"), ("Type", "type", 13, "text"),
        ("Room", "room", 14, "text"), ("Chairs", "chairs", 40, "text"), ("Hosts / joint with", "hosts", 30, "text"),
        ("# Presentations", "n_papers", 14, "text"), ("Description", "summary", 50, "text"), ("Link", "url", 30, "text"),
    ], sessions)

    sheet(wb.create_sheet("Authors"), [
        ("Author", "author", 26, "text"), ("Affiliation", "affiliation", 30, "text"), ("Role", "role", 12, "text"),
        ("Presenter", "presenter", 10, "text"), ("Author order", "order", 12, "text"),
        ("Date", "date", 15, "date"), ("Start", "time", 10, "time"),
        ("Paper #", "paper_number", 10, "text"), ("Paper title", "paper_title", 55, "text"),
        ("Session #", "session_number", 10, "text"), ("Link", "paper_url", 30, "text"),
    ], authors)

    about = wb.create_sheet("About")
    n_est = sum(1 for s in sessions if s.get("end_estimated"))
    for line in [
        ("Conference", conf_title), ("Source", source),
        ("Harvested", time.strftime("%Y-%m-%d %H:%M")),
        ("Sessions", len(sessions)), ("Presentations", len(presentations)), ("Author entries", len(authors)),
        ("Times", "Talk start times come from each paper's SlotData record in the Confex API. Session start = first talk."),
        ("End times", f"The API gives start times only, so ends are worked out: a talk ends when the next talk in its "
                      f"session starts. A session's end is set to the end of its time block (when the parallel "
                      f"sessions at that start time finish), {talk_minutes} min after the last talk starts. Ends that "
                      f"rest on that estimate are in grey italics ({n_est} sessions). Poster session ends are left blank."),
        ("Note", "Data via the public Confex Meeting API. Check the live program for last-minute changes."),
    ]:
        about.append(line)
    for row in about.iter_rows():
        row[0].font = Font(name="Arial", bold=True); row[1].font = body_font
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    about.column_dimensions["A"].width = 16; about.column_dimensions["B"].width = 100
    wb.save(path)

def write_outputs(out, conf_title, source, sessions, presentations, authors, talk):
    write_xlsx(out + ".xlsx", conf_title, sessions, presentations, authors, source, talk)
    with open(out + ".json", "w", encoding="utf-8") as f:
        json.dump({"conference": conf_title, "source": source,
                   "harvested": time.strftime("%Y-%m-%d %H:%M"), "talk_minutes": talk,
                   "sessions": sessions, "presentations": presentations, "authors": authors},
                  f, ensure_ascii=False)
    timed = sum(1 for p in presentations if p.get("time"))
    print(f"\nWrote {out}.xlsx and {out}.json: {len(sessions)} sessions, {len(presentations)} presentations "
          f"({timed} with times), {len(authors)} author entries.")

def slots_from_file(path):
    """Schedule records from an earlier output: SlotData pseudo-rows (first harvester version) or,
    failing that, each presentation's own date/time."""
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    by_paper, by_session_title = {}, {}
    for p in d.get("presentations", []):
        m = re.match(r"SlotData/Paper(\d+)", p.get("paper_id", ""))
        if m:
            rec = {"date": p.get("date", ""), "start": p.get("time", ""), "end": p.get("end", ""), "room": p.get("room", "")}
            by_paper[f"Paper/{m.group(1)}"] = rec
            by_session_title.setdefault(p.get("session_title", ""), []).append(rec)
    if not by_paper:
        for p in d.get("presentations", []):
            if p.get("paper_id", "").startswith("Paper/"):
                by_paper[p["paper_id"]] = {"date": p.get("date", ""), "start": p.get("time", ""),
                                           "end": p.get("end", ""), "room": p.get("room", "")}
    return by_paper, by_session_title

def repair(json_path, out, slots_from=()):
    """Rebuild outputs from an earlier 26AI_program.json (and .xlsx, for authors) without re-downloading.
    --slots-from files supply the schedule (SlotData) and take priority over the input's own times."""
    with open(json_path, encoding="utf-8") as f:
        d = json.load(f)
    sessions = d.get("sessions", [])
    title_to_sid = {s["title"]: s["session_id"] for s in sessions}
    for s in sessions:  # schedule is rebuilt from slot records below
        s["_slot"] = {}
    own, own_by_title = slots_from_file(json_path)
    slots_by_paper, slots_by_session, presentations = dict(own), {}, []
    for title, recs in own_by_title.items():
        slots_by_session.setdefault(title_to_sid.get(title, ""), []).extend(recs)
    for src in slots_from or ():
        by_paper, by_title = slots_from_file(src)
        for pid, rec in by_paper.items():
            base = slots_by_paper.get(pid, {})
            slots_by_paper[pid] = {k: rec.get(k) or base.get(k, "") for k in ("date", "start", "end", "room")}
        for title, recs in by_title.items():
            slots_by_session.setdefault(title_to_sid.get(title, ""), []).extend(recs)
        print(f"Schedule from {src}: {len(by_paper)} paper slots")
    for p in d.get("presentations", []):
        if p.get("paper_id", "").startswith("SlotData/") or p.get("title") == "Slot Data":
            continue
        p.setdefault("session_id", title_to_sid.get(p.get("session_title"), ""))
        for k in ("title", "presenter", "authors", "abstract"):
            p[k] = clean(p.get(k, ""))
        p["affiliations"] = "; ".join(x for x in (clean(y) for y in str(p.get("affiliations", "")).split(";")) if x)
        presentations.append(p)
    for s in sessions:
        s["n_papers"] = sum(1 for p in presentations if p["session_id"] == s["session_id"])
    authors = d.get("authors")
    xlsx_path = os.path.splitext(json_path)[0] + ".xlsx"
    if not authors and os.path.exists(xlsx_path):
        from openpyxl import load_workbook
        ws = load_workbook(xlsx_path, read_only=True)["Authors"]
        rows = list(ws.iter_rows(values_only=True))
        keymap = {"Author": "author", "Affiliation": "affiliation", "Role": "role", "Presenter": "presenter",
                  "Author order": "order", "Paper #": "paper_number", "Paper title": "paper_title",
                  "Session #": "session_number", "Link": "paper_url"}
        hdr = [keymap.get(h, h) for h in rows[0]]
        authors = [{k: ("" if v is None else v) for k, v in zip(hdr, r)} for r in rows[1:]]
        for a in authors:
            m = re.search(r"Paper/\d+", str(a.get("paper_url", "")))
            a["paper_id"] = m.group(0) if m else ""
    authors = authors or []
    for au in authors:
        au["affiliation"] = clean(au.get("affiliation", ""))
    talk = finalize_times(sessions, presentations, slots_by_paper, slots_by_session)
    sort_all(sessions, presentations, authors)
    write_outputs(out, d.get("conference", ""), d.get("source", ""), sessions, presentations, authors, talk)

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--client", default="ams")
    ap.add_argument("--meeting", default="107ANNUAL")
    ap.add_argument("--program", default="1955")
    ap.add_argument("--out", default="26AI_program")
    ap.add_argument("--cache", default="confex_cache")
    ap.add_argument("--refresh", action="store_true", help="ignore cached responses")
    ap.add_argument("--inspect", action="store_true", help="print field names and exit")
    ap.add_argument("--repair", metavar="OLD_JSON", help="rebuild outputs from an earlier 26AI_program.json")
    ap.add_argument("--slots-from", metavar="JSON", action="append",
                    help="with --repair: earlier output whose SlotData supplies dates/times (repeatable)")
    ap.add_argument("--api-base", help=argparse.SUPPRESS)  # for testing against a mock server
    a = ap.parse_args()

    if a.repair:
        repair(a.repair, a.out, a.slots_from); return
    cx = Confex(a.client, a.meeting, a.cache, refresh=a.refresh)
    if a.api_base:
        cx.api = a.api_base.rstrip("/")
    if a.inspect:
        inspect(cx, a.program); return

    conf_title, sessions, presentations, authors, talk = harvest(cx, a.program)
    source = f"{cx.app}/Program/{a.program}"
    write_outputs(a.out, conf_title, source, sessions, presentations, authors, talk)
    if not presentations:
        print("No presentations found - run with --inspect and check the field names.")
    elif not any(p.get("room") for p in presentations):
        print("Rooms came back empty - run with --inspect and look for the room field name on SlotData.")

if __name__ == "__main__":
    main()
