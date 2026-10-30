#!/usr/bin/env python3
"""
Find and harvest AI / machine-learning sessions from a Confex-hosted meeting (default: AGU26)
via the Confex Meeting API.

Step 1 - scan (fast): walk every section (Program) -> session, match session titles against
         AI/ML keyword rules, and write a review sheet:
             python confex_ai_sessions_harvest.py --scan
         Open AGU26_AI_sessions_review.xlsx, set include = N for any session you don't want.
Step 2 - harvest: download papers, schedule slots and authors for the included sessions:
             python confex_ai_sessions_harvest.py --harvest
         Writes AGU26_AI_program.xlsx / .json (load the .json into the AI Program Explorer page).

Options:
  --meeting-url URL      meeting app URL (default: https://studio.m-anage.com/agu/agu26/meetingapp.cgi)
  --scan-papers          also find AI/ML-titled papers inside sessions whose titles didn't match
  --scan-descriptions    also match on session descriptions (broader; tagged tier = "description")
  --core-only            skip sessions that only matched "related" rules (data-driven, digital twin, ...)
  --keywords FILE        JSON list of [label, regex, tier, case_sensitive] to replace the built-in rules
  --workers N            parallel requests (default 4; please be polite to the server)
  --inspect              print the field names of each object type, then exit
  --refresh              ignore the cache and re-download

Everything the API returns is cached in ./confex_cache_<meeting>/ so re-runs are fast and resumable.
Requires: pip install requests openpyxl
"""
import argparse
import concurrent.futures as cf
import html
import json
import os
import re
import sys
import threading
import time

try:
    import requests
except ImportError:
    sys.exit("Please `pip install requests openpyxl` first.")

# ----------------------------------------------------------------------------- keyword rules
# (label, regex, tier, case_sensitive). Acronym rules are case-sensitive so "ML" doesn't match "html".
KEYWORD_RULES = [
    ["artificial intelligence", r"(?<![A-Za-z])AIs?(?![A-Za-z])|GeoAI|\bAI4\w*|\bXAI\b", "core", True],
    ["artificial intelligence", r"artificial intelligence|computational intelligence", "core", False],
    ["machine learning", r"(?<![A-Za-z])MLs?(?![A-Za-z])|SciML|TinyML|GeoML|MLWP", "core", True],
    ["machine learning", r"machine[- ]learn|learning machine", "core", False],
    ["deep learning", r"deep[- ]learn|deep neural|deep generative", "core", False],
    ["neural networks", r"neural (network|net\b|operator|ode|field|emulat)|convolutional|recurrent neural|graph neural"
                        r"|transformer|u-?net\b|autoencoder|long short-term memory", "core", False],
    ["neural networks", r"\b(CNN|LSTM|GNN|RNN|PINN|FNO|GAN|VAE)s?\b", "core", True],
    ["emulators & surrogates", r"emulat|surrogate model|surrogate[- ]based|surrogates? for", "core", False],
    ["hybrid & physics-informed", r"hybrid (physics|model|modeling|modelling|machine|ml|ai|data|process|approach|framework|learning)"
                                  r"|physics[- ](informed|guided|aware|constrained|embedded)|differentiable|knowledge[- ]guided"
                                  r"|theory[- ]guided|scientific machine learning|grey[- ]box|gray[- ]box|neural[- ]physical", "core", False],
    ["explainable AI", r"explainab|interpretab|trustworthy (ai|ml|machine|model)|responsible ai|ethical ai", "core", False],
    ["generative / foundation / LLM", r"generative|diffusion model|foundation model|large language|language model|chatbot"
                                      r"|agentic|ai agent|vision[- ]language|multimodal model|world model", "core", False],
    ["generative / foundation / LLM", r"\b(LLM|GPT|LVM|VLM)s?\b", "core", True],
    ["learning methods", r"(reinforcement|unsupervised|self[- ]supervised|semi[- ]supervised|supervised|representation|transfer"
                         r"|federated|active|statistical|contrastive|few[- ]shot|zero[- ]shot|manifold) learning|random forest"
                         r"|gradient[- ]boost|xgboost|support vector|gaussian process regression|symbolic regression"
                         r"|equation discovery|computer vision|image (segmentation|recognition)|semantic segmentation"
                         r"|object detection|pattern recognition|clustering algorithm|self[- ]organizing map|kernel method", "core", False],
    ["data-driven & digital twins", r"data[- ]driven|data science|big data|digital twin|ai[- ]ready|ml[- ]ready"
                                    r"|automat\w* (\w+ )?(detection|classification|picking|mapping|identification|segmentation|tracking)"
                                    r"|phase picking|reduced[- ]order model|intelligent (system|monitor|sens)"
                                    r"|causal (discovery|inference|learning)|smart (sensor|monitor)", "related", False],
]

def compile_rules(rules):
    return [(lab, re.compile(rx, 0 if cs else re.I), tier) for lab, rx, tier, cs in rules]

def match_terms(text, rules):
    """Return (sorted labels, tier, matched snippets) for a piece of text."""
    labels, snippets, tiers = set(), set(), set()
    for lab, rx, tier in rules:
        for m in rx.finditer(text or ""):
            labels.add(lab); tiers.add(tier); snippets.add(m.group(0).strip().lower())
    tier = "core" if "core" in tiers else ("related" if tiers else "")
    return sorted(labels), tier, sorted(snippets)

# ----------------------------------------------------------------------------- HTTP

class Confex:
    def __init__(self, meeting_url, cache_dir, refresh=False, delay=0.15):
        m = re.match(r"(https?://[^/]+/[^/]+/[^/]+)", meeting_url.strip())
        if not m:
            sys.exit(f"Can't parse meeting URL: {meeting_url}")
        self.base = m.group(1)
        self.api = self.base + "/meetingapi.cgi"
        self.app = self.base + "/meetingapp.cgi"
        self.code = self.base.rstrip("/").split("/")[-1]
        self.cache_dir, self.refresh, self.delay = cache_dir, refresh, delay
        self._local = threading.local()
        self.requests_made = 0
        os.makedirs(cache_dir, exist_ok=True)

    def _session(self):
        if not hasattr(self._local, "s"):
            s = requests.Session()
            s.headers.update({"Content-Type": "application/json", "Accept": "application/json",
                              "User-Agent": "confex-ai-sessions-harvester (personal research use)"})
            self._local.s = s
        return self._local.s

    def _cache_path(self, key):
        return os.path.join(self.cache_dir, re.sub(r"[^A-Za-z0-9_.-]+", "_", key)[:180] + ".json")

    def get(self, path, config=None):
        key = path + ("__" + json.dumps(config, sort_keys=True) if config else "")
        cp = self._cache_path(key)
        if not self.refresh and os.path.exists(cp):
            with open(cp, encoding="utf-8") as f:
                return json.load(f)
        headers = {"X-Confex-MeetingApi-Config": json.dumps(config)} if config else {}
        for attempt in range(5):
            try:
                r = self._session().get(f"{self.api}/{path}", headers=headers, timeout=90)
                self.requests_made += 1
                if r.status_code == 404:
                    return None
                if r.status_code in (429, 500, 502, 503, 504):
                    raise requests.HTTPError(f"HTTP {r.status_code}")
                r.raise_for_status()
                data = r.json()
                tmp = cp + f".{threading.get_ident()}.tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False)
                os.replace(tmp, cp)
                time.sleep(self.delay)
                return data
            except (requests.RequestException, ValueError) as e:
                wait = 2 ** attempt
                print(f"  ! {path}: {e} - retrying in {wait}s", file=sys.stderr)
                time.sleep(wait)
        print(f"  ! giving up on {path}", file=sys.stderr)
        return None

    def get_children(self, obj, listname):
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

    def list_all(self, model, fields=None, size=500):
        """Paged listing of every object of a type (Confex 'size'/'page' config)."""
        out, page = [], 1
        while True:
            cfg = {"size": size, "page": page}
            if fields:
                cfg["Fields"] = fields
            batch = self.get(model, cfg)
            if not isinstance(batch, list) or not batch:
                break
            out.extend(b for b in batch if isinstance(b, dict))
            if len(batch) < size or page > 400:
                break
            page += 1
        return out

# ----------------------------------------------------------------------------- shared helpers
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
    aff = clean(pick(role, "Affiliation", "Person_Affiliation", "Institution", "Organization", "Company",
                     "Address_Affiliation", "Person_Institution"))
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
        if "poster" in (s.get("type") or "").lower() or "lightning" in (s.get("type") or "").lower():
            continue  # poster halls are legitimately shared by parallel poster sessions
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


# ----------------------------------------------------------------------------- discovery

def meeting_info(cx):
    home = cx.get("Home/0") or {}
    return {
        "title": clean(pick(home, "MeetingTitle", "Meeting_MeetingTitle", "Title", "label")) or cx.code,
        "start": norm_date(pick(home, "Meeting_MeetingStart", "MeetingStart", "StartDate")),
        "end": norm_date(pick(home, "Meeting_MeetingEnd", "MeetingEnd", "EndDate")),
        "home": home,
    }

def list_programs(cx, home):
    """All Programs (AGU sections), via the ModuleProgramBook index, else a paged listing."""
    refs = []
    for k, v in home.items():
        if k.startswith("ChildList_") and isinstance(v, list):
            refs += [x for x in v if isinstance(x, str)]
    for ref in [r for r in refs if r.startswith("ModuleProgramBook")] or ["ModuleProgramBook/0"]:
        book = cx.get(ref)
        if isinstance(book, dict) and book.get("ChildList_Program"):
            return cx.get_children(book, "ChildList_Program")
    return [p for p in cx.list_all("Program") if isinstance(p, dict)]

def sessions_in(cx, container, depth=0):
    """Sessions under a Program or Symposium (recursing into Symposia)."""
    out = []
    for key in list_keys(container, "Session"):
        if key.startswith("ChildList_"):
            out += [s for s in cx.get_children(container, key) if str(s.get("_url", "")).startswith("Session/")]
    if depth < 2:
        for key in list_keys(container, "Symposium"):
            if key.startswith("ChildList_"):
                for sym in cx.get_children(container, key):
                    out += sessions_in(cx, sym, depth + 1)
    return out

def all_sessions(cx, workers):
    info = meeting_info(cx)
    programs = list_programs(cx, info["home"])
    print(f"{info['title']}: {len(programs)} sections/programs")
    sessions, sections = {}, {}
    def per_program(p):
        return p, sessions_in(cx, p)
    with cf.ThreadPoolExecutor(workers) as ex:
        for p, ss in ex.map(per_program, programs):
            name = clean(pick(p, "Title", "Name", "ProgramName", "label"))
            for s in ss:
                sessions.setdefault(s["_url"], s)
                sections.setdefault(s["_url"], [])
                if name and name not in sections[s["_url"]]:
                    sections[s["_url"]].append(name)
    if not sessions:  # meetings without a program book
        for s in cx.list_all("Session"):
            if str(s.get("_url", "")).startswith("Session/"):
                sessions[s["_url"]] = s
                sections[s["_url"]] = []
    print(f"{len(sessions)} sessions in total")
    return info, list(sessions.values()), sections

def session_kind(s):
    title = clean(pick(s, "Title", "SessionTitle", "label"))
    stype = clean(pick(s, "SessionType", "Type", "Format", "GoodType", "PresentationType"))
    blob = (title + " " + stype).lower()
    if "lightning" in blob:
        return "eLightning"
    if "poster" in blob:
        return "Poster"
    if any(w in blob for w in ("break", "reception", "luncheon", "exhibit")) and "session" not in blob:
        return stype or "Event"
    return stype or "Oral"

def session_row(cx, s, sections):
    return {
        "session_id": s["_url"],
        "number": clean(pick(s, "FinalSessionNumber", "SessionNumber", "SessionNum", "Number", "SessionCode", "Index")),
        "title": clean(pick(s, "Title", "SessionTitle", "Name", "label")),
        "section": "; ".join(sections.get(s["_url"], [])),
        "type": session_kind(s),
        "n_papers": len([r for r in (s.get("ChildList_Paper") or []) if not str(r).startswith("SlotData")]),
        "description": clean(pick(s, "Description", "Summary", "Abstract", "Synopsis")),
        "url": f"{cx.app}/{s['_url']}",
    }

# ----------------------------------------------------------------------------- scan

def scan(cx, args, rules):
    info, sessions, sections = all_sessions(cx, args.workers)
    rows = []
    for s in sessions:
        r = session_row(cx, s, sections)
        labels, tier, snips = match_terms(r["title"], rules)
        source = "session title" if labels else ""
        if not labels and args.scan_descriptions:
            labels, tier, snips = match_terms(r["description"], rules)
            if labels:
                tier, source = "description", "session description"
        if labels:
            r.update(matched_rules="; ".join(labels), matched_terms="; ".join(snips), tier=tier, source=source,
                     include="N" if (args.core_only and tier != "core") else "Y", paper_ids="")
            rows.append(r)

    if args.scan_papers:
        print("Scanning paper titles in the remaining sessions ...")
        matched = {r["session_id"] for r in rows}
        rest = [s for s in sessions if s["_url"] not in matched]
        def per_session(s):
            hits = []
            for c in children_of(cx, s):
                if is_slot(c) or not str(c.get("_url", "")).startswith("Paper/"):
                    continue
                labels, tier, snips = match_terms(clean(pick(c, "Title", "label")), rules)
                if labels and tier == "core":
                    hits.append((c["_url"], labels, snips))
            return s, hits
        with cf.ThreadPoolExecutor(args.workers) as ex:
            for i, (s, hits) in enumerate(ex.map(per_session, rest), 1):
                if i % 200 == 0:
                    print(f"  {i}/{len(rest)} sessions checked")
                if hits:
                    r = session_row(cx, s, sections)
                    r.update(matched_rules="; ".join(sorted({l for _, ls, _ in hits for l in ls})),
                             matched_terms="; ".join(sorted({t for *_, ts in hits for t in ts})),
                             tier="paper titles", source=f"{len(hits)} paper title(s)", include="Y",
                             paper_ids=" ".join(h[0] for h in hits))
                    rows.append(r)

    rows.sort(key=lambda r: ({"core": 0, "paper titles": 1, "related": 2, "description": 3}.get(r["tier"], 4),
                             r["section"], r["number"], r["title"]))
    out = args.out + "_sessions_review"
    write_review(out + ".xlsx", info, rows, len(sessions))
    with open(out + ".json", "w", encoding="utf-8") as f:
        json.dump({"meeting": {k: v for k, v in info.items() if k != "home"}, "total_sessions": len(sessions),
                   "sessions": rows}, f, ensure_ascii=False, indent=1)
    by = {}
    for r in rows:
        by[r["tier"]] = by.get(r["tier"], 0) + 1
    print(f"\nMatched {len(rows)} of {len(sessions)} sessions: " + ", ".join(f"{v} {k}" for k, v in by.items()))
    print(f"Est. presentations: {sum(r['n_papers'] for r in rows if r['tier'] != 'paper titles')} in matched sessions"
          + (f" + {sum(len(r['paper_ids'].split()) for r in rows if r['tier'] == 'paper titles')} individually matched papers"
             if args.scan_papers else ""))
    print(f"Review: {out}.xlsx  (set include = N to drop a session), then run with --harvest")
    print(f"API requests this run: {cx.requests_made} (the rest came from cache)")

def write_review(path, info, rows, total):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook(); ws = wb.active; ws.title = "Sessions"
    cols = [("include", "include", 9), ("Tier", "tier", 13), ("Matched rules", "matched_rules", 30),
            ("Matched terms", "matched_terms", 26), ("Session #", "number", 11), ("Title", "title", 60),
            ("Section", "section", 28), ("Type", "type", 10), ("# Papers", "n_papers", 9),
            ("Source", "source", 18), ("Link", "url", 30), ("session_id", "session_id", 15), ("paper_ids", "paper_ids", 20)]
    ws.append([c[0] for c in cols])
    for i, (_, _, w) in enumerate(cols, 1):
        c = ws.cell(row=1, column=i)
        c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E5F")
        ws.column_dimensions[get_column_letter(i)].width = w
    fills = {"core": None, "paper titles": PatternFill("solid", fgColor="EAF3F8"),
             "related": PatternFill("solid", fgColor="FFF6E0"), "description": PatternFill("solid", fgColor="F3F3F3")}
    for r in rows:
        ws.append([r.get(k, "") for _, k, _ in cols])
        for c in ws[ws.max_row]:
            c.font = Font(name="Arial", size=10); c.alignment = Alignment(vertical="top", wrap_text=True)
            if fills.get(r["tier"]):
                c.fill = fills[r["tier"]]
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    ws.add_data_validation(dv); dv.add(f"A2:A{max(2, ws.max_row)}")
    ws.freeze_panes = "B2"; ws.auto_filter.ref = ws.dimensions
    about = wb.create_sheet("About")
    for line in [("Meeting", info["title"]), ("Dates", f"{info['start']} to {info['end']}"),
                 ("Sessions scanned", total), ("Sessions matched", len(rows)),
                 ("How to use", "Set include = N for sessions you don't want, save, then run with --harvest. "
                                "Tier 'core' = explicit AI/ML terms; 'related' = data-driven, digital twin, automated "
                                "detection, causal inference; 'paper titles' = only some papers matched, and only those "
                                "papers are harvested; 'description' = matched on the session description only."),
                 ("Rules", "; ".join(sorted({r[0] for r in KEYWORD_RULES})))]:
        about.append(line)
    about.column_dimensions["A"].width = 18; about.column_dimensions["B"].width = 110
    for row in about.iter_rows():
        row[0].font = Font(name="Arial", bold=True); row[1].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)

def read_review(path):
    """Included rows from the (possibly hand-edited) review sheet."""
    if path.endswith(".xlsx"):
        from openpyxl import load_workbook
        ws = load_workbook(path, read_only=True)["Sessions"]
        rows = list(ws.iter_rows(values_only=True))
        hdr = [str(h) for h in rows[0]]
        recs = [dict(zip(hdr, r)) for r in rows[1:]]
        get = lambda r, k: "" if r.get(k) is None else str(r.get(k))
        return [{"session_id": get(r, "session_id"), "include": get(r, "include").strip().upper() or "Y",
                 "tier": get(r, "Tier"), "matched_rules": get(r, "Matched rules"), "matched_terms": get(r, "Matched terms"),
                 "source": get(r, "Source"), "paper_ids": get(r, "paper_ids")} for r in recs if get(r, "session_id")]
    with open(path, encoding="utf-8") as f:
        return json.load(f)["sessions"]

# ----------------------------------------------------------------------------- harvest

def children_of(cx, s):
    kids = []
    for key in list_keys(s, "Paper") + list_keys(s, "SlotData"):
        if key.startswith("ChildList_"):
            kids += cx.get_children(s, key)
    return kids

def index_terms(p):
    vals = []
    for k, v in p.items():
        if re.search(r"keyword|indexterm|index_term|topic", k, re.I) and not k.startswith(("_", "ChildList_", "ParentList_")):
            if isinstance(v, list):
                vals += [clean(x) for x in v if isinstance(x, str)]
            elif isinstance(v, str):
                vals.append(clean(v))
    return "; ".join(x for x in vals if x)

def orcid(role):
    v = clean(pick(role, "ORCID", "Person_ORCIDiD", "ORCIDiD", "Person_ORCID", "OrcidID"))
    m = re.search(r"\d{4}-\d{4}-\d{4}-\d{3}[\dX]", v)
    return m.group(0) if m else ""

def harvest_session(cx, s, sections, review, rules):
    sid = s["_url"]
    row = session_row(cx, s, sections)
    only = set((review.get("paper_ids") or "").split())
    slot = session_slot(cx, s)
    chairs = []
    for r in sorted(roles_for(cx, s), key=position):
        if person_name(r):
            chairs.append(f"{person_name(r)} ({affiliation(r)})" if affiliation(r) else person_name(r))
    sess = {**{k: row[k] for k in ("session_id", "number", "title", "section", "type", "url")},
            "date": "", "start": "", "end": "", "room": "", "_slot": slot, "chairs": "; ".join(chairs),
            "summary": row["description"], "n_papers": 0,
            "tier": review.get("tier", ""), "matched_rules": review.get("matched_rules", ""),
            "matched_terms": review.get("matched_terms", ""), "source": review.get("source", "")}
    papers, slots, seen = [], {}, set()
    session_slots = []
    for c in children_of(cx, s):
        if is_slot(c):
            pid = slot_paper_id(c)
            if pid:
                slots[pid] = slot_fields(c)
            session_slots.append(slot_fields(c))
        elif str(c.get("_url", "")).startswith("Paper/") and c["_url"] not in seen:
            if only and c["_url"] not in only:
                continue
            if clean(pick(c, "GoodType")).lower() == "break":
                continue
            seen.add(c["_url"]); papers.append(c)
    pres, authors = [], []
    for p in papers:
        pid = p["_url"]
        roles = sorted(roles_for(cx, p), key=position)
        names = [person_name(r) for r in roles if person_name(r)]
        presenters = [person_name(r) for r in roles if is_presenter(r) and person_name(r)]
        affs = []
        for r in roles:
            a = affiliation(r)
            if a and a not in affs:
                affs.append(a)
        title = clean(pick(p, "Title", "label"))
        abstract = clean(pick(p, "Abstract", "AbstractText", "Body", "Summary"))
        labels, _, _ = match_terms(title + " " + abstract, rules)
        pres.append({
            "paper_id": pid,
            "paper_number": clean(pick(p, "FinalPaperNumber", "PaperNumber", "Number", "AbstractNumber")),
            "title": title, "presenter": "; ".join(presenters) or (names[0] if names else ""),
            "authors": "; ".join(names), "affiliations": "; ".join(affs),
            "author_affiliations": " | ".join(f"{person_name(r)}: {affiliation(r)}" for r in roles if person_name(r)),
            "orcids": "; ".join(f"{person_name(r)}: {orcid(r)}" for r in roles if orcid(r)),
            "abstract": abstract, "index_terms": index_terms(p),
            "invited": "Yes" if re.search(r"\(invited\)", title, re.I) or truthy(pick(p, "Invited", "IsInvited")) else "",
            "ai_terms_in_paper": "; ".join(labels),
            "session_id": sid, "session_number": row["number"], "session_title": row["title"],
            "session_type": row["type"], "section": row["section"],
            "date": "", "time": "", "end": "", "room": "",
            "withdrawn": "Yes" if str(pick(p, "Withdrawn")).lower() in ("w", "1", "y", "yes", "true") else "",
            "url": f"{cx.app}/{pid}",
        })
        for order, r in enumerate(roles, 1):
            if person_name(r):
                authors.append({"author": person_name(r), "affiliation": affiliation(r), "orcid": orcid(r),
                                "role": role_kind(r), "presenter": "Yes" if is_presenter(r) else "", "order": order,
                                "paper_number": pres[-1]["paper_number"], "paper_title": title,
                                "session_number": row["number"], "paper_url": pres[-1]["url"], "paper_id": pid})
    sess["n_papers"] = len(pres)
    return sess, pres, authors, slots, session_slots

def harvest(cx, args, rules):
    review_path = args.review or (args.out + "_sessions_review.xlsx")
    if not os.path.exists(review_path):
        sys.exit(f"No review file at {review_path}. Run with --scan first.")
    review = {r["session_id"]: r for r in read_review(review_path) if r.get("include", "Y").upper() != "N"}
    print(f"Harvesting {len(review)} sessions listed in {review_path}")
    info, sessions_all, sections = all_sessions(cx, args.workers)
    todo = [s for s in sessions_all if s["_url"] in review]
    missing = set(review) - {s["_url"] for s in todo}
    for sid in missing:  # sessions not reachable from the program book: fetch directly
        s = cx.get(sid)
        if isinstance(s, dict):
            todo.append(s)
    sessions, presentations, authors, slots_by_paper, slots_by_session = [], [], [], {}, {}
    done = 0
    with cf.ThreadPoolExecutor(args.workers) as ex:
        futs = {ex.submit(harvest_session, cx, s, sections, review[s["_url"]], rules): s for s in todo}
        for fut in cf.as_completed(futs):
            sess, pres, auth, slots, sslots = fut.result()
            sessions.append(sess); presentations += pres; authors += auth
            slots_by_paper.update(slots); slots_by_session[sess["session_id"]] = sslots
            done += 1
            if done % 10 == 0 or done == len(todo):
                print(f"  {done}/{len(todo)} sessions, {len(presentations)} presentations "
                      f"({cx.requests_made} API requests so far)")
    talk = finalize_times(sessions, presentations, slots_by_paper, slots_by_session)
    # Sessions harvested for a few AI-titled papers only: take the session start from all of its slots.
    for sess in sessions:
        if sess.get("tier") == "paper titles":
            sl = [x for x in slots_by_session.get(sess["session_id"], []) if valid_time(x.get("start")) and x.get("date")]
            if sl:
                first = min(sl, key=lambda x: (x["date"], x["start"]))
                if (first["date"], first["start"]) < (sess["date"] or "9999", sess["start"] or "99:99"):
                    sess["date"], sess["start"] = first["date"], first["start"]
                    for p in presentations:
                        if p["session_id"] == sess["session_id"]:
                            p["session_start"] = first["start"]
    sort_outputs(sessions, presentations, authors)
    write_outputs(args.out + "_program", info, sessions, presentations, authors, talk)

def sort_outputs(sessions, presentations, authors):
    sessions.sort(key=lambda s: (s["date"] or "9999", to_minutes(s["start"]), s["number"], s["title"]))
    order = {s["session_id"]: k for k, s in enumerate(sessions)}
    presentations.sort(key=lambda p: (order.get(p["session_id"], 9e9), p["date"] or "9999", to_minutes(p["time"]),
                                      p["paper_number"], p["title"]))
    when = {p["paper_id"]: p for p in presentations}
    for a in authors:
        p = when.get(a["paper_id"])
        if p:
            a["date"], a["time"] = p["date"], p["time"]
    authors.sort(key=lambda a: (a["author"].split()[-1].lower(), a["author"]))

# ----------------------------------------------------------------------------- output

def write_outputs(out, info, sessions, presentations, authors, talk):
    import datetime as dt
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    hf, hfill = Font(name="Arial", bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E5F")
    bf, ef = Font(name="Arial", size=10), Font(name="Arial", size=10, italic=True, color="7F7F7F")
    lf = Font(name="Arial", size=10, color="0563C1", underline="single")
    def as_date(v):
        try: return dt.date.fromisoformat(v) if v else None
        except ValueError: return v
    def as_time(v):
        try: h, m = map(int, v.split(":")[:2]); return dt.time(h, m)
        except (ValueError, AttributeError): return v or None
    def sheet(ws, cols, rows):
        ws.append([c[0] for c in cols])
        for i, c in enumerate(cols, 1):
            cell = ws.cell(row=1, column=i); cell.font, cell.fill = hf, hfill
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(i)].width = c[2]
        for ri, r in enumerate(rows, 2):
            for ci, (_, key, _, kind) in enumerate(cols, 1):
                v, cell = r.get(key, ""), ws.cell(row=ri, column=ci)
                cell.font = bf; cell.alignment = Alignment(vertical="top", wrap_text=kind == "text")
                if kind == "date":
                    cell.value = as_date(v); cell.number_format = "ddd d mmm yyyy"
                elif kind.startswith("time"):
                    cell.value = as_time(v); cell.number_format = "h:mm AM/PM"
                    if ":" in kind and r.get(kind.split(":", 1)[1]): cell.font = ef
                else:
                    cell.value = v
                    if isinstance(v, str) and v.startswith("http"): cell.hyperlink, cell.font = v, lf
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    ws = wb.active; ws.title = "Presentations"
    sheet(ws, [("Date", "date", 15, "date"), ("Start", "time", 10, "time"), ("End", "end", 10, "time:end_estimated"),
               ("Session #", "session_number", 11, "text"), ("Session", "session_title", 36, "text"),
               ("Section", "section", 22, "text"), ("Type", "session_type", 10, "text"), ("Room", "room", 16, "text"),
               ("Paper #", "paper_number", 13, "text"), ("Title", "title", 50, "text"),
               ("Presenter", "presenter", 22, "text"), ("Authors", "authors", 36, "text"),
               ("Affiliations", "affiliations", 34, "text"), ("Author: affiliation", "author_affiliations", 40, "text"),
               ("ORCIDs", "orcids", 24, "text"), ("Abstract", "abstract", 90, "text"),
               ("Index terms", "index_terms", 26, "text"), ("Invited", "invited", 8, "text"),
               ("AI terms in paper", "ai_terms_in_paper", 26, "text"), ("Withdrawn", "withdrawn", 10, "text"),
               ("Link", "url", 30, "text")], presentations)
    sheet(wb.create_sheet("Sessions"), [
        ("Date", "date", 15, "date"), ("Start", "start", 10, "time"), ("End", "end", 10, "time:end_estimated"),
        ("Session #", "number", 11, "text"), ("Title", "title", 50, "text"), ("Section", "section", 24, "text"),
        ("Type", "type", 10, "text"), ("Room", "room", 16, "text"), ("Chairs", "chairs", 40, "text"),
        ("# Presentations", "n_papers", 13, "text"), ("Tier", "tier", 12, "text"),
        ("Matched rules", "matched_rules", 28, "text"), ("Matched terms", "matched_terms", 24, "text"),
        ("Description", "summary", 60, "text"), ("Link", "url", 30, "text")], sessions)
    sheet(wb.create_sheet("Authors"), [
        ("Author", "author", 26, "text"), ("Affiliation", "affiliation", 34, "text"), ("ORCID", "orcid", 20, "text"),
        ("Role", "role", 12, "text"), ("Presenter", "presenter", 10, "text"), ("Author order", "order", 11, "text"),
        ("Date", "date", 15, "date"), ("Start", "time", 10, "time"), ("Paper #", "paper_number", 13, "text"),
        ("Paper title", "paper_title", 55, "text"), ("Session #", "session_number", 11, "text"),
        ("Link", "paper_url", 30, "text")], authors)
    about = wb.create_sheet("About")
    n_est = sum(1 for s in sessions if s.get("end_estimated"))
    for line in [("Meeting", info["title"]), ("Dates", f"{info['start']} to {info['end']}"),
                 ("Harvested", time.strftime("%Y-%m-%d %H:%M")), ("Sessions", len(sessions)),
                 ("Presentations", len(presentations)), ("Author entries", len(authors)),
                 ("Times", "Talk times come from each paper's SlotData record. A session starts with its first talk."),
                 ("End times", f"The API gives start times only. A talk ends when the next one starts; an oral session ends "
                               f"at its time block's end ({talk} min after its last talk). Inferred ends are grey italics "
                               f"({n_est} sessions). Poster ends are blank unless the API gives them."),
                 ("Note", "Data via the public Confex Meeting API. Check the live program for last-minute changes.")]:
        about.append(line)
    about.column_dimensions["A"].width = 16; about.column_dimensions["B"].width = 100
    for row in about.iter_rows():
        row[0].font = Font(name="Arial", bold=True); row[1].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(out + ".xlsx")
    with open(out + ".json", "w", encoding="utf-8") as f:
        json.dump({"conference": info["title"], "meeting_start": info["start"], "meeting_end": info["end"],
                   "source": info.get("source", ""), "harvested": time.strftime("%Y-%m-%d %H:%M"),
                   "talk_minutes": talk, "sessions": sessions, "presentations": presentations, "authors": authors},
                  f, ensure_ascii=False)
    timed = sum(1 for p in presentations if p["time"])
    print(f"\nWrote {out}.xlsx and {out}.json: {len(sessions)} sessions, {len(presentations)} presentations "
          f"({timed} with times), {len(authors)} author entries.")
    if presentations and not any(p["room"] for p in presentations):
        print("Rooms came back empty - run --inspect and look for the room field on SlotData.")
    if presentations and not any(p["affiliations"] for p in presentations):
        print("Affiliations came back empty - run --inspect and check the Role fields.")

# ----------------------------------------------------------------------------- inspect

def inspect(cx):
    def show(label, o):
        if isinstance(o, dict):
            print(f"\n== {label} ({o.get('_url')})")
            for k, v in o.items():
                print(f"   {k:34} {json.dumps(v, ensure_ascii=False)[:110]}")
    info = meeting_info(cx); show("Home", info["home"])
    programs = list_programs(cx, info["home"])
    if not programs:
        return print("No programs found.")
    show("Program", programs[0])
    for p in programs:
        ss = sessions_in(cx, p)
        if ss:
            show("Session", ss[0])
            kids = children_of(cx, ss[0])
            sl = next((c for c in kids if is_slot(c)), None); show("SlotData", sl)
            pa = next((c for c in kids if str(c.get("_url", "")).startswith("Paper/")), None); show("Paper", pa)
            if pa:
                rs = roles_for(cx, pa)
                if rs: show("Role", rs[0])
            break

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--meeting-url", default="https://studio.m-anage.com/agu/agu26/meetingapp.cgi")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--scan", action="store_true", help="find AI/ML sessions and write the review sheet")
    g.add_argument("--harvest", action="store_true", help="download the sessions marked include = Y")
    g.add_argument("--inspect", action="store_true", help="print field names of each object type")
    ap.add_argument("--scan-papers", action="store_true", help="also match paper titles in non-matching sessions")
    ap.add_argument("--scan-descriptions", action="store_true", help="also match session descriptions")
    ap.add_argument("--core-only", action="store_true", help="pre-set related-only sessions to include = N")
    ap.add_argument("--keywords", help="JSON file of [label, regex, tier, case_sensitive] rules")
    ap.add_argument("--review", help="review sheet to harvest from (default: <out>_sessions_review.xlsx)")
    ap.add_argument("--out", help="output file prefix (default: e.g. AGU26_AI)")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--refresh", action="store_true")
    ap.add_argument("--api-base", help=argparse.SUPPRESS)
    a = ap.parse_args()
    cx = Confex(a.meeting_url, f"confex_cache_{re.sub(r'[^A-Za-z0-9]+', '', a.meeting_url.split('/')[-2])}",
                refresh=a.refresh)
    if a.api_base:
        cx.api = a.api_base.rstrip("/")
    a.out = a.out or f"{cx.code.upper()}_AI"
    rules = compile_rules(json.load(open(a.keywords)) if a.keywords else KEYWORD_RULES)
    if a.inspect:
        inspect(cx)
    elif a.scan:
        scan(cx, a, rules)
    else:
        harvest(cx, a, rules)

if __name__ == "__main__":
    main()
