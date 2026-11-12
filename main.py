# =====================================================================
#  تقرير تحليل شكاوى العملاء المفصل - Zenput template 512247
#  GitHub Actions version (repo: complaints_analysis_report)
#
#  - Department from the form («نوع الشكوى» main type + sub-types)
#  - Qwen reads «محتوى شكوى العميل»: fills department when the form says
#    اخرى/empty, gives a fine-grained issue, flags form/text mismatches
#  - Branch from Zenput location ID, area manager from roster (08-11-2026)
#  - Period vs previous period; PDF + Excel emailed
#
#  Weekly: Sunday 00:00 -> Sunday 00:00 (the 7 days before the run-day
#  Sunday), compared with the week before. Triggered by the github alarm
#  Apps Script (workflow_dispatch) - nothing to set on the workflow.
#  Secrets: ZENPUT_API_KEY, GMAIL_APP_PASSWORD, GROQ_API_KEY (free, console.groq.com)
# =====================================================================
import os, re, ast, json, hashlib, html, smtplib, time
from collections import Counter
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import requests, pytz
import pandas as pd

# ------------------------------- CONFIG -------------------------------
ZENPUT_TOKEN = os.environ.get("ZENPUT_API_KEY", "")
GMAIL_PASS   = os.environ.get("GMAIL_APP_PASSWORD", "")
TEMPLATE_ID  = 512247
BASE         = "https://www.zenput.com"
TZ           = pytz.timezone("Asia/Riyadh")       # KSA (UTC+3)

WEEK_START   = ""          # backfill only: a Sunday "YYYY-MM-DD" -> that week; empty = normal run
TEST_MODE    = False       # True -> only a.alsalem (TO) + o.salahaddin (CC)
USE_QWEN     = True        # Qwen via Groq free API (GROQ_API_KEY secret); no key -> keyword rules
EMAIL_TITLE  = "تحليل الشكاوي الخاص بتجربة العميل"
COMPARE_PREV = True
MAX_RECORDS  = 30000

SENDER       = "aof.group.auto@gmail.com"
SENDER_NAME  = "Business Intelligence"
TO_EMAILS    = ["n.alkheraiji@aofgroup.com"]
CC_EMAILS    = ["omar@aofgroup.com", "o.salahaddin@aofgroup.com",
                "a.alsalem@aofgroup.com", "lena@aofgroup.com"]
TEST_TO      = ["a.alsalem@aofgroup.com"]
TEST_CC      = ["o.salahaddin@aofgroup.com"]

GROQ_API_KEY      = os.environ.get("GROQ_API_KEY", "")
GROQ_URL          = "https://api.groq.com/openai/v1"
GROQ_MODELS       = ["qwen/qwen3-32b", "qwen/qwen3.6-27b"]   # first one available on the account is used
LLM_BATCH         = 10       # complaints per request
MAX_CHARS         = 500
MAX_LLM_PER_RUN   = 1500
LLM_TIME_BUDGET_MIN = 30
CACHE_FILE        = "data/qwen_cache_v2.json"   # committed back to the repo by the workflow

REPEAT_MIN   = 2
TOP_N        = 15
COMP_MAX     = 1000
OUT_DIR      = "output"

DEPTS = ["تشغيل", "جودة", "توصيل", "تقني", "صيانة", "أخرى"]
DEPT_COLORS = {"تشغيل": "#2F6DB5", "جودة": "#2E9E6B", "توصيل": "#E08A1E",
               "تقني": "#7A4FC2", "صيانة": "#C2453D", "أخرى": "#8A94A3"}

# --------------------------- BRANCH MASTER ----------------------------
LOCATION_MAP = {
    2155652: "NURUH B01", 2155654: "AQRUH B13", 2164013: "KHRUH B02", 2164014: "GHRUH B03",
    2164016: "RWRUH B05", 2164017: "DARUH B06", 2164019: "SWRUH B08", 2164020: "AZRUH B09",
    2164021: "SHRUH B10", 2164022: "NRRUH B11", 2164023: "TWRUH B12", 2164025: "RBRUH B14",
    2164026: "NDRUH B15", 2164027: "BDRUH B16", 2164028: "QRRUH B17", 2164030: "MURUH B19",
    2164031: "SFJED B24", 2164032: "KRRUH B21", 2185452: "OBJED B22", 2190657: "SLAHS B23",
    2197297: "NSRUH B04", 2197298: "TKRUH B18", 2197299: "LBRUH B07", 2199002: "RWAHS B25",
    2199835: "HAJED B26", 2203271: "SARUH B27", 2210205: "MAJED B28", 2211854: "QARUH B30",
    2235670: "ANRUH B31", 2239240: "FYJED B32", 2242934: "HIRJED B33", 2243963: "URRUH B34",
    2250799: "IRRUH B35", 2257790: "SHWMAK B37", 2258220: "PSJED B36", 2260889: "UHDMM B38",
    2263062: "HSRUH B39", 2281339: "MZDMM B40", 2283017: "AQKBR B41", 2283896: "KFJUB B42",
    2169459: "لبدة العقيق LB01", 2222802: "لبدة الخليج LB02",
    2232755: "قراطيس السويدي QB01", 2235805: "قراطيس النرجس QB02", 2254072: "قراطيس اليرموك QB03",
    2256386: "قراطيس العقيق QB04", 2268360: "قراطيس الخافجي QB05", 2270650: "قراطيس الخليج QB06",
    2274188: "قراطيس الندوة QB07", 2276794: "قراطيس المنصورة QB08", 2282774: "قراطيس النزهة QB09",
    2171883: "تويستي TW01",
}
NON_BRANCH_IDS = {2175245, 2230615, 2256173}

ROSTER = {
    "عدنان العربي":   ["B21", "B17", "B15", "QB07", "B30", "B19", "B04", "B39"],
    "علي اسماعيل":    ["QB04", "B13", "B01", "TW01", "QB09"],
    "عمر الجوهري":    ["B18", "B16", "B12", "B07", "B06"],
    "علام محمود":     ["B26", "B32", "B22", "B37"],
    "اسلام مصطفى":    ["B35", "B34", "B14", "B11", "B31", "QB02", "QB03"],
    "مقداد علي":      ["QB05", "QB06", "B02", "B27", "B05"],
    "ناريش جوشي":     ["QB01", "B09", "B10", "B08", "QB08"],
    "عبدالله الغانمي": ["LB01", "LB02"],
    "صلاح الدين":     ["B23", "B25"],
    "محمد سهيل":      ["B38", "B40", "B41", "B42"],
    "سوزان بادول":    ["B28", "B24", "B33", "B36"],
}
CODE_TO_MGR = {c: m for m, cs in ROSTER.items() for c in cs}
CITY = {"RUH": "الرياض", "JED": "جدة", "AHS": "الأحساء", "DMM": "الدمام", "KBR": "الخبر", "JUB": "الجبيل", "MAK": "مكة"}

# ---------------------- ARABIC NORMALISATION -------------------------
def norm(s):
    s = str(s or "")
    s = re.sub(r"[إأآا]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", s).strip()

def clean_text(s):
    s = str(s or "")
    s = re.sub(r"[,،.]{3,}", " ", s)            # trailing ,,,,,,,
    s = re.sub(r"(.)\1{2,}", r"\1", s)          # بارررررد -> بارد
    return re.sub(r"\s+", " ", s).strip()

# ----------------------- FORM TYPE -> DEPARTMENT ----------------------
MAIN_TYPES = {"تشغيل": "تشغيل", "جوده": "جودة", "جودة": "جودة", "توصيل": "توصيل",
              "تقنى": "تقني", "تقني": "تقني", "صيانه": "صيانة", "صيانة": "صيانة",
              "اخرى": "أخرى", "أخرى": "أخرى"}
SUB_TYPES = {   # sub-type -> department (edit here if new values appear)
    "خطأ في الطلب من الفرع": "تشغيل", "تاخير فى تحضير الطلب": "تشغيل", "سلوك موظف فرع": "تشغيل",
    "خطاء فى تسجيل الطلب": "تشغيل", "طلب مبعثر من الفرع": "تشغيل", "الشكل العام ونظافه الفرع": "تشغيل",
    "تاخير فى توصيل الطلب": "توصيل", "تأخير شركة شروق": "توصيل", "خطأ في الطلب من قبل السائق": "توصيل",
    "اغلاق الطلب قبل التوصيل": "توصيل", "طلب مبعثر خطأ سائق": "توصيل",
    "ملاحظات واقتراحات": "أخرى", "الأسعار": "أخرى", "كول سنتر": "أخرى",
}
MAIN_N = {norm(k): v for k, v in MAIN_TYPES.items()}
SUB_N = {norm(k): (k, v) for k, v in SUB_TYPES.items()}
UNMAPPED_TYPES = Counter()

def split_multi(v):
    if v is None: return []
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("[") and s.endswith("]"):
            try: v = ast.literal_eval(s)
            except Exception: v = [s]
        else:
            v = [x for x in re.split(r"\s*\|\s*", s)]
    if not isinstance(v, (list, tuple)): v = [v]
    out = []
    for x in v:
        if isinstance(x, dict): x = x.get("name") or x.get("label") or x.get("value") or ""
        x = re.sub(r"\s+", " ", str(x)).strip()
        if x: out.append(x)
    return out

def form_departments(type_tokens, quality_tokens):
    depts, subs = [], []
    for t in type_tokens:
        n = norm(t)
        if n in MAIN_N:
            d = MAIN_N[n]
            if d not in depts: depts.append(d)
        elif n in SUB_N:
            label, d = SUB_N[n]
            subs.append(label)
            if d not in depts: depts.append(d)
        else:
            UNMAPPED_TYPES[t] += 1
            subs.append(t)
    if quality_tokens and "جودة" not in depts:
        depts.append("جودة")
    real = [d for d in depts if d != "أخرى"]
    return (real or depts), subs

# ----------------------------- HELPERS --------------------------------
def now_ksa(): return datetime.now(TZ)

def parse_dt(v):
    try: return pd.to_datetime(v, utc=True).tz_convert(TZ).to_pydatetime()
    except Exception: return None

def scalar(a):
    v = a.get("display_value")
    if v in (None, "", [], "[]"): v = a.get("value")
    if v is None: return ""
    if isinstance(v, (list, tuple, dict)) or (isinstance(v, str) and v.strip().startswith("[")):
        return " | ".join(split_multi(v))
    return re.sub(r"[ \t]+", " ", str(v)).strip()

def answers_by_title(ans):
    d = {}
    for a in ans or []:
        t = re.sub(r"<[^>]+>", "", a.get("title") or "").strip()
        d[t] = a
    return d

def get(d, title, multi=False):
    a = d.get(title)
    if a is None: return [] if multi else ""
    if multi:
        v = a.get("display_value")
        if v in (None, "", [], "[]"): v = a.get("value")
        return split_multi(v)
    return scalar(a)

def branch_info(md, form_branch):
    loc = md.get("location") or {}
    lid = loc.get("id") if isinstance(loc, dict) else None
    try: lid = int(lid)
    except Exception: lid = None
    if lid in NON_BRANCH_IDS: return None
    name = LOCATION_MAP.get(lid) or (f"UNMAPPED-{lid} " + re.sub(r"\s+", " ", form_branch) if lid else re.sub(r"\s+", " ", form_branch))
    m = re.search(r"\b(QB\d+|LB\d+|TW\d+|B\d+)\s*$", name)
    code = m.group(1) if (m and lid in LOCATION_MAP) else (m.group(1) if m else "")
    if code.startswith("QB"): brand = "قراطيس"
    elif code.startswith("LB"): brand = "لبدة"
    elif code.startswith("TW"): brand = "تويستي"
    else: brand = "شاورما كلاسك"
    city = "-"
    pm = re.match(r"([A-Z]+)\s", name)
    if pm: city = CITY.get(pm.group(1)[-3:], "-")
    return {"الفرع": name, "الكود": code, "البراند": brand, "المدينة": city,
            "مدير المنطقة": CODE_TO_MGR.get(code, "")}

def validity(review):
    r = norm(review)
    if not r: return "لم يحدد"
    if "غير صحيح" in r: return "غير صحيحة"
    if "صحيح" in r: return "صحيحة"
    return "لم يحدد"

def to_num(v):
    try: return float(str(v).replace(",", ""))
    except Exception: return 0.0

# ------------------------------- FETCH --------------------------------
def fetch_since(since_dt):
    H = {"X-API-TOKEN": ZENPUT_TOKEN}
    out, start = [], 0
    while start < MAX_RECORDS:
        r = requests.get(f"{BASE}/api/v3/submissions/", headers=H,
                         params={"form_template_id": TEMPLATE_ID, "limit": 100, "start": start}, timeout=90)
        if r.status_code != 200:
            print("HTTP", r.status_code, r.text[:300]); break
        data = r.json().get("data", [])
        if not data: break
        out.extend(data)
        dts = [d for d in (parse_dt((s.get("smetadata") or {}).get("date_submitted")) for s in data) if d]
        print(f"  start={start:>5}  got={len(data):>3}  oldest={min(dts):%Y-%m-%d %H:%M}" if dts else "")
        if dts and min(dts) < since_dt: break          # newest-first ordering confirmed
        if len(data) < 100: break
        start += 100
    return out

def build_df(subs, s, e):
    rows = []
    for sub in subs:
        md = sub.get("smetadata") or {}
        dt = parse_dt(md.get("date_submitted"))
        if not dt or dt < s or dt > e: continue
        d = answers_by_title(sub.get("answers"))
        bi = branch_info(md, get(d, "اختر الفرع"))
        if bi is None: continue
        types = get(d, "نوع الشكوى", multi=True)
        qual = get(d, "فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى", multi=True)
        depts, subtypes = form_departments(types, qual)
        products = [p for p in get(d, "الشكوى على اي منتج؟", multi=True) if "لا يوجد" not in p]
        comp = to_num(get(d, "قيمة التعويض"))
        legacy = sub.get("legacy_submission_id") or ""
        rows.append({
            "التاريخ": dt.strftime("%Y-%m-%d"), "الوقت": dt.strftime("%H:%M"),
            **bi,
            "مدير المنطقة (النموذج)": get(d, "مدير المنطقة المسؤول"),
            "مصدر الشكوى": get(d, "مصدر الشكوى") or "غير محدد",
            "تم الطلب من خلال": get(d, "تم الطلب من خلال") or "غير محدد",
            "نوع الطلب": get(d, "نوع الطلب") or "غير محدد",
            "نوع الشكوى (النموذج)": " | ".join(types),
            "_depts_form": depts,
            "_subtypes": subtypes,
            "_quality": qual,
            "_products": products,
            "تفصيل الجودة (النموذج)": " | ".join(qual),
            "المنتج": " | ".join(products),
            "محتوى الشكوى": clean_text(get(d, "محتوى شكوى العميل")),
            "تكرار الشكوى": get(d, "تكرار الشكوى") or "غير محدد",
            "قيمة التعويض": comp if comp <= COMP_MAX else 0.0,
            "تعويض غير منطقي": comp if comp > COMP_MAX else "",
            "الاجراء المتخذ": get(d, "الاجراء المتخذ"),
            "مدى الاجراء المتخذ": get(d, "مدى الاجراء المتخذ") or "لم يراجع",
            "رد المراجعة": get(d, "رد قسم المراجعة و رد قسم المتابعة"),
            "صحة الشكوى": validity(get(d, "رد قسم المراجعة و رد قسم المتابعة")),
            "الإجراء التشغيلي": re.sub(r"\s*\n\s*", " / ", get(d, "الإجراء التشغيلي المتخذ")),
            "رابط": f"https://ops-execution.crunchtime.com/reports/#form_id/{legacy}" if legacy else "",
        })
    return pd.DataFrame(rows)

# ---------------------------- QWEN LAYER ------------------------------
ISSUES = {
    "تشغيل": ["خطأ في الطلب (صنف مختلف)", "نقص في الطلب", "تأخير في التحضير", "سلوك الموظف",
              "خطأ في تسجيل الطلب", "نظافة ومظهر الفرع", "أخرى - تشغيل"],
    "جودة":  ["طعام بارد", "كميات وأحجام", "جودة الدجاج", "جودة اللحوم", "جودة البطاطس",
              "خبز غير طازج", "جودة الصوصات", "كمية الزيت", "جودة المشروبات",
              "غير مستوي أو محروق", "تلوث أو جسم غريب", "أخرى - جودة"],
    "توصيل": ["تأخير التوصيل", "خطأ أو نقص من السائق", "إغلاق الطلب قبل التوصيل",
              "سلوك المندوب", "أخرى - توصيل"],
    "تقني":  ["مشكلة التطبيق", "مشكلة الدفع", "الكاشير الذاتي أو الأنظمة", "أخرى - تقني"],
    "صيانة": ["التكييف", "المعدات", "دورات المياه", "الإنارة والمرافق", "أخرى - صيانة"],
    "أخرى":  ["ملاحظات واقتراحات", "الأسعار", "خدمة العملاء والكول سنتر", "غير واضحة"],
}
ISSUE_N = {d: {norm(c): c for c in cs} for d, cs in ISSUES.items()}

def allowed_depts(depts):
    return depts if depts and depts != ["أخرى"] else list(ISSUES.keys())

def issue_menu(depts):
    return "\n".join(f"- {d}: " + " | ".join(ISSUES[d]) for d in allowed_depts(depts))

SYS_PROMPT = (
    "أنت محلل شكاوى عملاء لسلسلة مطاعم شاورما في السعودية. اقرأ نص الشكوى وحدد المشكلة الأساسية.\n"
    "اختر إدارة واحدة ومشكلة واحدة حرفيا من القائمة المرسلة مع الشكوى فقط.\n"
    "تعريفات: تشغيل = ما يحدث داخل الفرع (دقة الطلب، النواقص، سرعة التحضير، تعامل الموظفين، نظافة الفرع). "
    "جودة = المنتج نفسه (حرارته، طعمه، استواؤه، كميته، طزاجته، أجسام غريبة). "
    "توصيل = بعد استلام المندوب للطلب (تأخر الوصول، أخطاء المندوب). "
    "تقني = التطبيق والدفع والأنظمة. صيانة = المعدات والمبنى. "
    "إذا اشتكى العميل من أكثر من شيء اختر الأخطر أو الأكثر تفصيلا.\n"
    'أجب بـ JSON فقط: {"department":"...","issue":"...","summary":"ملخص لا يتجاوز 10 كلمات"}'
)

def kw_issue(text, depts):
    t = norm(text)
    rules = [("جودة", "تلوث أو جسم غريب", ["شعر", "حشره", "صرصور", "بلاستيك", "جسم غريب", "ذبابه"]),
             ("جودة", "طعام بارد", ["بارد"]), ("جودة", "غير مستوي أو محروق", ["محروق", "ني ", "نيء", "مو مستوي"]),
             ("جودة", "كميات وأحجام", ["صغير", "قليل", "كميه", "حجم"]), ("جودة", "خبز غير طازج", ["خبز", "خبزه"]),
             ("جودة", "جودة البطاطس", ["بطاطس", "بطاط"]),
             ("توصيل", "تأخير التوصيل", ["تاخر التوصيل", "وصل متاخر", "المندوب تاخر"]),
             ("توصيل", "سلوك المندوب", ["المندوب", "السائق"]),
             ("تشغيل", "نقص في الطلب", ["ناقص", "نقص", "ما حطو", "بدون", "نسو"]),
             ("تشغيل", "خطأ في الطلب (صنف مختلف)", ["غلط", "بدل", "خطا"]),
             ("تشغيل", "تأخير في التحضير", ["تاخير", "تاخر", "انتظار", "بطي"]),
             ("تشغيل", "سلوك الموظف", ["تعامل", "اسلوب", "الموظف", "الكاشير"]),
             ("تقني", "مشكلة الدفع", ["الدفع", "مدي", "خصم"]), ("تقني", "مشكلة التطبيق", ["التطبيق"]),
             ("صيانة", "التكييف", ["مكيف", "تكييف"])]
    open_ = not depts or depts == ["أخرى"]
    for d, c, kws in rules:
        if (open_ or d in depts) and any(k in t for k in kws): return d, c
    d = depts[0] if depts else "أخرى"
    return d, ISSUES[d][-1]

def parse_llm(txt):
    m = re.search(r"\{.*?\}", txt, re.S)
    if not m: return None
    try: j = json.loads(m.group(0))
    except Exception: return None
    d = norm(j.get("department", ""))
    dept = next((x for x in DEPTS if norm(x) == d or norm(x) in d), None)
    if not dept: return None
    iss = ISSUE_N[dept].get(norm(j.get("issue", "")))
    if not iss:
        ni = norm(j.get("issue", ""))
        iss = next((v for k, v in ISSUE_N[dept].items() if ni and (ni in k or k in ni)), None)
        if not iss and ni:                      # best word overlap
            stop = {"في", "من", "او", "و", "ال"}
            w = set(ni.replace("ب", " ").split()) | set(ni.split())
            score = {v: len((set(k.split()) - stop) & w) for k, v in ISSUE_N[dept].items()}
            best = max(score, key=score.get)
            iss = best if score[best] else None
        iss = iss or ISSUES[dept][-1]
    return dept, iss, str(j.get("summary", "")).strip()

def load_cache():
    if os.path.exists(CACHE_FILE):
        try: return json.load(open(CACHE_FILE, encoding="utf-8"))
        except Exception: pass
    return {}

def save_cache(cache):
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    json.dump(cache, open(CACHE_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=0)

def groq_post(payload, tries=6):
    H = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    for i in range(tries):
        r = requests.post(f"{GROQ_URL}/chat/completions", headers=H, json=payload, timeout=120)
        if r.status_code == 200:
            return r.json()["choices"][0]["message"]["content"] or ""
        if r.status_code == 400 and "reasoning_effort" in payload and "reasoning" in r.text.lower():
            payload = {k: v for k, v in payload.items() if k != "reasoning_effort"}
            continue
        if r.status_code in (429, 500, 502, 503):
            wait = float(r.headers.get("retry-after") or 0) or min(60, 5 * (i + 1))
            print(f"  Groq {r.status_code} - waiting {wait:.0f}s")
            time.sleep(wait); continue
        print(f"  Groq HTTP {r.status_code}: {r.text[:300]}")
        return None
    return None

def pick_groq_model():
    try:
        r = requests.get(f"{GROQ_URL}/models", headers={"Authorization": f"Bearer {GROQ_API_KEY}"}, timeout=30)
        ids = {m["id"] for m in r.json().get("data", [])} if r.status_code == 200 else set()
    except Exception:
        ids = set()
    for m in GROQ_MODELS:
        if m in ids: return m
    qwen = sorted(i for i in ids if "qwen" in i.lower())
    if qwen: return qwen[0]
    print("Groq models available:", sorted(ids) or "none (check GROQ_API_KEY)")
    return None

BATCH_SYS = (SYS_PROMPT.split('أجب بـ JSON')[0] +
    "ستصلك عدة شكاوى، لكل شكوى رقم وقائمة الإدارات المسموح بها لها. اختر من تلك الإدارات فقط.\n"
    "قائمة المشاكل لكل إدارة (اكتب المشكلة حرفيا):\n" +
    "\n".join(f"- {d}: " + " | ".join(c) for d, c in ISSUES.items()) +
    '\n\nأجب بمصفوفة JSON فقط بدون أي نص آخر: '
    '[{"id":1,"department":"...","issue":"...","summary":"ملخص لا يتجاوز 10 كلمات"}, ...]')

def parse_batch(txt):
    txt = re.sub(r"<think>.*?</think>", "", txt or "", flags=re.S)
    m = re.search(r"\[.*\]", txt, re.S)
    if not m: return {}
    try: arr = json.loads(m.group(0))
    except Exception: return {}
    out = {}
    for o in arr if isinstance(arr, list) else []:
        try: out[int(o.get("id"))] = o
        except Exception: pass
    return out

def run_qwen(items):
    """Qwen (Groq free API). Returns key -> [dept, issue, summary, source].
    Only real Qwen answers are cached; keyword fallbacks are retried next run."""
    cache = load_cache()
    res = dict(cache)
    todo = [it for it in items if it["key"] not in cache][:MAX_LLM_PER_RUN]
    print(f"Qwen: {len(items)} unique complaints, {len(items) - len(todo)} cached, {len(todo)} new")
    model = pick_groq_model() if (todo and GROQ_API_KEY) else None
    if todo and not GROQ_API_KEY: print("GROQ_API_KEY not set -> keyword rules")
    if model:
        print(f"Using Groq model {model}")
        t0 = time.time(); done = 0
        for b in range(0, len(todo), LLM_BATCH):
            if time.time() - t0 > LLM_TIME_BUDGET_MIN * 60:
                print(f"LLM time budget ({LLM_TIME_BUDGET_MIN} min) reached"); break
            batch = todo[b:b + LLM_BATCH]
            lines = []
            for i, it in enumerate(batch, 1):
                lines.append(f"[{i}] الإدارات المسموح بها: {' / '.join(allowed_depts(it['depts']))}\n"
                             f"نوع الشكوى حسب خدمة العملاء: {it['form_type'] or '-'} | تفصيل الجودة: {it['quality'] or '-'} | المنتج: {it['product'] or '-'}\n"
                             f"نص الشكوى: {it['text'][:MAX_CHARS]}")
            payload = {"model": model, "temperature": 0, "max_tokens": 120 * len(batch) + 200,
                       "reasoning_effort": "none",
                       "messages": [{"role": "system", "content": BATCH_SYS},
                                    {"role": "user", "content": "\n\n".join(lines) + "\n\n/no_think"}]}
            ans = parse_batch(groq_post(payload))
            for i, it in enumerate(batch, 1):
                o = ans.get(i)
                if not o: continue
                p = parse_llm(json.dumps(o, ensure_ascii=False))
                if p and p[0] in allowed_depts(it["depts"]):
                    cache[it["key"]] = res[it["key"]] = [p[0], p[1], p[2], "qwen"]
            done += len(batch)
            print(f"  {done}/{len(todo)}  ({(time.time() - t0) / done:.1f}s per complaint)")
            if (b // LLM_BATCH) % 5 == 0: save_cache(cache)
            time.sleep(2.5)                      # stay under the free per-minute limits
        save_cache(cache)
    for it in items:
        if it["key"] not in res:
            d, c = kw_issue(it["text"], it["depts"]); res[it["key"]] = [d, c, "", "keywords"]
    return res

def enrich(df, use_llm):
    df["_key"] = [hashlib.md5(f"{t}|{ft}".encode()).hexdigest()
                  for t, ft in zip(df["محتوى الشكوى"], df["نوع الشكوى (النموذج)"])]
    if use_llm:
        items = []
        for _, r in df.drop_duplicates("_key").iterrows():
            items.append({"key": r["_key"], "text": r["محتوى الشكوى"], "form_type": r["نوع الشكوى (النموذج)"],
                          "quality": r["تفصيل الجودة (النموذج)"], "product": r["المنتج"], "depts": r["_depts_form"]})
        res = run_qwen(items)
    else:
        res = {}
        for _, r in df.drop_duplicates("_key").iterrows():
            d, c = kw_issue(r["محتوى الشكوى"], r["_depts_form"]); res[r["_key"]] = [d, c, "", "keywords"]
    q = df["_key"].map(res)
    df["إدارة (تحليل النص)"] = q.map(lambda x: x[0])
    df["المشكلة (تحليل النص)"] = q.map(lambda x: x[1])
    df["ملخص الشكوى"] = q.map(lambda x: x[2])
    df["مصدر التحليل"] = q.map(lambda x: x[3])

    finals, srcs, mism = [], [], []
    for fd, qd in zip(df["_depts_form"], df["إدارة (تحليل النص)"]):
        if fd and fd != ["أخرى"]:
            finals.append(fd); srcs.append("النموذج")
            mism.append("نعم" if qd not in fd and qd != "أخرى" else "")
        else:
            finals.append([qd]); srcs.append("تحليل النص"); mism.append("")
    df["_depts"] = finals
    df["الإدارة الرئيسية"] = [x[0] for x in finals]
    df["الإدارات"] = [" | ".join(x) for x in finals]
    df["مصدر الإدارة"] = srcs
    df["تصنيف مختلف عن النص"] = mism
    return df.drop(columns="_key")

def form_only(df):
    df["_depts"] = [x if x else ["أخرى"] for x in df["_depts_form"]]
    df["الإدارة الرئيسية"] = [x[0] for x in df["_depts"]]
    return df

# ------------------------------ ANALYSIS ------------------------------
def pct(a, b): return round(a / b * 100, 1) if b else 0.0
def chg(cur, prev): return "جديد" if not prev and cur else (f"{(cur - prev) / prev * 100:+.0f}%" if prev else "-")

def explode(df, col, name):
    return df[[c for c in df.columns if not c.startswith("_") and c != name] + [col]].explode(col).rename(columns={col: name}).dropna(subset=[name]).reset_index(drop=True)

def analyse(df, prev):
    N, A = len(df), {}
    P = len(prev) if prev is not None else 0
    de = explode(df, "_depts", "الإدارة")
    dc = de["الإدارة"].value_counts().reindex(DEPTS, fill_value=0)
    pc = explode(prev, "_depts", "الإدارة")["الإدارة"].value_counts().reindex(DEPTS, fill_value=0) if P else None
    A["dept"] = pd.DataFrame({
        "الإدارة": DEPTS, "عدد الشكاوى": dc.values, "النسبة %": [pct(v, N) for v in dc.values],
        "إدارة رئيسية": df["الإدارة الرئيسية"].value_counts().reindex(DEPTS, fill_value=0).values,
        **({"الفترة السابقة": pc.values, "التغير": [chg(a, b) for a, b in zip(dc.values, pc.values)]} if P else {}),
    })
    A["dept"] = A["dept"][A["dept"]["عدد الشكاوى"] > 0]

    iss = df.groupby(["المشكلة (تحليل النص)", "إدارة (تحليل النص)"]).agg(
        عدد=("الفرع", "size"), فروع=("الفرع", "nunique")).reset_index().sort_values("عدد", ascending=False)
    iss.columns = ["المشكلة", "الإدارة", "عدد الشكاوى", "عدد الفروع"]
    iss.insert(3, "النسبة %", iss["عدد الشكاوى"].map(lambda v: pct(v, N)))
    A["issues"] = iss.reset_index(drop=True)

    st = explode(df, "_subtypes", "النوع")["النوع"].value_counts()
    qt = explode(df, "_quality", "النوع")["النوع"].value_counts()
    A["form_sub"] = pd.concat([
        pd.DataFrame({"المصدر": "نوع الشكوى", "التفصيل": st.index, "عدد الشكاوى": st.values}),
        pd.DataFrame({"المصدر": "تفصيل الجودة", "التفصيل": qt.index, "عدد الشكاوى": qt.values})], ignore_index=True)
    if P:
        pst = explode(prev, "_subtypes", "النوع")["النوع"].value_counts()
        pqt = explode(prev, "_quality", "النوع")["النوع"].value_counts()
        A["form_sub"]["الفترة السابقة"] = [ (pst if s == "نوع الشكوى" else pqt).get(t, 0)
                                             for s, t in zip(A["form_sub"]["المصدر"], A["form_sub"]["التفصيل"])]
        A["form_sub"]["التغير"] = [chg(a, b) for a, b in zip(A["form_sub"]["عدد الشكاوى"], A["form_sub"]["الفترة السابقة"])]

    piv = pd.crosstab(de["الفرع"], de["الإدارة"]).reindex(columns=[d for d in DEPTS], fill_value=0)
    base = df.groupby("الفرع").agg(البراند=("البراند", "first"), المدينة=("المدينة", "first"),
                                    مدير=("مدير المنطقة", "first"), الإجمالي=("الفرع", "size"),
                                    مكررة=("تكرار الشكوى", lambda s: (s == "مكررة").sum()),
                                    صحيحة=("صحة الشكوى", lambda s: (s == "صحيحة").sum()),
                                    تعويض=("قيمة التعويض", "sum"),
                                    أكثر_مشكلة=("المشكلة (تحليل النص)", lambda s: s.value_counts().index[0]))
    br = base.join(piv).reset_index()
    if P:
        pb = prev.groupby("الفرع").size()
        br["الفترة السابقة"] = br["الفرع"].map(pb).fillna(0).astype(int)
        br["التغير"] = [chg(a, b) for a, b in zip(br["الإجمالي"], br["الفترة السابقة"])]
    br = br.rename(columns={"مدير": "مدير المنطقة", "مكررة": "مكررة (النموذج)", "صحيحة": "صحيحة (المراجعة)",
                            "تعويض": "إجمالي التعويض", "أكثر_مشكلة": "أكثر مشكلة"})
    A["branches"] = br.sort_values("الإجمالي", ascending=False).reset_index(drop=True)

    rep = df.groupby(["الفرع", "إدارة (تحليل النص)", "المشكلة (تحليل النص)"]).agg(
        العدد=("التاريخ", "size"), أول=("التاريخ", "min"), آخر=("التاريخ", "max")).reset_index()
    rep = rep[rep["العدد"] >= REPEAT_MIN].sort_values(["العدد", "الفرع"], ascending=[False, True]).reset_index(drop=True)
    rep.columns = ["الفرع", "الإدارة", "المشكلة", "العدد", "أول تاريخ", "آخر تاريخ"]
    A["repeat"] = rep

    pe = explode(df, "_products", "المنتج")
    if len(pe):
        pr = pe.groupby("المنتج").agg(عدد=("الفرع", "size"),
                                      أكثر_مشكلة=("المشكلة (تحليل النص)", lambda s: s.value_counts().index[0])).reset_index()
        pr.columns = ["المنتج", "عدد الشكاوى", "أكثر مشكلة"]
        A["products"] = pr.sort_values("عدد الشكاوى", ascending=False).reset_index(drop=True)
    else:
        A["products"] = pd.DataFrame(columns=["المنتج", "عدد الشكاوى", "أكثر مشكلة"])

    mg = pd.crosstab(de["مدير المنطقة"].replace("", "غير محدد"), de["الإدارة"]).reindex(columns=DEPTS, fill_value=0)
    mg.insert(0, "الإجمالي", df.groupby(df["مدير المنطقة"].replace("", "غير محدد")).size())
    mg["عدد الفروع"] = df.groupby(df["مدير المنطقة"].replace("", "غير محدد"))["الفرع"].nunique()
    A["managers"] = mg.sort_values("الإجمالي", ascending=False).reset_index().rename(columns={"مدير المنطقة": "مدير المنطقة"})

    brand = df.groupby("البراند").size().rename("عدد الشكاوى").reset_index()
    A["brand"] = brand.sort_values("عدد الشكاوى", ascending=False)
    for k, col in [("source", "مصدر الشكوى"), ("channel", "تم الطلب من خلال"), ("otype", "نوع الطلب"),
                   ("action", "الاجراء المتخذ"), ("status", "مدى الاجراء المتخذ"), ("valid", "صحة الشكوى"),
                   ("repeatflag", "تكرار الشكوى")]:
        A[k] = df[col].replace("", "غير محدد").value_counts().rename_axis(col).reset_index(name="عدد الشكاوى")

    A["mismatch"] = df[df["تصنيف مختلف عن النص"] == "نعم"][
        ["التاريخ", "الفرع", "نوع الشكوى (النموذج)", "إدارة (تحليل النص)", "المشكلة (تحليل النص)", "محتوى الشكوى"]]
    A["daily"] = df.groupby("التاريخ").size().rename("عدد الشكاوى").reset_index()
    A["kpi"] = {"N": N, "P": P, "branches": df["الفرع"].nunique(),
                "repeat_pct": pct((df["تكرار الشكوى"] == "مكررة").sum(), N),
                "valid": (df["صحة الشكوى"] == "صحيحة").sum(), "invalid": (df["صحة الشكوى"] == "غير صحيحة").sum(),
                "comp": df["قيمة التعويض"].sum(), "unreviewed": (df["مدى الاجراء المتخذ"] == "لم يراجع").sum(),
                "mismatch": len(A["mismatch"])}
    return A

# -------------------------------- EXCEL -------------------------------
def write_excel(df, A, path):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    raw_cols = [c for c in df.columns if not c.startswith("_")]
    sheets = [("الإدارات", A["dept"]), ("أكثر المشاكل (النص)", A["issues"]),
              ("أنواع الشكوى (النموذج)", A["form_sub"]), ("الفروع", A["branches"]),
              ("المشاكل المتكررة بالفرع", A["repeat"]), ("المنتجات", A["products"]),
              ("مديرو المناطق", A["managers"]), ("البراند", A["brand"]),
              ("مصدر الشكوى", A["source"]), ("قناة الطلب", A["channel"]), ("نوع الطلب", A["otype"]),
              ("الإجراء والحالة", pd.concat([A["action"], A["status"], A["valid"]], axis=1)),
              ("تصنيف مختلف", A["mismatch"]), ("يومي", A["daily"]), ("البيانات", df[raw_cols])]
    hdr = PatternFill("solid", fgColor="1F3A5F"); thin = Side(style="thin", color="D0D7E2")
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, d in sheets:
            d.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            ws.sheet_view.rightToLeft = True; ws.freeze_panes = "A2"
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF"); c.fill = hdr
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                    c.alignment = Alignment(vertical="top", wrap_text=True)
            for col in ws.columns:
                L = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(11, L + 2), 60)
    print("Excel saved:", path)

# -------------------------------- HTML --------------------------------
E = lambda x: html.escape(str(x))

def bars(items, color="#2F6DB5", colors=None, total=None):
    mx = max([v for _, v in items] or [1]) or 1
    out = []
    for lab, v in items:
        col = (colors or {}).get(lab, color)
        p = f' <span class="pp">{pct(v, total)}%</span>' if total else ""
        out.append(f'<div class="bar"><div class="bl">{E(lab)}</div><div class="bt"><div class="bf" '
                   f'style="width:{v / mx * 100:.1f}%;background:{col}"></div></div><div class="bv">{v}{p}</div></div>')
    return "".join(out)

def table(d, maxrows=None):
    d = d.head(maxrows) if maxrows else d
    th = "".join(f"<th>{E(c)}</th>" for c in d.columns)
    tr = "".join("<tr>" + "".join(f"<td>{E(v)}</td>" for v in r) + "</tr>" for r in d.itertuples(index=False))
    return f"<table><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table>"

def build_html(df, A, s, e, ps, pe):
    K, N = A["kpi"], A["kpi"]["N"]
    delta = f' <small>({chg(K["N"], K["P"])} عن الفترة السابقة)</small>' if K["P"] else ""
    real = A["issues"][~A["issues"]["المشكلة"].str.startswith("أخرى") & (A["issues"]["المشكلة"] != "غير واضحة")]
    top_iss = real.iloc[0] if len(real) else (A["issues"].iloc[0] if len(A["issues"]) else None)
    top_br = A["branches"].iloc[0] if len(A["branches"]) else None
    kpis = [("إجمالي الشكاوى", f'{N}{delta}'), ("عدد الفروع", K["branches"]),
            ("أكثر مشكلة", f'{top_iss["المشكلة"]} ({top_iss["عدد الشكاوى"]})' if top_iss is not None else "-"),
            ("أعلى فرع", f'{top_br["الفرع"]} ({top_br["الإجمالي"]})' if top_br is not None else "-"),
            ("شكاوى مكررة (النموذج)", f'{K["repeat_pct"]}%'),
            ("صحيحة / غير صحيحة", f'{K["valid"]} / {K["invalid"]}'),
            ("إجمالي التعويض", f'{K["comp"]:,.0f} ر.س'), ("لم تراجع من التشغيل", K["unreviewed"])]
    kpi_html = "".join(f'<div class="kpi"><div class="kl">{E(a)}</div><div class="kv">{b}</div></div>' for a, b in kpis)

    dept_cards = ""
    de = df[["الفرع", "المشكلة (تحليل النص)", "_depts"]].explode("_depts").reset_index(drop=True)
    for d in DEPTS:
        sub = de[de["_depts"] == d]
        if sub.empty: continue
        iss = sub["المشكلة (تحليل النص)"].value_counts().head(8)
        brs = sub["الفرع"].value_counts().head(6)
        dept_cards += (f'<div class="card"><h3 style="border-color:{DEPT_COLORS[d]}">{d} '
                       f'<span>{len(sub)} شكوى - {pct(len(sub), N)}%</span></h3><div class="two">'
                       f'<div><h4>المشاكل (تحليل النص)</h4>{bars(list(iss.items()), DEPT_COLORS[d])}</div>'
                       f'<div><h4>أكثر الفروع</h4>{bars(list(brs.items()), DEPT_COLORS[d])}</div></div></div>')

    br = A["branches"].head(TOP_N)
    mx = max(br["الإجمالي"].max() if len(br) else 1, 1)
    stacked = ""
    prim = df.groupby(["الفرع", "الإدارة الرئيسية"]).size().unstack(fill_value=0)
    for _, r in br.iterrows():
        row = prim.loc[r["الفرع"]] if r["الفرع"] in prim.index else {}
        segs = "".join(f'<div style="width:{row.get(d, 0) / mx * 100:.2f}%;background:{DEPT_COLORS[d]}"></div>'
                       for d in DEPTS if row.get(d, 0))
        stacked += (f'<div class="bar"><div class="bl">{E(r["الفرع"])}</div><div class="bt st">{segs}</div>'
                    f'<div class="bv">{r["الإجمالي"]}</div></div>')
    legend = "".join(f'<span class="lg"><i style="background:{DEPT_COLORS[d]}"></i>{d}</span>' for d in DEPTS)

    br_cols = ["الفرع", "مدير المنطقة", "الإجمالي"] + [d for d in DEPTS if d in A["branches"].columns] + \
              ["مكررة (النموذج)", "صحيحة (المراجعة)", "أكثر مشكلة"] + (["الفترة السابقة", "التغير"] if K["P"] else [])

    samples = ""
    for iss in A["issues"]["المشكلة"].head(6):
        ex = df[df["المشكلة (تحليل النص)"] == iss].head(3)
        li = "".join(f'<li><b>{E(r["الفرع"])}</b> ({E(r["التاريخ"])}): {E(r["محتوى الشكوى"][:230])}</li>' for _, r in ex.iterrows())
        samples += f'<div class="card"><h4>{E(iss)}</h4><ul>{li}</ul></div>'

    def mini(title, d, col):
        return f'<div><h4>{title}</h4>{bars(list(zip(d[col], d["عدد الشكاوى"])), total=N)}</div>'

    prev_txt = f" &nbsp;|&nbsp; المقارنة: الأحد {ps:%Y-%m-%d} إلى الأحد {pe + timedelta(seconds=1):%Y-%m-%d}" if K["P"] else ""
    return f"""<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
<style>
body{{font-family:'Cairo',Tahoma,sans-serif;background:#F4F6FA;color:#1E2733;margin:0;padding:24px;width:1240px}}
.hd{{background:#1F3A5F;color:#fff;border-radius:14px;padding:20px 26px;margin-bottom:18px}}
.hd h1{{margin:0;font-size:26px}} .hd p{{margin:4px 0 0;opacity:.85;font-size:14px}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}}
.kpi{{background:#fff;border-radius:12px;padding:12px 14px;box-shadow:0 1px 3px rgba(0,0,0,.06)}}
.kl{{font-size:12px;color:#6B7684}} .kv{{font-size:18px;font-weight:700;margin-top:4px}} .kv small{{font-size:12px;color:#6B7684;font-weight:400}}
.card{{background:#fff;border-radius:12px;padding:16px 18px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.06)}}
h2{{font-size:19px;margin:24px 0 10px;color:#1F3A5F}} h3{{margin:0 0 10px;border-right:5px solid;padding-right:10px}}
h3 span{{font-size:13px;color:#6B7684;font-weight:400;margin-right:8px}} h4{{margin:6px 0 8px;font-size:14px;color:#44505E}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:22px}} .three{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:22px}}
.bar{{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:13px}}
.bl{{width:210px;text-align:right}} .bt{{flex:1;background:#EEF1F6;border-radius:6px;height:16px;overflow:hidden;display:flex}}
.bf{{height:100%;border-radius:6px}} .st div{{height:100%}} .bv{{min-width:44px;font-weight:700;text-align:left}} .pp{{color:#6B7684;font-weight:400;font-size:11px}}
.lg{{display:inline-flex;align-items:center;gap:5px;margin-left:16px;font-size:13px}} .lg i{{width:12px;height:12px;border-radius:3px;display:inline-block}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}} th{{background:#1F3A5F;color:#fff;padding:7px;font-weight:600}}
td{{padding:6px 7px;border-bottom:1px solid #E3E8EF;vertical-align:top}} tr:nth-child(even) td{{background:#F8FAFC}}
ul{{margin:0;padding-right:18px;font-size:13px}} li{{margin:4px 0}} .note{{font-size:12px;color:#6B7684;margin-top:6px}}
</style></head><body>
<div class="hd"><h1>{EMAIL_TITLE}</h1>
<p>الفترة: الأحد {s:%Y-%m-%d} إلى الأحد {e + timedelta(seconds=1):%Y-%m-%d}{prev_txt} &nbsp;|&nbsp; نموذج {TEMPLATE_ID} &nbsp;|&nbsp; الإصدار: {now_ksa():%Y-%m-%d %H:%M}</p></div>
<div class="kpis">{kpi_html}</div>

<h2>توزيع الشكاوى على الإدارات</h2>
<div class="card">{bars(list(zip(A["dept"]["الإدارة"], A["dept"]["عدد الشكاوى"])), colors=DEPT_COLORS, total=N)}
<div class="note">الإدارة حسب اختيار خدمة العملاء في النموذج؛ عند «اخرى» أو عدم التحديد تؤخذ من تحليل نص الشكوى. الشكوى قد تخص أكثر من إدارة لذلك مجموع النسب قد يتجاوز 100%.</div></div>
<div class="card">{table(A["dept"])}</div>

<h2>أكثر الشكاوى تكرارا (تحليل نص الشكوى)</h2>
<div class="card">{bars([(r["المشكلة"], r["عدد الشكاوى"]) for _, r in A["issues"].head(TOP_N).iterrows()], total=N)}</div>
<div class="card">{table(A["issues"], TOP_N * 2)}</div>

<h2>أنواع الشكوى كما سجلتها خدمة العملاء</h2>
<div class="card">{table(A["form_sub"])}</div>

<h2>أكثر الفروع شكاوى</h2>
<div class="card"><div style="margin-bottom:8px">{legend}</div>{stacked}</div>
<div class="card">{table(A["branches"][br_cols], 60)}</div>

<h2>تفصيل كل إدارة</h2>{dept_cards}

<h2>مشاكل متكررة في نفس الفرع (≥ {REPEAT_MIN})</h2>
<div class="card">{table(A["repeat"], 40) if len(A["repeat"]) else "لا توجد"}</div>

<h2>المنتجات الأكثر شكاوى</h2>
<div class="card">{table(A["products"], TOP_N)}</div>

<h2>مديرو المناطق</h2><div class="card">{table(A["managers"])}</div>

<h2>المصادر والقنوات</h2>
<div class="card"><div class="three">{mini("مصدر الشكوى", A["source"], "مصدر الشكوى")}
{mini("تم الطلب من خلال", A["channel"], "تم الطلب من خلال")}{mini("نوع الطلب", A["otype"], "نوع الطلب")}</div></div>
<div class="card"><div class="three">{mini("الاجراء المتخذ", A["action"], "الاجراء المتخذ")}
{mini("مدى الاجراء المتخذ", A["status"], "مدى الاجراء المتخذ")}{mini("صحة الشكوى (رد المراجعة)", A["valid"], "صحة الشكوى")}</div></div>

<h2>أمثلة من نص الشكاوى</h2>{samples}
</body></html>"""

# ------------------------------ PERIODS -------------------------------
def periods():
    """Week = [Sunday 00:00, next Sunday 00:00) KSA.
    Normal run: the week that ends at the most recent Sunday 00:00 (today if today is Sunday).
    Returns (start, end_inclusive, end_sunday) for current and previous week."""
    if WEEK_START:
        ws = datetime.strptime(WEEK_START, "%Y-%m-%d").date()
        if ws.weekday() != 6: raise SystemExit("WEEK_START must be a Sunday")
        end_sun = ws + timedelta(days=7)
    else:
        today = now_ksa().date()
        end_sun = today - timedelta(days=(today.weekday() - 6) % 7)   # Sunday = 6
    start_sun = end_sun - timedelta(days=7)
    mk = lambda d: TZ.localize(datetime(d.year, d.month, d.day))
    cur  = (mk(start_sun), mk(end_sun) - timedelta(seconds=1), end_sun)
    prev = (mk(start_sun - timedelta(days=7)), mk(start_sun) - timedelta(seconds=1), start_sun)
    return cur, prev

# -------------------------------- PDF ---------------------------------
def html_to_pdf(page_html, pdf_path):
    from playwright.sync_api import sync_playwright
    from PIL import Image
    from fpdf import FPDF
    Image.MAX_IMAGE_PIXELS = None
    png = pdf_path.replace(".pdf", ".png")
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 1290, "height": 1000}, device_scale_factor=2)
        pg.set_content(page_html, wait_until="networkidle")
        pg.wait_for_timeout(1500)
        pg.screenshot(path=png, full_page=True)
        b.close()
    im = Image.open(png); w, h = im.size
    # split very tall screenshots into A4-ratio pages so the PDF stays readable
    page_h = int(w * 1.414)
    pdf = FPDF(unit="pt", format=(w / 2, page_h / 2)); pdf.set_auto_page_break(False)
    y = 0; i = 0
    while y < h:
        crop = im.crop((0, y, w, min(y + page_h, h)))
        part = png.replace(".png", f"_{i}.png")
        canvas = Image.new("RGB", (w, page_h), (244, 246, 250)); canvas.paste(crop, (0, 0)); canvas.save(part)
        pdf.add_page(); pdf.image(part, 0, 0, w / 2, page_h / 2)
        os.remove(part); y += page_h; i += 1
    pdf.output(pdf_path); os.remove(png)
    print("PDF saved:", pdf_path, f"({i} pages)")

# -------------------------------- EMAIL -------------------------------
def email_body(A, s, end_sun):
    K = A["kpi"]
    td = 'style="padding:6px 10px;border:1px solid #D9DEE7;text-align:right"'
    th = 'style="padding:6px 10px;border:1px solid #1F3A5F;background:#1F3A5F;color:#fff;text-align:right"'
    def tbl(d, cols):
        h = "".join(f"<th {th}>{E(c)}</th>" for c in cols)
        r = "".join("<tr>" + "".join(f"<td {td}>{E(v)}</td>" for v in row) + "</tr>"
                    for row in d[cols].itertuples(index=False))
        return f'<table style="border-collapse:collapse;font-size:13px;margin:6px 0 16px">{h and "<tr>"+h+"</tr>"}{r}</table>'
    dcols = ["الإدارة", "عدد الشكاوى", "النسبة %"] + (["الفترة السابقة", "التغير"] if K["P"] else [])
    bcols = ["الفرع", "مدير المنطقة", "الإجمالي", "أكثر مشكلة"] + (["التغير"] if K["P"] else [])
    prev_line = f" مقارنة بـ {K['P']} شكوى في الأسبوع السابق ({chg(K['N'], K['P'])})" if K["P"] else ""
    return f"""<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;color:#1E2733;font-size:14px;line-height:1.7">
<h2 style="color:#1F3A5F;margin:0 0 10px">{EMAIL_TITLE}</h2>
<p>السادة الإدارة،</p>
<p>مقدم لسيادتكم تحليل الشكاوي الخاص بتجربة العميل للأسبوع من
<b>الأحد {s:%Y-%m-%d}</b> إلى <b>الأحد {end_sun:%Y-%m-%d}</b>، مقارنة بالأسبوع السابق.</p>
<p>إجمالي الشكاوى: <b>{K['N']}</b>{prev_line}، في {K['branches']} فرع.
الشكاوى المكررة حسب النموذج: <b>{K['repeat_pct']}%</b>. إجمالي التعويض: <b>{K['comp']:,.0f} ر.س</b>.
شكاوى لم تراجع من التشغيل: <b>{K['unreviewed']}</b>.</p>
<h3 style="color:#1F3A5F;margin:14px 0 4px">توزيع الشكاوى على الإدارات</h3>{tbl(A['dept'], dcols)}
<h3 style="color:#1F3A5F;margin:14px 0 4px">أكثر الشكاوى تكرارا</h3>{tbl(A['issues'].head(8), ['المشكلة', 'الإدارة', 'عدد الشكاوى', 'النسبة %', 'عدد الفروع'])}
<h3 style="color:#1F3A5F;margin:14px 0 4px">أكثر الفروع شكاوى</h3>{tbl(A['branches'].head(10), bcols)}
<p>التقرير الكامل مرفق بصيغة PDF، والبيانات التفصيلية مرفقة بملف Excel.</p>
<p style="color:#6B7684;font-size:12px">تم إعداد هذا التقرير آليا بواسطة فريق ذكاء الأعمال.</p>
</div>"""

def send_email(subject, body_html, attachments):
    to, cc = (TEST_TO, TEST_CC) if TEST_MODE else (TO_EMAILS, CC_EMAILS)
    msg = MIMEMultipart("mixed")
    msg["From"] = f"{SENDER_NAME} <{SENDER}>"
    msg["To"] = ", ".join(to); msg["Cc"] = ", ".join(cc)
    msg["Subject"] = ("[TEST] " if TEST_MODE else "") + subject
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    for path in attachments:
        if not path or not os.path.exists(path): continue
        part = MIMEBase("application", "octet-stream")
        part.set_payload(open(path, "rb").read()); encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(SENDER, GMAIL_PASS)
        srv.sendmail(SENDER, to + cc, msg.as_string())
    print(f"Email sent -> TO {to} CC {cc}")

# --------------------------------- MAIN -------------------------------
def main():
    if not ZENPUT_TOKEN: raise SystemExit("ZENPUT_API_KEY secret is missing")
    os.makedirs(OUT_DIR, exist_ok=True)
    (s, e, end_sun), (ps, pe, _) = periods()
    print(f"Week    : Sunday {s:%Y-%m-%d} -> Sunday {end_sun:%Y-%m-%d}  | test={TEST_MODE}  qwen={USE_QWEN}")
    print(f"Previous: Sunday {ps:%Y-%m-%d} -> Sunday {s:%Y-%m-%d}")

    subs = fetch_since(ps if COMPARE_PREV else s)
    print(f"Fetched {len(subs)} submissions")
    prev = build_df(subs, ps, pe) if COMPARE_PREV else None
    UNMAPPED_TYPES.clear()
    df = build_df(subs, s, e)
    print(f"Current: {len(df)} complaints | Previous: {len(prev) if prev is not None else '-'}")

    subject_base = f"{EMAIL_TITLE} - الأحد {s:%Y-%m-%d} إلى الأحد {end_sun:%Y-%m-%d}"
    if df.empty:
        send_email(subject_base, f'<div dir="rtl" style="font-family:Tahoma">لا توجد شكاوى مسجلة خلال الأسبوع من الأحد {s:%Y-%m-%d} إلى الأحد {end_sun:%Y-%m-%d}.</div>', [])
        return

    if UNMAPPED_TYPES:
        print("\nWARNING - «نوع الشكوى» values not mapped (add to SUB_TYPES):")
        for k, v in UNMAPPED_TYPES.most_common(): print(f"   {v:>4}  {k}")
    unm = df[df["الفرع"].str.startswith("UNMAPPED")]["الفرع"].value_counts()
    if len(unm): print("\nWARNING - unmapped locations:\n", unm.to_string())
    bad = df.loc[df["تعويض غير منطقي"] != "", "تعويض غير منطقي"].tolist()
    if bad: print(f"\nIgnored compensation values > {COMP_MAX}: {bad}")

    df = enrich(df, USE_QWEN)
    prev = form_only(prev) if prev is not None and len(prev) else None
    print("\nAnalysis source:", df["مصدر التحليل"].value_counts().to_dict())
    print("Department source:", df["مصدر الإدارة"].value_counts().to_dict())

    A = analyse(df, prev)
    print("\n", A["dept"].to_string(index=False))
    print("\nTop issues:\n", A["issues"].head(12).to_string(index=False))
    print(f"\nMismatch form vs text: {A['kpi']['mismatch']}")

    tag = f"week_{s:%Y%m%d}_{end_sun:%Y%m%d}"
    xlsx = f"{OUT_DIR}/complaints_analysis_{tag}.xlsx"
    htmlp = f"{OUT_DIR}/complaints_analysis_{tag}.html"
    pdfp = f"{OUT_DIR}/complaints_analysis_{tag}.pdf"
    write_excel(df, A, xlsx)
    page = build_html(df, A, s, e, ps, pe)
    open(htmlp, "w", encoding="utf-8").write(page)
    try:
        html_to_pdf(page, pdfp)
    except Exception as ex:
        print("PDF failed:", ex); pdfp = None

    send_email(subject_base, email_body(A, s, end_sun), [pdfp, xlsx])

if __name__ == "__main__":
    main()
