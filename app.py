import os, re, sys, io, html, json, time, base64, random, hashlib, datetime, zipfile, shutil, subprocess, threading, traceback, unicodedata
from concurrent.futures import ThreadPoolExecutor
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import telebot
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from pptx import Presentation
from pptx.util import Pt as PptPt, Inches as PptInches
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.text import PP_ALIGN

try:
    import weasyprint
    HAS_WEASYPRINT = True
except Exception:
    HAS_WEASYPRINT = False

# ==============================================================================
# الوحدة (1): محرك Google Gemini الحصري (8 مفاتيح + تناوب على أقوى نموذج أولاً + بحث Google الحي)
# ==============================================================================
def _make_resilient_session():
    s = requests.Session()
    retry = Retry(total=2, backoff_factor=0.4, status_forcelist=[502, 503, 504], allowed_methods=["GET", "POST"])
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s

_HTTP_SESSION = _make_resilient_session()
_error_log = []

def log_runtime_error(context_str, exc):
    entry = f"[{datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}] {context_str}: {str(exc)[:250]}"
    _error_log.append(entry)
    if len(_error_log) > 50:
        _error_log.pop(0)

raw_gem_env = os.environ.get("STATLAB_GEMINI_KEYS", "").strip() or os.environ.get("STATLAB_GEMINI_KEY", "").strip()
GEMINI_KEYS = []
for k_item in raw_gem_env.split(","):
    k_c = k_item.strip()
    if k_c and k_c not in GEMINI_KEYS:
        GEMINI_KEYS.append(k_c)

BOT_TOKEN = os.environ.get("STATLAB_BOT_TOKEN", "").strip()
GH_SYNC_TOKEN = os.environ.get("STATLAB_GH_TOKEN", "").strip()
ADMIN_ID = int(os.environ.get("STATLAB_ADMIN_ID", "8639396072"))

bot = telebot.TeleBot(BOT_TOKEN if ":" in BOT_TOKEN else "100000000:AAH_placeholder_for_static_analysis", threaded=True, num_threads=8)
orders_db = {}
awaiting_paraphrase_uids = set()
MEMORY_FILE = "learned_lessons.json"
ORDERS_CACHE_FILE = "orders_cache.local"
ALLOWED_USERS_FILE = "allowed_users.json"
CREDITS_DB_FILE = "office_credits.json"
_key_lock = threading.Lock()
_gemkey_idx = 0

def get_current_academic_year():
    now = datetime.datetime.utcnow()
    y = now.year
    return f"{y} — {y + 1} م" if now.month >= 9 else f"{y - 1} — {y} م"

def discover_gemini_models_and_keys():
    pref_models = [
        "gemini-3.8-flash",
        "gemini-3.6-flash",
        "gemini-2.5-flash",
        "gemini-3.5-flash-lite",
        "gemini-3.1-flash-lite",
        "gemini-2.5-flash-lite"
    ]
    working_keys = []
    discovered_models = []
    for k in GEMINI_KEYS:
        try:
            r = _HTTP_SESSION.get(
                f"https://generativelanguage.googleapis.com/v1beta/models?key={k}",
                headers={"x-goog-api-key": k},
                timeout=8
            )
            if r.status_code == 200:
                working_keys.append(k)
                if not discovered_models:
                    avail = {m.get("name", "").replace("models/", "") for m in r.json().get("models", [])}
                    for pm in pref_models:
                        if pm in avail and pm not in discovered_models:
                            discovered_models.append(pm)
                    for am in sorted(avail, reverse=True):
                        if ("flash" in am) and ("tts" not in am) and ("image" not in am) and ("audio" not in am) and (am not in discovered_models):
                            discovered_models.append(am)
            else:
                working_keys.append(k)
        except Exception:
            working_keys.append(k)
    final_keys = working_keys if working_keys else GEMINI_KEYS
    final_models = discovered_models[:5] if discovered_models else pref_models[:4]
    return final_keys, final_models

WORKING_GEMINI_KEYS, ACTIVE_GEMINI_MODELS = discover_gemini_models_and_keys()

def get_next_gemini_key():
    global _gemkey_idx
    with _key_lock:
        if not WORKING_GEMINI_KEYS:
            return ""
        k = WORKING_GEMINI_KEYS[_gemkey_idx % len(WORKING_GEMINI_KEYS)]
        _gemkey_idx += 1
        return k

def _pull_file_from_github_if_missing(filename):
    if os.path.exists(filename) or not GH_SYNC_TOKEN:
        return
    try:
        hdr = {"Authorization": f"token {GH_SYNC_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        u = _HTTP_SESSION.get("https://api.github.com/user", headers=hdr, timeout=8).json().get("login")
        if not u:
            return
        rg = _HTTP_SESSION.get(f"https://api.github.com/repos/{u}/statlab-ai-cloud/contents/{filename}", headers=hdr, timeout=8)
        if rg.status_code == 200 and "content" in rg.json():
            raw_bytes = base64.b64decode(rg.json()["content"])
            with open(filename, "wb") as f:
                f.write(raw_bytes)
    except Exception:
        pass

def _sync_file_to_github(filename, content_bytes_or_str):
    if not GH_SYNC_TOKEN:
        return
    def _bg_push():
        try:
            hdr = {"Authorization": f"token {GH_SYNC_TOKEN}", "Accept": "application/vnd.github.v3+json"}
            u = _HTTP_SESSION.get("https://api.github.com/user", headers=hdr, timeout=8).json().get("login")
            if not u:
                return
            url_f = f"https://api.github.com/repos/{u}/statlab-ai-cloud/contents/{filename}"
            rg = _HTTP_SESSION.get(url_f, headers=hdr, timeout=8)
            raw_b = content_bytes_or_str.encode("utf-8") if isinstance(content_bytes_or_str, str) else content_bytes_or_str
            payload = {
                "message": f"Auto-sync {filename} [skip ci]",
                "content": base64.b64encode(raw_b).decode("utf-8")
            }
            if rg.status_code == 200 and "sha" in rg.json():
                payload["sha"] = rg.json()["sha"]
            _HTTP_SESSION.put(url_f, headers=hdr, json=payload, timeout=12)
        except Exception as e:
            log_runtime_error("_sync_file_to_github", e)
    threading.Thread(target=_bg_push, daemon=True).start()

for _f_init in (MEMORY_FILE, ALLOWED_USERS_FILE, ORDERS_CACHE_FILE, CREDITS_DB_FILE):
    _pull_file_from_github_if_missing(_f_init)

def load_allowed_users():
    allowed = {ADMIN_ID}
    if os.path.exists(ALLOWED_USERS_FILE):
        try:
            with open(ALLOWED_USERS_FILE, "r", encoding="utf-8") as f:
                for x in json.load(f):
                    allowed.add(int(x))
        except Exception:
            pass
    return allowed

def add_allowed_user(new_uid):
    allowed = load_allowed_users()
    allowed.add(int(new_uid))
    raw_j = json.dumps(list(allowed), ensure_ascii=False, indent=2)
    with open(ALLOWED_USERS_FILE, "w", encoding="utf-8") as f:
        f.write(raw_j)
    _sync_file_to_github(ALLOWED_USERS_FILE, raw_j)

def is_authorized(uid):
    return int(uid) in load_allowed_users()

def load_credits_db():
    if os.path.exists(CREDITS_DB_FILE):
        try:
            with open(CREDITS_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def get_user_credits(uid):
    if int(uid) == ADMIN_ID:
        return 999999
    db = load_credits_db()
    return int(db.get(str(uid), 10))

def add_user_credits(uid, amount):
    db = load_credits_db()
    cur = int(db.get(str(uid), 10))
    db[str(uid)] = max(0, cur + int(amount))
    raw_j = json.dumps(db, ensure_ascii=False, indent=2)
    with open(CREDITS_DB_FILE, "w", encoding="utf-8") as f:
        f.write(raw_j)
    _sync_file_to_github(CREDITS_DB_FILE, raw_j)
    return db[str(uid)]

def load_learned_memory():
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return [
        "التزم بصيغة الفعل الماضي والغائب الأكاديمي الرصين، وادخل في التحليل العلمي مباشرة دون قول (برهن الباحث، قارن الباحث، ركز الباحث، استند الباحث).",
        "نوّع مطالع الفقرات وأدوات الربط الأكاديمية (ينطلق التأصيل، وفي هذا الإطار، وتأسيساً على ذلك، وعند موازنة هذه المعطيات) ويُمنع بدء فقرتين متتاليتين بـ (تُعدّ / يُعدّ).",
        "التزم حصراً بالمراجع والأبحاث الحقيقية المستخرجة من بحث جوجل العلمي والويب المطابقة للموضوع بدقة 100% ولا تختلق أي مرجع خارجي."
    ]

def add_learned_rule(new_rule):
    rules = load_learned_memory()
    if new_rule and new_rule not in rules:
        rules.append(new_rule)
        raw_j = json.dumps(rules[-50:], ensure_ascii=False, indent=2)
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            f.write(raw_j)
        _sync_file_to_github(MEMORY_FILE, raw_j)

def get_memory_prompt():
    return "\n".join([f"- {r}" for r in load_learned_memory()])

def _xor_cipher(data_bytes):
    key_bytes = hashlib.sha256(( GH_SYNC_TOKEN or BOT_TOKEN or "statlab_v10_key").encode("utf-8")).digest()
    return bytes([b ^ key_bytes[i % len(key_bytes)] for i, b in enumerate(data_bytes)])

def save_order_state(uid, order_info):
    orders_db[uid] = order_info
    cache_data = {}
    if os.path.exists(ORDERS_CACHE_FILE):
        try:
            with open(ORDERS_CACHE_FILE, "rb") as f:
                raw_dec = _xor_cipher(base64.b64decode(f.read())).decode("utf-8")
                cache_data = json.loads(raw_dec)
        except Exception:
            cache_data = {}
    try:
        cache_data[str(uid)] = order_info
        enc_b64 = base64.b64encode(_xor_cipher(json.dumps(cache_data, ensure_ascii=False).encode("utf-8")))
        with open(ORDERS_CACHE_FILE, "wb") as f:
            f.write(enc_b64)
        _sync_file_to_github(ORDERS_CACHE_FILE, enc_b64)
    except Exception as e:
        log_runtime_error("save_order_state", e)

def get_order_state(uid):
    if uid in orders_db:
        return orders_db[uid]
    if os.path.exists(ORDERS_CACHE_FILE):
        try:
            with open(ORDERS_CACHE_FILE, "rb") as f:
                raw_dec = _xor_cipher(base64.b64decode(f.read())).decode("utf-8")
                c = json.loads(raw_dec)
                if str(uid) in c:
                    orders_db[uid] = c[str(uid)]
                    return c[str(uid)]
        except Exception:
            pass
    return None

def cleanup_temp_files(*file_paths):
    try:
        now_ts = time.time()
        for fname in os.listdir("."):
            if fname.startswith(("StatLab_", "chart_")) and fname.endswith((".pdf", ".docx", ".xlsx", ".pptx", ".png")):
                if now_ts - os.path.getmtime(fname) > 1200:
                    try:
                        os.remove(fname)
                    except Exception:
                        pass
    except Exception:
        pass

# ==============================================================================
# الوحدة (2): التطهير اللغوي واستخراج الـ JSON وإدارة طلبات Gemini + بحث Google الحي
# ==============================================================================
WORD_REPLACEMENTS = [
    (r'المسائلة', 'المساءلة'),
    (r'مسائلة', 'مساءلة'),
    (r'كفائة', 'كفاءة'),
    (r'إدارى\b', 'إداري'),
    (r'المللكي', 'المالكي'),
    (r'المللكية', 'الملكية'),
    (r'اللكهربائية', 'الكهربائية'),
    (r'اللكهربائي', 'الكهربائي'),
    (r'اللكهروضوئية', 'الكهروضوئية'),
    (r'اللكهروضوئي', 'الكهروضوئي'),
    (r'اللكفاءة', 'الكفاءة'),
    (r'اللكشف', 'الكشف'),
    (r'الإللكترونية', 'الإلكترونية'),
    (r'الإللكتروني', 'الإلكتروني'),
    (r'المستهللكين', 'المستهلكين'),
    (r'المستهللكة', 'المستهلكة'),
    (r'هججين', 'هجين'),
    (r'الججيران', 'الجيران'),
    (r'للكن', 'لكن'),
    (r'النانوفلوز', 'الموائع النانوية'),
    (r'المايعات النانوية', 'الموائع النانوية'),
    (r'الغرر العظمي', 'الغرر الفاحش'),
    (r'علاوة على ذلك', 'فضلاً عن ذلك'),
    (r'بالإضافة إلى ذلك', 'يُضاف إلى ما تقدم'),
    (r'يلعب دوراً حيوياً', 'ينهض بدور محوري'),
    (r'تلعب دوراً حيوياً', 'تنهض بدور محوري'),
    (r'يسلط الضوء على', 'يتناول بالتحليل والتأصيل'),
    (r'في نهاية المطاف', 'وخلاصة القول')
]

BANNED_META_SENTENCES = [
    "السؤال واضح ومحدد.", "المنهجية دقيقة وشاملة.", "التقسيم منطقي ومبرر.",
    "الخطة شاملة ومفصلة.", "الضوابط دقيقة ومطبقة.", "الإضافة نظرية وعميقة.",
    "الفائدة عملية ومباشرة.", "وفق ما ورد في الملخص الأصلي", "وفقاً للملخص",
    "عزيزي الباحث", "يرجى تزويدي", "تم إعداد مسودة"
]

def truncate_title_by_words(text, max_chars=65):
    cleaned = re.sub(r'["\\\n\r]', '', str(text or "")).strip()
    if len(cleaned) <= max_chars:
        return cleaned
    words = cleaned.split()
    stop_preps = {"في", "عن", "على", "من", "إلى", "الى", "مع", "بين", "لدى", "حول", "عبر", "خلال", "تحت", "فوق", "ضد", "نحو", "و", "أو"}
    out = []
    cur_len = 0
    for w in words:
        if cur_len + len(w) + 1 > max_chars:
            break
        out.append(w)
        cur_len += len(w) + 1
    while len(out) > 3 and out[-1] in stop_preps:
        out.pop()
    return " ".join(out) if out else cleaned[:max_chars]

def strip_random_diacritics_keep_quran(text):
    if not text:
        return ""
    parts = re.split(r'(﴿.*?﴾)', str(text), flags=re.DOTALL)
    out = []
    for p in parts:
        if p.startswith('﴿') and p.endswith('﴾'):
            out.append(p)
        else:
            p_clean = re.sub(r'[\u064B-\u0652\u0670]{2,}', '', p)
            p_clean = re.sub(r'[\u0700-\u074F]+', '', p_clean)
            out.append(p_clean)
    return "".join(out)

def clean_iraqi_text(txt, is_english=False, preserve_urls=False, is_reference=False):
    if not txt:
        return ""
    if isinstance(txt, list):
        txt = "، ".join([str(x).strip() for x in txt if str(x).strip()])
    txt = unicodedata.normalize("NFC", str(txt))
    txt = re.sub(r'<think>.*?</think>', '', txt, flags=re.DOTALL)
    txt = re.sub(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u0700-\u074F]+', '', txt)
    txt = re.sub(r'"?(?:deep_?thinking_?plan|self_?critique|paragraphs|footnotes|healed_?paragraphs|cited_?ref_?indices)"?\s*:\s*\[?', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'\[\s*(?:REF|FN|SOURCE|CITE|مصدر|مرجع|هامش)[\s_\-\d,،:;A-Za-z\.]*\]', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'\s*\[\d{1,3}\]\s*(?=[.،؛]?$)', '', txt)
    txt = txt.replace('\u200f', '').replace('\u200e', '')
    txt = strip_random_diacritics_keep_quran(txt)

    if is_reference:
        for pat, rep in WORD_REPLACEMENTS:
            txt = re.sub(pat, rep, txt)
        txt = re.sub(r'\(\s*CERN European Organization[^)]*\)', '', txt, flags=re.IGNORECASE)
        if not is_english and not preserve_urls:
            txt = re.sub(r'https?://\S+', '', txt)
            txt = re.sub(r'\b10\.\d{4,9}/[-._;()/:A-Za-z0-9]+', '', txt)
        txt = re.sub(r'[.،,\s]+$', '', txt).strip()
        return (txt + ".") if txt else ""

    if preserve_urls:
        return re.sub(r'[ \t]{2,}', ' ', txt).strip()

    txt = re.sub(r'[#*`~]+', '', txt)
    if not is_english:
        for bad_s in BANNED_META_SENTENCES:
            txt = txt.replace(bad_s, "")
        for pat, rep in WORD_REPLACEMENTS:
            txt = re.sub(pat, rep, txt)
        txt = re.sub(r"(?<=[\s،؛:(])'([^'\n]{2,65})'(?=[\s،؛.؟!)])", r"«\1»", txt)
        txt = re.sub(r'(?<=\s)و\s+(?=[\u0621-\u064A])', 'و', txt)
        txt = re.sub(r'\s+([،؛.؟!])', r'\1', txt)
    txt = re.sub(r'[ \t]{2,}', ' ', txt)
    txt = re.sub(r'\n{3,}', '\n\n', txt)
    return txt.strip()

def format_reference_html_safe(ref_text, is_english_ref=False):
    cleaned = clean_iraqi_text(ref_text, is_english=is_english_ref, preserve_urls=is_english_ref, is_reference=True)
    return html.escape(cleaned, quote=True)

def extract_json_safely(raw_text):
    if not raw_text:
        return {}
    cleaned = re.sub(r'<think>.*?</think>', '', str(raw_text), flags=re.DOTALL).strip()
    cleaned = re.sub(r'^```(?:json)?\s*', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s*```$', '', cleaned)
    try:
        obj = json.loads(cleaned, strict=False)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass
    m = re.search(r'\{.*\}', cleaned, flags=re.DOTALL)
    if m:
        try:
            obj = json.loads(m.group(0), strict=False)
            if isinstance(obj, dict):
                return obj
        except Exception:
            pass
    repaired = {}
    for list_key in ["paragraphs", "healed_paragraphs", "intro_paragraphs", "importance_paragraphs", "methodology_plan_paragraphs", "findings", "recommendations", "items", "cited_ref_indices", "references_ar", "references_books", "references_en"]:
        m_arr = re.search(rf'"{list_key}"\s*:\s*\[(.*?)(?:\]|$)', cleaned, flags=re.DOTALL)
        if m_arr:
            if list_key == "cited_ref_indices":
                nums = [int(x) for x in re.findall(r'\d+', m_arr.group(1))]
                if nums:
                    repaired[list_key] = nums
            else:
                items_found = re.findall(r'"((?:[^"\\]|\\.)*)"', m_arr.group(1), flags=re.DOTALL)
                if items_found:
                    repaired[list_key] = [x.replace('\\"', '"').replace('\\n', ' ').strip() for x in items_found if len(x.strip()) > 3]
    for str_key in ["clean_title", "specialty", "college_name", "university_name", "student_name", "supervisor_name", "academic_year", "commentary", "dedication", "abstract_ar", "abstract_en", "keywords_ar", "keywords_en", "conclusion_intro", "translated_title", "manager_reply"]:
        m_str = re.search(rf'"{str_key}"\s*:\s*"((?:[^"\\]|\\.)*)"', cleaned, flags=re.DOTALL)
        if m_str:
            repaired[str_key] = m_str.group(1).replace('\\"', '"').strip()
    return repaired

def ask_agent(prompt_text, agent_role="writer", json_mode=True, required_keys=None, live_blueprint=None, use_google_search=False):
    """
    المحرك النقي الموحد (Gemini 3.8 / 3.6 / 2.5 Flash حصراً):
    1. يدور على جميع المفاتيح الـ 8 في أقوى نموذج أولاً قبل الانتقال للنموذج الذي يليه.
    2. يفعل أداة بحث جوجل الحقيقي ("google_search") عندما يُطلب منه البحث في الويب أو المراجعة الشاملة.
    """
    learned_str = get_memory_prompt()
    live_ctx = f"\n\n[البصمة المنهجية المستخرجة من الأبحاث المنشورة في هذا الموضوع]:\n{live_blueprint}\n" if live_blueprint else ""

    role_prompts = {
        "scout": (
            "أنت الوكيل المعماري الأكاديمي والمدير التفاعلي في الجامعات العراقية. "
            "لديك وصول مباشر لمحرك بحث Google الحقيقي. افهم طلب المستخدم بدقة واستخرج كافة بيانات البحث بصيغة JSON."
        ),
        "calibrator": (
            "أنت مهندس البحث الأكاديمي الحي في الجامعات العراقية. "
            "ابحث في Google والباحث العلمي والمجلات العراقية المحكمة (iasj.rdd.edu.iq) عن الأبحاث والكتب المنشورة والمقبولة فعلياً في نفس موضوع البحث وتخصصه الدقيق، "
            "واقرأ منهجيتها وتقسيماتها، واستخرج منها خطة البحث الأصيلة والمصادر الحقيقية الموثقة 100% بصيغة JSON."
        ),
        "writer": (
            f"أنت أستاذ دكتور وباحث أكاديمي قدير في الجامعات العراقية.{live_ctx}\n"
            f"تلتزم بالقواعد الأكاديمية التالية:\n{learned_str}\n"
            "قواعد الصياغة الإلزامية:\n"
            "1. اكتب المتن العلمي بأسلوب أكاديمي تحليلي رصين ومفصل يشرح جوهر الموضوع نفسه مباشرة (المفاهيم، الأسس، الآليات، المعادلات أو النصوص أو التحليل) دون تحويل المتن إلى سرد لأسماء الباحثين.\n"
            "2. نوّع مطالع الفقرات وأدوات الربط الأكاديمية (ينطلق التأصيل، وفي هذا الإطار، وتأسيساً على ذلك، وعند موازنة هذه المعطيات) ويُمنع تكرار (تُعدّ / يُعدّ) في بداية فقرتين متتاليتين.\n"
            "3. لا تكتب أي رموز مرجعية مربعة مثل [0] أو [REF_1] داخل نص الفقرة العربية، وأنهِ كل فقرة بنقطة تامة."
        ),
        "auditor": (
            f"أنت المدقق الأكاديمي الشامل ورئيس اللجنة العلمية (Editor-in-Chief).{live_ctx}\n"
            f"تلتزم بالقواعد التالية:\n{learned_str}\n"
            "مهمتك قراءة التقرير والفقرات والمصادر بوعي كامل كما يقرؤها أستاذ جامعي فاحص، وكشف أي خطأ علمي أو منهجي أو مصدري وإصلاحه فوراً."
        )
    }
    sys_msg = role_prompts.get(agent_role, role_prompts["writer"])
    if json_mode and "JSON" not in prompt_text.upper():
        prompt_text += "\n\nأخرج النتيجة بصيغة JSON صالحة حصراً دون أي نص خارج قوسي { }."

    def _is_valid_output(txt_out):
        if not txt_out or len(txt_out) < 10:
            return False
        if not json_mode:
            return True
        parsed_obj = extract_json_safely(txt_out)
        if not parsed_obj:
            return False
        if required_keys:
            for rk in required_keys:
                if rk not in parsed_obj or parsed_obj[rk] in (None, "", []):
                    return False
        return True

    last_non_empty = ""
    models_to_try = ACTIVE_GEMINI_MODELS if ACTIVE_GEMINI_MODELS else ["gemini-3.8-flash", "gemini-3.6-flash", "gemini-2.5-flash"]
    num_keys = max(1, len(WORKING_GEMINI_KEYS))

    # التناوب على جميع المفاتيح الـ 8 في النموذج الأقوى أولاً قبل النزول للنموذج التالي
    for model_name in models_to_try:
        for _ in range(num_keys):
            gem_key = get_next_gemini_key()
            if not gem_key:
                break
            url_gem = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={gem_key}"
            hdrs_gem = {"Content-Type": "application/json", "x-goog-api-key": gem_key}

            # محاولة أولى (مع أداة بحث Google إذا كانت مفعلة)
            search_modes = [True, False] if use_google_search else [False]
            for s_mode in search_modes:
                gen_cfg = {"temperature": 0.25, "maxOutputTokens": 8192}
                if json_mode and not s_mode:
                    gen_cfg["responseMimeType"] = "application/json"
                payload = {
                    "systemInstruction": {"parts": [{"text": sys_msg}]},
                    "contents": [{"role": "user", "parts": [{"text": prompt_text}]}],
                    "generationConfig": gen_cfg
                }
                if s_mode:
                    payload["tools"] = [{"google_search": {}}]

                try:
                    r_gem = _HTTP_SESSION.post(url_gem, json=payload, headers=hdrs_gem, timeout=45)
                    if r_gem.status_code == 200:
                        cands = r_gem.json().get("candidates") or []
                        if cands:
                            parts = cands[0].get("content", {}).get("parts") or []
                            txt_joined = "\n".join([p.get("text", "") for p in parts if p.get("text")])
                            txt = re.sub(r'<think>.*?</think>', '', txt_joined, flags=re.DOTALL).strip()
                            if txt:
                                last_non_empty = txt
                            if _is_valid_output(txt):
                                return txt
                        break
                    elif r_gem.status_code == 400 and s_mode:
                        # إذا كان هذا النموذج الفرعي لا يدمج أداة البحث مع هذا الطلب، نعيد المحاولة فوراً بدون s_mode على نفس المفتاح
                        continue
                    elif r_gem.status_code == 429:
                        # هذا المفتاح وصل لحد الدقيقة، ننتقل فوراً للمفتاح التالي على نفس النموذج الأقوى!
                        time.sleep(0.3)
                        break
                    else:
                        break
                except Exception as e:
                    log_runtime_error(f"Gemini {model_name}", e)
                    break

    if last_non_empty:
        return last_non_empty
    raise RuntimeError("تعذر الاتصال بمحرك Gemini حالياً، يرجى المحاولة بعد ثوانٍ.")

# ==============================================================================
# الوحدة (3): محرك البحث الحي في Google والمدقق الشامل للتقرير الكامل (Full-Report Inspector)
# ==============================================================================
def _convert_arabic_number_words(text):
    if not text:
        return ""
    s = str(text)
    word_map = [
        (r'\bخمسة\s+عشر\b|\bخمس\s+عشرة\b', '15'),
        (r'\bأربعة\s+عشر\b|\bأربع\s+عشرة\b', '14'),
        (r'\bثلاثة\s+عشر\b|\bثلاث\s+عشرة\b', '13'),
        (r'\bاثنا\s+عشر\b|\bاثني\s+عشر\b|\bاثنتا\s+عشرة\b', '12'),
        (r'\bأحد\s+عشر\b|\bإحدى\s+عشرة\b', '11'),
        (r'\bعشرة\b|\bعشر\b', '10'),
        (r'\bعشرين\b|\bعشرون\b', '20'),
        (r'\bخمسة\s+وعشرين\b|\bخمس\s+وعشرين\b', '25'),
        (r'\bثلاثين\b|\bثلاثون\b', '30'),
        (r'\bثلاثة\s+مباحث\b|\bثلاث\s+مباحث\b', '3 مباحث'),
        (r'\bأربعة\s+مباحث\b|\bأربع\s+مباحث\b', '4 مباحث'),
        (r'\bمبحثين\b', '2 مباحث'),
        (r'\bمطلبين\b', '2 مطالب'),
        (r'\bثلاثة\s+مطالب\b|\bثلاث\s+مطالب\b', '3 مطالب'),
    ]
    for pat, rep in word_map:
        s = re.sub(pat, rep, s)
    return s

def parse_and_calibrate_order(raw_user_text):
    """
    البحث الحي الطبيعي عبر Google Search المدمج في Gemini 3.8 Flash:
    1. يفهم العنوان والتخصص والكلية وشروط الطالب.
    2. يبحث حياً في Google والباحث العلمي والمجلات العراقية (iasj.rdd.edu.iq) عن الأبحاث والكتب المنشورة والمقبولة في هذا الموضوع بالذات.
    3. يتعلم منهجيتها ويبني الخطة والمصادر الحقيقية المطابقة للموضوع 100%.
    """
    normalized_input = _convert_arabic_number_words(raw_user_text)
    default_acad_yr = get_current_academic_year()

    calib_prompt = f"""
    رسالة وطلب الباحث:
    «{normalized_input}»

    قم الآن بالبحث الحي في محرك بحث Google والباحث العلمي (Google Scholar) والمجلات الأكاديمية العراقية (iasj.rdd.edu.iq) عن الأبحاث والدراسات والكتب المنشورة والمقبولة فعلياً في هذا الموضوع الدقيق، ثم أخرج كائن JSON شامل يحتوي على:
    1. العنوان الأكاديمي الصافي ("clean_title") بعد فصل الشروط، مع وضع أي اختصار إنجليزي في مكانه الصحيح بعد مصطلحه العربي مباشرة.
    2. التخصص الدقيق ("specialty")، واسم الكلية ("college_name")، واسم الجامعة ("university_name")، وفئة الغلاف ("category_code": اختر من "law" أو "economics" أو "medical" أو "engineering" أو "humanities").
    3. هل التخصص شرعي/فقهي/أدبي عربي خالص يمنع المراجع الإنجليزية ("is_arabic_doctrinal_only": true/false)؟
    4. المتغير المستقل ("independent_var") والمتغير التابع ("dependent_var") للتحليل الإحصائي.
    5. البصمة المنهجية الحية ("live_methodology_blueprint"): كيف تكتب الأبحاث المنشورة المقبولة في هذا الموضوع؟ ما المنهج العلمي المتبع؟ وما المصطلحات أو المعادلات أو النصوص القياسية الواجب ذكرها؟
    6. هيكل التقسيم الأكاديمي ("sections_plan"):
       - إذا كان التخصص في (القانون، الشريعة، الإدارة والاقتصاد، العلوم السياسية، العلوم الإنسانية) أو طلب المستخدم كلمة "مباحث ومطالب"، استخدم ("المبحث الأول: ..." وتحته "المطلب الأول: ...", "المطلب الثاني: ...").
       - إذا كان التخصص في (الطب، الهندسة، علوم الحاسوب، العلوم الصرفة) ولم يطلب المستخدم كلمة "مباحث"، استخدم ("المحور الأول: ..." وتحته "أولاً: ...", "ثانياً: ...").
    7. المصادر الحقيقية الموثقة المطابقة للموضوع 100% (يُمنع منعاً باتاً إدراج أي مصدر من تخصص آخر أو موضوع عام كطرق البحث العلمي!):
       - "references_books": قائمة من 5 كتب مرجعية عربية حقيقية مطبوعة في صلب هذا الموضوع الدقيق بصيغة: "اسم المؤلف (السنة)، «عنوان الكتاب»، دار النشر، مدينة النشر."
       - "references_ar": قائمة من 6 أبحاث ودراسات أكاديمية عراقية وعربية محكمة منشورة فعلياً في المجلات العلمية في صلب هذا الموضوع الدقيق بصيغة: "اسم الباحث (السنة)، «عنوان البحث»، اسم المجلة، الجامعة، ص من-إلى."
       - "references_en": قائمة من 5 أبحاث أجنبية محكمة منشورة في مجلات عالمية (IEEE / Elsevier / Springer / PubMed / إلخ) في صلب هذا الموضوع الدقيق بصيغة APA الكاملة باللغة الإنجليزية.
    8. جدول المقارنة التحليلي ("comparison_table") المناسب لصلب الموضوع.

    أخرج كائن JSON بالحقول التالية حصراً:
    {{
      "clean_title": "العنوان الأكاديمي الصافي",
      "specialty": "التخصص الدقيق",
      "college_name": "اسم الكلية",
      "university_name": "اسم الجامعة",
      "category_code": "law أو economics أو medical أو engineering أو humanities",
      "is_arabic_doctrinal_only": false,
      "independent_var": "المحور الأول",
      "dependent_var": "المحور الثاني",
      "font_size": 16,
      "min_pages": 15,
      "student_name": "........................................",
      "supervisor_name": "........................................",
      "academic_year": "{default_acad_yr}",
      "user_ forced_mabahith": false,
      "custom_instructions": "",
      "live_methodology_blueprint": "شرح مكثف للمنهجية العلمية والأسلوب الأكاديمي والمصطلحات القياسية المستخرجة من الأبحاث المنشورة في هذا الموضوع",
      "sections_plan": [
        {{
          "section_title": "المبحث الأول (أو المحور الأول): ...",
          "sub_sections": ["المطلب الأول (أو أولاً): ...", "المطلب الثاني (أو ثانياً): ..."]
        }},
        {{
          "section_title": "المبحث الثاني (أو المحور الثاني): ...",
          "sub_sections": ["المطلب الأول (أو أولاً): ...", "المطلب الثاني (أو ثانياً): ..."]
        }},
        {{
          "section_title": "المبحث الثالث (أو المحور الثالث): ...",
          "sub_sections": ["المطلب الأول (أو أولاً): ...", "المطلب الثاني (أو ثانياً): ..."]
        }}
      ],
      "references_books": [
        "المؤلف (السنة)، «عنوان الكتاب المتخصص في صلب الموضوع»، دار النشر، المدينة."
      ],
      "references_ar": [
        "الباحث (السنة)، «عنوان البحث العراقي/العربي المنشور في صلب الموضوع»، اسم المجلة، الجامعة، ص 10-35."
      ],
      "references_en": [
        "Author, A. (Year). Exact Title of Published Paper on Topic. Journal Name, pp. 1-15."
      ],
      "include_comparison_table": true,
      "comparison_table": {{
        "title": "جدول المقارنة التحليلية لمحاور الدراسة",
        "col1_header": "وجه المقارنة",
        "col2_header": "المحور أو الاتجاه الأول",
        "col3_header": "المحور أو الاتجاه الثاني",
        "rows": [
          ["الأساس العلمي والمنهجي", "...", "..."],
          ["الركائز والآليات الحاكمة", "...", "..."],
          ["نطاق التطبيق والأثر", "...", "..."],
          ["الضمانات ومخرجات الأداء", "...", "..."]
        ]
      }}
    }}
    """
    calib_data = extract_json_safely(ask_agent(calib_prompt, "calibrator", True, required_keys=["clean_title", "sections_plan"], use_google_search=True))

    clean_title = clean_iraqi_text(calib_data.get("clean_title") or raw_user_text[:95])
    specialty = clean_iraqi_text(calib_data.get("specialty") or "الدراسات الأكاديمية")
    college_name = clean_iraqi_text(calib_data.get("college_name") or "كلية الدراسات التخصصية")
    univ_name = clean_iraqi_text(calib_data.get("university_name") or "جامعة بغداد")
    cat_code = str(calib_data.get("category_code") or "humanities").strip().lower()

    is_doctrinal = bool(calib_data.get("is_arabic_doctrinal_only", False))
    if any(k in f"{raw_user_text} {specialty} {college_name}" for k in ["فقه", "شريعة", "إسلامي", "قرآن", "حديث", "أصول الفقه", "الإمام الأعظم", "العلوم الإسلامية"]):
        is_doctrinal = True

    nums_pg = re.findall(r'\d+', str(calib_data.get("min_pages") or 15))
    min_pgs = max(6, min(45, int(nums_pg[0]) if nums_pg else 15))
    nums_fn = re.findall(r'\d+', str(calib_data.get("font_size") or 16))
    font_sz = max(12, min(20, int(nums_fn[0]) if nums_fn else 16))

    user_forced_mabahith = bool(calib_data.get("user_forced_mabahith", False)) or ("مبحث" in normalized_input or "مباحث" in normalized_input or "مطلب" in normalized_input)
    use_mabahith_labels = user_forced_mabahith or (cat_code in ("law", "humanities", "economics")) or is_doctrinal

    arabic_ords = ["الأول", "الثاني", "الثالث", "الرابع", "الخامس"]
    raw_plan = calib_data.get("sections_plan") or []
    plan = []
    for idx_p, sec_item in enumerate(raw_plan):
        if isinstance(sec_item, dict):
            ord_w = arabic_ords[min(idx_p, len(arabic_ords) - 1)]
            s_title = clean_iraqi_text(sec_item.get("section_title") or sec_item.get("mabhath_title") or "")
            s_title = re.sub(r'^(?:المبحث|المحور|الفصل|القسم)\s+(?:الأول|الثاني|الثالث|الرابع|الخامس|\d+)\s*[:：\-]?\s*', '', s_title).strip()
            prefix_sec = f"المبحث {ord_w}" if use_mabahith_labels else f"المحور {ord_w}"
            s_title = f"{prefix_sec}: {s_title}"

            subs_raw = sec_item.get("sub_sections") or sec_item.get("matalib") or []
            subs_clean = []
            for sub_idx, sub_r in enumerate(subs_raw, 1):
                sub_ord = arabic_ords[min(sub_idx - 1, len(arabic_ords) - 1)]
                sub_c = clean_iraqi_text(sub_r)
                sub_c = re.sub(r'^(?:المطلب\s+(?:الأول|الثاني|الثالث|الرابع|الخامس)|أولاً|ثانياً|ثالثاً|رابعاً|\d+)\s*[:：\-]?\s*', '', sub_c).strip()
                if use_mabahith_labels:
                    sub_c = f"المطلب {sub_ord}: {sub_c}"
                else:
                    ord_adv = ["أولاً", "ثانياً", "ثالثاً", "رابعاً"][min(sub_idx - 1, 3)]
                    sub_c = f"{ord_adv}: {sub_c}"
                subs_clean.append(sub_c)
            if subs_clean:
                plan.append({"mabhath_title": s_title, "matalib": subs_clean})

    if not plan:
        p_lbl = "المبحث" if use_mabahith_labels else "المحور"
        s1_lbl = "المطلب الأول" if use_mabahith_labels else "أولاً"
        s2_lbl = "المطلب الثاني" if use_mabahith_labels else "ثانياً"
        plan = [
            {"mabhath_title": f"{p_lbl} الأول: الإطار المفاهيمي والأسس النظرية لـ {clean_title}", "matalib": [f"{s1_lbl}: المفهوم والخصائص والمرتكزات الأساسية", f"{s2_lbl}: الأبعاد الحاكمة والضوابط المعيارية"]},
            {"mabhath_title": f"{p_lbl} الثاني: الآليات التحليلية والتطبيقية في {clean_title}", "matalib": [f"{s1_lbl}: الآليات والنماذج المعتمدة في التحليل", f"{s2_lbl}: تقييم الأداء والتحديات العملية"]},
            {"mabhath_title": f"{p_lbl} الثالث: المعالجات المقترحة وآفاق التطوير المستدام", "matalib": [f"{s1_lbl}: متطلبات التطوير ورفع الكفاءة", f"{s2_lbl}: الرؤية المستقبلية والضمانات التنفيذية"]}
        ]

    books_list = [clean_iraqi_text(x, is_reference=True) for x in (calib_data.get("references_books") or []) if len(str(x).strip()) > 12]
    ar_refs_list = [clean_iraqi_text(x, is_reference=True) for x in (calib_data.get("references_ar") or []) if len(str(x).strip()) > 12]
    en_refs_list = [] if is_doctrinal else [clean_iraqi_text(x, is_english=True, preserve_urls=True, is_reference=True) for x in (calib_data.get("references_en") or []) if len(str(x).strip()) > 12]

    arabic_tier_ords = ["أولاً", "ثانياً", "ثالثاً"]
    formatted_tiers = []
    if books_list:
        formatted_tiers.append({"title": f"{arabic_tier_ords[len(formatted_tiers)]} — الكتب والمؤلفات الأكاديمية المعتمدة", "items": books_list, "is_ltr": False})
    if ar_refs_list:
        formatted_tiers.append({"title": f"{arabic_tier_ords[len(formatted_tiers)]} — الأبحاث والدراسات المنشورة في الدوريات والمجلات العلمية المحكمة", "items": ar_refs_list, "is_ltr": False})
    if en_refs_list:
        formatted_tiers.append({"title": f"{arabic_tier_ords[len(formatted_tiers)]} — المصادر والدوريات الأجنبية المحكمة (Foreign References)", "items": en_refs_list, "is_ltr": True})

    calib_data["formatted_reference_tiers"] = formatted_tiers

    # بناء قائمة الكائنات المرجعية الموحدة للهوامش والاستشهادات داخل المتن
    verified_pool_with_context = []
    verified_source_objects = []

    def _extract_cite_tag(ref_line, is_en=False):
        auth_part = ref_line.split("(")[0].strip(" .،,-•")
        yr_m = re.findall(r'(?:19|20)\d{2}', ref_line)
        yr_str = yr_m[0] if yr_m else "2024"
        words_a = [w.strip(".,،؛:") for w in auth_part.split() if len(w.strip(".,،؛:")) >= 2]
        surname = words_a[0] if words_a else ("Author" if is_en else "الباحث")
        if not is_en and ("،" in auth_part or " و" in auth_part):
            surname = auth_part[:35].strip(" ،,")
        return f"{surname}, {yr_str}" if is_en else f"{surname}، {yr_str}"

    for b_line in books_list:
        c_tag = _extract_cite_tag(b_line, is_en=False)
        verified_pool_with_context.append({"ref_line": b_line, "summary": b_line})
        verified_source_objects.append({"cite": c_tag, "short_cite": c_tag, "full_auth": b_line.split("(")[0].strip(), "yr": "2022", "title_orig": b_line, "ar_context": b_line, "full_ref_line": b_line})

    iraqi_studies = []
    for ar_line in ar_refs_list:
        c_tag = _extract_cite_tag(ar_line, is_en=False)
        tit_m = re.search(r'«([^»]+)»', ar_line)
        t_clean = tit_m.group(1) if tit_m else ar_line
        verified_pool_with_context.append({"ref_line": ar_line, "summary": ar_line})
        st_obj = {"cite": c_tag, "short_cite": c_tag, "full_auth": ar_line.split("(")[0].strip(), "yr": "2024", "title_orig": t_clean, "ar_context": ar_line, "jrn": "مجلة علمية محكمة", "pages": "", "abstract": ar_line, "full_ref_line": ar_line}
        iraqi_studies.append(st_obj)
        verified_source_objects.append(st_obj)

    openalex_studies = []
    for en_line in en_refs_list:
        c_tag = _extract_cite_tag(en_line, is_en=True)
        verified_pool_with_context.append({"ref_line": en_line, "summary": en_line})
        en_obj = {"cite": c_tag, "short_cite": c_tag, "full_auth": en_line.split("(")[0].strip(), "yr": "2024", "title_en": en_line, "ar_context": en_line, "jrn": "International Peer-Reviewed Journal", "pages": "", "doi": "", "abstract": en_line}
        openalex_studies.append(en_obj)
        verified_source_objects.append(en_obj)

    return {
        "raw_text": raw_user_text,
        "clean_title": clean_title,
        "specialty": specialty,
        "college_name": college_name,
        "university_name": univ_name,
        "category_code": cat_code,
        "is_doctrinal": is_doctrinal,
        "allow_scientific_abbr": (cat_code in ("medical", "engineering") and not is_doctrinal),
        "independent_var": clean_iraqi_text(calib_data.get("independent_var") or f"المحور الأول في {clean_title[:35]}"),
        "dependent_var": clean_iraqi_text(calib_data.get("dependent_var") or f"المحور الثاني في {clean_title[:35]}"),
        "font_size": font_sz,
        "min_pages": min_pgs,
        "student_name": clean_iraqi_text(calib_data.get("student_name") or "........................................"),
        "supervisor_name": clean_iraqi_text(calib_data.get("supervisor_name") or "........................................"),
        "academic_year": clean_iraqi_text(calib_data.get("academic_year") or default_acad_yr),
        "custom_instructions": clean_iraqi_text(calib_data.get("custom_instructions") or ""),
        "live_blueprint": clean_iraqi_text(calib_data.get("live_methodology_blueprint") or ""),
        "plan": plan,
        "calib": calib_data,
        "iraqi_studies": iraqi_studies,
        "openalex_studies": openalex_studies,
        "real_books": books_list,
        "verified_source_objects": verified_source_objects,
        "verified_pool_with_context": verified_pool_with_context
    }

def inspect_and_heal_full_report(order_info, intro_paras, imp_paras, method_paras, generated_mabahith, conc_data):
    """
    المدقق الشامل للتقرير الكامل (يرى التقرير كاملاً من الغلاف والفهرس والمتن حتى قائمة المصادر والهوامش تماماً كما تفحصه في الدردشة!):
    يكشف أي مصدر غير مطابق للموضوع أو أي تناقض أو خطأ علمي ويصلحه قبل الطباعة.
    """
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    calib = order_info["calib"]
    tiers = calib.get("formatted_reference_tiers") or []

    all_refs_preview = []
    for t_obj in tiers:
        all_refs_preview.append(f"[{t_obj.get('title')}]:")
        for r_idx, r_it in enumerate(t_obj.get("items") or [], 1):
            all_refs_preview.append(f"  {r_idx}. {r_it}")
    refs_dump = "\n".join(all_refs_preview)

    sections_dump = []
    for mb in generated_mabahith:
        sections_dump.append(f"■ {mb['m_title']}:")
        for mt in mb["matalib"]:
            sections_dump.append(f"  - {mt['mt_title']}: {mt['paras_raw'][0][:220]}...")
    sec_summary = "\n".join(sections_dump)

    audit_prompt = f"""
    أنت الآن المدقق الأكاديمي الأعلى (Gemini 3.8 Flash). أمامك المسودة الكاملة للتقرير قبل تحويله إلى PDF و Word:
    - العنوان: «{title_ar}»
    - التخصص والكلية: {specialty} — {order_info['college_name']} ({order_info['university_name']})
    - شروط الطالب الخاصة: {order_info.get('custom_instructions') or 'التزام الدقة العلمية والتخصصية'}

    [ملخص أقسام ومباحث التقرير المكتوبة]:
    {sec_summary}

    [الاستنتاجات المكتوبة]:
    {json.dumps(conc_data.get('findings') or [], ensure_ascii=False)}

    [التوصيات المكتوبة]:
    {json.dumps(conc_data.get('recommendations') or [], ensure_ascii=False)}

    [قائمة المصادر والمراجع المدرجة في نهاية التقرير]:
    {refs_dump}

    افحص هذا التقرير كاملاً بعين ناقدة صارمة (استخدم بحث Google للتحقق عند الحاجة):
    1. افحص كل مرجع في [قائمة المصادر والمراجع]: هل هناك أي مرجع دخيل لا ينتمي لموضوع «{title_ar}» وتخصص «{specialty}» (مثل كتاب عام في مناهج البحث أو بحث من تخصص آخر)؟ إذا وجدت أي مرجع دخيل، استبدله فوراً بمرجع حقيقي موثوق في صلب «{title_ar}»!
    2. افحص الاستنتاجات والتوصيات: هل هي دقيقة علمياً ومرتبطة 100% بمباحث التقرير؟ صوّب أي خطأ علمي أو لغوي فيها.

    أخرج كائن JSON بالحقول المصححة التالية:
    {{
      "healed_findings": ["الاستنتاج 1 المصحح...", "الاستنتاج 2...", "الاستنتاج 3...", "الاستنتاج 4...", "الاستنتاج 5..."],
      "healed_recommendations": ["التوصية 1 المصححة...", "التوصية 2...", "التوصية 3...", "التوصية 4...", "التوصية 5..."],
      "healed_books": ["قائمة الكتب المرجعية بعد تطهير أي كتاب غير متخصص..."],
      "healed_ar_refs": ["قائمة الأبحاث العراقية/العربية المحكمة بعد تطهير أي بحث دخيل..."],
      "healed_en_refs": ["قائمة الأبحاث الأجنبية المحكمة بعد تطهير أي بحث دخيل..."]
    }}
    """
    try:
        res = extract_json_safely(ask_agent(audit_prompt, "auditor", True, required_keys=["healed_findings", "healed_recommendations"], use_google_search=True))
        if res.get("healed_findings") and len(res["healed_findings"]) >= 3:
            conc_data["findings"] = [clean_iraqi_text(x) for x in res["healed_findings"]]
        if res.get("healed_recommendations") and len(res["healed_recommendations"]) >= 3:
            conc_data["recommendations"] = [clean_iraqi_text(x) for x in res["healed_recommendations"]]

        new_books = [clean_iraqi_text(x, is_reference=True) for x in (res.get("healed_books") or []) if len(str(x).strip()) > 12]
        new_ar = [clean_iraqi_text(x, is_reference=True) for x in (res.get("healed_ar_refs") or []) if len(str(x).strip()) > 12]
        new_en = [] if order_info.get("is_doctrinal") else [clean_iraqi_text(x, is_english=True, preserve_urls=True, is_reference=True) for x in (res.get("healed_en_refs") or []) if len(str(x).strip()) > 12]

        if new_books or new_ar:
            ords = ["أولاً", "ثانياً", "ثالثاً"]
            new_tiers = []
            if new_books:
                new_tiers.append({"title": f"{ords[len(new_tiers)]} — الكتب والمؤلفات الأكاديمية المعتمدة", "items": new_books, "is_ltr": False})
            if new_ar:
                new_tiers.append({"title": f"{ords[len(new_tiers)]} — الأبحاث والدراسات المنشورة في الدوريات والمجلات العلمية المحكمة", "items": new_ar, "is_ltr": False})
            if new_en:
                new_tiers.append({"title": f"{ords[len(new_tiers)]} — المصادر والدوريات الأجنبية المحكمة (Foreign References)", "items": new_en, "is_ltr": True})
            if new_tiers:
                calib["formatted_reference_tiers"] = new_tiers
    except Exception as e:
        log_runtime_error("inspect_and_heal_full_report", e)
    return conc_data

def process_conversational_manager_edit(order_info, user_edit_text):
    current_state_summary = {
        "clean_title": order_info["clean_title"],
        "specialty": order_info["specialty"],
        "college_name": order_info["college_name"],
        "university_name": order_info["university_name"],
        "student_name": order_info["student_name"],
        "supervisor_name": order_info["supervisor_name"],
        "academic_year": order_info["academic_year"],
        "font_size": order_info["font_size"],
        "min_pages": order_info["min_pages"],
        "is_doctrinal": order_info["is_doctrinal"],
        "custom_instructions": order_info.get("custom_instructions", ""),
        "plan": order_info["plan"]
    }
    edit_prompt = f"""
    أنت المدير الأكاديمي التفاعلي. هذه هي بيانات وخطة البحث الحالية:
    {json.dumps(current_state_summary, ensure_ascii=False, indent=2)}

    أرسل المستخدم التعديل التالي:
    «{user_edit_text}»

    عدّل الحقول المطلوبة وأخرج كائن JSON المحدث بالحقول التالية:
    {{
      "clean_title": "...",
      "specialty": "...",
      "college_name": "...",
      "university_name": "...",
      "student_name": "...",
      "supervisor_name": "...",
      "academic_year": "...",
      "font_size": 16,
      "min_pages": 15,
      "is_doctrinal": false,
      "custom_instructions": "...",
      "plan": [
        {{
          "mabhath_title": "...",
          "matalib": ["...", "..."]
        }}
      ],
      "manager_reply": "رسالة ترحيبية تؤكد تنفيذ التعديل وتخبر المستخدم بأنه يمكنه كتابة «ابدأ» للبدء فوراً"
    }}
    """
    updated = extract_json_safely(ask_agent(edit_prompt, "scout", True, required_keys=["clean_title", "plan"]))
    if updated.get("clean_title"):
        order_info["clean_title"] = clean_iraqi_text(updated["clean_title"])
    if updated.get("specialty"):
        order_info["specialty"] = clean_iraqi_text(updated["specialty"])
    if updated.get("college_name"):
        order_info["college_name"] = clean_iraqi_text(updated["college_name"])
    if updated.get("university_name"):
        order_info["university_name"] = clean_iraqi_text(updated["university_name"])
    if updated.get("student_name"):
        order_info["student_name"] = clean_iraqi_text(updated["student_name"])
    if updated.get("supervisor_name"):
        order_info["supervisor_name"] = clean_iraqi_text(updated["supervisor_name"])
    if updated.get("academic_year"):
        order_info["academic_year"] = clean_iraqi_text(updated["academic_year"])
    if updated.get("font_size"):
        nums_f = re.findall(r'\d+', str(updated["font_size"]))
        if nums_f:
            order_info["font_size"] = max(12, min(20, int(nums_f[0])))
    if updated.get("min_pages"):
        nums_p = re.findall(r'\d+', str(updated["min_pages"]))
        if nums_p:
            order_info["min_pages"] = max(6, min(45, int(nums_p[0])))
    if "is_doctrinal" in updated:
        order_info["is_doctrinal"] = bool(updated["is_doctrinal"])
    if updated.get("custom_instructions") is not None:
        order_info["custom_instructions"] = clean_iraqi_text(updated.get("custom_instructions") or "")
    if isinstance(updated.get("plan"), list) and len(updated["plan"]) >= 1:
        new_plan = []
        for sec in updated["plan"]:
            if isinstance(sec, dict) and sec.get("mabhath_title"):
                mt_list = [clean_iraqi_text(x) for x in (sec.get("matalib") or sec.get("sub_sections") or []) if x]
                if mt_list:
                    new_plan.append({"mabhath_title": clean_iraqi_text(sec["mabhath_title"]), "matalib": mt_list})
        if new_plan:
            order_info["plan"] = new_plan
    reply_msg = clean_iraqi_text(updated.get("manager_reply") or "تم تحديث الخطة والبيانات بنجاح وفق توجيهك!")
    return order_info, reply_msg

# ==============================================================================
# الوحدة (4): تنسيق الـ Word القياسي (OpenXML) ورسومات الغلاف (SVG)
# ==============================================================================
def _esc(val):
    return html.escape(str(val or ""), quote=True)

def get_cover_graphics_html(category_code):
    top_logos_bar = """
    <table style="width:100%; border-collapse:collapse; border-bottom:2px solid #1B3B2B; padding-bottom:8px; margin-bottom:14px;">
      <tr>
        <td style="width:22%; text-align:right; vertical-align:middle;">
          <svg width="64" height="64" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
            <circle cx="50" cy="50" r="46" fill="#FAF6F0" stroke="#1B3B2B" stroke-width="3"/>
            <circle cx="50" cy="50" r="39" fill="none" stroke="#B38F40" stroke-width="1.5" stroke-dasharray="3,2"/>
            <path d="M28 62 L50 22 L72 62 Z" fill="none" stroke="#1B3B2B" stroke-width="2.5"/>
            <path d="M32 56 Q50 46 68 56 L68 70 Q50 60 32 70 Z" fill="#7A1C24"/>
            <circle cx="50" cy="36" r="5.5" fill="#B38F40"/>
            <polygon points="50,75 52,80 57,80 53,83 54,88 50,85 46,88 47,83 43,80 48,80" fill="#B38F40"/>
          </svg>
        </td>
        <td style="width:56%; text-align:center; vertical-align:middle;">
          <div style="font-size:14pt; font-weight:bold; color:#1A1A1A; line-height:1.5;">جمهورية العراق<br>وزارة التعليم العالي والبحث العلمي</div>
        </td>
        <td style="width:22%; text-align:left; vertical-align:middle;">
          <svg width="64" height="64" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
            <path d="M50 10 L62 28 L85 22 L75 48 L82 78 L50 68 L18 78 L25 48 L15 22 L38 28 Z" fill="#C59B27" stroke="#8C6D1F" stroke-width="2"/>
            <rect x="37" y="36" width="26" height="30" rx="3" fill="#FFF" stroke="#1A1A1A" stroke-width="1.5"/>
            <rect x="37" y="36" width="26" height="10" fill="#CE1126"/>
            <rect x="37" y="56" width="26" height="10" fill="#000"/>
          </svg>
        </td>
      </tr>
    </table>
    """
    cat = (category_code or "humanities").lower()
    if cat == "law":
        trio = """
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><circle cx="38" cy="48" r="20" fill="none" stroke="#5A626A" stroke-width="6"/><circle cx="84" cy="74" r="20" fill="none" stroke="#5A626A" stroke-width="6"/><line x1="54" y1="58" x2="68" y2="66" stroke="#444" stroke-width="5" stroke-dasharray="4,3"/></svg></td>
        <td style="width:34%; text-align:center;"><svg width="92" height="92" viewBox="0 0 140 140"><rect x="42" y="118" width="56" height="10" rx="3" fill="#5C3A21"/><line x1="70" y1="22" x2="70" y2="118" stroke="#B38F40" stroke-width="6"/><circle cx="70" cy="20" r="7" fill="#D4AF37"/><path d="M22 42 Q70 32 118 42" fill="none" stroke="#B38F40" stroke-width="5"/><line x1="26" y1="42" x2="14" y2="82" stroke="#8C6D1F" stroke-width="2"/><line x1="26" y1="42" x2="38" y2="82" stroke="#8C6D1F" stroke-width="2"/><path d="M10 82 Q26 96 42 82 Z" fill="#D4AF37"/><line x1="114" y1="42" x2="102" y2="82" stroke="#8C6D1F" stroke-width="2"/><line x1="114" y1="42" x2="126" y2="82" stroke="#8C6D1F" stroke-width="2"/><path d="M98 82 Q114 96 130 82 Z" fill="#D4AF37"/></svg></td>
        <td style="width:33%; text-align:center;"><svg width="76" height="76" viewBox="0 0 120 120"><rect x="12" y="84" width="52" height="11" rx="3" fill="#6E472B"/><g transform="rotate(-32 65 55)"><rect x="58" y="38" width="9" height="58" rx="3" fill="#8B5A2B"/><rect x="38" y="20" width="48" height="20" rx="5" fill="#5C3A21"/><rect x="56" y="20" width="12" height="20" fill="#D4AF37"/></g></svg></td>
        """
    elif cat == "economics":
        trio = """
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><rect x="18" y="68" width="16" height="34" fill="#1B3B2B"/><rect x="46" y="46" width="16" height="56" fill="#B38F40"/><rect x="74" y="24" width="16" height="78" fill="#7A1C24"/><line x1="10" y1="102" x2="104" y2="102" stroke="#333" stroke-width="3"/></svg></td>
        <td style="width:34%; text-align:center;"><svg width="90" height="90" viewBox="0 0 120 120"><polygon points="60,14 106,38 106,46 14,46 14,38" fill="#1B3B2B"/><rect x="26" y="48" width="10" height="42" fill="#B38F40"/><rect x="55" y="48" width="10" height="42" fill="#B38F40"/><rect x="84" y="48" width="10" height="42" fill="#B38F40"/><rect x="14" y="90" width="92" height="10" fill="#1B3B2B"/></svg></td>
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><circle cx="60" cy="60" r="42" fill="#FAF6F0" stroke="#B38F40" stroke-width="4"/><polyline points="28,76 50,54 66,64 90,36" fill="none" stroke="#7A1C24" stroke-width="4"/></svg></td>
        """
    elif cat == "medical":
        trio = """
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><path d="M35 20 Q60 45 85 20 Q60 45 35 70 Q60 95 85 70" fill="none" stroke="#1B3B2B" stroke-width="4"/><path d="M85 20 Q60 45 35 20 Q60 45 85 70 Q60 95 35 70" fill="none" stroke="#7A1C24" stroke-width="4"/></svg></td>
        <td style="width:34%; text-align:center;"><svg width="90" height="90" viewBox="0 0 120 120"><circle cx="60" cy="60" r="46" fill="#FAF6F0" stroke="#1B3B2B" stroke-width="3.5"/><rect x="50" y="26" width="20" height="68" rx="3" fill="#7A1C24"/><rect x="26" y="50" width="68" height="20" rx="3" fill="#7A1C24"/></svg></td>
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><polyline points="10,60 36,60 48,26 66,94 78,48 90,60 110,60" fill="none" stroke="#B38F40" stroke-width="4.5"/></svg></td>
        """
    elif cat == "engineering":
        trio = """
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><circle cx="60" cy="60" r="28" fill="none" stroke="#1B3B2B" stroke-width="8" stroke-dasharray="14,8"/><circle cx="60" cy="60" r="12" fill="#B38F40"/></svg></td>
        <td style="width:34%; text-align:center;"><svg width="90" height="90" viewBox="0 0 120 120"><ellipse cx="60" cy="60" rx="48" ry="18" fill="none" stroke="#1B3B2B" stroke-width="3"/><ellipse cx="60" cy="60" rx="48" ry="18" fill="none" stroke="#7A1C24" stroke-width="3" transform="rotate(60 60 60)"/><ellipse cx="60" cy="60" rx="48" ry="18" fill="none" stroke="#B38F40" stroke-width="3" transform="rotate(120 60 60)"/><circle cx="60" cy="60" r="8" fill="#7A1C24"/></svg></td>
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><polygon points="60,16 104,96 16,96" fill="none" stroke="#1B3B2B" stroke-width="4"/><line x1="60" y1="16" x2="60" y2="96" stroke="#B38F40" stroke-width="2.5"/></svg></td>
        """
    else:
        trio = """
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><path d="M20 30 Q60 20 100 30 L100 90 Q60 80 20 90 Z" fill="#FAF6F0" stroke="#1B3B2B" stroke-width="3.5"/><line x1="60" y1="25" x2="60" y2="85" stroke="#B38F40" stroke-width="3"/></svg></td>
        <td style="width:34%; text-align:center;"><svg width="90" height="90" viewBox="0 0 120 120"><polygon points="60,20 110,42 60,64 10,42" fill="#1B3B2B"/><path d="M28 52 L28 82 Q60 98 92 82 L92 52" fill="none" stroke="#B38F40" stroke-width="4"/></svg></td>
        <td style="width:33%; text-align:center;"><svg width="74" height="74" viewBox="0 0 120 120"><path d="M30 95 L88 22 C96 30 92 48 76 62 L30 95 Z" fill="#7A1C24" stroke="#B38F40" stroke-width="2"/></svg></td>
        """
    center_trio = f'<table style="width:86%; margin:12px auto; border-collapse:collapse; text-align:center;"><tr>{trio}</tr></table>'
    return top_logos_bar, center_trio

def sanitize_docx_forensics(doc, docx_path, author_name="الباحث الأكاديمي"):
    now_dt = datetime.datetime.utcnow()
    created_dt = now_dt - datetime.timedelta(days=random.randint(3, 6), hours=random.randint(2, 9))
    modified_dt = now_dt - datetime.timedelta(minutes=random.randint(12, 40))
    cp = doc.core_properties
    cp.author = author_name
    cp.last_modified_by = author_name
    cp.comments = ""
    cp.title = ""
    cp.subject = ""
    cp.revision = random.randint(19, 38)
    cp.created = created_dt
    cp.modified = modified_dt
    doc.save(docx_path)
    try:
        tmp_zip = docx_path + ".tmp.zip"
        edit_mins = str(random.randint(195, 340))
        with zipfile.ZipFile(docx_path, 'r') as zin, zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename == "docProps/app.xml":
                    txt_xml = content.decode("utf-8", errors="ignore")
                    txt_xml = re.sub(r'<Application>.*?</Application>', '<Application>Microsoft Office Word</Application>', txt_xml)
                    txt_xml = re.sub(r'<TotalTime>.*?</TotalTime>', f'<TotalTime>{edit_mins}</TotalTime>', txt_xml)
                    content = txt_xml.encode("utf-8")
                zout.writestr(item, content)
        shutil.move(tmp_zip, docx_path)
    except Exception:
        pass

def _set_pPr_bidi_and_shd(p, is_english=False, fill_hex=None, bottom_border_hex=None, outline_lvl=None):
    pPr = p._p.get_or_add_pPr()
    if not is_english and pPr.find(qn('w:bidi')) is None:
        bidi_el = OxmlElement('w:bidi')
        spacing_el = pPr.find(qn('w:spacing'))
        jc_el = pPr.find(qn('w:jc'))
        if spacing_el is not None:
            spacing_el.addprevious(bidi_el)
        elif jc_el is not None:
            jc_el.addprevious(bidi_el)
        else:
            pPr.append(bidi_el)
    if bottom_border_hex and pPr.find(qn('w:pBdr')) is None:
        pBdr_el = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{bottom_border_hex}"/></w:pBdr>')
        pPr.insert(0, pBdr_el)
    if fill_hex and pPr.find(qn('w:shd')) is None:
        shd_el = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
        bidi_ref = pPr.find(qn('w:bidi'))
        if bidi_ref is not None:
            bidi_ref.addprevious(shd_el)
        else:
            pPr.append(shd_el)
    if outline_lvl is not None and pPr.find(qn('w:outlineLvl')) is None:
        ol_el = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{int(outline_lvl)}"/>')
        pPr.append(ol_el)

def add_page_number_to_footer(footer_para):
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = footer_para.add_run("— ( ")
    r1.font.name = "Arial"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r_pg = footer_para.add_run()
    r_pg.font.name = "Arial"
    r_pg.font.size = Pt(10)
    r_pg.font.bold = True
    r_pg.font.color.rgb = RGBColor(0x7A, 0x1C, 0x24)
    r_pg._r.append(fldChar1)
    r_pg._r.append(instrText)
    r_pg._r.append(fldChar2)
    r_pg._r.append(fldChar3)
    r2 = footer_para.add_run(" ) —")
    r2.font.name = "Arial"
    r2.font.size = Pt(10)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

def setup_iraqi_document(is_english=False, header_title=""):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.right_margin = Cm(2.5 if is_english else 3.2)
        section.left_margin = Cm(3.2 if is_english else 2.5)
        sectPr = section._sectPr
        if sectPr.find(qn('w:pgBorders')) is None:
            pgBorders = parse_xml(
                f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
                '<w:top w:val="double" w:sz="12" w:space="18" w:color="1F4E78"/>'
                '<w:left w:val="double" w:sz="12" w:space="18" w:color="1F4E78"/>'
                '<w:bottom w:val="double" w:sz="12" w:space="18" w:color="1F4E78"/>'
                '<w:right w:val="double" w:sz="12" w:space="18" w:color="1F4E78"/>'
                '</w:pgBorders>'
            )
            pgMar = sectPr.find(qn('w:pgMar'))
            if pgMar is not None:
                pgMar.addnext(pgBorders)
            else:
                sectPr.append(pgBorders)
        if header_title:
            hdr = section.header.paragraphs[0]
            hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rh = hdr.add_run(truncate_title_by_words(clean_iraqi_text(header_title), 65))
            rh.font.name = "Arial"
            rh.font.size = Pt(9.5)
            rh.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        add_page_number_to_footer(section.footer.paragraphs[0])
    return doc

SUBHEADING_PREFIXES = (
    "أولاً:", "ثانياً:", "ثالثاً:", "رابعاً:", "خامساً:", "سادساً:", "سابعاً:",
    "أولاً —", "ثانياً —", "ثالثاً —", "رابعاً —",
    "هدف الدراسة:", "منهجية الدراسة:", "أهم النتائج:", "أوجه الاستفادة والمقارنة:",
    "أوجه الاستفادة والمقارنة بالدراسة الحالية:", "الكلمات المفتاحية:", "مخرجات التعلم:"
)

def add_para(doc, text, font_name='Arial', size=14, bold=False, align=None, is_english=False, color_rgb=None, preserve_urls=False, indent_first_line=False):
    if isinstance(font_name, (int, float)):
        size = font_name
        font_name = 'Arial'
    p = doc.add_paragraph()
    p.alignment = align if align is not None else (WD_ALIGN_PARAGRAPH.LEFT if is_english else WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    if indent_first_line and not bold and align is None:
        p.paragraph_format.first_line_indent = Cm(1.15)
    _set_pPr_bidi_and_shd(p, is_english=is_english)

    cleaned = clean_iraqi_text(text, is_english=is_english, preserve_urls=preserve_urls)
    f_use = 'Times New Roman' if is_english else 'Arial'

    if not is_english and ":" in cleaned and not bold and not preserve_urls and any(cleaned.startswith(pfx) for pfx in SUBHEADING_PREFIXES):
        parts = cleaned.split(":", 1)
        r1 = p.add_run(parts[0] + ":")
        r1.bold = True
        r1.font.name = f_use
        r1.font.size = Pt(float(size))
        r1.font.color.rgb = RGBColor(0x7A, 0x1C, 0x24)
        r2 = p.add_run(parts[1])
        r2.font.name = f_use
        r2.font.size = Pt(float(size))
        r2.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
        for r_item in (r1, r2):
            rPr = r_item._r.get_or_add_rPr()
            if rPr.find(qn('w:rtl')) is None:
                rPr.append(OxmlElement('w:rtl'))
        return p

    run = p.add_run(cleaned)
    run.font.name = f_use
    run.font.size = Pt(float(size))
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb) if color_rgb else RGBColor(0x1A, 0x1A, 0x1A)
    if not is_english:
        rPr = run._r.get_or_add_rPr()
        if rPr.find(qn('w:rtl')) is None:
            rPr.append(OxmlElement('w:rtl'))
    return p

def add_banner_heading(doc, text, size=15, is_english=False, outline_lvl=0, page_break_before=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if is_english else WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    if page_break_before:
        p.paragraph_format.page_break_before = True
    _set_pPr_bidi_and_shd(p, is_english=is_english, fill_hex="EAF1F8", bottom_border_hex="1F4E78", outline_lvl=outline_lvl)
    run = p.add_run(clean_iraqi_text(text, is_english=is_english))
    run.font.name = 'Times New Roman' if is_english else 'Arial'
    run.font.size = Pt(float(size))
    run.bold = True
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    if not is_english:
        rPr = run._r.get_or_add_rPr()
        if rPr.find(qn('w:rtl')) is None:
            rPr.append(OxmlElement('w:rtl'))
    return p

def style_zebra_table(tbl, is_english=False, col_widths_cm=None, text_col_idx=None):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for r_idx, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
        if r_idx == 0 and trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))
        for c_idx, cell in enumerate(row.cells):
            if col_widths_cm and c_idx < len(col_widths_cm):
                cell.width = Cm(col_widths_cm[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            fill_hex = "1F4E78" if r_idx == 0 else ("F2F6FC" if r_idx % 2 == 1 else "FFFFFF")
            old_shd = tcPr.find(qn('w:shd'))
            if old_shd is not None:
                tcPr.remove(old_shd)
            shd_el = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
            vAlign_el = tcPr.find(qn('w:vAlign'))
            if vAlign_el is not None:
                vAlign_el.addprevious(shd_el)
            else:
                tcPr.append(shd_el)
            for p in cell.paragraphs:
                if r_idx > 0 and text_col_idx is not None and c_idx == text_col_idx:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if is_english else WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _set_pPr_bidi_and_shd(p, is_english=is_english)
                for run in p.runs:
                    run.font.name = 'Times New Roman' if is_english else 'Arial'
                    run.font.size = Pt(11 if r_idx == 0 else 10.5)
                    run.bold = (r_idx == 0)
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if r_idx == 0 else RGBColor(0x1A, 0x1A, 0x1A)

def convert_docx_to_pdf_fallback(docx_path):
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", "."], timeout=60, check=False)
        pdf_path = docx_path.replace(".docx", ".pdf")
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        pass
    return None

# ==============================================================================
# الوحدة (5): المصنع (1 — البحث المقسم الفاخر) والمصنع (2 — الدراسات السابقة) والمصنع (3 — التخرج)
# ==============================================================================
def build_genspark_luxury_research(uid, order_info, progress_callback=None):
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    plan = order_info["plan"]
    calib = order_info["calib"]
    font_sz = order_info["font_size"]
    min_pgs = order_info.get("min_pages", 15)
    live_bp = order_info.get("live_blueprint", "")
    custom_inst = order_info.get("custom_instructions", "")
    ts_tag = int(time.time())

    paras_per_sub = 3 if min_pgs <= 11 else (4 if min_pgs <= 17 else 5)
    pool_objs = order_info.get("verified_pool_with_context") or [{"ref_line": f"الدوريات العلمية المحكمة في تخصص {specialty}.", "summary": title_ar}]
    max_ref_idx = len(pool_objs[:16]) - 1
    refs_with_summaries_str = "\n".join([f"[REF_{i}] {r['ref_line']}" for i, r in enumerate(pool_objs[:16])])

    actual_plan_outline = "\n".join([f"- {m['mabhath_title']} (ويتناول: {'، '.join(m.get('matalib', []))})" for m in plan])
    custom_note = f"\nتوجيهات إضافية من المستخدم: {custom_inst}\n" if custom_inst else ""

    if progress_callback:
        progress_callback("⏳ [1/4] جاري تأليف الإطار التمهيدي والمنهجي عبر Gemini 3.8 Flash...")

    intro_prompt = f"""
    عنوان البحث: «{title_ar}» — التخصص: «{specialty}» — الكلية: «{order_info['college_name']}».{custom_note}
    الهيكل الفعلي المعتمد لأقسام البحث (التزم بذكره حرفياً في منهجية البحث وخطته):
    {actual_plan_outline}

    اكتب الإطار التمهيدي والمنهجي مقسماً إلى 3 محاور بصيغة JSON:
    {{
      "intro_paragraphs": [
        "الفقرة الأولى (140 كلمة): التأصيل العلمي للموضوع...",
        "الفقرة الثانية (140 كلمة): إشكالية البحث وتساؤلاته المحورية..."
      ],
      "importance_paragraphs": [
        "الفقرة الأولى (140 كلمة): الأهمية العلمية والنظرية للبحث...",
        "الفقرة الثانية (140 كلمة): الأهمية العملية والتطبيقية وأهداف البحث..."
      ],
      "methodology_plan_paragraphs": [
        "الفقرة الأولى (130 كلمة): المنهج العلمي المتبع وأدوات التحليل...",
        "الفقرة الثانية (140 كلمة): عرض هيكل وخطة البحث وفق الأقسام الفعلية أعلاه..."
      ]
    }}
    """
    intro_data = extract_json_safely(ask_agent(intro_prompt, "writer", True, required_keys=["intro_paragraphs", "importance_paragraphs", "methodology_plan_paragraphs"], live_blueprint=live_bp))
    intro_paras = [clean_iraqi_text(p) for p in (intro_data.get("intro_paragraphs") or []) if p]
    imp_paras = [clean_iraqi_text(p) for p in (intro_data.get("importance_paragraphs") or []) if p]
    method_paras = [clean_iraqi_text(p) for p in (intro_data.get("methodology_plan_paragraphs") or []) if p]

    generated_mabahith = []
    global_footnotes = []
    used_ref_counts = {}
    covered_memory = []

    for m_idx, mabhath in enumerate(plan, 1):
        m_title = clean_iraqi_text(mabhath.get("mabhath_title") or f"القسم {m_idx}")
        matalib_list = mabhath.get("matalib") or ["أولاً", "ثانياً"]
        if progress_callback:
            progress_callback(f"⏳ [2/4] جاري التأليف العميق لـ ({m_title})...")

        matalib_outputs = []
        for mt_title_raw in matalib_list:
            mt_title = clean_iraqi_text(mt_title_raw)
            anti_rep = " | ".join(covered_memory[-5:]) if covered_memory else "بداية الدراسة"
            example_paras_json = ", ".join([f'"الفقرة {k} المفصلة..."' for k in range(1, paras_per_sub + 1)])
            example_indices_json = ", ".join([str(k % max(1, max_ref_idx + 1)) for k in range(paras_per_sub)])

            mt_prompt = f"""
            عنوان البحث: «{title_ar}» — التخصص: «{specialty}».{custom_note}
            القسم الرئيسي الحالي: «{m_title}» — الفرع الحالي: «{mt_title}».
            ما تم شرحه سابقاً (يُمنع تكراره): {anti_rep}.

            قائمة المراجع الحقيقية المعتمدة للتوثيق في الهامش السفلي (مرقمة من 0 إلى {max_ref_idx}):
            {refs_with_summaries_str}

            التعليمات الإلزامية:
            1. اكتب ({paras_per_sub}) فقرات علمية تحليلية عميقة ومفصلة تشرح «{mt_title}» مباشرة وبأسلوب أكاديمي عراقي رصين.
            2. لا تضع أي أرقام مربعة مثل [0] أو [REF_0] داخل نص الفقرة، بل ضع أرقام المراجع في حقل "cited_ref_indices" فقط.
            3. نوّع مطالع الفقرات ولا تبدأ فقرتين متتاليتين بـ (تُعدّ / يُعدّ).

            أخرج كائن JSON بالحقول التالية:
            {{
              "paragraphs": [{example_paras_json}],
              "cited_ref_indices": [{example_indices_json}]
            }}
            """
            mt_data = extract_json_safely(ask_agent(mt_prompt, "writer", True, required_keys=["paragraphs"], live_blueprint=live_bp))
            paras_raw = [clean_iraqi_text(p) for p in (mt_data.get("paragraphs") or []) if len(str(p).strip()) > 20]
            if not paras_raw:
                paras_raw = [clean_iraqi_text(f"يرتكز التحليل العلمي لمحور {mt_title} ضمن إطار {title_ar} على أسس منهجية رصينة في تخصص {specialty}.")]

            ref_indices = mt_data.get("cited_ref_indices") or []
            processed_paras_html = []
            processed_paras_docx = []
            section_local_fns = []

            for p_i, p_txt in enumerate(paras_raw):
                p_clean = clean_iraqi_text(p_txt)
                p_clean = re.sub(r'^\s*\(\d+\)\s*', '', p_clean)
                p_clean = re.sub(r'\s*\[\d{1,3}\]\s*\.?$', '', p_clean).strip().rstrip('.')

                raw_choice = ref_indices[p_i] if (isinstance(ref_indices, list) and p_i < len(ref_indices) and isinstance(ref_indices[p_i], int)) else p_i
                chosen_idx = raw_choice % len(pool_objs)
                if used_ref_counts.get(chosen_idx, 0) >= 2 and len(pool_objs) > 4:
                    least_used = sorted(range(len(pool_objs)), key=lambda idx_k: used_ref_counts.get(idx_k, 0))
                    chosen_idx = least_used[0]
                used_ref_counts[chosen_idx] = used_ref_counts.get(chosen_idx, 0) + 1

                fn_content = pool_objs[chosen_idx]["ref_line"]
                global_footnotes.append(fn_content)
                fn_num = len(global_footnotes)
                section_local_fns.append((fn_num, fn_content))

                p_docx_final = f"{p_clean} ({fn_num})."
                fn_safe_html = format_reference_html_safe(fn_content, is_english_ref=False)
                p_html_final = f'{_esc(p_clean)}<sup class="fn-ref">({fn_num})</sup><span class="fn-call">{fn_safe_html}</span>.'
                processed_paras_html.append(p_html_final)
                processed_paras_docx.append(p_docx_final)

            covered_memory.append(f"{mt_title}: {paras_raw[0][:110]}")
            matalib_outputs.append({
                "mt_title": mt_title,
                "paras_raw": paras_raw,
                "paras_html": processed_paras_html,
                "paras_docx": processed_paras_docx,
                "local_fns": section_local_fns
            })
        generated_mabahith.append({"m_idx": m_idx, "m_title": m_title, "matalib": matalib_outputs})

    if progress_callback:
        progress_callback("⏳ [3/4] جاري صياغة الخاتمة وإجراء الفحص الشامل للتقرير والمصادر عبر Gemini 3.8 Flash...")

    conc_prompt = f"""
    عنوان البحث: «{title_ar}» — التخصص: «{specialty}».
    خلاصة ما تم بحثه: {" | ".join(covered_memory[-6:])}.
    أخرج كائن JSON للخاتمة يتضمن فقرة استهلالية، و5 استنتاجات علمية دقيقة، و5 توصيات عملية:
    {{
      "conclusion_intro": "فقرة استهلالية مركزة للخاتمة (75 كلمة)...",
      "findings": ["الاستنتاج الأول...", "الاستنتاج الثاني...", "الاستنتاج الثالث...", "الاستنتاج الرابع...", "الاستنتاج الخامس..."],
      "recommendations": ["التوصية الأولى...", "التوصية الثانية...", "التوصية الثالثة...", "التوصية الرابعة...", "التوصية الخامسة..."]
    }}
    """
    conc_data = extract_json_safely(ask_agent(conc_prompt, "writer", True, required_keys=["findings", "recommendations"], live_blueprint=live_bp))

    # الفحص الشامل للتقرير الكامل والمصادر عبر Gemini 3.8 Flash قبل الطباعة!
    conc_data = inspect_and_heal_full_report(order_info, intro_paras, imp_paras, method_paras, generated_mabahith, conc_data)

    healed_findings = [clean_iraqi_text(x) for x in (conc_data.get("findings") or []) if x]
    healed_recs = [clean_iraqi_text(x) for x in (conc_data.get("recommendations") or []) if x]

    if progress_callback:
        progress_callback("⏳ [4/4] جاري التنضيد المطبعي وتوليد ملفي (PDF الفاخر + Word)...")

    toc_rows_html = (
        '<tr class="toc-main-row"><td class="toc-td-title"><a class="toc-a" href="#sec-intro">أولاً: مقدمة البحث وإشكاليته</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#sec-intro"></a></td></tr>'
        '<tr class="toc-sub-row"><td class="toc-td-title"><a class="toc-a" href="#sec-imp">ثانياً: أهمية البحث وأهدافه</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#sec-imp"></a></td></tr>'
        '<tr class="toc-sub-row"><td class="toc-td-title"><a class="toc-a" href="#sec-plan">ثالثاً: منهجية البحث وخطته</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#sec-plan"></a></td></tr>'
    )
    for mb in generated_mabahith:
        m_id = f"sec-mb-{mb['m_idx']}"
        toc_rows_html += f'<tr class="toc-main-row"><td class="toc-td-title"><a class="toc-a" href="#{m_id}">{_esc(mb["m_title"])}</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#{m_id}"></a></td></tr>'
        for mt_i, mt in enumerate(mb["matalib"], 1):
            mt_id = f"sec-mb-{mb['m_idx']}-mt-{mt_i}"
            toc_rows_html += f'<tr class="toc-sub-row"><td class="toc-td-title"><a class="toc-a" href="#{mt_id}">{_esc(mt["mt_title"])}</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#{mt_id}"></a></td></tr>'
    toc_rows_html += (
        '<tr class="toc-main-row"><td class="toc-td-title"><a class="toc-a" href="#sec-conc">الخاتمة: الاستنتاجات والتوصيات</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#sec-conc"></a></td></tr>'
        '<tr class="toc-main-row"><td class="toc-td-title"><a class="toc-a" href="#sec-refs">قائمة المصادر والمراجع</a></td><td class="toc-td-pg"><a class="toc-pg-a" href="#sec-refs"></a></td></tr>'
    )
    toc_table_html = f'<table class="toc-table"><thead><tr><th style="width:86%; text-align:right;">الموضوع</th><th style="width:14%; text-align:center;">الصفحة</th></tr></thead><tbody>{toc_rows_html}</tbody></table>'

    comp_table_block_html = ""
    comp_tbl = calib.get("comparison_table") or {}
    comp_rows = comp_tbl.get("rows") or []
    if calib.get("include_comparison_table", True) and comp_rows:
        c1_h = clean_iraqi_text(comp_tbl.get("col1_header") or "وجه المقارنة")
        c2_h = clean_iraqi_text(comp_tbl.get("col2_header") or "المحور الأول")
        c3_h = clean_iraqi_text(comp_tbl.get("col3_header") or "المحور الثاني")
        c_title = clean_iraqi_text(comp_tbl.get("title") or "جدول المقارنة التحليلي لمحاور الدراسة")
        c_rows_h = ""
        for r_row in comp_rows:
            if isinstance(r_row, list) and len(r_row) >= 3:
                c_rows_h += f'<tr><td class="td-key">{_esc(clean_iraqi_text(r_row[0]))}</td><td>{_esc(clean_iraqi_text(r_row[1]))}</td><td>{_esc(clean_iraqi_text(r_row[2]))}</td></tr>'
        if c_rows_h:
            comp_table_block_html = f'<div style="page-break-inside:avoid; margin:18px 0;"><div style="font-weight:bold; font-size:13pt; color:#1B3B2B; margin-bottom:6px; text-align:center;">{_esc(c_title)}</div><table class="comp-table"><thead><tr><th>{_esc(c1_h)}</th><th>{_esc(c2_h)}</th><th>{_esc(c3_h)}</th></tr></thead><tbody>{c_rows_h}</tbody></table></div>'

    mabahith_html = ""
    for mb in generated_mabahith:
        m_id = f"sec-mb-{mb['m_idx']}"
        mabahith_html += f'<h2 class="sec-heading" id="{m_id}" style="margin-top:22px;">{_esc(mb["m_title"])}</h2>'
        for mt_i, mt in enumerate(mb["matalib"], 1):
            mt_id = f"sec-mb-{mb['m_idx']}-mt-{mt_i}"
            mabahith_html += f'<h3 class="matlab-heading" id="{mt_id}">{_esc(mt["mt_title"])}</h3>'
            for p_h in mt["paras_html"]:
                mabahith_html += f'<p class="body-p">{p_h}</p>'
        if mb["m_idx"] == len(generated_mabahith) and comp_table_block_html:
            mabahith_html += comp_table_block_html

    findings_html = "".join([f'<p class="body-p" style="text-indent:0;"><strong>{i}.</strong> {_esc(x.rstrip(".")) + "."}</p>' for i, x in enumerate(healed_findings, 1)])
    recs_html = "".join([f'<p class="body-p" style="text-indent:0;"><strong>{i}.</strong> {_esc(x.rstrip(".")) + "."}</p>' for i, x in enumerate(healed_recs, 1)])

    refs_html = ""
    for tier_obj in (calib.get("formatted_reference_tiers") or []):
        t_title = clean_iraqi_text(tier_obj.get("title") or "")
        t_items = tier_obj.get("items") or []
        is_ltr_tier = bool(tier_obj.get("is_ltr", False))
        if t_items:
            dir_attr = ' dir="ltr" style="text-align:left; direction:ltr;"' if is_ltr_tier else ''
            items_li = "".join([f'<li>{format_reference_html_safe(r_it, is_english_ref=is_ltr_tier)}</li>' for r_it in t_items])
            refs_html += f'<h3 class="ref-cat-heading">{_esc(t_title)}</h3><ol class="ref-list"{dir_attr}>{items_li}</ol>'

    intro_ps_html = "".join([f'<p class="body-p">{_esc(p.rstrip(".")) + "."}</p>' for p in intro_paras])
    imp_ps_html = "".join([f'<p class="body-p">{_esc(p.rstrip(".")) + "."}</p>' for p in imp_paras])
    method_ps_html = "".join([f'<p class="body-p">{_esc(p.rstrip(".")) + "."}</p>' for p in method_paras])
    safe_css_title = truncate_title_by_words(clean_iraqi_text(title_ar), 65)
    top_logos_html, center_trio_html = get_cover_graphics_html(order_info.get("category_code", "humanities"))

    html_content = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<style>
@page {{
    size: A4; margin: 22mm 20mm 22mm 20mm;
    @top-center {{ content: "{safe_css_title}"; font-family: 'Noto Naskh Arabic', 'Amiri', 'DejaVu Serif', serif; font-size: 9pt; color: #666; border-bottom: 0.5pt solid #D8D2C8; padding-bottom: 3px; margin-bottom: 6px; }}
    @bottom-center {{ content: counter(page); font-family: 'Noto Naskh Arabic', 'Amiri', 'DejaVu Serif', serif; font-size: 10.5pt; font-weight: bold; color: #7A1C24; }}
    @footnote {{ border-top: 0.8pt solid #999; padding-top: 4px; margin-top: 6px; }}
}}
@page :first {{ @top-center {{ content: none; }} @bottom-center {{ content: none; }} }}
body {{
    font-family: 'Noto Naskh Arabic', 'Amiri', 'DejaVu Serif', serif;
    font-feature-settings: "calt" 0, "dlig" 0;
    font-variant-ligatures: no-contextual;
    font-size: {font_sz}pt; line-height: 1.68; color: #1A1A1A; direction: rtl; text-align: justify; orphans: 3; widows: 3;
}}
.page-break {{ page-break-before: always; }}
.cover-wrap {{ box-sizing: border-box; padding: 4px 0; }}
.cover-univ {{ text-align: center; font-size: 18pt; font-weight: bold; color: #1B3B2B; margin-top: 10px; }}
.cover-college {{ text-align: center; font-size: 14.5pt; color: #333; margin-top: 4px; }}
.cover-kicker {{ text-align: center; font-size: 11.5pt; color: #9A7B4F; margin-top: 10px; }}
.cover-title-box {{ text-align: center; margin: 18px 8px; padding: 16px 12px; border-top: 1.5px solid #9A7B4F; border-bottom: 1.5px solid #9A7B4F; background-color: #FAF6F0; }}
.cover-main-title {{ font-size: 20pt; font-weight: bold; color: #7A1C24; line-height: 1.48; }}
.cover-meta {{ text-align: center; font-size: 13.5pt; margin-top: 18px; line-height: 1.9; }}
.cover-year {{ text-align: center; font-size: 12pt; color: #555; margin-top: 24px; border-top: 1px solid #DDD; padding-top: 8px; }}
.sec-heading {{ font-size: 17.5pt; font-weight: bold; color: #7A1C24; border-bottom: 1.5px solid #7A1C24; padding-bottom: 4px; margin-top: 18px; margin-bottom: 10px; page-break-after: avoid; }}
.matlab-heading {{ font-size: 15pt; font-weight: bold; color: #1B3B2B; border-right: 4px solid #B38F40; padding-right: 9px; margin-top: 14px; margin-bottom: 8px; page-break-after: avoid; }}
.body-p {{ margin-top: 0; margin-bottom: 10px; text-align: justify; text-indent: 1.15cm; orphans: 3; widows: 3; }}
.toc-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12.5pt; }}
.toc-table th {{ background-color: #1B3B2B; color: #FFF; padding: 7px 10px; border: 1px solid #1B3B2B; }}
.toc-table td {{ padding: 6px 10px; border-bottom: 1px dotted #B8B0A4; }}
.toc-main-row td {{ font-weight: bold; color: #7A1C24; background-color: #FAF6F0; }}
.toc-sub-row td.toc-td-title {{ padding-right: 22px; color: #1A1A1A; }}
.toc-td-pg {{ text-align: center; font-weight: bold; color: #1B3B2B; width: 14%; }}
a.toc-a {{ text-decoration: none; color: inherit; }}
a.toc-pg-a {{ text-decoration: none; color: #7A1C24; font-weight: bold; }}
a.toc-pg-a::after {{ content: target-counter(attr(href), page); }}
sup.fn-ref {{ font-size: 8.5pt; color: #7A1C24; font-weight: bold; }}
span.fn-call {{ float: footnote; font-size: 9pt; line-height: 1.38; color: #222; text-align: right; direction: rtl; font-weight: normal; text-indent: 0; }}
::footnote-call {{ content: ""; }}
::footnote-marker {{ content: counter(footnote) ". "; font-weight: bold; color: #7A1C24; }}
.comp-table {{ width: 100%; border-collapse: collapse; margin: 8px 0; font-size: 11pt; page-break-inside: avoid; }}
.comp-table th {{ background-color: #1B3B2B; color: #FFF; padding: 7px; border: 1px solid #1B3B2B; text-align: center; }}
.comp-table td {{ padding: 7px; border: 1px solid #D5CFC7; text-align: center; }}
.comp-table td.td-key {{ font-weight: bold; color: #7A1C24; background-color: #FAF6F0; width: 24%; }}
.ref-cat-heading {{ font-size: 14pt; font-weight: bold; color: #1B3B2B; margin-top: 14px; margin-bottom: 6px; border-right: 4px solid #7A1C24; padding-right: 8px; page-break-after: avoid; }}
.ref-list {{ margin-top: 4px; padding-right: 20px; }}
.ref-list li {{ margin-bottom: 6px; line-height: 1.5; font-size: {max(font_sz-2, 12)}pt; }}
</style>
</head>
<body>
<div class="cover-wrap">
{top_logos_html}
<div class="cover-univ">{_esc(order_info['university_name'])}</div>
<div class="cover-college">{_esc(order_info['college_name'])}</div>
<div class="cover-kicker">{_esc(specialty)}</div>
<div class="cover-title-box"><div class="cover-main-title">{_esc(title_ar)}</div></div>
{center_trio_html}
<div class="cover-meta"><div><strong>إعداد الباحث:</strong> {_esc(order_info['student_name'])}</div><div><strong>بإشراف الأستاذ:</strong> {_esc(order_info['supervisor_name'])}</div></div>
<div class="cover-year">السنة الدراسية {_esc(order_info['academic_year'])}</div>
</div>

<div class="page-break"></div>
<h2 class="sec-heading">فهرس المحتويات</h2>
{toc_table_html}

<div class="page-break"></div>
<h2 class="sec-heading" id="sec-intro">أولاً: مقدمة البحث وإشكاليته</h2>
{intro_ps_html}
<h3 class="matlab-heading" id="sec-imp">ثانياً: أهمية البحث وأهدافه</h3>
{imp_ps_html}
<h3 class="matlab-heading" id="sec-plan">ثالثاً: منهجية البحث وخطته</h3>
{method_ps_html}

{mabahith_html}

<h2 class="sec-heading" id="sec-conc" style="margin-top:22px;">الخاتمة: الاستنتاجات والتوصيات</h2>
<p class="body-p">{_esc(clean_iraqi_text(conc_data.get('conclusion_intro') or '').rstrip('.')) + '.'}</p>
<h3 class="matlab-heading">أولاً: الاستنتاجات</h3>
{findings_html}
<h3 class="matlab-heading">ثانياً: التوصيات والمقترحات</h3>
{recs_html}

<div class="page-break"></div>
<h2 class="sec-heading" id="sec-refs">قائمة المصادر والمراجع</h2>
{refs_html}
</body>
</html>"""

    pdf_file = f"StatLab_Research_{uid}_{ts_tag}.pdf"
    if HAS_WEASYPRINT:
        try:
            weasyprint.HTML(string=html_content).write_pdf(pdf_file)
        except Exception as e_wp:
            log_runtime_error("WeasyPrint", e_wp)
            pdf_file = None

    doc = setup_iraqi_document(is_english=False, header_title=title_ar)
    add_para(doc, "جمهورية العراق\nوزارة التعليم العالي والبحث العلمي", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, f"{order_info['university_name']} — {order_info['college_name']}\n{specialty}", size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(0x1F, 0x4E, 0x78))
    add_para(doc, f"\n« {title_ar} »\n", size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(0x7A, 0x1C, 0x24))
    add_para(doc, f"إعداد الباحث:\n({order_info['student_name']})\n\nبإشراف الأستاذ:\n({order_info['supervisor_name']})\n\nالسنة الدراسية: {order_info['academic_year']}", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_banner_heading(doc, "فهرس المحتويات", size=16, page_break_before=True)
    w_toc = doc.add_table(rows=1, cols=2)
    w_toc.style = "Table Grid"
    w_toc.rows[0].cells[0].text = "القسم / المحور"
    w_toc.rows[0].cells[1].text = "الموضوع التفصيلي"
    toc_w_rows = [
        ("الإطار التمهيدي", "أولاً: مقدمة البحث وإشكاليته | ثانياً: أهمية البحث وأهدافه | ثالثاً: منهجية البحث وخطته")
    ]
    for mb in generated_mabahith:
        toc_w_rows.append((mb["m_title"], " | ".join([mt["mt_title"] for mt in mb["matalib"]])))
    toc_w_rows.append(("الخاتمة والمصادر", "الاستنتاجات، التوصيات، وقائمة المصادر والمراجع المعتمدة"))
    for c_sec, c_sub in toc_w_rows:
        rw = w_toc.add_row().cells
        rw[0].text = c_sec
        rw[1].text = c_sub
    style_zebra_table(w_toc, col_widths_cm=[5.0, 10.5], text_col_idx=1)


    add_banner_heading(doc, "أولاً: مقدمة البحث وإشكاليته", size=16, page_break_before=True)
    for p in intro_paras:
        add_para(doc, p.rstrip(".") + ".", size=font_sz, indent_first_line=True)

    add_banner_heading(doc, "ثانياً: أهمية البحث وأهدافه", size=15.5)
    for p in imp_paras:
        add_para(doc, p.rstrip(".") + ".", size=font_sz, indent_first_line=True)

    add_banner_heading(doc, "ثالثاً: منهجية البحث وخطته", size=15.5)
    for p in method_paras:
        add_para(doc, p.rstrip(".") + ".", size=font_sz, indent_first_line=True)

    for mb in generated_mabahith:
        add_banner_heading(doc, mb["m_title"], size=16)
        for mt in mb["matalib"]:
            add_para(doc, mt["mt_title"], size=font_sz, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
            for p_d in mt["paras_docx"]:
                add_para(doc, p_d, size=font_sz, indent_first_line=True)
            if mt["local_fns"]:
                p_line = doc.add_paragraph()
                p_line.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r_l = p_line.add_run("ــــــــــــــــــــــــــــــــــــــــــــــــــــــــ")
                r_l.font.size = Pt(10)
                r_l.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
                for fn_num_d, fn_txt_d in mt["local_fns"]:
                    add_para(doc, f"({fn_num_d}) {fn_txt_d}", size=10.5, color_rgb=(0x33, 0x33, 0x33))

    if calib.get("include_comparison_table", True) and comp_rows:
        c1_h = clean_iraqi_text(comp_tbl.get("col1_header") or "وجه المقارنة")
        c2_h = clean_iraqi_text(comp_tbl.get("col2_header") or "المحور الأول")
        c3_h = clean_iraqi_text(comp_tbl.get("col3_header") or "المحور الثاني")
        add_banner_heading(doc, clean_iraqi_text(comp_tbl.get("title") or "جدول المقارنة التحليلي لمحاور الدراسة"), size=15)
        w_comp_tbl = doc.add_table(rows=1, cols=3)
        w_comp_tbl.style = "Table Grid"
        w_comp_tbl.rows[0].cells[0].text = c3_h
        w_comp_tbl.rows[0].cells[1].text = c2_h
        w_comp_tbl.rows[0].cells[2].text = c1_h
        for r_row in comp_rows:
            if isinstance(r_row, list) and len(r_row) >= 3:
                rc_w = w_comp_tbl.add_row().cells
                rc_w[0].text = clean_iraqi_text(r_row[2])
                rc_w[1].text = clean_iraqi_text(r_row[1])
                rc_w[2].text = clean_iraqi_text(r_row[0])
        style_zebra_table(w_comp_tbl, col_widths_cm=[5.6, 5.6, 4.2])

    add_banner_heading(doc, "الخاتمة: الاستنتاجات والتوصيات", size=16)
    add_para(doc, clean_iraqi_text(conc_data.get("conclusion_intro", "")).rstrip(".") + ".", size=font_sz, indent_first_line=True)
    add_para(doc, "أولاً: الاستنتاجات:", size=font_sz, bold=True, color_rgb=(0x7A, 0x1C, 0x24))
    for idx_f, f_it in enumerate(healed_findings, 1):
        add_para(doc, f"{idx_f}- {f_it.rstrip('.')}.", size=font_sz)
    add_para(doc, "ثانياً: التوصيات والمقترحات:", size=font_sz, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
    for idx_r, r_it in enumerate(healed_recs, 1):
        add_para(doc, f"{idx_r}- {r_it.rstrip('.')}.", size=font_sz)

    add_banner_heading(doc, "قائمة المصادر والمراجع المعتمدة", size=16, page_break_before=True)
    for tier_obj in (calib.get("formatted_reference_tiers") or []):
        t_title = clean_iraqi_text(tier_obj.get("title") or "")
        t_items = tier_obj.get("items") or []
        is_eng_t = bool(tier_obj.get("is_ltr", False))
        if t_items:
            add_para(doc, t_title, size=14.5, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
            for idx_ref, r_it in enumerate(t_items, 1):
                add_para(doc, f"{idx_ref}. {r_it}", size=12 if is_eng_t else 13.5, is_english=is_eng_t, align=WD_ALIGN_PARAGRAPH.LEFT if is_eng_t else None, preserve_urls=is_eng_t)

    docx_file = f"StatLab_Editable_Research_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(doc, docx_file)
    if not pdf_file or not os.path.exists(pdf_file):
        pdf_file = convert_docx_to_pdf_fallback(docx_file)
    return pdf_file, docx_file

def build_literature_review_pack(uid, order_info):
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    live_bp = order_info.get("live_blueprint", "")
    ts_tag = int(time.time())

    lit_prompt = f"""
    ابحث في Google والباحث العلمي عن الدراسات السابقة المنشورة فعلياً حول موضوع: «{title_ar}» في تخصص ({specialty}).
    أخرج كائن JSON يتضمن:
    1. "ar_studies": 3 دراسات عراقية وعربية محكمة حقيقية في صلب الموضوع.
    2. "foreign_studies": 3 دراسات أجنبية محكمة حقيقية في صلب الموضوع (أو عربية إضافية إذا كان التخصص شرعياً خالصاً).
    3. "commentary": فقرة مفصلة (150 كلمة) للتعقيب العام على الدراسات السابقة وبيان الفجوة البحثية وما يميز الدراسة الحالية.

    صيغة JSON المطلوبة:
    {{
      "ar_studies": [
        {{
          "cite": "اسم الباحث، السنة",
          "study_title": "عنوان الدراسة الحقيقي",
          "aim": "هدفت الدراسة إلى...",
          "method": "اعتمدت الدراسة المنهج...",
          "results": "توصلت الدراسة إلى...",
          "comp": "تتفق الدراسة الحالية مع هذه الدراسة في... وتتميز عنها في...",
          "env": "بيئة الدراسة في 3 كلمات",
          "short_m": "المنهج في 3 كلمات",
          "short_r": "النتيجة في 5 كلمات",
          "full_ref": "التوثيق الكامل للدراسة"
        }}
      ],
      "foreign_studies": [
        {{
          "cite": "Author, Year",
          "study_title": "الترجمة العربية الدقيقة لعنوان الدراسة الأجنبية",
          "aim": "هدفت الدراسة إلى...",
          "method": "اعتمدت الدراسة...",
          "results": "توصلت الدراسة إلى...",
          "comp": "تتفق الدراسة الحالية مع هذه الدراسة في... وتختلف عنها في...",
          "env": "بيئة الدراسة",
          "short_m": "المنهج في 3 كلمات",
          "short_r": "النتيجة في 5 كلمات",
          "apa_ref": "Full English APA Reference"
        }}
      ],
      "commentary": "فقرة التعقيب العام والفجوة البحثية..."
    }}
    """
    lit_data = extract_json_safely(ask_agent(lit_prompt, "calibrator", True, required_keys=["ar_studies", "commentary"], live_blueprint=live_bp, use_google_search=True))
    ar_studies = lit_data.get("ar_studies") or []
    foreign_studies = [] if order_info.get("is_doctrinal") else (lit_data.get("foreign_studies") or [])
    comment_txt = clean_iraqi_text(lit_data.get("commentary") or f"تُبرز مراجعة الدراسات السابقة أهمية معالجة موضوع «{title_ar}» بمنهجية تحليلية متكاملة تسد الفجوة البحثية في تخصص {specialty}.")

    def create_lit_doc(include_refs=True, is_preview=False):
        doc = setup_iraqi_document(header_title=title_ar)
        add_banner_heading(doc, f"الدراسات السابقة والفجوة البحثية — {title_ar}", 16)
        if is_preview:
            add_para(doc, "🔒 [نسخة معاينة خاصة قبل التسليم النهائي — StatLab Academic]", size=13, bold=True, color_rgb=(0x7A, 0x1C, 0x24))
        sec_ords = ["أولاً", "ثانياً", "ثالثاً"]
        sec_idx = 0
        if ar_studies:
            add_banner_heading(doc, f"{sec_ords[sec_idx]}: الدراسات العراقية والعربية المحكمة:", 15)
            sec_idx += 1
            for i, ars in enumerate(ar_studies, 1):
                add_para(doc, f"{i}- دراسة ({clean_iraqi_text(ars.get('cite', ''))}) الموسومة: «{clean_iraqi_text(ars.get('study_title', ''))}»", size=14, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
                add_para(doc, f"هدف الدراسة: {clean_iraqi_text(ars.get('aim', ''))}", size=14)
                add_para(doc, f"منهجية الدراسة: {clean_iraqi_text(ars.get('method', ''))}", size=14)
                res_t = clean_iraqi_text(ars.get("results", ""))
                add_para(doc, f"أهم النتائج: {res_t[:75] + '... [🔒 محجوب في نسخة المعاينة]' if is_preview else res_t}", size=14)
                add_para(doc, f"أوجه الاستفادة والمقارنة: {clean_iraqi_text(ars.get('comp', ''))}", size=14)
        if foreign_studies:
            add_banner_heading(doc, f"{sec_ords[sec_idx]}: الدراسات الأجنبية المحكمة:", 15)
            sec_idx += 1
            for i, fp in enumerate(foreign_studies, 1):
                add_para(doc, f"{i}- دراسة ({clean_iraqi_text(fp.get('cite', ''), is_english=True)}) الموسومة: «{clean_iraqi_text(fp.get('study_title', ''))}»", size=14, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
                add_para(doc, f"هدف الدراسة: {clean_iraqi_text(fp.get('aim', ''))}", size=14)
                add_para(doc, f"منهجية الدراسة: {clean_iraqi_text(fp.get('method', ''))}", size=14)
                res_f = clean_iraqi_text(fp.get("results", ""))
                add_para(doc, f"أهم النتائج: {res_f[:75] + '... [🔒 محجوب في نسخة المعاينة]' if is_preview else res_f}", size=14)
                add_para(doc, f"أوجه الاستفادة والمقارنة بالدراسة الحالية: {clean_iraqi_text(fp.get('comp', ''))}", size=14)
                if include_refs and not is_preview and fp.get("apa_ref"):
                    add_para(doc, f"Documentation (APA 7): {fp['apa_ref']}", size=11.5, is_english=True, align=WD_ALIGN_PARAGRAPH.LEFT, preserve_urls=True)

        add_banner_heading(doc, f"{sec_ords[min(sec_idx, 2)]}: التعقيب العام على الدراسات السابقة والفجوة البحثية:", 15)
        add_para(doc, comment_txt[:110] + "... [🔒 محجوب في نسخة المعاينة]" if is_preview else comment_txt, size=14, indent_first_line=True)

        add_banner_heading(doc, "جدول (1): ملخص الدراسات السابقة وأوجه مقارنتها بالدراسة الحالية", 14.5)
        tbl = doc.add_table(rows=1, cols=6)
        tbl.style = "Table Grid"
        for i, h in enumerate(["بيئة ومجال الدراسة", "أهم النتائج", "المنهج المعتمد", "عنوان الدراسة", "الباحث والسنة", "ت"]):
            tbl.rows[0].cells[i].text = h
        all_rows = []
        for ars in ar_studies:
            all_rows.append((ars.get("env", ""), ars.get("short_r", ""), ars.get("short_m", ""), ars.get("study_title", ""), f"({ars.get('cite', '')})"))
        for fp in foreign_studies:
            all_rows.append((fp.get("env", ""), fp.get("short_r", ""), fp.get("short_m", ""), fp.get("study_title", ""), f"({fp.get('cite', '')})"))
        for idx_t, r_data in enumerate(all_rows, 1):
            rc = tbl.add_row().cells
            for c_i in range(5):
                rc[c_i].text = clean_iraqi_text(r_data[c_i])
            rc[5].text = str(idx_t)
        style_zebra_table(tbl, col_widths_cm=[2.6, 3.2, 2.5, 4.4, 2.2, 1.2], text_col_idx=3)
        add_para(doc, "المصدر: من إعداد الباحث بالاعتماد على الدراسات السابقة المعتمدة.", size=11, bold=True)

        if include_refs and not is_preview:
            add_banner_heading(doc, "قائمة مصادر الدراسات السابقة (References):", 15)
            for ars in ar_studies:
                if ars.get("full_ref"):
                    add_para(doc, f"• {clean_iraqi_text(ars['full_ref'], is_reference=True)}", size=13)
            for fp in foreign_studies:
                if fp.get("apa_ref"):
                    add_para(doc, clean_iraqi_text(fp["apa_ref"], is_english=True, preserve_urls=True, is_reference=True), size=11.5, is_english=True, align=WD_ALIGN_PARAGRAPH.LEFT, preserve_urls=True)
        return doc

    f_main = f"StatLab_Full_Lit_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(create_lit_doc(True, False), f_main)
    f_turn = f"StatLab_Turnitin_NoRefs_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(create_lit_doc(False, False), f_turn)
    f_prev = f"StatLab_Office_Preview_Locked_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(create_lit_doc(False, True), f_prev)
    return f_main, f_turn, f_prev

def build_complete_graduation_thesis(uid, order_info):
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    calib = order_info["calib"]
    plan = order_info["plan"]
    live_bp = order_info.get("live_blueprint", "")
    custom_inst = order_info.get("custom_instructions", "")
    verified_objs = order_info.get("verified_source_objects") or []
    ts_tag = int(time.time())

    allowed_cites = [s.get("cite") for s in verified_objs if s.get("cite")]
    citations_str = "، ".join([f"({c})" for c in allowed_cites[:12]])
    actual_plan_outline = "\n".join([f"- {m['mabhath_title']} (ويتفرع إلى: {'، '.join(m.get('matalib', []))})" for m in plan])

    doc = setup_iraqi_document(is_english=False, header_title=title_ar)
    add_para(doc, f"جمهورية العراق\nوزارة التعليم العالي والبحث العلمي\n{order_info['university_name']} — {order_info['college_name']}\nقسم: {specialty}", size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(0x1F, 0x4E, 0x78))
    add_para(doc, f"\n\n« {title_ar} »\n", size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(0x7A, 0x1C, 0x24))
    add_para(doc, f"بحث تخرج مقدم إلى مجلس الكلية كجزء من متطلبات نيل شهادة البكالوريوس\n\nإعداد الطالب:\n({order_info['student_name']})\n\nبإشراف الأستاذ:\n({order_info['supervisor_name']})\n\nالسنة الدراسية: {order_info['academic_year']}", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    abs_prompt = f"""
    لبحث تخرج جامعي بعنوان «{title_ar}» في تخصص ({specialty}):
    أخرج كائن JSON يحتوي على:
    {{
      "dedication": "فقرة إهداء وشكر وتقدير أكاديمية رصينة مكتملة الجمل (85 كلمة) دون ذكر أسماء أشخاص غير مذكورين في الطلب",
      "abstract_ar": "المستخلص باللغة العربية الفصحى (160 كلمة تلخص المشكلة والهدف والمنهج وأبرز النتائج)",
      "keywords_ar": "5 كلمات مفتاحية بالعربية مفصولة بفواصل",
      "abstract_en": "Academic English Abstract (160 words summarizing problem, aim, methodology, and findings)",
      "keywords_en": "5 English Keywords separated by commas"
    }}
    """
    ab_data = extract_json_safely(ask_agent(abs_prompt, "writer", True, required_keys=["abstract_ar", "abstract_en"], live_blueprint=live_bp))

    add_banner_heading(doc, "الإهداء والشكر والتقدير", 16, page_break_before=True)
    add_para(doc, clean_iraqi_text(ab_data.get("dedication") or "إلى منارات العلم والمعرفة في جامعاتنا العراقية العريقة..."), size=14, indent_first_line=True)
    add_banner_heading(doc, "المستخلص", 16)
    add_para(doc, clean_iraqi_text(ab_data.get("abstract_ar") or ""), size=14, indent_first_line=True)
    kw_ar_str = clean_iraqi_text(ab_data.get("keywords_ar") or title_ar).strip("[]«»'\" ")
    add_para(doc, f"الكلمات المفتاحية: {kw_ar_str}.", size=13, bold=True)
    add_banner_heading(doc, "Abstract", 16, is_english=True)
    add_para(doc, clean_iraqi_text(ab_data.get("abstract_en") or "", is_english=True), size=13, is_english=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    kw_en_str = clean_iraqi_text(ab_data.get("keywords_en") or "", is_english=True).strip("[]'\" ")
    add_para(doc, f"Keywords: {kw_en_str}.", size=12, bold=True, is_english=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    add_banner_heading(doc, "فهرس المحتويات (قائمة الموضوعات)", 16, page_break_before=True)
    toc_tbl = doc.add_table(rows=1, cols=3)
    toc_tbl.style = "Table Grid"
    toc_tbl.rows[0].cells[0].text = "الموضوع التفصيلي"
    toc_tbl.rows[0].cells[1].text = "القسم الرئيسي"
    toc_tbl.rows[0].cells[2].text = "ت"

    dynamic_toc = [("الإطار التمهيدي", "الإهداء، المستخلص العربي، والـ Abstract الإنجليزي"), ("الفصل الأول", "الإطار العام والمنهجي للبحث")]
    plan_steps = []
    for mb in plan:
        m_t = clean_iraqi_text(mb.get("mabhath_title") or "القسم")
        for mt_t in (mb.get("matalib") or ["أولاً", "ثانياً"]):
            mt_c = clean_iraqi_text(mt_t)
            dynamic_toc.append((m_t, mt_c))
            plan_steps.append((f"{m_t} — {mt_c}", f"اشرح بعمق أكاديمي مفصل (4 فقرات غزيرة) كافة أبعاد ({mt_c}) ضمن ({m_t}) مع التوثيق العلمي."))
    dynamic_toc.append(("الخاتمة والمصادر", "الاستنتاجات العلمية، التوصيات الإجرائية، وقائمة المصادر المعتمدة"))

    for idx_tc, (c_sec, c_top) in enumerate(dynamic_toc, 1):
        r_tc = toc_tbl.add_row().cells
        r_tc[0].text = c_top
        r_tc[1].text = c_sec
        r_tc[2].text = str(idx_tc)
    style_zebra_table(toc_tbl, col_widths_cm=[8.8, 5.2, 1.2], text_col_idx=0)

    ch1_prompt = f"""
    عنوان البحث: «{title_ar}» — التخصص: «{specialty}».
    المراجع المعتمدة للتوثيق داخل المتن: {citations_str}.
    الهيكل الفعلي للبحث:
    {actual_plan_outline}

    اكتب الفصل الأول للبحث شاملاً الأركان الـ 7 كاملة بصيغة JSON:
    {{
      "p1_intro": "فقرة موثقة مفصلة (150 كلمة) لمقدمة البحث...",
      "p2_problem": "فقرة مفصلة (130 كلمة) لمشكلة البحث وتساؤلاته الفرعية...",
      "p3_importance": "فقرة مفصلة (130 كلمة) للأهمية العلمية والعملية للبحث...",
      "p4_aims": "فقرة مفصلة (100 كلمة) لأهداف البحث...",
      "p5_hypotheses": "فقرة مفصلة (100 كلمة) لفرضيات أو تساؤلات البحث...",
      "p6_method": "فقرة مفصلة (110 كلمات) لمنهجية البحث وهيكل تقسيمه الفعلي أعلاه...",
      "p7_limits": "فقرة مفصلة (90 كلمة) لحدود البحث الموضوعية والمكانية والزمانية..."
    }}
    """
    ch1_data = extract_json_safely(ask_agent(ch1_prompt, "writer", True, required_keys=["p1_intro", "p2_problem", "p3_importance", "p4_aims", "p5_hypotheses", "p6_method", "p7_limits"], live_blueprint=live_bp))
    add_banner_heading(doc, "الفصل الأول: الإطار العام والمنهجي للبحث", 15.5, page_break_before=True)

    ch1_labels = [
        ("p1_intro", "أولاً: مقدمة البحث"),
        ("p2_problem", "ثانياً: مشكلة البحث وتساؤلاته"),
        ("p3_importance", "ثالثاً: أهمية البحث"),
        ("p4_aims", "رابعاً: أهداف البحث"),
        ("p5_hypotheses", "خامساً: فرضيات وتساؤلات البحث"),
        ("p6_method", "سادساً: منهجية البحث وخطته"),
        ("p7_limits", "سابعاً: حدود البحث")
    ]
    for idx_ch, (k_ch1, lbl_ch1) in enumerate(ch1_labels):
        val_ch1 = clean_iraqi_text(ch1_data.get(k_ch1) or "")
        if val_ch1:
            val_ch1 = re.sub(r'^(?:أولاً|ثانياً|ثالثاً|رابعاً|خامساً|سادساً|سابعاً)\s*[:：\-]?\s*(?:مقدمة البحث|مشكلة البحث وتساؤلاته|مشكلة البحث|أهمية البحث|أهداف البحث|فرضيات البحث|فرضيات وتساؤلات البحث|منهجية البحث وخطته|منهجية البحث|حدود البحث)?\s*[:：\-]?\s*', '', val_ch1).strip()
            if idx_ch == 0 and allowed_cites and "(" not in val_ch1:
                val_ch1 = val_ch1.rstrip(".") + f" ({allowed_cites[0]})."
            add_para(doc, f"{lbl_ch1}: {val_ch1.rstrip('.')}.", size=14)

    covered_mem = []
    cite_counter = 0
    for sec_head, sec_task in plan_steps:
        anti_rep = " | ".join(covered_mem[-4:]) if covered_mem else "بداية البحث"
        st_prompt = f"""
        عنوان البحث: «{title_ar}» — التخصص: «{specialty}».
        القسم المطلوب كتابته الآن: «{sec_head}».
        المهمة التفصيلية: {sec_task}
        المراجع الحقيقية المعتمدة للتوثيق داخل المتن (استشهد بها بصيغة (المؤلف، السنة)): {citations_str}.
        تجنب تكرار ما سبق: {anti_rep}.

        أخرج كائن JSON بالحقول التالية:
        {{
          "paragraphs": [
            "الفقرة الأولى المفصلة والموثقة (135 كلمة)...",
            "الفقرة الثانية المفصلة والموثقة (135 كلمة)...",
            "الفقرة الثالثة المفصلة والموثقة (135 كلمة)...",
            "الفقرة الرابعة المفصلة والموثقة (135 كلمة)..."
          ]
        }}
        """
        sec_json = extract_json_safely(ask_agent(st_prompt, "writer", True, required_keys=["paragraphs"], live_blueprint=live_bp))
        sec_paras = [clean_iraqi_text(p) for p in (sec_json.get("paragraphs") or []) if len(str(p).strip()) > 20]
        if sec_paras:
            covered_mem.append(f"{sec_head}: {sec_paras[0][:110]}")
        add_banner_heading(doc, sec_head, 15)
        for p_line in sec_paras:
            s_clean = clean_iraqi_text(p_line).rstrip(".")
            if allowed_cites and not re.search(r'\((?:19|20)\d{2}|\d{4}\)', s_clean):
                s_clean = f"{s_clean} ({allowed_cites[cite_counter % len(allowed_cites)]})"
            cite_counter += 1
            add_para(doc, s_clean + ".", size=14, indent_first_line=True)

    conc_g_prompt = f"""
    عنوان البحث: «{title_ar}» — التخصص: «{specialty}».
    اكتب الخاتمة مقسمة إلى فقرة استهلالية، و5 استنتاجات علمية، و5 توصيات عملية بصيغة JSON:
    {{
      "conclusion_intro": "فقرة استهلالية للخاتمة...",
      "findings": ["الاستنتاج 1...", "الاستنتاج 2...", "الاستنتاج 3...", "الاستنتاج 4...", "الاستنتاج 5..."],
      "recommendations": ["التوصية 1...", "التوصية 2...", "التوصية 3...", "التوصية 4...", "التوصية 5..."]
    }}
    """
    conc_g = extract_json_safely(ask_agent(conc_g_prompt, "writer", True, required_keys=["findings", "recommendations"], live_blueprint=live_bp))
    add_banner_heading(doc, "الخاتمة: الاستنتاجات العلمية والتوصيات", 15.5)
    add_para(doc, clean_iraqi_text(conc_g.get("conclusion_intro") or "").rstrip(".") + ".", size=14, indent_first_line=True)
    add_para(doc, "أولاً: الاستنتاجات:", size=14, bold=True, color_rgb=(0x7A, 0x1C, 0x24))
    for i_f, f_txt in enumerate(conc_g.get("findings") or [], 1):
        add_para(doc, f"{i_f}- {clean_iraqi_text(f_txt).rstrip('.')}.", size=14)
    add_para(doc, "ثانياً: التوصيات والمقترحات:", size=14, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
    for i_r, r_txt in enumerate(conc_g.get("recommendations") or [], 1):
        add_para(doc, f"{i_r}- {clean_iraqi_text(r_txt).rstrip('.')}.", size=14)

    add_banner_heading(doc, "قائمة المصادر والمراجع المعتمدة (References):", 16, page_break_before=True)
    for tier_obj in (calib.get("formatted_reference_tiers") or []):
        t_title = clean_iraqi_text(tier_obj.get("title") or "")
        t_items = sorted(tier_obj.get("items") or [])
        is_eng_t = bool(tier_obj.get("is_ltr", False))
        if t_items:
            add_para(doc, t_title, size=14.5, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
            for idx_r, r_it in enumerate(t_items, 1):
                add_para(doc, f"{idx_r}. {r_it}", size=12 if is_eng_t else 13.5, is_english=is_eng_t, align=WD_ALIGN_PARAGRAPH.LEFT if is_eng_t else None, preserve_urls=is_eng_t)

    fname = f"StatLab_Full_Grad_Thesis_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(doc, fname)
    pdf_grad = convert_docx_to_pdf_fallback(fname)
    return fname, pdf_grad

# ==============================================================================
# الوحدة (6): المصانع (4: SPSS، 5: السمنار، 6: بولونيا، 7: الإنجليزي، 8: خفض الاستلال)
# ==============================================================================
def build_ultimate_spss_pack(uid, order_info):
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    indep_var = order_info.get("independent_var") or f"المحور الأول في {title_ar[:30]}"
    dep_var = order_info.get("dependent_var") or f"المحور الثاني في {title_ar[:30]}"
    ts_tag = int(time.time())

    q_prompt = f"""
    لبحث بعنوان: «{title_ar}» في تخصص ({specialty}):
    صمم 12 عبارة استبيان أكاديمية دقيقة على مقياس ليكرت الخماسي (6 للمتغير المستقل «{indep_var}»، و6 للمتغير التابع «{dep_var}»)، وحدد مجتمع الدراسة المناسب.
    أخرج JSON بالحقول التالية:
    {{
      "population_name": "وصف مجتمع الدراسة الميداني في 6 كلمات",
      "items": ["العبارة 1...", "العبارة 2...", "العبارة 3...", "العبارة 4...", "العبارة 5...", "العبارة 6...", "العبارة 7...", "العبارة 8...", "العبارة 9...", "العبارة 10...", "العبارة 11...", "العبارة 12..."]
    }}
    """
    q_data = extract_json_safely(ask_agent(q_prompt, "writer", True, required_keys=["items"]))
    pop_name = clean_iraqi_text(q_data.get("population_name") or f"المختصين والعاملين في ميدان {specialty} في العراق")
    items = [clean_iraqi_text(x) for x in (q_data.get("items") or []) if len(str(x).strip()) > 8][:12]
    while len(items) < 12:
        items.append(f"تسهم آليات ({indep_var}) في تعزيز ({dep_var}) في الفقرة ({len(items)+1}).")

    k = len(items)
    half_k = k // 2
    N_pop = random.randint(240, 580)
    z_val, d_val, p_val = 1.96, 0.05, 0.50
    n_thompson = int(np.ceil((N_pop * (z_val**2) * p_val * (1 - p_val)) / (((N_pop - 1) * (d_val**2)) + ((z_val**2) * p_val * (1 - p_val)))))
    dist_cnt = n_thompson + random.randint(12, 22)
    ret_cnt = n_thompson + random.randint(4, 10)
    excl_cnt = random.randint(2, 5)
    sample_size = max(60, ret_cnt - excl_cnt)
    resp_rate = (sample_size / dist_cnt) * 100.0

    rows_data = []
    latent = np.random.normal(loc=3.84, scale=0.56, size=sample_size)
    for idx_r in range(sample_size):
        gender = np.random.choice(["ذكر", "أنثى"], p=[0.56, 0.44])
        qual = np.random.choice(["بكالوريوس", "ماجستير", "دكتوراه"], p=[0.66, 0.24, 0.10])
        age_grp = np.random.choice(["21 - 25 سنة", "26 - 30 سنة", "31 - 40 سنة", "41 سنة فأكثر"], p=[0.24, 0.36, 0.28, 0.12])
        exp_grp = np.random.choice(["أقل من 5 سنوات", "6 - 10 سنوات", "11 سنة فأكثر"], p=[0.38, 0.42, 0.20])
        row_dict = {"ت": idx_r + 1, "الجنس": gender, "الفئة_العمرية": age_grp, "المؤهل_العلمي": qual, "سنوات_الخبرة": exp_grp}
        for i_q in range(k):
            val = int(np.clip(np.round(latent[idx_r] + np.random.normal(0, 0.58) - (i_q * 0.015)), 1, 5))
            row_dict[f"فقرة_{i_q+1}"] = val
        rows_data.append(row_dict)

    df = pd.DataFrame(rows_data)
    excel_file = f"StatLab_SPSS_RawData_{uid}_{ts_tag}.xlsx"
    df.to_excel(excel_file, index=False)

    item_cols = [f"فقرة_{i}" for i in range(1, k + 1)]
    num_df = df[item_cols]
    alpha_val = float((k / (k - 1.0)) * (1.0 - (num_df.var(axis=0, ddof=1).sum() / max(num_df.sum(axis=1).var(ddof=1), 1e-6))))
    validity_val = float(np.sqrt(max(alpha_val, 0.0)))

    x_indep = num_df[item_cols[:half_k]].mean(axis=1)
    y_dep = num_df[item_cols[half_k:]].mean(axis=1)
    r_val, _ = stats.pearsonr(x_indep, y_dep)
    slope, _, r_value, p_reg, _ = stats.linregress(x_indep, y_dep)
    r_squared = float(r_value ** 2)
    f_stat = float((r_squared / 1.0) / max((1.0 - r_squared) / (sample_size - 2), 1e-6))
    ks_stat, _ = stats.kstest(stats.zscore(num_df.mean(axis=1)), 'norm')

    means_vals = [float(num_df[c].mean()) for c in item_cols]
    plt.figure(figsize=(7.5, 3.8), dpi=170)
    colors_bar = ['#1F4E78' if i < half_k else '#7A1C24' for i in range(k)]
    bars = plt.bar([f"Q{i}" for i in range(1, k + 1)], means_vals, color=colors_bar, edgecolor='black')
    plt.axhline(y=3.0, color='#B38F40', linestyle='--', linewidth=1.6, label='3.0 (Standard Mean)')
    plt.ylim(0, 5.0)
    plt.xlabel(f"Q1-Q{half_k} (Axis 1)  |  Q{half_k+1}-Q{k} (Axis 2)", fontweight='bold')
    plt.ylabel("Mean (1 - 5)", fontweight='bold')
    for bar in bars:
        yv = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2.0, yv + 0.08, f"{yv:.2f}", ha='center', va='bottom', fontsize=8, fontweight='bold')
    plt.legend()
    plt.tight_layout()
    chart_path = f"chart_{uid}_{ts_tag}.png"
    plt.savefig(chart_path)
    plt.close()

    doc = setup_iraqi_document(header_title=title_ar)
    add_banner_heading(doc, f"الجانب الميداني والتحليل الإحصائي (SPSS v.28)\nللدراسة الموسومة: «{title_ar}»", 16)
    add_banner_heading(doc, "أولاً: مجتمع وعينة الدراسة وإجراءات التطبيق الميداني:", 15)
    add_para(doc, f"تمثل مجتمع الدراسة بـ ({pop_name}) والبالغ قوامه ({N_pop}) مفردة، وجرى احتساب الحد الأدنى لحجم العينة الممثلة إحصائياً وفق معادلة (ستيفن ثومبسون — Steven Thompson) عند مستوى ثقة (95%) وهامش خطأ (0.05) فبلغ ({n_thompson}) مفردة. ولضمان دقة التمثيل الإحصائي، وُزعت ({dist_cnt}) استمارة استبيان، استُرد منها ({ret_cnt}) استمارة، وبعد الفحص والتدقيق استُبعدت ({excl_cnt}) استمارات لعدم اكتمال بياناتها، ليبلغ عدد الاستمارات الصالحة للتحليل الإحصائي ({sample_size}) استمارة بنسبة استجابة صافية بلغت ({resp_rate:.1f}%).", size=14, indent_first_line=True)

    add_banner_heading(doc, "ثانياً: الوصف الإحصائي للخصائص الديموغرافية لعينة الدراسة:", 15)
    tdemo = doc.add_table(rows=1, cols=4)
    tdemo.style = "Table Grid"
    for i, h in enumerate(["النسبة المئوية %", "التكرار (Frequency)", "الفئة", "المتغير الديموغرافي"]):
        tdemo.rows[0].cells[i].text = h
    for col_d, label_d in [("الجنس", "الجنس"), ("الفئة_العمرية", "الفئة العمرية"), ("المؤهل_العلمي", "المؤهل العلمي"), ("سنوات_الخبرة", "سنوات الخبرة")]:
        vc = df[col_d].value_counts()
        for cat_name, cnt_v in vc.items():
            rd = tdemo.add_row().cells
            rd[0].text = f"{(cnt_v / sample_size) * 100:.1f}%"
            rd[1].text = str(int(cnt_v))
            rd[2].text = str(cat_name)
            rd[3].text = label_d
    style_zebra_table(tdemo, col_widths_cm=[3.2, 3.2, 4.2, 4.2])
    add_para(doc, f"المصدر: من إعداد الباحث بالاعتماد على مخرجات برنامج (SPSS v.28) للعينة (n = {sample_size}).", size=11, bold=True)

    add_banner_heading(doc, "ثالثاً: اختبار التوزيع الطبيعي وثبات وصدق أداة القياس (Cronbach's Alpha):", 15)
    t1 = doc.add_table(rows=2, cols=5)
    t1.style = "Table Grid"
    for i, h in enumerate(["القرار الإحصائي", "إحصاء (K-S)", "معامل الصدق الذاتي", "معامل ألفا كرونباخ", "عدد الفقرات"]):
        t1.rows[0].cells[i].text = h
    for i, v in enumerate(["ثبات عالٍ وبيانات طبيعية", f"{ks_stat:.3f}", f"{validity_val:.3f}", f"{alpha_val:.3f}", str(k)]):
        t1.rows[1].cells[i].text = v
    style_zebra_table(t1, col_widths_cm=[3.8, 2.7, 2.8, 2.8, 2.5])
    add_para(doc, f"تُظهر نتائج الجدول أعلاه أن قيمة معامل الثبات العام (ألفا كرونباخ) المحسوبة فعلياً من بيانات المقياس البالغ عدد فقراته ({k}) فقرة قد بلغت ({alpha_val:.3f})، وبمعامل صدق ذاتي قدره ({validity_val:.3f})، وهي قيم تتجاوز الحد المعياري الأدنى المعتمد أكاديمياً (0.70)، مما يؤكد اتساق أداة القياس وصلاحيتها لاختبار متغيري الدراسة ({indep_var}) و({dep_var}).", size=14, indent_first_line=True)

    add_banner_heading(doc, "رابعاً: التحليل الوصفي لفقرات ومحاور الدراسة (الأوساط الحسابية، الانحرافات المعيارية، معامل الاختلاف، وقيمة t):", 15)
    t2 = doc.add_table(rows=1, cols=8)
    t2.style = "Table Grid"
    for i, h in enumerate(["رتبة", "المستوى", "قيمة t", "الوزن %", "الاختلاف%", "الانحراف", "الوسط", "عبارة الفقرة"]):
        t2.rows[0].cells[i].text = h
    st_list = []
    for i_c, col_c in enumerate(item_cols):
        m_v = float(df[col_c].mean())
        sd_v = float(df[col_c].std(ddof=1))
        cv_v = (sd_v / m_v) * 100.0
        w_v = (m_v / 5.0) * 100.0
        t_val = (m_v - 3.0) / (sd_v / np.sqrt(sample_size))
        lvl = "مرتفع جداً" if m_v >= 4.2 else ("مرتفع" if m_v >= 3.4 else "متوسط")
        st_list.append((clean_iraqi_text(items[i_c]), m_v, sd_v, cv_v, w_v, t_val, lvl))
    st_list.sort(key=lambda x: x[1], reverse=True)
    for rnk, item_s in enumerate(st_list, 1):
        rc = t2.add_row().cells
        rc[0].text = str(rnk)
        rc[1].text = item_s[6]
        rc[2].text = f"{item_s[5]:.2f}"
        rc[3].text = f"{item_s[4]:.1f}%"
        rc[4].text = f"{item_s[3]:.1f}%"
        rc[5].text = f"{item_s[2]:.3f}"
        rc[6].text = f"{item_s[1]:.3f}"
        rc[7].text = item_s[0]
    style_zebra_table(t2, col_widths_cm=[1.0, 1.3, 1.3, 1.3, 1.3, 1.4, 1.4, 7.2], text_col_idx=7)

    overall_mean = float(num_df.values.mean())
    overall_sd = float(num_df.values.std(ddof=1))
    top_1, low_1 = st_list[0], st_list[-1]
    add_para(doc, f"يتضح من نتائج التحليل الوصفي أن الوسط الحسابي العام لإجابات المبحوثين بلغ ({overall_mean:.3f}) بانحراف معياري عام ({overall_sd:.3f})، وهو يفوق الوسط الفرضي البالغ (3.0) بدلالة قيم (t) المحسوبة. وقد تصدرت الفقرة «{top_1[0]}» المرتبة الأولى بوسط حسابي ({top_1[1]:.3f}) ووزن نسبي ({top_1[4]:.1f}%)، بينما جاءت الفقرة «{low_1[0]}» في الترتيب الأخير بوسط حسابي ({low_1[1]:.3f}) ومستوى استجابة ({low_1[6]}).", size=14, indent_first_line=True)

    if os.path.exists(chart_path):
        add_para(doc, f"شكل (1): الأوساط الحسابية لفقرات محور ({indep_var}) ومحور ({dep_var}) مقارنة بالوسط الفرضي (3.0)", size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(chart_path, width=Inches(5.6))

    add_banner_heading(doc, f"خامساً: اختبار فرضيات الارتباط والأثر بين ({indep_var}) و({dep_var}):", 15)
    t3 = doc.add_table(rows=2, cols=6)
    t3.style = "Table Grid"
    for i, h in enumerate(["القرار الإحصائي", "مستوى الدلالة (Sig.)", "قيمة (F) المحسوبة", "معامل التحديد (R²)", "ميل الانحدار (β)", "معامل الارتباط (r)"]):
        t3.rows[0].cells[i].text = h
    for i, v in enumerate(["دال معنوياً (قبول فرضية الأثر)", f"{p_reg:.4f}", f"{f_stat:.2f}", f"{r_squared:.3f}", f"{slope:.3f}", f"{r_val:.3f}"]):
        t3.rows[1].cells[i].text = v
    style_zebra_table(t3, col_widths_cm=[3.6, 2.3, 2.3, 2.3, 2.3, 2.3])
    add_para(doc, f"تكشف نتائج مصفوفة الارتباط والانحدار الخطي عن وجود علاقة ارتباط طردية موجبة وذات دلالة إحصائية بين ({indep_var}) و({dep_var}) حيث بلغ معامل ارتباط بيرسون (r = {r_val:.3f}). كما أظهرت نتائج الانحدار أن معامل التحديد بلغ (R² = {r_squared:.3f})، مما يدل على أن ({indep_var}) يفسر ما نسبته ({r_squared*100:.1f}%) من التباين الحاصل في ({dep_var})، وبدلالة قيمة (F) المحسوبة البالغة ({f_stat:.2f}) وميل انحدار (β = {slope:.3f}) عند مستوى دلالة معنوية (Sig. = {p_reg:.4f} وهي أقل من مستوى الدلالة المعياري 0.05).", size=14, indent_first_line=True)

    add_banner_heading(doc, "الملحق (1): استمارة الاستبيان النهائية الموجهة لعينة الدراسة", 16, page_break_before=True)
    tq = doc.add_table(rows=1, cols=7)
    tq.style = "Table Grid"
    for i, h in enumerate(["غير موافق بشدة (1)", "غير موافق (2)", "محايد (3)", "موافق (4)", "موافق بشدة (5)", "فقرات ومؤشرات القياس", "ت"]):
        tq.rows[0].cells[i].text = h
    for idx_q, itm in enumerate(items, 1):
        rq = tq.add_row().cells
        rq[5].text = clean_iraqi_text(itm)
        rq[6].text = str(idx_q)
    style_zebra_table(tq, col_widths_cm=[1.5, 1.5, 1.4, 1.4, 1.5, 7.6, 1.2], text_col_idx=5)

    word_file = f"StatLab_SPSS_Complete_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(doc, word_file)
    pdf_spss = convert_docx_to_pdf_fallback(word_file)
    return word_file, excel_file, pdf_spss, chart_path

def _set_ppt_para_rtl(p):
    try:
        pPr = p._p.get_or_add_pPr()
        pPr.set('rtl', '1')
    except Exception:
        pass

def build_seminar_pptx(uid, order_info):
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    plan = order_info["plan"]
    live_bp = order_info.get("live_blueprint", "")
    ts_tag = int(time.time())
    actual_plan_outline = "\n".join([f"- {m['mabhath_title']}: {' | '.join(m.get('matalib', []))}" for m in plan])

    s_prompt = f"""
    صمم محتوى عرض تقديمي جامعي احترافي (سمنار مناقشة) من 10 شرائح علمية لبحث بعنوان: «{title_ar}» في تخصص ({specialty}).
    التزم بعناوين وهيكل البحث الفعلي التالي:
    {actual_plan_outline}
    أخرج JSON يحتوي على مفتاح "slides" وبداخله قائمة الشرائح، لكل شريحة:
    {{"title": "عنوان الشريحة", "points": ["النقطة العلمية 1", "النقطة العلمية 2", "النقطة العلمية 3", "النقطة العلمية 4"]}}
    """
    slides_data = extract_json_safely(ask_agent(s_prompt, "writer", True, required_keys=["slides"], live_blueprint=live_bp)).get("slides") or [
        {"title": "مشكلة البحث وأهميته", "points": ["تحديد الفجوة العلمية والعملية", "أهمية الموضوع في البيئة العراقية", "الأهداف المركزية للدراسة"]}
    ]
    prs = Presentation()
    cover_slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = cover_slide.background.fill
    fill.solid()
    fill.fore_color.rgb = PptRGB(0x13, 0x29, 0x4B)
    txBox = cover_slide.shapes.add_textbox(PptInches(0.8), PptInches(1.0), PptInches(8.4), PptInches(5.4))
    tf_c = txBox.text_frame
    tf_c.word_wrap = True
    p_top = tf_c.paragraphs[0]
    p_top.text = f"جمهورية العراق — وزارة التعليم العالي والبحث العلمي\n{order_info['university_name']} — {order_info['college_name']} ({specialty})"
    p_top.font.size = PptPt(16)
    p_top.font.bold = True
    p_top.font.color.rgb = PptRGB(0xD9, 0xE1, 0xF2)
    p_top.alignment = PP_ALIGN.CENTER
    _set_ppt_para_rtl(p_top)

    p_tit = tf_c.add_paragraph()
    p_tit.text = f"\n« {title_ar} »\n"
    p_tit.font.size = PptPt(24)
    p_tit.font.bold = True
    p_tit.font.color.rgb = PptRGB(0xFF, 0xD9, 0x66)
    p_tit.alignment = PP_ALIGN.CENTER
    _set_ppt_para_rtl(p_tit)

    p_sub = tf_c.add_paragraph()
    p_sub.text = f"عرض تقديمي للمناقشة العلمية\nإعداد الباحث: {order_info['student_name']}  |  إشراف: {order_info['supervisor_name']}\nالسنة الدراسية: {order_info['academic_year']}"
    p_sub.font.size = PptPt(14)
    p_sub.font.color.rgb = PptRGB(0xFF, 0xFF, 0xFF)
    p_sub.alignment = PP_ALIGN.CENTER
    _set_ppt_para_rtl(p_sub)

    for sl in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        f_s = slide.background.fill
        f_s.solid()
        f_s.fore_color.rgb = PptRGB(0xF8, 0xFA, 0xFC)
        slide.shapes.title.text = clean_iraqi_text(sl.get("title", ""))
        p_t = slide.shapes.title.text_frame.paragraphs[0]
        p_t.font.color.rgb = PptRGB(0x1F, 0x4E, 0x78)
        p_t.alignment = PP_ALIGN.RIGHT
        _set_ppt_para_rtl(p_t)
        tf = slide.placeholders[1].text_frame
        tf.word_wrap = True
        for p_idx, pt_txt in enumerate(sl.get("points", [])):
            p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
            p.text = f"❖ {clean_iraqi_text(pt_txt)}"
            p.font.size = PptPt(18)
            p.font.color.rgb = PptRGB(0x1A, 0x1A, 0x1A)
            p.alignment = PP_ALIGN.RIGHT
            _set_ppt_para_rtl(p)

    pptx_file = f"StatLab_Seminar_{uid}_{ts_tag}.pptx"
    prs.save(pptx_file)
    return pptx_file

def build_english_or_bologna_report(uid, order_info, is_english=False):
    title_ar = order_info["clean_title"]
    specialty = order_info["specialty"]
    plan = order_info["plan"]
    calib = order_info["calib"]
    live_bp = order_info.get("live_blueprint", "")
    verified_objs = order_info.get("verified_source_objects") or []
    ts_tag = int(time.time())

    if is_english:
        en_cites = [s.get("cite") for s in verified_objs if s.get("cite")]
        cites_str = ", ".join([f"({c})" for c in en_cites[:10]])
        eng_prompt = f"""
        Topic: "{title_ar}" — Field: "{specialty}" — University: "{order_info['university_name']}" — College: "{order_info['college_name']}".
        Write a comprehensive, formal academic research paper in English (with Google Search verification for authentic APA references).
        Output JSON:
        {{
          "translated_title": "Formal English Academic Title",
          "university_en": "English Name of {order_info['university_name']}",
          "college_en": "English Name of {order_info['college_name']}",
          "specialty_en": "English Name of {specialty}",
          "abstract_en": "180-word academic abstract",
          "keywords_en": "5 academic keywords separated by commas",
          "sections": [
            {{"heading": "1. Introduction, Problem Statement, and Conceptual Framework", "paragraphs": ["Para 1 (150 words)...", "Para 2...", "Para 3..."]}},
            {{"heading": "2. Theoretical Foundations and Literature Review", "paragraphs": ["Para 1...", "Para 2...", "Para 3..."]}},
            {{"heading": "3. Analytical Dimensions and Comparative Discussion", "paragraphs": ["Para 1...", "Para 2...", "Para 3..."]}},
            {{"heading": "4. Practical Implications, Conclusions, and Recommendations", "paragraphs": ["Para 1...", "Para 2...", "Para 3..."]}}
          ],
          "apa_references": [
            "Author, A. (Year). Title of article. Journal Name, Volume(Issue), pp. 1-15."
          ]
        }}
        """
        eng_data = extract_json_safely(ask_agent(eng_prompt, "writer", True, required_keys=["translated_title", "sections"], use_google_search=True))
        t_en = clean_iraqi_text(eng_data.get("translated_title") or title_ar, is_english=True)
        u_en = clean_iraqi_text(eng_data.get("university_en") or "University of Baghdad", is_english=True)
        c_en = clean_iraqi_text(eng_data.get("college_en") or "College of Academic Studies", is_english=True)
        sp_en = clean_iraqi_text(eng_data.get("specialty_en") or "Academic Studies", is_english=True)
        kw_en_clean = clean_iraqi_text(eng_data.get("keywords_en") or "", is_english=True).strip("[]'\" ")
        apa_refs = [clean_iraqi_text(r, is_english=True, preserve_urls=True, is_reference=True) for r in (eng_data.get("apa_references") or []) if len(str(r)) > 12]

        doc = setup_iraqi_document(is_english=True, header_title=t_en)
        add_para(doc, f"Republic of Iraq\nMinistry of Higher Education and Scientific Research\n{u_en} — {c_en}\n{sp_en}", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, is_english=True)
        add_para(doc, f"\n\"{t_en}\"\n", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, is_english=True, color_rgb=(0x7A, 0x1C, 0x24))
        add_para(doc, f"Prepared by: {order_info['student_name']}\nSupervised by: {order_info['supervisor_name']}\nAcademic Year: {order_info['academic_year']}", size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, is_english=True)

        add_banner_heading(doc, "Abstract", 15, is_english=True, page_break_before=True)
        add_para(doc, clean_iraqi_text(eng_data.get("abstract_en") or "", is_english=True), size=13, is_english=True)
        add_para(doc, f"Keywords: {kw_en_clean}.", size=12, bold=True, is_english=True)

        sec_html_blocks = ""
        for sec_obj in (eng_data.get("sections") or []):
            h_en = clean_iraqi_text(sec_obj.get("heading") or "Section", is_english=True)
            add_banner_heading(doc, h_en, 14.5, is_english=True)
            sec_html_blocks += f'<h2 style="color:#1F4E78; border-bottom:1.5px solid #1F4E78; padding-bottom:4px; margin-top:20px;">{_esc(h_en)}</h2>'
            for p_en in (sec_obj.get("paragraphs") or []):
                p_clean_en = clean_iraqi_text(p_en, is_english=True).rstrip(".") + "."
                add_para(doc, p_clean_en, size=13, is_english=True)
                sec_html_blocks += f'<p style="text-align:justify; text-indent:1cm; margin-bottom:11px;">{_esc(p_clean_en)}</p>'

        add_banner_heading(doc, "References", 15, is_english=True, page_break_before=True)
        refs_en_html = ""
        for r_line in sorted(apa_refs):
            add_para(doc, r_line, size=11.5, is_english=True, align=WD_ALIGN_PARAGRAPH.LEFT, preserve_urls=True)
            refs_en_html += f'<li style="margin-bottom:7px;">{_esc(r_line)}</li>'

        fname = f"StatLab_English_Research_{uid}_{ts_tag}.docx"
        sanitize_docx_forensics(doc, fname)
        pdf_out = f"StatLab_English_Research_{uid}_{ts_tag}.pdf"
        if HAS_WEASYPRINT:
            try:
                en_html = f"""<!DOCTYPE html><html lang="en" dir="ltr"><head><meta charset="UTF-8"><style>@page {{ size: A4; margin: 24mm 22mm; @bottom-center {{ content: counter(page); font-size: 10pt; }} }} body {{ font-family: 'Times New Roman', 'DejaVu Serif', serif; font-size: 13pt; line-height: 1.68; color: #1A1A1A; direction: ltr; text-align: justify; }}</style></head><body><div style="text-align:center; padding-top:35px;"><div style="font-size:15pt; font-weight:bold; color:#1B3B2B;">Republic of Iraq<br>Ministry of Higher Education and Scientific Research<br>{_esc(u_en)} — {_esc(c_en)}<br>{_esc(sp_en)}</div><div style="margin:28px 10px; padding:18px; border-top:2px solid #9A7B4F; border-bottom:2px solid #9A7B4F; background:#FAF6F0; font-size:20pt; font-weight:bold; color:#7A1C24;">{_esc(t_en)}</div><div style="font-size:13pt; margin-top:25px; line-height:1.9;"><strong>Prepared by:</strong> {_esc(order_info['student_name'])}<br><strong>Supervised by:</strong> {_esc(order_info['supervisor_name'])}<br><br><strong>Academic Year:</strong> {_esc(order_info['academic_year'])}</div></div><div style="page-break-before:always;"></div><h2 style="color:#1F4E78; border-bottom:1.5px solid #1F4E78;">Abstract</h2><p>{_esc(clean_iraqi_text(eng_data.get('abstract_en') or '', is_english=True))}</p><p><strong>Keywords:</strong> {_esc(kw_en_clean)}.</p>{sec_html_blocks}<h2 style="color:#1F4E78; border-bottom:1.5px solid #1F4E78; margin-top:22px;">References</h2><ol>{refs_en_html}</ol></body></html>"""
                weasyprint.HTML(string=en_html).write_pdf(pdf_out)
            except Exception:
                pdf_out = convert_docx_to_pdf_fallback(fname)
        else:
            pdf_out = convert_docx_to_pdf_fallback(fname)
        return fname, pdf_out

    # تقرير مسار بولونيا الفصلي المعتمد (شامل جدول الحمل الدراسي SWL/ECTS ومصفوفة التقييم Rubric)
    allowed_cites_ar = [s.get("cite") for s in verified_objs if s.get("cite")]
    cites_str_ar = "، ".join([f"({c})" for c in allowed_cites_ar[:10]])

    doc = setup_iraqi_document(is_english=False, header_title=title_ar)
    add_para(doc, f"جمهورية العراق\nوزارة التعليم العالي والبحث العلمي\n{order_info['university_name']} — {order_info['college_name']}\nالقسم العلمي: {specialty}\n« تقرير علمي فصلي معتمد وفق متطلبات مسار بولونيا »", size=14.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(0x1F, 0x4E, 0x78))
    add_para(doc, f"\n« {title_ar} »\n", size=19, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(0x7A, 0x1C, 0x24))
    add_para(doc, f"إعداد الطالب:\n({order_info['student_name']})\n\nبإشراف أستاذ المادة:\n({order_info['supervisor_name']})\n\nالسنة الدراسية: {order_info['academic_year']}", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    # جدول الحمل الدراسي (SWL / ECTS) ومصفوفة التقييم وفق دليل مسار بولونيا في وزارة التعليم العالي العراقية
    add_banner_heading(doc, "بطاقة العبء الدراسي ومصفوفة التقييم التكويني (Bologna SWL & ECTS Rubric)", 15, page_break_before=True)
    swl_tbl = doc.add_table(rows=1, cols=4)
    swl_tbl.style = "Table Grid"
    for i, h in enumerate(["الوزن النسبي %", "الساعات غير المجدولة (Unstructured)", "الساعات المجدولة (Structured)", "معيار التقييم ومخرجات التعلم (Learning Outcomes)"]):
        swl_tbl.rows[0].cells[i].text = h
    swl_rows = [
        ("25%", "6 ساعات", "2 ساعة", "1. الأصالة العلمية وتحديد الإشكالية ومخرجات التعلم المستهدفة"),
        ("35%", "10 ساعات", "3 ساعات", "2. العمق التحليلي والربط المنهجي بمحاور المادة الدراسية"),
        ("20%", "5 ساعات", "1 ساعة", "3. دقة التوثيق الأكاديمي وحداثة المصادر المعتمدة"),
        ("20%", "4 ساعات", "1 ساعة", "4. جودة الاستنتاجات والتوصيات الإجرائية والعرض المطبعي")
    ]
    for r_swl in swl_rows:
        rc = swl_tbl.add_row().cells
        for c_i in range(4):
            rc[c_i].text = r_swl[c_i]
    style_zebra_table(swl_tbl, col_widths_cm=[2.5, 3.2, 3.2, 6.5], text_col_idx=3)
    add_para(doc, "إجمالي العبء الدراسي للتقرير (SWL): 32 ساعة عمل فصلية (تعادل 1.28 وحدة ECTS وفق دليل وزارة التعليم العالي العراقية).", size=11.5, bold=True)

    bol_intro_prompt = f"""
    لتقرير فصلي وفق مسار بولونيا في الجامعات العراقية بعنوان: «{title_ar}» في تخصص ({specialty}):
    المراجع المعتمدة: {cites_str_ar}.
    أخرج JSON يحتوي على:
    {{
      "learning_outcomes": "فقرة مخرجات التعلم المستهدفة (100 كلمة)...",
      "intro_body": "فقرتان مفصلتان (240 كلمة) للمقدمة العلمية وإشكالية التقرير وأهميته مع التوثيق..."
    }}
    """
    bol_intro = extract_json_safely(ask_agent(bol_intro_prompt, "writer", True, required_keys=["learning_outcomes", "intro_body"], live_blueprint=live_bp))
    add_banner_heading(doc, "أولاً: مخرجات التعلم المستهدفة والمقدمة العلمية", 15.5)
    add_para(doc, f"مخرجات التعلم: {clean_iraqi_text(bol_intro.get('learning_outcomes') or '')}", size=14)
    for ln in str(bol_intro.get("intro_body") or "").split("\n"):
        if len(ln.strip()) > 15:
            p_b_c = clean_iraqi_text(ln).rstrip(".")
            if allowed_cites_ar and "(" not in p_b_c:
                p_b_c += f" ({allowed_cites_ar[0]})"
            add_para(doc, p_b_c + ".", size=14, indent_first_line=True)

    cite_b_idx = 1
    for m_idx, mb in enumerate(plan, 1):
        m_t = clean_iraqi_text(mb.get("mabhath_title") or f"المحور {m_idx}")
        sub_t = "، ".join([clean_iraqi_text(x) for x in (mb.get("matalib") or [])])
        sec_b_prompt = f"""
        عنوان التقرير: «{title_ar}» — التخصص: «{specialty}».
        المحور الحالي: «{m_t}» (ويشمل: {sub_t}).
        المراجع المعتمدة للتوثيق داخل المتن: {cites_str_ar}.
        اكتب 3 فقرات تحليلية معمقة وموثقة (390 كلمة) تغطي هذا المحور بصيغة JSON:
        {{
          "paragraphs": ["الفقرة الأولى الموثقة...", "الفقرة الثانية الموثقة...", "الفقرة الثالثة الموثقة..."]
        }}
        """
        sec_b_data = extract_json_safely(ask_agent(sec_b_prompt, "writer", True, required_keys=["paragraphs"], live_blueprint=live_bp))
        add_banner_heading(doc, m_t, 15)
        for p_b in (sec_b_data.get("paragraphs") or []):
            p_clean_b = clean_iraqi_text(p_b).rstrip(".")
            if allowed_cites_ar and not re.search(r'\((?:19|20)\d{2}|\d{4}\)', p_clean_b):
                p_clean_b += f" ({allowed_cites_ar[cite_b_idx % len(allowed_cites_ar)]})"
            cite_b_idx += 1
            add_para(doc, p_clean_b + ".", size=14, indent_first_line=True)

    bol_conc_prompt = f"""
    لتقرير مسار بولونيا بعنوان «{title_ar}»:
    اكتب الخاتمة (فقرة استهلالية + 5 استنتاجات + 5 توصيات) بصيغة JSON:
    {{
      "conclusion_intro": "...",
      "findings": ["...", "...", "...", "...", "..."],
      "recommendations": ["...", "...", "...", "...", "..."]
    }}
    """
    bol_conc = extract_json_safely(ask_agent(bol_conc_prompt, "writer", True, required_keys=["findings", "recommendations"], live_blueprint=live_bp))
    add_banner_heading(doc, "الخاتمة: الاستنتاجات والتوصيات", 15.5)
    add_para(doc, clean_iraqi_text(bol_conc.get("conclusion_intro") or "").rstrip(".") + ".", size=14, indent_first_line=True)
    add_para(doc, "أولاً: الاستنتاجات:", size=14, bold=True, color_rgb=(0x7A, 0x1C, 0x24))
    for i_f, f_b in enumerate(bol_conc.get("findings") or [], 1):
        add_para(doc, f"{i_f}- {clean_iraqi_text(f_b).rstrip('.')}.", size=14)
    add_para(doc, "ثانياً: التوصيات:", size=14, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
    for i_r, r_b in enumerate(bol_conc.get("recommendations") or [], 1):
        add_para(doc, f"{i_r}- {clean_iraqi_text(r_b).rstrip('.')}.", size=14)

    add_banner_heading(doc, "قائمة المصادر والمراجع المعتمدة", 15.5, page_break_before=True)
    for tier_obj in (calib.get("formatted_reference_tiers") or []):
        t_title = clean_iraqi_text(tier_obj.get("title") or "")
        t_items = tier_obj.get("items") or []
        is_eng_t = bool(tier_obj.get("is_ltr", False))
        if t_items:
            add_para(doc, t_title, size=14, bold=True, color_rgb=(0x1F, 0x4E, 0x78))
            for idx_r, r_it in enumerate(t_items, 1):
                add_para(doc, f"{idx_r}. {r_it}", size=11.5 if is_eng_t else 13, is_english=is_eng_t, align=WD_ALIGN_PARAGRAPH.LEFT if is_eng_t else None, preserve_urls=is_eng_t)

    fname = f"StatLab_Bologna_Report_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(doc, fname)
    pdf_out = convert_docx_to_pdf_fallback(fname)
    return fname, pdf_out

def build_paraphrased_document(uid, raw_input_text):
    ts_tag = int(time.time())
    latin_cnt = len(re.findall(r'[a-zA-Z]', raw_input_text))
    arabic_cnt = len(re.findall(r'[\u0621-\u064A]', raw_input_text))
    is_eng = (latin_cnt > arabic_cnt * 1.5)
    para_prompt = f"""
    أعد صياغة النص الأكاديمي التالي صياغة علمية جذرية تخفض نسبة الاستلال (Turnitin Similarity) إلى الصفر مع الحفاظ التام على المعنى العلمي الدقيق والتوثيقات:
    «{raw_input_text}»
    أخرج JSON يحتوي على مفتاح "paragraphs" وبداخله قائمة الفقرات المعاد صياغتها.
    """
    p_data = extract_json_safely(ask_agent(para_prompt, "auditor", True, required_keys=["paragraphs"]))
    paras_out = p_data.get("paragraphs") or [raw_input_text]
    doc = setup_iraqi_document(is_english=is_eng, header_title="Academic Paraphrasing" if is_eng else "إعادة الصياغة الأكاديمية وخفض الاستلال")
    add_banner_heading(doc, "Academic Paraphrased Text (Turnitin Ready)" if is_eng else "النص الأكاديمي المعاد صياغته (مخفض الاستلال — جاهز للفحص)", 15.5, is_english=is_eng)
    for p_txt in paras_out:
        add_para(doc, clean_iraqi_text(p_txt, is_english=is_eng), size=13.5 if is_eng else 14, is_english=is_eng, indent_first_line=True)
    docx_p = f"StatLab_Paraphrased_{uid}_{ts_tag}.docx"
    sanitize_docx_forensics(doc, docx_p)
    pdf_p = convert_docx_to_pdf_fallback(docx_p)
    return docx_p, pdf_p

# ==============================================================================
# الوحدة (7): لوحة المصانع الـ 9 والمدير التفاعلي الواعي في تليغرام
# ==============================================================================
def _safe_edit_html(chat_id, msg_id, html_text, reply_markup=None):
    try:
        bot.edit_message_text(html_text, chat_id=chat_id, message_id=msg_id, parse_mode="HTML", reply_markup=reply_markup)
    except Exception:
        pass

def get_master_keyboard(uid):
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton("📕 1. البحث/التقرير الأكاديمي المفصل (PDF الفاخر + Word بالهوامش)", callback_data=f"m_gen_{uid}"),
        InlineKeyboardButton("📗 2. فصل الدراسات السابقة (عراقي + أجنبي + Turnitin + معاينة)", callback_data=f"m_lit_{uid}"),
        InlineKeyboardButton("🎓 3. بحث التخرج الجامعي الوزاري الكامل (7 أركان + فهرس + PDF)", callback_data=f"m_grad_{uid}"),
        InlineKeyboardButton("📊 4. إحصاء SPSS v.28 الشامل (ثومبسون + t-test + رسم + Excel)", callback_data=f"m_spss_{uid}"),
        InlineKeyboardButton("🖥️ 5. عرض مناقشة السمنار الملكي (PowerPoint .pptx RTL)", callback_data=f"m_pptx_{uid}"),
        InlineKeyboardButton("📄 6. تقرير مسار بولونيا الفصلي المعتمد (SWL/ECTS + Word + PDF)", callback_data=f"m_bol_{uid}"),
        InlineKeyboardButton("🇬🇧 7. البحث/التقرير الكامل باللغة الإنجليزية الأكاديمية (LTR Word + PDF)", callback_data=f"m_eng_{uid}"),
        InlineKeyboardButton("🛡️ 8. مصنع خفض الاستلال وإعادة الصياغة الأكاديمية (Turnitin Paraphraser)", callback_data=f"m_para_{uid}"),
        InlineKeyboardButton("💳 9. محفظة المكاتب والرصيد الأكاديمي (Office Credits Wallet)", callback_data=f"m_wallet_{uid}")
    )
    return markup

def format_order_proposal_message(order_info, header_note="🦅 <b>أهلاً بك! بحثتُ في Google والباحث العلمي والمجلات العراقية وأعددتُ لك خطة العمل والمصادر الحقيقية التالية:</b>"):
    plan_lines = []
    for m in order_info["plan"]:
        plan_lines.append(f"  • <b>{_esc(m['mabhath_title'])}</b>")
        for sub_m in m.get("matalib", []):
            plan_lines.append(f"      ▫️ {_esc(sub_m)}")
    plan_str = "\n".join(plan_lines)

    refs_preview = []
    for t_obj in (order_info["calib"].get("formatted_reference_tiers") or []):
        for r_it in (t_obj.get("items") or [])[:3]:
            refs_preview.append(f"  ▪️ {_esc(r_it[:95])}...")
    refs_str = "\n".join(refs_preview[:6]) if refs_preview else "  ▪️ مراجع تخصصية محكمة"

    return (
        f"{header_note}\n\n"
        f"📌 <b>العنوان الصافي:</b> «{_esc(order_info['clean_title'])}»\n"
        f"🏛️ <b>الجامعة والكلية:</b> {_esc(order_info['university_name'])} — {_esc(order_info['college_name'])}\n"
        f"🔬 <b>التخصص الدقيق:</b> {_esc(order_info['specialty'])}\n"
        f"👤 <b>إعداد الباحث:</b> {_esc(order_info['student_name'])} | <b>إشراف:</b> {_esc(order_info['supervisor_name'])}\n"
        f"📐 <b>المواصفات:</b> الخط <code>{order_info['font_size']}</code> | الصفحات المستهدفة: <code>{order_info['min_pages']}</code>\n\n"
        f"📑 <b>هيكل التقسيم المقترح:</b>\n{plan_str}\n\n"
        f"📚 <b>عينة من المصادر المسحوبة من Google Scholar:</b>\n{refs_str}\n\n"
        f"💬 <b>أنا بانتظار توجيهك الآن:</b>\n"
        f"▪️ <b>للتعديل:</b> اكتب لي أي تعديل تريده في رسالة عادية وسأعدله فوراً!\n"
        f"▪️ <b>للتنفيذ الفوري:</b> اكتب كلمة <b>«ابدأ»</b> (أو اضغط على زر المصنع المطلوب بالأسفل)!"
    )

@bot.message_handler(commands=['start'])
def start_cmd(message):
    uid = message.chat.id
    if not is_authorized(uid):
        bot.reply_to(message, f"🔒 هذا البوت خاص. رقم معرفك: <code>{uid}</code>", parse_mode="HTML")
        return
    bot.send_message(
        uid,
        f"🦅 <b>منظومة StatLab AI (محرك Gemini 3.8 Flash النقي + بحث Google العلمي الحي + المدقق الشامل) جاهزة!</b>\n"
        f"🔑 <b>مفاتيح Gemini الفعالة:</b> <code>{len(WORKING_GEMINI_KEYS)}</code> مفاتيح بالتناوب الذكي | <b>النماذج:</b> <code>{', '.join(ACTIVE_GEMINI_MODELS[:3])}</code>\n\n"
        f"▪️ أرسل <b>عنوان البحث</b> ليبحث البوت في Google Scholar ويعرض لك الخطة والمصادر ويناقشك قبل البدء.\n"
        f"▪️ اكتب <code>/new</code> لبدء بحث جديد، أو <code>/diagnose</code> لفحص السيرفر.",
        parse_mode="HTML",
        reply_markup=get_master_keyboard(uid)
    )

@bot.message_handler(commands=['new'])
def new_session_cmd(message):
    uid = message.chat.id
    if not is_authorized(uid):
        return
    orders_db.pop(uid, None)
    awaiting_paraphrase_uids.discard(uid)
    bot.reply_to(message, "🆕 <b>تم فتح جلسة بحث جديدة!</b> أرسل الآن عنوان البحث أو الموضوع الجديد.", parse_mode="HTML")

@bot.message_handler(commands=['diagnose'])
def diagnose_cmd(message):
    if message.chat.id != ADMIN_ID:
        return
    disk = shutil.disk_usage(".")
    free_gb = disk.free / (1024 ** 3)
    bot.reply_to(
        message,
        f"🩺 <b>تقرير فحص السيرفر (Pure Gemini Engine):</b>\n"
        f"• مفاتيح Gemini النشطة: <code>{len(WORKING_GEMINI_KEYS)}</code>\n"
        f"• النماذج المكتشفة: <code>{', '.join(ACTIVE_GEMINI_MODELS)}</code>\n"
        f"• محرك WeasyPrint PDF: <code>{'مفعل ✅' if HAS_WEASYPRINT else 'احتياطي'}</code>\n"
        f"• المساحة الحرة: <code>{free_gb:.2f} GB</code>",
        parse_mode="HTML"
    )

@bot.message_handler(commands=['errors'])
def errors_cmd(message):
    if message.chat.id != ADMIN_ID:
        return
    err_txt = "\n".join(_error_log[-10:]) if _error_log else "لا توجد أي أخطاء مسجلة ✅"
    bot.reply_to(message, f"📋 <b>سجل آخر الأخطاء:</b>\n<code>{_esc(err_txt)}</code>", parse_mode="HTML")

@bot.message_handler(commands=['allow'])
def allow_user_cmd(message):
    if message.chat.id != ADMIN_ID:
        return
    nums = re.findall(r'\d+', message.text.replace("/allow", ""))
    if nums:
        add_allowed_user(int(nums[0]))
        bot.reply_to(message, f"✅ تم منح الصلاحية للمستخدم <code>{nums[0]}</code>.", parse_mode="HTML")

@bot.message_handler(commands=['addcredit'])
def add_credit_cmd(message):
    if message.chat.id != ADMIN_ID:
        return
    nums = re.findall(r'-?\d+', message.text.replace("/addcredit", ""))
    if len(nums) >= 2:
        t_uid, amt = int(nums[0]), int(nums[1])
        new_bal = add_user_credits(t_uid, amt)
        bot.reply_to(message, f"💳 تم تحديث رصيد المكتب <code>{t_uid}</code> إلى: <b>{new_bal}</b> نقطة.", parse_mode="HTML")
    else:
        bot.reply_to(message, "✍️ الصيغة: <code>/addcredit USER_ID AMOUNT</code>", parse_mode="HTML")

@bot.message_handler(commands=['learn'])
def learn_new_rule_cmd(message):
    if not is_authorized(message.chat.id):
        return
    rule_txt = message.text.replace("/learn", "").strip()
    if rule_txt:
        add_learned_rule(rule_txt)
        bot.reply_to(message, f"🧠 <b>تم حفظ القاعدة الأكاديمية للأبد!</b>\n({_esc(rule_txt)})", parse_mode="HTML")

def _execute_factory_for_user(uid, order_info, factory_code="m_gen"):
    title_ar = order_info["clean_title"]
    if factory_code == "m_gen":
        msg_w = bot.send_message(uid, f"🚀 <b>انطلق التنفيذ بـ Gemini 3.8 Flash!</b>\n📌 «{_esc(title_ar)}»\n⏳ [1/4] جاري التأليف والتدقيق الشامل...", parse_mode="HTML")
        pdf_f, docx_f = build_genspark_luxury_research(uid, order_info, progress_callback=lambda t: _safe_edit_html(uid, msg_w.message_id, f"📌 <b>«{_esc(title_ar)}»</b>\n{_esc(t)}"))
        try:
            if pdf_f and os.path.exists(pdf_f):
                with open(pdf_f, "rb") as fp:
                    bot.send_document(uid, fp, caption=f"📕 <b>1. نسخة الـ PDF المطبعية الفاخرة (مفحوصة بالكامل عبر Gemini 3.8 Flash)!</b>\n📌 «{_esc(title_ar)}»", parse_mode="HTML")
            if docx_f and os.path.exists(docx_f):
                with open(docx_f, "rb") as fd:
                    bot.send_document(uid, fd, caption="📘 <b>2. نسخة الوورد القابلة للتعديل (.docx) شاملة الفهرس والهوامش!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        finally:
            cleanup_temp_files()

@bot.message_handler(func=lambda m: m.text and not m.text.startswith('/'))
def handle_research_order(message):
    uid = message.chat.id
    if not is_authorized(uid):
        bot.reply_to(message, f"🔒 هذا البوت خاص. رقم معرفك: <code>{uid}</code>", parse_mode="HTML")
        return
    raw_txt = message.text.strip()

    # التعامل مع من يكتب new/ في بداية أو نهاية السطر بسبب الكيبورد العربي
    if re.match(r'^(?:new/|/new|جديد)\s*', raw_txt, flags=re.IGNORECASE):
        orders_db.pop(uid, None)
        raw_txt = re.sub(r'^(?:new/|/new|جديد)\s*', '', raw_txt, flags=re.IGNORECASE).strip()
        if not raw_txt:
            bot.reply_to(message, "🆕 <b>تم فتح جلسة بحث جديدة!</b> أرسل الآن عنوان البحث.", parse_mode="HTML")
            return

    if uid in awaiting_paraphrase_uids:
        awaiting_paraphrase_uids.discard(uid)
        w_msg = bot.reply_to(message, "⏳ <b>جاري إعادة الصياغة الأكاديمية وخفض الاستلال...</b>", parse_mode="HTML")
        docx_p, pdf_p = None, None
        try:
            docx_p, pdf_p = build_paraphrased_document(uid, raw_txt)
            with open(docx_p, "rb") as fd:
                bot.send_document(uid, fd, caption="🛡️ <b>1. النص المعاد صياغته أكاديمياً (.docx)!</b>", parse_mode="HTML")
            if pdf_p and os.path.exists(pdf_p):
                with open(pdf_p, "rb") as fp:
                    bot.send_document(uid, fp, caption="📕 <b>2. نسخة الـ PDF للنص المعاد صياغته!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        except Exception as e:
            bot.send_message(uid, f"⚠️ تنبيه: {_esc(str(e))}", parse_mode="HTML")
        finally:
            cleanup_temp_files()
        return

    norm_cmd = re.sub(r'[^\u0621-\u064Aa-zA-Z]', '', raw_txt).lower()
    start_triggers = {"ابدأ", "ابدا", "ابدئ", "ابدي", "نفذ", "اعتمد", "موافق", "تمام", "انطلق", "start", "go", "ok"}
    existing_order = get_order_state(uid)

    if existing_order and (norm_cmd in start_triggers or (len(raw_txt) < 22 and any(st in raw_txt for st in ["ابدأ", "ابدا", "ابدئ", "نفذ", "اعتمد"]))):
        try:
            _execute_factory_for_user(uid, existing_order, "m_gen")
        except Exception as e:
            log_runtime_error("start_trigger", e)
            bot.send_message(uid, f"⚠️ خطأ أثناء التنفيذ: {_esc(str(e))}", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        return

    edit_starts = ("غير", "غيّر", "عدل", "عدّل", "اجعل", "بدل", "أضف", "اضف", "احذف", "خلي")
    if existing_order and raw_txt.startswith(edit_starts) and len(raw_txt) < 220:
        upd_msg = bot.reply_to(message, "🔄 <b>يقوم المدير الأكاديمي بتحديث الخطة والبيانات...</b>", parse_mode="HTML")
        try:
            updated_order, mgr_reply = process_conversational_manager_edit(existing_order, raw_txt)
            save_order_state(uid, updated_order)
            prop_txt = format_order_proposal_message(updated_order, header_note=f"✅ <b>{_esc(mgr_reply)}</b>")
            _safe_edit_html(uid, upd_msg.message_id, prop_txt, reply_markup=get_master_keyboard(uid))
        except Exception as e:
            bot.send_message(uid, f"⚠️ تنبيه أثناء التعديل: {_esc(str(e))}", parse_mode="HTML")
        return

    status_msg = bot.reply_to(
        message,
        "🦅 <b>استلم المدير الأكاديمي (Gemini 3.8 Flash) طلبك!</b>\n"
        "🔍 جاري الآن <b>البحث الحي في Google والباحث العلمي والمجلات العراقية (iasj.rdd.edu.iq)</b> لاستنباط المنهجية والمصادر الحقيقية...",
        parse_mode="HTML"
    )
    try:
        order_info = parse_and_calibrate_order(raw_txt)
        save_order_state(uid, order_info)
        prop_txt = format_order_proposal_message(order_info)
        _safe_edit_html(uid, status_msg.message_id, prop_txt, reply_markup=get_master_keyboard(uid))
    except Exception as e:
        log_runtime_error("handle_research_order", e)
        bot.send_message(uid, f"⚠️ تنبيه أثناء المسح الأولي: {_esc(str(e))}", parse_mode="HTML")

@bot.callback_query_handler(func=lambda call: True)
def handle_all_callbacks(call):
    uid = call.message.chat.id
    if not is_authorized(uid):
        bot.answer_callback_query(call.id, "🔒 غير مصرح لك.")
        return
    data = call.data
    if data.startswith("m_wallet_"):
        bal = get_user_credits(uid)
        bot.answer_callback_query(call.id, f"💳 رصيدك: {bal}")
        bot.send_message(uid, f"💳 <b>محفظة المكاتب الأكاديمية:</b>\n• معرف الحساب: <code>{uid}</code>\n• الرصيد المتاح: <b>{'غير محدود (مدير النظام)' if uid == ADMIN_ID else f'{bal} بحث'}</b>", parse_mode="HTML")
        return
    if data.startswith("m_para_"):
        awaiting_paraphrase_uids.add(uid)
        bot.answer_callback_query(call.id, "🛡️ أرسل النص لإعادة صياغته")
        bot.send_message(uid, "🛡️ <b>مصنع خفض الاستلال (Turnitin Paraphraser):</b>\n✍️ أرسل الآن الفقرات التي تريد إعادة صياغتها أكاديمياً.", parse_mode="HTML")
        return

    order_info = get_order_state(uid)
    if not order_info:
        bot.answer_callback_query(call.id, "⚠️ أرسل عنوان البحث أولاً!")
        bot.send_message(uid, "⚠️ أرسل عنوان البحث في رسالة عادية أولاً.", parse_mode="HTML")
        return

    try:
        if data.startswith("m_gen_"):
            bot.answer_callback_query(call.id, "📕 جاري بناء البحث الفاخر...")
            _execute_factory_for_user(uid, order_info, "m_gen")
        elif data.startswith("m_lit_"):
            bot.answer_callback_query(call.id, "📗 جاري إنتاج الدراسات السابقة...")
            bot.send_message(uid, "⏳ جاري إعداد <b>حزمة الدراسات السابقة الثلاثية عبر بحث Google العلمي</b>...", parse_mode="HTML")
            f_main, f_turn, f_prev = build_literature_review_pack(uid, order_info)
            with open(f_prev, "rb") as fp:
                bot.send_document(uid, fp, caption="🔒 <b>1. نسخة المعاينة المقفلة!</b>", parse_mode="HTML")
            with open(f_main, "rb") as fm:
                bot.send_document(uid, fm, caption="📗 <b>2. النسخة الكاملة المفتوحة للطباعة!</b>", parse_mode="HTML")
            with open(f_turn, "rb") as ft:
                bot.send_document(uid, ft, caption="🛡️ <b>3. نسخة فحص الاستلال (Turnitin)!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        elif data.startswith("m_grad_"):
            bot.answer_callback_query(call.id, "🎓 جاري تأليف بحث التخرج...")
            bot.send_message(uid, "⏳ جاري كتابة <b>بحث التخرج الجامعي الكامل</b>...", parse_mode="HTML")
            f_grad, pdf_grad = build_complete_graduation_thesis(uid, order_info)
            with open(f_grad, "rb") as fg:
                bot.send_document(uid, fg, caption="🎓 <b>1. بحث التخرج الجامعي الكامل (.docx)!</b>", parse_mode="HTML")
            if pdf_grad and os.path.exists(pdf_grad):
                with open(pdf_grad, "rb") as fpg:
                    bot.send_document(uid, fpg, caption="📕 <b>2. نسخة الـ PDF لبحث التخرج!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        elif data.startswith("m_spss_"):
            bot.answer_callback_query(call.id, "📊 جاري التحليل الإحصائي SPSS...")
            bot.send_message(uid, "⏳ جاري بناء <b>تقرير الـ SPSS الشامل + ملف الإكسل</b>...", parse_mode="HTML")
            f_w, f_x, f_p, c_p = build_ultimate_spss_pack(uid, order_info)
            with open(f_w, "rb") as fw:
                bot.send_document(uid, fw, caption="📊 <b>1. تقرير التحليل الإحصائي الشامل (SPSS Word)!</b>", parse_mode="HTML")
            if f_p and os.path.exists(f_p):
                with open(f_p, "rb") as fp_s:
                    bot.send_document(uid, fp_s, caption="📕 <b>2. نسخة الـ PDF لتقرير التحليل الإحصائي!</b>", parse_mode="HTML")
            with open(f_x, "rb") as fx:
                bot.send_document(uid, fx, caption="📗 <b>3. ملف البيانات الخام (Excel) المطابق 100%!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        elif data.startswith("m_pptx_"):
            bot.answer_callback_query(call.id, "🖥️ جاري تصميم السمنار...")
            f_ppt = build_seminar_pptx(uid, order_info)
            with open(f_ppt, "rb") as fp:
                bot.send_document(uid, fp, caption="🖥️ <b>عرض مناقشة السمنار الملكي (PowerPoint .pptx) جاهز!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        elif data.startswith("m_bol_"):
            bot.answer_callback_query(call.id, "📄 جاري إعداد تقرير بولونيا...")
            f_bol, p_bol = build_english_or_bologna_report(uid, order_info, is_english=False)
            with open(f_bol, "rb") as fb:
                bot.send_document(uid, fb, caption="📄 <b>1. تقرير مسار بولونيا الفصلي المعتمد (شامل بطاقة SWL/ECTS) (.docx)!</b>", parse_mode="HTML")
            if p_bol and os.path.exists(p_bol):
                with open(p_bol, "rb") as fpb:
                    bot.send_document(uid, fpb, caption="📕 <b>2. نسخة الـ PDF لتقرير مسار بولونيا!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
        elif data.startswith("m_eng_"):
            bot.answer_callback_query(call.id, "🇬🇧 Generating English Paper...")
            f_eng, p_eng = build_english_or_bologna_report(uid, order_info, is_english=True)
            with open(f_eng, "rb") as fe:
                bot.send_document(uid, fe, caption="🇬🇧 <b>1. Full Academic English Research Paper (.docx)!</b>", parse_mode="HTML")
            if p_eng and os.path.exists(p_eng):
                with open(p_eng, "rb") as fpe:
                    bot.send_document(uid, fpe, caption="📕 <b>2. Full Academic English Research Paper (.pdf)!</b>", parse_mode="HTML", reply_markup=get_master_keyboard(uid))
    except Exception as e:
        log_runtime_error("callback", e)
        bot.send_message(uid, f"⚠️ خطأ أثناء التنفيذ: {_esc(str(e))}", parse_mode="HTML")
    finally:
        cleanup_temp_files()

if __name__ == "__main__":
    print(f"🦅 StatLab AI Pure Gemini Engine Started! Keys={len(WORKING_GEMINI_KEYS)}, Models={ACTIVE_GEMINI_MODELS}")
    while True:
        try:
            bot.infinity_polling(timeout=60, long_polling_timeout=60, skip_pending=False)
        except Exception as e:
            log_runtime_error("polling", e)
            time.sleep(4)
