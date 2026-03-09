"""
validation_core.py  –  APA 7th Edition Citation Validator
==========================================================
[see inline docstrings for full spec coverage list]
"""
from __future__ import annotations
import re, logging, difflib, threading, unicodedata
from collections import defaultdict
from typing import Any, Dict, List, Optional, Set, Tuple
from xml.etree import ElementTree as ET
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ── Constants ────────────────────────────────────────────────────────────────
CITE_STYLE         = "cite_bib"
GREEN              = WD_COLOR_INDEX.BRIGHT_GREEN
YELLOW             = WD_COLOR_INDEX.YELLOW
COMMENT_AUTHOR     = "Ref Validator"
COMMENT_INITIALS   = "RV"
FUZZY_THRESHOLD    = 0.80
NEAR_DUP_THRESHOLD = 0.97
ET_AL_MIN          = 3
BIB_TAG_OPEN       = "<ref-open>"
BIB_TAG_CLOSE      = "<ref-close>"
_NAME_PREFIXES: Set[str] = {
    "van","von","de","del","della","di","du",
    "le","la","los","las","den","der","ter","al",
}

# ── Year helpers ─────────────────────────────────────────────────────────────
_Y             = r"(?:19|20)\d{2}[a-z]?"
_YEAR_SPECIAL  = r"(?:n\.d\.|in\s+press)"
_YEAR_ANY      = rf"(?:{_Y}|{_YEAR_SPECIAL})"
_YEAR_RANGE_P  = r"(?:(?:19|20)\d{2})/(?:(?:19|20)\d{2})"

# ── Regex patterns ───────────────────────────────────────────────────────────
_RE_PAREN      = re.compile(r'\(([A-ZÀ-Ö][^()]{1,120}?' + _YEAR_ANY + r')\)', re.UNICODE|re.IGNORECASE)
_RE_NARRATIVE  = re.compile(r'(?<!\()([A-ZÀ-Ö][a-zà-ö]+(?:\s+et\s+al\.)?)\s+\((' + _YEAR_ANY + r')\)', re.UNICODE|re.IGNORECASE)
_RE_MULTI_Y_P  = re.compile(r'\(([A-ZÀ-Ö][^()]{1,80}?),\s*((?:' + _Y + r')(?:\s*,\s*(?:' + _Y + r')){1,4})\)', re.UNICODE)
_RE_MULTI_Y_N  = re.compile(r'(?<!\()([A-ZÀ-Ö][a-zà-ö]+(?:\s+et\s+al\.)?)\s+\(((?:' + _Y + r')(?:\s*,\s*(?:' + _Y + r')){1,4})\)', re.UNICODE)
_RE_SECONDARY  = re.compile(r'\(([^()]+?),\s*(' + _YEAR_ANY + r')\s*,\s*as\s+cited\s+in\s+([^()]+?),\s*(' + _YEAR_ANY + r')\)', re.IGNORECASE)
_RE_CITE_UNIT  = re.compile(r'^(.*?)\s*,?\s*(' + _YEAR_ANY + r'|' + _YEAR_RANGE_P + r')$', re.DOTALL|re.IGNORECASE)
_RE_LOCATOR    = re.compile(r',\s*(pp?\.\s*\d+(?:[–—\-]\d+)?)', re.IGNORECASE)
_RE_AMA        = re.compile(r'\b([A-Z][a-z]+(?:\s+et\s+al\.)?)\s+((?:19|20)\d{2})\b(?!\s*[,;])')
_RE_BAD_ETAL   = re.compile(r'\(([A-Z][a-z]+)\s+et\s+al\s+((?:19|20)\d{2}[a-z]?)\)')
_RE_MISS_COMMA = re.compile(r'\(([A-Z][^(),]{0,80}?)\s+((?:19|20)\d{2}[a-z]?)\)')
_RE_YR_RANGE   = re.compile(r'\b((?:19|20)\d{2})/((?:19|20)\d{2})\b')
_RE_BAD_ND     = re.compile(r'\(\s*([A-ZÀ-Ö][^()]{0,80}?),?\s*(nd\.?|n\.d(?!\.)|N\.D\.?)\s*\)')
_RE_BAD_INPRES = re.compile(r'\(\s*([A-ZÀ-Ö][^()]{0,80}?),\s*(In\s+Press|IN\s+PRESS|In\s+press)\s*\)')
_RE_ETAL_NOPER = re.compile(r'\bet\s+al(?!\.)\b')
_RE_BIB_ABBREV = re.compile(r'\[([A-Z]{2,8})\]')
_HEADING_RE    = re.compile(r"heading\s*\d|title", re.IGNORECASE)
_FOOTNOTE_RE   = re.compile(r"footnote|endnote", re.IGNORECASE)
_CAPTION_RE    = re.compile(r"caption|figure|table\s*(?:title|caption)", re.IGNORECASE)
_BQUOTE_RE     = re.compile(r"block\s*quote|blockquote|quote", re.IGNORECASE)

# ── Advanced normalisation ────────────────────────────────────────────────────
def _strip_diacritics(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _norm_apos(s: str) -> str:
    return re.sub(r"['\u2019\u02bc\u2018]", "'", s)

def _sort_key_name(s: str) -> str:
    s = _strip_diacritics(_norm_apos(s))
    words = s.split()
    return " ".join(w for w in words if w.lower().rstrip(".,") not in _NAME_PREFIXES).lower().strip()

def _norm(s: str) -> str:
    s = _strip_diacritics(_norm_apos(s))
    s = re.sub(r"\s*et\s+al\.?\s*", " etal ", s, flags=re.IGNORECASE)
    s = re.sub(r"[.,;&]", " ", s)
    return re.sub(r"\s+", " ", s).strip().lower()

def _first_surname(a: str) -> str:
    p = _norm(a).split(); return p[0] if p else ""

def _surname_set(a: str) -> Set[str]:
    return set(re.findall(r"[a-z\u00c0-\u024f]{2,}", _norm(a))) - {"et","al","etal","and","the"}

def _strip_suffix(year: str) -> str:
    m = re.match(r"^((?:19|20)\d{2})[a-z]?$", year); return m.group(1) if m else year

def _acronym_of(text: str) -> str:
    return "".join(w[0].upper() for w in text.split() if w and w[0].isupper())

def _is_organization(full: str) -> bool:
    org_keywords = {
        'association', 'board', 'committee', 'department', 'society', 'institute', 'institutes',
        'organization', 'organisation', 'council', 'federation', 'center', 'centre',
        'national', 'international', 'academy', 'academies', 'university', 'college',
        'school', 'hospital', 'ministry', 'agency', 'commission', 'foundation', 'group',
        'corporation', 'inc', 'ltd', 'company', 'union', 'administration', 'authority',
        'office', 'bureau', 'world', 'network', 'consortium', 'collaborators', 'alliance',
        'task force', 'working group', 'services', 'congress', 'assembly', 'trust',
        'disease control', 'food and', 'u.s.', 'uk', 'united states', 'health'
    }
    return any(re.search(rf"\b{kw}\b", full, re.IGNORECASE) for kw in org_keywords)

def _count_authors(full: str) -> int:
    if _is_organization(full):
        return 1
    
    parts = re.split(r'\s*&\s*|\s*,\s*', full)
    count = 0
    has_et_al = False
    
    for part in parts:
        part = part.strip()
        if not part: continue
        
        lower_part = part.lower()
        if 'et al' in lower_part:
            has_et_al = True
            part = re.sub(r'(?i)\bet\s+al\.?\b', '', part).strip()
            if not part: continue
            
        # exclude parts that are just initials, like 'J. S.' or 'A.' or 'J'
        # Also handle hyphenated initials like 'A.-B.'
        if re.match(r'^([A-Z]\.?\s*-?)+[A-Z]?\.?$', part):
            continue
            
        count += 1
        
    if has_et_al:
        # If it has et al., it implies there are more authors than listed. 
        # Make sure count is at least ET_AL_MIN if et al is present.
        count = max(count + 1, 3) 
        
    return max(count, 1)

def _is_org_match(cite_auth: str, bib_auth: str) -> bool:
    c_parts = [p.strip() for p in re.split(r',|&', cite_auth) if p.strip()]
    if not c_parts: return False
    
    b_norm = _norm(bib_auth)
    
    def match_part(cp):
        if _norm(cp) in b_norm: return True
        cp_up = cp.upper().replace('.', '').replace(' ', '')
        
        b_words = re.sub(r'U\.S\.', 'US ', bib_auth)
        b_words = re.sub(r'U\.K\.', 'UK ', b_words)
        b_words = re.split(r'[\s.,;]+', b_words)
        
        strict = ''; loose = ''
        for w in b_words:
            if w == 'US': strict += 'US'; loose += 'US'
            elif w == 'UK': strict += 'UK'; loose += 'UK'
            elif w:
                if w[0].isupper() and w.lower() not in {'of', 'and', 'the', 'for', 'in', 'on'}: 
                    strict += w[0].upper()
                if w[0].isalpha():
                    loose += w[0].upper()
                    
        if cp_up in strict or cp_up in loose: return True
        
        b_all = [w for w in b_words if w.lower() not in {'and', 'for', 'of', 'the', 'in', 'on'}]
        ac2 = ''.join(w[0].upper() for w in b_all if w)
        if cp_up in ac2: return True
        
        if len(cp_up) >= 2 and any(a.startswith(cp_up) for a in [strict, loose, ac2]): return True
        return False
        
    return all(match_part(cp) for cp in c_parts)

# ── Locator validator ─────────────────────────────────────────────────────────
def check_locator(loc: str) -> List[str]:
    probs: List[str] = []
    is_range   = bool(re.search(r"\d+[–—\-]\d+", loc))
    uses_pp    = loc.lower().startswith("pp")
    uses_p     = loc.lower().startswith("p.") and not uses_pp
    has_endash = "–" in loc or "—" in loc
    has_hyphen = bool(re.search(r"\d-\d", loc))
    has_space  = bool(re.match(r"pp?\.\s\d", loc, re.IGNORECASE))
    if is_range  and uses_p:  probs.append("use 'pp.' (not 'p.') for ranges")
    if not is_range and uses_pp: probs.append("use 'p.' (not 'pp.') for single page")
    if is_range  and has_hyphen and not has_endash: probs.append("use en-dash (–) not hyphen for ranges")
    if not has_space:
        spaced_loc = re.sub(r"\.(?=\d)", ". ", loc)
        probs.append(f"add space after period: e.g. '{spaced_loc}'")
    return probs

# ── APA Fixer ─────────────────────────────────────────────────────────────────
class ApaFixer:
    @staticmethod
    def needs_fix(text: str) -> bool:
        return bool(_RE_AMA.search(text) or
                    _RE_BAD_ETAL.search(text) or _RE_MISS_COMMA.search(text) or
                    _RE_YR_RANGE.search(text) or _RE_BAD_ND.search(text) or
                    _RE_BAD_INPRES.search(text) or _RE_ETAL_NOPER.search(text))

    @staticmethod
    def fix(text: str) -> Tuple[str, List[Dict]]:
        changes: List[Dict] = []
        r = text
        def _chg(orig, fixed, ft):
            if orig != fixed: changes.append({"original":orig,"fixed":fixed,"fix_type":ft})
            return fixed
        def _ama(m):   return _chg(m.group(0), f"{m.group(1)} ({m.group(2)})", "ama_to_apa_narrative")
        def _comma(m):
            orig, inner, year = m.group(0), m.group(1).rstrip(), m.group(2)
            if inner.endswith(",") or re.search(r"et\s+al", inner, re.IGNORECASE): return orig
            return _chg(orig, f"({inner}, {year})", "missing_comma")
        def _etal(m):  return _chg(m.group(0), f"({m.group(1).strip()} et al., {m.group(2)})", "etal_punctuation")
        def _etalp(m): changes.append({"original":m.group(0),"fixed":"et al.","fix_type":"etal_missing_period"}); return "et al."
        def _nd(m):    return _chg(m.group(0), f"({m.group(1).strip().rstrip(',')}, n.d.)", "nd_format")
        def _inp(m):   return _chg(m.group(0), f"({m.group(1).strip().rstrip(',')}, in press)", "inpress_capitalisation")
        r = _RE_AMA.sub(_ama, r)
        r = _RE_MISS_COMMA.sub(_comma, r)
        r = _RE_BAD_ETAL.sub(_etal, r)
        r = _RE_ETAL_NOPER.sub(_etalp, r)
        r = _RE_BAD_ND.sub(_nd, r)
        r = _RE_BAD_INPRES.sub(_inp, r)
        for m in _RE_YR_RANGE.finditer(r):
            changes.append({"original":m.group(0),"fixed":m.group(0),"fix_type":"year_range_not_apa"})
        return r, changes

    @staticmethod
    def fix_etal_expansion(cite_author: str, bib: Dict) -> Optional[str]:
        if bib.get("author_count",1) < ET_AL_MIN: return None
        if re.search(r"\bet\s+al\b", cite_author, re.IGNORECASE): return None
        first = bib.get("display","").split(" et al.")[0].split(" &")[0].strip()
        return f"{first} et al." if first else None

# ── Citation Extractor ────────────────────────────────────────────────────────
class CitationExtractor:
    @staticmethod
    def extract(text: str) -> List[Dict]:
        hits: List[Dict] = []
        occ:  List[Tuple[int,int]] = []
        def _over(s,e): return any(s<oe and e>os for os,oe in occ)

        # Secondary
        for m in _RE_SECONDARY.finditer(text):
            occ.append((m.start(),m.end()))
            hits.append({"raw":m.group(0),"author":m.group(3).strip(),"year":m.group(4).strip(),
                         "original_author":m.group(1).strip(),"original_year":m.group(2).strip(),
                         "cite_type":"secondary","start":m.start(),"end":m.end()})

        # Multi-year parenthetical
        for m in _RE_MULTI_Y_P.finditer(text):
            if _over(m.start(),m.end()): continue
            author = m.group(1).strip().rstrip(",")
            years  = [y.strip() for y in m.group(2).split(",") if y.strip()]
            hits.append({"raw":m.group(0),"author":author,"years":years,
                         "cite_type":"multi_year","start":m.start(),"end":m.end()})
            occ.append((m.start(),m.end()))

        # Multi-year narrative
        for m in _RE_MULTI_Y_N.finditer(text):
            if _over(m.start(),m.end()): continue
            years = [y.strip() for y in m.group(2).split(",") if y.strip()]
            hits.append({"raw":m.group(0),"author":m.group(1).strip(),"years":years,
                         "cite_type":"multi_year","start":m.start(),"end":m.end()})
            occ.append((m.start(),m.end()))

        # Standard parenthetical (may be multi-citation block)
        for m in _RE_PAREN.finditer(text):
            if _over(m.start(),m.end()): continue
            segs  = [s.strip() for s in m.group(1).split(";")]
            auths: List[str] = []
            block: List[Dict] = []
            for seg in segs:
                um = _RE_CITE_UNIT.search(seg)
                if um:
                    sa = um.group(1).strip().rstrip(",")
                    sy = um.group(2).strip()
                    lm = _RE_LOCATOR.search(seg)
                    block.append({"raw": f"({seg})" if len(segs)>1 else m.group(0),
                                  "author":sa,"year":sy,"locator": lm.group(1) if lm else None,
                                  "cite_type":"parenthetical","start":m.start(),"end":m.end(),
                                  "block_raw":m.group(0),"block_size":len(segs),
                                  "block_order_ok":True})
                    auths.append(_first_surname(sa))
            order_ok = all(auths[i]>=auths[i-1] for i in range(1,len(auths)))
            for bc in block: bc["block_order_ok"] = order_ok
            hits.extend(block)
            occ.append((m.start(),m.end()))

        # Narrative
        for m in _RE_NARRATIVE.finditer(text):
            if _over(m.start(),m.end()): continue
            lm = _RE_LOCATOR.search(m.group(0))
            hits.append({"raw":m.group(0),"author":m.group(1).strip(),"year":m.group(2).strip(),
                         "locator": lm.group(1) if lm else None,
                         "cite_type":"narrative","start":m.start(),"end":m.end()})
            occ.append((m.start(),m.end()))

        return hits

# ── Bibliography Parser ───────────────────────────────────────────────────────
class BibliographyParser:
    _RE_APA = re.compile(r"^(?P<authors>.+?)\s*\((?P<year>(?:19|20)\d{2}[a-z]?|n\.d\.|in\s+press)(?:,\s*[^)]+)?\)\.", re.DOTALL|re.IGNORECASE)
    _RE_AMA_Y = re.compile(r"\b((?:19|20)\d{2})\b")

    @classmethod
    def parse_entry(cls, raw: str) -> Optional[Dict]:
        m = cls._RE_APA.match(raw.strip())
        if m:
            authors_raw, year = m.group("authors").strip().rstrip("."), m.group("year")
        else:
            ym = cls._RE_AMA_Y.search(raw)
            if not ym: return None
            year = ym.group(1)
            authors_raw = raw.split(".",1)[0].strip()
        full   = re.sub(r"\s+"," ", authors_raw).strip().rstrip(",.")
        
        is_org = _is_organization(full)
        disp   = cls._display(full, is_org)
        ac     = 1 if is_org else _count_authors(full)
        
        abm    = _RE_BIB_ABBREV.search(full)
        abbrev = abm.group(1) if abm else None
        ca     = re.sub(r"\[.*?\]","",full).strip()
        auto_ac= None
        if " " in ca and not re.search(r",\s*[A-Z]\.?", ca):
            a2 = _acronym_of(ca)
            if len(a2)>=2: auto_ac=a2
        return {"full_author":full,"display":disp,"year":year,"raw":raw,"cited":False,
                "author_count":ac,"abbrev":abbrev,"auto_acronym":auto_ac, "is_org": is_org,
                "sort_key":_sort_key_name(full.split(",")[0])}

    @staticmethod
    def _display(full: str, is_org: bool = False) -> str:
        fc = re.sub(r'\[.*?\]','',full).strip()
        if is_org:
            return fc
            
        parts = re.split(r'\s*&\s*|\s*,\s*', fc)
        sn = []
        has_et_al = False
        
        for part in parts:
            part = part.strip()
            if not part: continue
            
            lower_part = part.lower()
            if 'et al' in lower_part:
                has_et_al = True
                part = re.sub(r'(?i)\bet\s+al\.?\b', '', part).strip()
                if not part: continue
                
            # exclude parts that are just initials
            if re.match(r'^([A-Z]\.?\s*-?)+[A-Z]?\.?$', part):
                continue
                
            sn.append(part)
            
        if not sn: sn = [w for w in fc.split() if w[0].isupper() and not re.match(r'^[A-Z]\.$', w)]
        if not sn: return fc.split()[0] if fc.split() else fc
        
        if len(sn) == 1:
            return sn[0] if not has_et_al else f"{sn[0]} et al."
        if has_et_al or len(sn) >= 3:
            return f"{sn[0]} et al."
        return f"{sn[0]} & {sn[1]}"

# ── Matcher ───────────────────────────────────────────────────────────────────
def match_citation(cite_author: str, cite_year: str, bibliography: Dict[str,dict]) -> Tuple[Optional[str],str]:
    cn = _norm(cite_author); cf = _first_surname(cite_author); cw = _surname_set(cite_author)
    cb = _strip_suffix(cite_year); cu = cite_author.strip().upper()
    ym: List[str]=[]; fh: List[Tuple[float,str,str]]=[]; sh: List[str]=[]
    for key, ref in bibliography.items():
        rfull=ref["full_author"]; rn=_norm(rfull); ry=ref["year"]; rb=_strip_suffix(ry)
        rf=_first_surname(rfull); yok=(cite_year==ry)
        def _sfx_or_miss():
            if cb==rb and cite_year==cb: sh.append(key)
            else: ym.append(key)
        if cn==rn:
            if yok: return key,"exact"
            _sfx_or_miss(); continue
        if cf and rf and (cf==rf or difflib.SequenceMatcher(None, cf, rf).ratio() >= FUZZY_THRESHOLD):
            if yok: return key,"smart"
            _sfx_or_miss(); continue
            
        bib_surnames = _surname_set(rfull)
        if cw and (cw.issubset(bib_surnames) or all(any(w1==w2 or difflib.SequenceMatcher(None, w1, w2).ratio() >= FUZZY_THRESHOLD for w2 in bib_surnames) for w1 in cw)):
            if yok: return key,"smart"
            _sfx_or_miss(); continue
        if _is_org_match(cite_author, rfull):
            if yok: return key,"org_abbrev"
            _sfx_or_miss(); continue
        ratio=difflib.SequenceMatcher(None,cn,rn).ratio()
        if ratio>=FUZZY_THRESHOLD: fh.append((ratio,key,ry))
    if sh:
        unique={bibliography[k]["raw"] for k in sh}
        return sh[0], "suffix_mismatch" if len(unique)==1 else "suffix_ambiguous"
    if ym: return ym[0],"year_mismatch"
    if fh: fh.sort(reverse=True); return fh[0][1],"spelling_mismatch"
    return None,"not_found"

# ── et al. checker ────────────────────────────────────────────────────────────
def check_etal_enforcement(cite_author: str, bib: Dict) -> Optional[str]:
    n = bib.get("author_count",1)
    has = bool(re.search(r"\bet\s+al\b", cite_author, re.IGNORECASE))
    
    if bib.get("is_org"):
        if has:
            return f"et al. INCORRECT: '{bib['display']}' is an organization — list full name, do not use et al."
        return None

    if n>=ET_AL_MIN and not has:
        return f"et al. REQUIRED: entry has {n} authors ('{bib['display']}') — APA 7th requires et al."
    if n<ET_AL_MIN and has:
        return f"et al. INCORRECT: entry has {n} author(s) ('{bib['display']}') — list all names."
    if has and not re.search(r"\bet\s+al\.", cite_author, re.IGNORECASE):
        return "et al. PUNCTUATION: missing period — must be 'et al.' (with period)."
    return None

# ── Bibliography structural checks ───────────────────────────────────────────
def check_bibliography_structure(entries: List[Dict]) -> List[Dict]:
    issues: List[Dict] = []
    seen_r: Dict[str,int]={}; seen_ay: Dict[str,int]={}
    # a/b: duplicates
    for e in entries:
        raw=e["raw"]; pidx=e.get("para_idx",-1)
        ay=f"{_norm(e['full_author'])}|{e['year']}"
        if raw in seen_r:
            issues.append({"type":"duplicate_entry","para_idx":pidx,"raw":e["display"],"para":None,
                "message":f"📋 DUPLICATE ENTRY: '{e['display']} ({e['year']})' at para {pidx+1} (first at {seen_r[raw]+1})."})
        else: seen_r[raw]=pidx
        if ay in seen_ay and seen_ay[ay]!=pidx:
            issues.append({"type":"duplicate_authyear","para_idx":pidx,"raw":e["display"],"para":None,
                "message":f"📋 SAME AUTHOR+YEAR: '{e['display']} ({e['year']})' — add a/b/c suffix."})
        seen_ay[ay]=pidx
    # c: alphabetical order
    pk=""; pd=""; pp=-1
    for e in entries:
        sk=e.get("sort_key","")
        if sk and pk and sk<pk:
            issues.append({"type":"order_error","para_idx":e.get("para_idx",-1),"raw":e["display"],"para":None,
                "message":f"🔤 ORDER: '{e['display']}' should precede '{pd}' — bibliography must be alphabetical."})
        pk=sk; pd=e["display"]; pp=e.get("para_idx",-1)
    # d: suffix ordering
    sg: Dict[Tuple[str,str],List] = {}
    for e in entries:
        bare=_strip_suffix(e["year"]); sfx=e["year"][len(bare):]
        if sfx: sg.setdefault((_norm(e["full_author"]),bare),[]).append((sfx,e.get("para_idx",-1),e["display"]))
    for (auth,bare),items in sg.items():
        actual=[s for s,_,_ in items]; exp=sorted(actual)
        if actual!=exp:
            for i,(s,pidx,disp) in enumerate(items):
                if actual[i]!=exp[i]:
                    issues.append({"type":"suffix_order_error","para_idx":pidx,"raw":disp,"para":None,
                        "message":f"🔡 SUFFIX ORDER: '{disp} ({bare}{s})' out of order — expected {', '.join(bare+x for x in exp)}."})
    # e: near-duplicate fuzzy
    rl=list({e["raw"] for e in entries})
    for i in range(len(rl)):
        for j in range(i+1,len(rl)):
            r=difflib.SequenceMatcher(None,rl[i].lower(),rl[j].lower()).ratio()
            if r>=NEAR_DUP_THRESHOLD:
                for e in entries:
                    if e["raw"]==rl[j]:
                        issues.append({"type":"near_duplicate","para_idx":e.get("para_idx",-1),"raw":e["display"],"para":None,
                            "message":f"📋 NEAR-DUPLICATE ({r:.0%}): '{rl[j][:80]}' may duplicate '{rl[i][:60]}'."})
                        break
    return issues

# ── Org tracker ───────────────────────────────────────────────────────────────
class OrgTracker:
    def __init__(self): self._introduced: Dict[str,Tuple[str,int]] = {}
    def record(self, abbrev: str, full: str, pidx: int):
        if abbrev not in self._introduced: self._introduced[abbrev]=(full,pidx)
    def check(self, abbrev: str) -> Optional[str]:
        if abbrev not in self._introduced:
            return f"ORG ABBREV FIRST USE: '{abbrev}' used without prior introduction. Spell out full name first: 'Full Name [{abbrev}]'."
        return None
    def is_known(self, abbrev: str) -> bool: return abbrev in self._introduced

# ── Context filter ────────────────────────────────────────────────────────────
def _para_context(para) -> str:
    sn = getattr(getattr(para,"style",None),"name","") or ""
    if _HEADING_RE.search(sn): return "heading"
    if _FOOTNOTE_RE.search(sn): return "footnote"
    if _CAPTION_RE.search(sn): return "caption"
    if _BQUOTE_RE.search(sn): return "blockquote"
    return "body"

# ── Word XML helpers ──────────────────────────────────────────────────────────
def _ensure_style(doc):
    if CITE_STYLE not in doc.styles:
        try: doc.styles.add_style(CITE_STYLE, WD_STYLE_TYPE.CHARACTER)
        except: pass

def _full_text(p): return "".join(r.text for r in p.runs)
def safe_splice(para, start, end, new_text, highlight, style):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    import copy
    
    runs = list(para.runs)
    if not runs: return None
    
    curr = 0; s_r_idx = -1; s_off = -1; e_r_idx = -1; e_off = -1
    for i, r in enumerate(runs):
        l = len(r.text)
        if s_r_idx == -1 and curr <= start < curr + l: s_r_idx, s_off = i, start - curr
        if e_r_idx == -1 and curr < end <= curr + l: e_r_idx, e_off = i, end - curr
        curr += l
    if e_r_idx == -1 and end == curr: e_r_idx = len(runs) - 1; e_off = len(runs[-1].text)
    if s_r_idx == -1 or e_r_idx == -1: return None
    
    st_text = runs[s_r_idx].text; end_text = runs[e_r_idx].text
    pre = st_text[:s_off]; post = end_text[e_off:]
    
    for i in range(s_r_idx + 1, e_r_idx): runs[i].text = ""
    
    base_rPr = copy.deepcopy(runs[s_r_idx]._r.rPr) if hasattr(runs[s_r_idx]._r, 'rPr') and runs[s_r_idx]._r.rPr is not None else None
    
    st = para.add_run(new_text)
    st._r.clear_content()
    if base_rPr is not None: st._r.append(copy.deepcopy(base_rPr))
    t = OxmlElement('w:t'); t.text = new_text
    if new_text.startswith(' ') or new_text.endswith(' '): t.set(qn('xml:space'), 'preserve')
    st._r.append(t)
    if style is not None:
        try: st.style = style
        except: pass
    if highlight is not None: st.font.highlight_color = highlight
    
    pe = runs[s_r_idx]._r.getparent()
    ai = list(pe).index(runs[s_r_idx]._r)
    pe.remove(st._r)
    
    if s_r_idx == e_r_idx:
        runs[s_r_idx].text = pre
        post_r = OxmlElement('w:r')
        if base_rPr is not None: post_r.append(copy.deepcopy(base_rPr))
        pt = OxmlElement('w:t'); pt.text = post
        if post.startswith(' ') or post.endswith(' '): pt.set(qn('xml:space'), 'preserve')
        post_r.append(pt)
        pe.insert(ai + 1, st._r)
        pe.insert(ai + 2, post_r)
    else:
        runs[s_r_idx].text = pre
        runs[e_r_idx].text = post
        pe.insert(ai + 1, st._r)
        
    return st._r

def isolate_target_run(p, txt):
    if not txt: return None
    full = _full_text(p); pos = full.find(txt)
    if pos == -1: pos = full.lower().find(txt.lower())
    if pos != -1: return safe_splice(p, pos, pos+len(txt), full[pos:pos+len(txt)], None, None)
    return None

def apply_style_to_text(p, txt, hl):
    if not txt: return None
    offset = 0
    res = None
    while True:
        full = _full_text(p)
        pos = full.find(txt, offset)
        if pos == -1: pos = full.lower().find(txt.lower(), offset)
        if pos == -1: break
        
        st = safe_splice(p, pos, pos+len(txt), txt, hl, CITE_STYLE)
        if st is not None: res = st
        
        offset = pos + len(txt)
    return res

def replace_and_style(p, old, new, hl):
    if not old: return None
    offset = 0
    res = None
    while True:
        full = _full_text(p)
        pos = full.find(old, offset)
        if pos == -1: break
        
        st = safe_splice(p, pos, pos+len(old), new, hl, CITE_STYLE)
        if st is not None: res = st
        
        offset = pos + len(new)
    return res

def _get_comments_part(doc):
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    from lxml import etree as _lxml_etree
    REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    
    for rel in doc.part.rels.values():
        if rel.reltype == REL: return rel.target_part
        
    for p in doc.part.package.parts:
        if p.partname == "/word/comments.xml":
            doc.part.relate_to(p, REL)
            return p
            
    _W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    _W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
    
    root = _lxml_etree.Element(f"{{{_W_NS}}}comments", nsmap={"w": _W_NS, "mc": _MC_NS, "w14": _W14_NS})
    root.set(f"{{{_MC_NS}}}Ignorable", "w14 wp14")
    
    blob = _lxml_etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    puri = PackURI("/word/comments.xml")
    part = Part(puri, CT, blob, doc.part.package)
    doc.part.package.add_part(part)
    doc.part.relate_to(part, REL)
    return part

def insert_comment(doc, para, text, target_run=None, target_text=None):
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from lxml import etree as _lxml_etree
        if not para.runs: return False
        
        if target_run is None and target_text is not None:
            # Try exact text first
            target_run = isolate_target_run(para, target_text)
            # If not found, try stripping surrounding parentheses
            if target_run is None:
                stripped = target_text.strip("() ")
                if stripped and stripped != target_text:
                    target_run = isolate_target_run(para, stripped)
            # If still not found, try anchoring on just the first word/token (e.g. first author name)
            if target_run is None:
                first_token = (target_text.strip("() ").split(",")[0]).strip()
                if first_token and len(first_token) > 2:
                    target_run = isolate_target_run(para, first_token)

             
        cp = _get_comments_part(doc)
        _W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        _NSMAP = {"w": _W_NS}
        
        try:
            tree = _lxml_etree.fromstring(cp._blob)
        except Exception:
            tree = _lxml_etree.Element(f"{{{_W_NS}}}comments", nsmap=_NSMAP)
            
        existing = [int(c.get(f"{{{_W_NS}}}id")) for c in tree.findall(f"{{{_W_NS}}}comment") if c.get(f"{{{_W_NS}}}id")]
        nid = max(existing) + 1 if existing else 0
        
        cel = _lxml_etree.SubElement(tree, f"{{{_W_NS}}}comment")
        cel.set(f"{{{_W_NS}}}id", str(nid))
        cel.set(f"{{{_W_NS}}}author", COMMENT_AUTHOR)
        cel.set(f"{{{_W_NS}}}initials", COMMENT_INITIALS)
        
        pel = _lxml_etree.SubElement(cel, f"{{{_W_NS}}}p")
        rel = _lxml_etree.SubElement(pel, f"{{{_W_NS}}}r")
        tel = _lxml_etree.SubElement(rel, f"{{{_W_NS}}}t")
        tel.text = text
        
        cp._blob = _lxml_etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        
        cs=OxmlElement("w:commentRangeStart"); cs.set(qn("w:id"),str(nid))
        ce=OxmlElement("w:commentRangeEnd"); ce.set(qn("w:id"),str(nid))
        rr=OxmlElement("w:r"); rf=OxmlElement("w:commentReference"); rf.set(qn("w:id"),str(nid)); rr.append(rf)
        
        pe = para._p
        if target_run is not None and target_run.getparent() == pe:
            idx = list(pe).index(target_run)
            pe.insert(idx, cs)
            pe.insert(idx + 2, ce)
            pe.insert(idx + 3, rr)
        else:
            # Fallback: place a point-anchor at the end of the last run
            # (commentRangeStart immediately before commentRangeEnd with no text between)
            # This avoids wrapping the entire paragraph which highlights everything in Word.
            runs_in_p = [child for child in pe if child.tag == qn("w:r")]
            if runs_in_p:
                last_run = runs_in_p[-1]
                last_run.addnext(rr)
                last_run.addnext(ce)
                last_run.addnext(cs)
            else:
                pe.append(cs)
                pe.append(ce)
                pe.append(rr)
        return True
    except: return False

# ── Validation Report ─────────────────────────────────────────────────────────
class ValidationReport:
    def __init__(self,issues,stats,total_refs,total_cites):
        self.issues=issues; self.stats=stats
        self.total_refs=total_refs; self.total_cites=total_cites

    def summary(self) -> str:
        s=self.stats
        def R(lbl,k,w=44): return f"  {lbl:<{w}}: {s.get(k,0)}"
        lines=["="*68,"APA 7th CITATION VALIDATION REPORT","="*68,
               f"  {'Total in-text citations':<44}: {self.total_cites}",
               f"  {'Total bibliography entries':<44}: {self.total_refs}","-"*68,
               R("Matched (green)","matched"), R("Missing references","missing"),
               R("Year mismatches","year_mismatch"), R("Suffix missing","suffix_mismatch"),
               R("Suffix ambiguous","suffix_ambiguous"), R("Spelling mismatches","spelling_mismatch"),
               R("et al. violations","etal_enforcement"), R("et al. auto-fixes","etal_autofixes"),
               R("Secondary citations","secondary"),
               R("Multi-citation blocks","multi_citation_blocks"),
               R("Duplicate citations in block","duplicate_citations"),
               R("Multi-year citations","multi_year"), R("Multi-year mismatches","multi_year_mismatches"),
               R("Block order violations","alphabetical_violations"),
               R("Org-author cases","organization_cases"),
               R("Org-abbrev first-use errors","org_abbrev_first_use"),
               R("Page locator errors","page_locator_errors"),
               R("Format auto-fixes","format_fixed"), R("Year-range (manual)","year_range"),
               R("Unused references","unused"),
               R("Duplicate bib entries","duplicate_entry"),
               R("Same author+year (needs suffix)","duplicate_authyear"),
               R("Near-duplicate bib entries","near_duplicate"),
               R("Bib order errors","order_error"), R("Suffix order errors","suffix_order_error"),
               "-"*68]
        _SEC=[("missing","MISSING REFERENCES"),("year_mismatch","YEAR MISMATCHES"),
              ("suffix_mismatch","SUFFIX MISSING"),("suffix_ambiguous","SUFFIX AMBIGUOUS"),
              ("spelling_mismatch","SPELLING MISMATCHES"),("etal_enforcement","et al. VIOLATIONS"),
              ("secondary","SECONDARY CITATIONS"),("duplicate_citations","DUPLICATE WITHIN BLOCK"),
              ("multi_year_mismatches","MULTI-YEAR MISMATCHES"),
              ("alphabetical_violations","BLOCK ORDER VIOLATIONS"),
              ("org_abbrev_first_use","ORG FIRST-USE ERRORS"),
              ("organization_cases","ORG ABBREV MATCHES"),("page_locator_errors","PAGE LOCATOR ERRORS"),
              ("format_fixed","FORMAT AUTO-FIXES"),("year_range","YEAR-RANGE CITATIONS"),
              ("unused","UNUSED REFERENCES"),
              ("duplicate_entry","DUPLICATE BIB ENTRIES"),
              ("duplicate_authyear","SAME AUTHOR+YEAR"),("near_duplicate","NEAR-DUPLICATE BIB"),
              ("order_error","BIB ORDER ERRORS"),("suffix_order_error","SUFFIX ORDER ERRORS")]
        for itype,heading in _SEC:
            items=[i for i in self.issues if i["type"]==itype]
            if items:
                lines+=[f"",heading,"-"*52]
                for item in items:
                    lines.append(f"  Para {item.get('para_idx','?')+1:>4}: {item['message']}")
        lines+=["","="*68,"END OF REPORT","="*68]
        return "\n".join(lines)

    def to_dict(self) -> dict:
        return {"stats":dict(self.stats),"total_refs":self.total_refs,
                "total_cites":self.total_cites,
                "issues":[{k:v for k,v in i.items() if k!="para"} for i in self.issues]}

# ── Main Processor ────────────────────────────────────────────────────────────
class CitationProcessor:
    """
    Usage:
        proc   = CitationProcessor("paper.docx", job_id="job-1",
                                   include_contexts={"body","footnote"})
        report = proc.run()
        proc.save("paper_out.docx")
        print(report.summary())
    """
    def __init__(self, doc_path: str, job_id: str="default",
                 include_contexts: Optional[Set[str]]=None):
        self.doc_path         = doc_path
        self.doc              = Document(doc_path)
        self.job_id           = job_id
        self.include_contexts = include_contexts or {"body"}
        self.log              = logging.getLogger(f"{__name__}.{job_id}")
        _ensure_style(self.doc)
        self.bibliography:  Dict[str,dict] = {}
        self._bib_ordered:  List[dict]     = []
        self._issues:       List[dict]     = []
        self._stats:        Dict[str,int]  = defaultdict(int)
        self._cited_keys:   set            = set()
        self._org_tracker:  OrgTracker     = OrgTracker()
        self._seen_blocks:  Set[str]       = set()
        self._seen_cites:   Dict[int,Set[str]] = {}
        self._lock                         = threading.Lock()

    def _parse_bibliography(self):
        in_bib=False
        for idx,para in enumerate(self.doc.paragraphs):
            txt=para.text.strip()
            if BIB_TAG_OPEN  in txt.lower(): in_bib=True;  continue
            if BIB_TAG_CLOSE in txt.lower(): in_bib=False; continue
            if not (in_bib and txt): continue
            e=BibliographyParser.parse_entry(txt)
            if not e: continue
            e["para_idx"]=idx; self._bib_ordered.append(e)
            if e.get("abbrev"):     self._org_tracker.record(e["abbrev"], e["full_author"], idx)
            if e.get("auto_acronym"): self._org_tracker.record(e["auto_acronym"], e["full_author"], idx)
            k1=f"{_norm(e['display'])}|{e['year']}"
            k2=f"{_norm(e['full_author'])}|{e['year']}"
            if k1 not in self.bibliography: self.bibliography[k1]=e
            if k2 not in self.bibliography: self.bibliography[k2]=e
        for iss in check_bibliography_structure(self._bib_ordered):
            pidx=iss.get("para_idx")
            iss["para"]=self.doc.paragraphs[pidx] if pidx is not None and pidx>=0 else None
            self._issues.append(iss); self._stats[iss["type"]]+=1

    def _process_body(self):
        for idx,para in enumerate(self.doc.paragraphs):
            txt=para.text
            if not txt.strip(): continue
            if BIB_TAG_OPEN in txt.lower(): break
            if _para_context(para) not in self.include_contexts: continue
            if ApaFixer.needs_fix(txt):
                _,changes=ApaFixer.fix(txt)
                for ch in changes:
                    ft=ch["fix_type"]
                    if ft=="year_range_not_apa":
                        self._add_issue("year_range",idx,para,ch["original"],f"📅 YEAR RANGE: '{ch['original']}' not valid APA.")
                        self._stats["year_range"]+=1
                    else:
                        replace_and_style(para,ch["original"],ch["fixed"],YELLOW)
                        self._add_issue("format_fixed",idx,para,ch["fixed"],f"🔧 FORMAT FIXED: '{ch['original']}' → '{ch['fixed']}'.")
                        self._stats["format_fixed"]+=1
                txt=_full_text(para)
            # Org introduction scan
            for om in re.finditer(r'([A-Z][A-Za-z\s]+?)\s*[\[(]([A-Z]{2,8})[\])]',txt):
                self._org_tracker.record(om.group(2),om.group(1).strip(),idx)
            for cite in CitationExtractor.extract(txt):
                self._validate_citation(cite,idx,para)

        # Scan all table cell paragraphs as well
        para_offset = len(self.doc.paragraphs)
        for tbl_idx, table in enumerate(self.doc.tables):
            seen_cell_ids = set()
            for row in table.rows:
                for cell in row.cells:
                    cid = id(cell._tc)
                    if cid in seen_cell_ids: continue
                    seen_cell_ids.add(cid)
                    for cell_para in cell.paragraphs:
                        txt = cell_para.text
                        if not txt.strip(): continue
                        virtual_idx = para_offset + tbl_idx
                        if ApaFixer.needs_fix(txt):
                            _,changes=ApaFixer.fix(txt)
                            for ch in changes:
                                ft=ch["fix_type"]
                                if ft=="year_range_not_apa":
                                    self._add_issue("year_range",virtual_idx,cell_para,ch["original"],
                                        f"YEAR RANGE (table {tbl_idx+1}): '{ch['original']}' not valid APA.")
                                    self._stats["year_range"]+=1
                                else:
                                    replace_and_style(cell_para,ch["original"],ch["fixed"],YELLOW)
                                    self._add_issue("format_fixed",virtual_idx,cell_para,ch["fixed"],
                                        f"FORMAT FIXED (table {tbl_idx+1}): '{ch['original']}' to '{ch['fixed']}'.")
                                    self._stats["format_fixed"]+=1
                            txt=_full_text(cell_para)
                        for om in re.finditer(r'([A-Z][A-Za-z\s]+?)\s*[\[(]([A-Z]{2,8})[\])]',txt):
                            self._org_tracker.record(om.group(2),om.group(1).strip(),virtual_idx)
                        for cite in CitationExtractor.extract(txt):
                            self._validate_citation(cite,virtual_idx,cell_para)

    def _validate_citation(self, cite, para_idx, para):
        ct=cite.get("cite_type")

        if ct=="secondary":
            raw=cite["raw"]; oa=cite.get("original_author","?"); oy=cite.get("original_year","?")
            auth=cite["author"]; year=cite["year"]
            rk,mt=match_citation(auth,year,self.bibliography)
            if rk and mt in("exact","smart","org_abbrev"):
                self.bibliography[rk]["cited"]=True; self._cited_keys.add(rk)
            apply_style_to_text(para,raw,YELLOW)
            self._add_issue("secondary",para_idx,para,raw,
                f"SECONDARY: '({oa},{oy}, as cited in {auth},{year})' — only '{auth},{year}' needs bib entry.")
            self._stats["secondary"]+=1; return

        if ct=="multi_year":
            raw=cite["raw"]; auth=cite["author"]; years=cite.get("years",[])
            self._stats["multi_year"]+=1
            any_miss=False
            for year in years:
                rk,mt=match_citation(auth,year,self.bibliography)
                if mt in("exact","smart"):
                    self.bibliography[rk]["cited"]=True; self._cited_keys.add(rk)
                else:
                    any_miss=True
                    self._add_issue("multi_year_mismatches",para_idx,para,raw,
                        f"MULTI-YEAR: '{auth}, {year}' — no matching bib entry.")
                    self._stats["multi_year_mismatches"]+=1
            apply_style_to_text(para,raw,YELLOW if any_miss else GREEN); return

        raw=cite["raw"]; auth=cite["author"]; year=cite.get("year","")
        blk_raw=cite.get("block_raw",raw); blk_sz=cite.get("block_size",1)

        if blk_sz>1 and blk_raw not in self._seen_blocks:
            self._seen_blocks.add(blk_raw); self._stats["multi_citation_blocks"]+=1

        if not cite.get("block_order_ok",True):
            self._add_issue("alphabetical_violations",para_idx,para,blk_raw,
                f"BLOCK ORDER: '{blk_raw}' not alphabetically sorted — APA requires alphabetical order.")
            self._stats["alphabetical_violations"]+=1

        loc=cite.get("locator")
        if loc:
            for prob in check_locator(loc):
                self._add_issue("page_locator_errors",para_idx,para,raw,f"LOCATOR '{loc}': {prob}.")
                self._stats["page_locator_errors"]+=1

        rk,mt=match_citation(auth,year,self.bibliography)

        if mt in("suffix_ambiguous","suffix_mismatch"):
            apply_style_to_text(para,raw,YELLOW)
            if mt=="suffix_ambiguous":
                self._add_issue("suffix_ambiguous",para_idx,para,raw,
                    f"SUFFIX AMBIGUOUS: '{raw}' — multiple entries; add a/b/c suffix.")
            else:
                ref=self.bibliography[rk]
                self._add_issue("suffix_mismatch",para_idx,para,raw,
                    f"SUFFIX MISSING: '{raw}' — entry is '{ref['display']} ({ref['year']})'. Add suffix.")
            self._stats[mt]+=1; return

        if mt in("exact","smart","org_abbrev"):
            ref=self.bibliography[rk]; ref["cited"]=True; self._cited_keys.add(rk)
            ew=check_etal_enforcement(auth,ref)
            if ew:
                ff=ApaFixer.fix_etal_expansion(auth,ref)
                if ff and ff!=auth:
                    replace_and_style(para,auth,ff,YELLOW)
                    self._add_issue("etal_enforcement",para_idx,para,raw,f"👥 {ew} AUTO-FIXED → '{ff}'.")
                    self._stats["etal_autofixes"]+=1
                else:
                    apply_style_to_text(para,raw,YELLOW)
                    self._add_issue("etal_enforcement",para_idx,para,raw,f"👥 {ew}")
                self._stats["etal_enforcement"]+=1
            else:
                apply_style_to_text(para,raw,GREEN); self._stats["matched"]+=1
            if mt=="org_abbrev":
                abbr=auth.strip().upper(); warn=self._org_tracker.check(abbr)
                if warn:
                    self._add_issue("org_abbrev_first_use",para_idx,para,raw,f"🏢 {warn}")
                    self._stats["org_abbrev_first_use"]+=1
                else:
                    self._add_issue("organization_cases",para_idx,para,raw,
                        f"ORG MATCHED: '{auth}' → '{ref['display']}' — verify.")
                    self._stats["organization_cases"]+=1
            # Duplicate within same paragraph/block
            sig=f"{_norm(auth)}|{year}"
            if sig in self._seen_cites.get(para_idx,set()):
                self._add_issue("duplicate_citations",para_idx,para,raw,
                    f"DUPLICATE: '{raw}' already cited in this paragraph/block.")
                self._stats["duplicate_citations"]+=1
            self._seen_cites.setdefault(para_idx,set()).add(sig)

        elif mt=="year_mismatch":
            ref=self.bibliography[rk]; ref["cited"]=True; self._cited_keys.add(rk)
            new_raw = raw.replace(year, ref['year'])
            if new_raw != raw:
                replace_and_style(para,raw,new_raw,YELLOW)
            else:
                apply_style_to_text(para,raw,YELLOW)
            self._add_issue("year_mismatch",para_idx,para,new_raw,
                f"AQ: Note that the citation of reference “{auth}, {year}” has been changed to “{auth}, {ref['year']}” to match with the reference list. Please confirm.")
            self._stats["year_mismatch"]+=1
        elif mt=="spelling_mismatch":
            ref=self.bibliography[rk]
            apply_style_to_text(para,raw,YELLOW)
            self._add_issue("spelling_mismatch",para_idx,para,raw,
                f"SPELLING: cited '{auth}', bib has '{ref['display']}'.")
            self._stats["spelling_mismatch"]+=1
        else:
            apply_style_to_text(para,raw,YELLOW)
            self._add_issue("missing",para_idx,para,raw,
                f"AQ: The reference “{raw}” is cited in the text but not given in the list. Please provide complete publication details of this reference in the list or delete the citation from the text.")
            self._stats["missing"]+=1

    def _flag_unused(self):
        seen: set=set()
        for k,ref in self.bibliography.items():
            raw=ref["raw"]
            if raw in seen: continue
            seen.add(raw)
            if ref.get("cited"): continue
            if any(self.bibliography.get(x,{}).get("raw")==raw for x in self._cited_keys): continue
            pidx=ref.get("para_idx"); para=self.doc.paragraphs[pidx] if pidx is not None else None
            self._add_issue("unused",pidx,para,ref["display"],
                f"AQ: The reference “{ref['display']}, {ref['year']}” is given in the list but not cited in the text. Please cite the reference in the text or delete from the list.")
            self._stats["unused"]+=1

    def _insert_comments(self):
        for iss in self._issues:
            p=iss.get("para")
            tt = iss.get("target_text") or iss.get("raw")
            if p is not None: insert_comment(self.doc,p,iss["message"], target_text=tt)

    def _add_issue(self, itype, para_idx, para, raw, message, target_text=None):
        if target_text is None: target_text = raw
        with self._lock:
            self._issues.append({"type":itype,"para_idx":para_idx if para_idx is not None else -1,
                                  "para":para,"raw":raw,"message":message,"target_text":target_text})

    def run(self) -> ValidationReport:
        self._parse_bibliography(); self._process_body()
        self._flag_unused(); self._insert_comments()
        return ValidationReport(
            issues=self._issues, stats=dict(self._stats),
            total_refs=len({r["raw"] for r in self.bibliography.values()}),
            total_cites=sum(self._stats.get(k,0) for k in(
                "matched","missing","year_mismatch","suffix_mismatch","suffix_ambiguous",
                "spelling_mismatch","etal_enforcement","secondary","multi_year")))

    def save(self, output_path: str):
        self.doc.save(output_path); self.log.info("Saved: %s", output_path)

