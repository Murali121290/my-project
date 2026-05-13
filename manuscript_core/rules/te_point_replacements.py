import re
def get_te_replacement(rule_id: str, m: re.Match) -> str | None:
    """
    Given a detection rule and its regex match, dynamically calculate what the 
    replacement string should be. If None, it means the rule cannot be safely 
    auto-fixed and requires manual human review.
    """
    surface = m.group(0)
    # 1. Percent
    if rule_id == "percent_symbol":
        # "\b\d+(?:\.\d+)?\s*%"  -> numbers percent
        v = re.sub(r'\s*%$', '', surface)
        return f"{v} percent"
    if rule_id in ("percent_word", "per_cent_word"):
        return "%"
    # 2. Ellipsis
    if rule_id in ("ellipsis_3dots", "ellipsis_spaced"):
        return "…"
    # 3. AM/PM
    if rule_id in ("ampm_am_upper", "ampm_am_upper_dots", "ampm_am_lower"):
        return "a.m."
    if rule_id in ("ampm_pm_upper", "ampm_pm_upper_dots", "ampm_pm_lower"):
        return "p.m."
    # 4. Leading zero
    if rule_id == "leading_zero_missing":
        # "(?<!\d)\.(?=\d)" matches "."
        return "0."
    # 5. Spaced locators
    if rule_id == "spaced_hyphens":
        return "—"
    # 6. Dimensions (x)
    if rule_id == "times_x_char":
        return " × "
    # 7. Versus
    if rule_id in ("versus_vs", "versus_v", "versus_vs_dot", "versus_v_dot"):
        return "versus"
    # 8. Era
    if rule_id in ("era_ad_nodot", "era_ad_dots", "era_ad_spaced", "era_ad_dots_spaced"):
        return "A.D."
    if rule_id in ("era_bc_nodot", "era_bc_dots", "era_bc_spaced", "era_bc_dots_spaced"):
        return "B.C."
    # 9. Centuries numeric
    cent_map = {"21st": "twenty-first", "20th": "twentieth", "19th": "nineteenth", "18th": "eighteenth"}
    if rule_id == "century_numeric":
        for n, w in cent_map.items():
            if n in surface.lower():
                return f"{w} century"
    # 10. Table/Figure/Box casing (e.g. figure -> Figure)
    if rule_id in ("ref_figure_lower", "ref_table_lower", "ref_box_lower", "ref_chapter_lower"):
        return surface.capitalize()
    # Plural expansions (e.g. Figs -> Figures)
    if rule_id == "ref_figs_single": return "Figures"
    if rule_id == "ref_fig_single": return "Figure"
    if rule_id == "ref_figs_and": return "Figures"
    if rule_id == "ref_fig_and": return "Figure"
    if rule_id == "ref_tabs_single": return "Tables"
    if rule_id == "ref_tab_single": return "Table"
    if rule_id == "ref_tabs_and": return "Tables"
    if rule_id == "ref_tab_and": return "Table"
    if rule_id == "ref_boxes_single": return "Boxes"
    if rule_id == "ref_box_single": return "Box"
    if rule_id == "ref_chapters_single": return "Chapters"
    if rule_id == "ref_chapter_single": return "Chapter"
    # Abbreviated plural dotted forms (e.g. Tabs 3.2 -> Tables 3.2, Figs 1.2 -> Figures 1.2)
    if rule_id in ("ref_tabs_dotted", "cap_tabs_dotted"):
        return re.sub(r'\bTabs\.?\b', 'Tables', surface, flags=re.IGNORECASE)
    if rule_id in ("ref_tabs_single_num", "cap_tabs_single_num"):
        return re.sub(r'\bTabs\.?\b', 'Tables', surface, flags=re.IGNORECASE)
    if rule_id in ("ref_figs_dotted", "cap_figs_dotted"):
        return re.sub(r'\bFigs\.?\b', 'Figures', surface, flags=re.IGNORECASE)
    if rule_id in ("ref_figs_single_num", "cap_figs_single_num"):
        return re.sub(r'\bFigs\.?\b', 'Figures', surface, flags=re.IGNORECASE)
    if rule_id in ("ref_boxes_dotted", "cap_boxes_dotted"):
        return re.sub(r'\bBoxes?\.?\b', 'Boxes', surface, flags=re.IGNORECASE)
    if rule_id in ("ref_boxes_single_num", "cap_boxes_single_num"):
        return re.sub(r'\bBoxes?\.?\b', 'Boxes', surface, flags=re.IGNORECASE)
    # 11. Ordinals Spelled Out -> Numeric
    ordinal_map = {
        "first": "1st", "second": "2nd", "third": "3rd", "fourth": "4th",
        "fifth": "5th", "sixth": "6th", "seventh": "7th", "eighth": "8th",
        "ninth": "9th", "tenth": "10th", "eleventh": "11th", "twelfth": "12th",
        "thirteenth": "13th", "fourteenth": "14th", "fifteenth": "15th",
        "sixteenth": "16th", "seventeenth": "17th", "eighteenth": "18th",
        "nineteenth": "19th", "twentieth": "20th", "twenty-first": "21st",
        "twenty-second": "22nd", "twenty-third": "23rd", "twenty-fourth": "24th",
        "twenty-fifth": "25th", "thirtieth": "30th", "fortieth": "40th",
        "fiftieth": "50th", "sixtieth": "60th", "seventieth": "70th",
        "eightieth": "80th", "ninetieth": "90th", "hundredth": "100th",
        "thousandth": "1000th",
    }
    if rule_id == "ordinal_spelled":
        w = surface.lower().strip()
        if w in ordinal_map:
            return ordinal_map[w]
    # 12. Fractions
    if rule_id == "fract_symbol_half": return "1/2"
    if rule_id == "fract_symbol_quarter": return "1/4"
    if rule_id == "fract_symbol_three_quarters": return "3/4"
    # 13. Degree -> Symbol
    if rule_id == "deg_celsius_word": return "°C"
    if rule_id == "deg_fahrenheit_word": return "°F"
    if rule_id in ("deg_word_degrees", "deg_word_degree"): return "°"
    # 14. Symbols
    if rule_id == "sym_trademark_text": return "™"
    if rule_id == "sym_registered_text": return "®"
    if rule_id == "sym_copyright_text": return "©"
    # 15. Ranges En-Dash
    if rule_id == "range_hyphen":
        try:
            pt = re.compile(r'(\d)\s*-\s*(\d)')
            return pt.sub(r'\1–\2', surface)
        except Exception:
            return None
    if rule_id == "range_to":
        try:
            pt = re.compile(r'(\d)\s+to\s+(\d)')
            return pt.sub(r'\1–\2', surface)
        except Exception:
            return None
    # 16. Quotes
    if rule_id == "quote_single_straight":
        # "'hello'" -> "‘hello’"
        try:
            pt = re.compile(r"'([^']+)'")
            return pt.sub(r'‘\1’', surface)
        except Exception:
            return None
    if rule_id == "quote_double_straight":
        try:
            pt = re.compile(r'"([^"]+)"')
            return pt.sub(r'“\1”', surface)
        except Exception:
            return None
    # 17. Fold (Num fold -> Num-fold)
    if rule_id == "fold_numeral_open":
        return surface.replace(' ', '-')
    if rule_id == "fold_numeral_closed":
        v = re.findall(r'\d+', surface)
        if v: return f"{v[0]}-fold"
    # 18. Caps after colon
    if rule_id == "caps_after_colon":
        return surface[:2] + surface[2:].lower()
    if rule_id == "lowercase_after_colon":
        return surface[:2] + surface[2:].upper()
    # 19. Latin abbreviations
    if rule_id in ("latin_eg_nodots", "latin_eg_upper", "latin_eg_spaced"): return "e.g.,"
    if rule_id in ("latin_ie_nodots", "latin_ie_upper", "latin_ie_spaced"): return "i.e.,"
    if rule_id == "latin_etc_nodot": return "etc."
    # 20. Thousand separated
    if rule_id == "thous_sep_missing":
        try:
            pt = re.compile(r'(?<!\d)(\d{1,3})(?=(?:\d{3})+(?!\d))')
            return pt.sub(r'\1,', surface)
        except Exception:
            pass
    if rule_id in ("thous_sep_space", "thous_sep_nbsp"):
        return surface.replace(' ', ',').replace('\u00A0', ',')
    # 21. Time unit abbrevs (table mode filters happen naturally, rule_id still applies)
    if rule_id in ("unit_hour_h", "unit_hour_hr"):
        return "hours"
    if rule_id == "unit_hour_spelled":
        return "hr"
    if rule_id == "unit_min_abbr":
        return "minutes"
    if rule_id == "unit_min_spelled":
        return "min"
    if rule_id in ("unit_sec_s", "unit_sec_abbr"):
        return "seconds"
    if rule_id == "unit_sec_spelled":
        return "sec"
    if rule_id == "time_y_word":
        v = re.sub(r'\s*(?:y|yr|yrs)\b', '', surface, flags=re.IGNORECASE)
        return f"{v} years"
    # 22. Numbers (0-9, 0-99 mapping numerals <-> spelled)
    if rule_id in ("num_single_numeral", "num_double_numeral", "num_zero_numeral"):
        # numeral to spelled
        val = int(surface)
        ones = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]
        teens = ["ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
        tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
        if val < 10: return ones[val]
        elif val < 20: return teens[val-10]
        else:
            t, o = divmod(val, 10)
            return tens[t] if o == 0 else f"{tens[t]}-{ones[o]}"
    if rule_id in ("num_single_spelled", "num_double_spelled", "num_zero_spelled"):
        # spelled to numeral
        w = surface.lower().strip()
        ones_map = {"zero":0,"one":1,"two":2,"three":3,"four":4,"five":5,"six":6,"seven":7,"eight":8,"nine":9}
        teens_map = {"ten":10,"eleven":11,"twelve":12,"thirteen":13,"fourteen":14,"fifteen":15,"sixteen":16,"seventeen":17,"eighteen":18,"nineteen":19}
        tens_map = {"twenty":20,"thirty":30,"forty":40,"fifty":50,"sixty":60,"seventy":70,"eighty":80,"ninety":90}
        if w in ones_map: return str(ones_map[w])
        if w in teens_map: return str(teens_map[w])
        parts = re.split(r'[-\s]+', w)
        if len(parts) == 1:
            if parts[0] in tens_map: return str(tens_map[parts[0]])
        elif len(parts) == 2:
            if parts[0] in tens_map and parts[1] in ones_map:
                return str(tens_map[parts[0]] + ones_map[parts[1]])
    return None
# ────────────────────────────────────────────────────────────────────────────
# Multi-option replacements for dropdown UI in editor_review.html
# ────────────────────────────────────────────────────────────────────────────
_NUM_WORDS = {1:"one",2:"two",3:"three",4:"four",5:"five",6:"six",7:"seven",8:"eight",9:"nine",10:"ten",11:"eleven",12:"twelve"}
_ORD_WORDS = {1:"first",2:"second",3:"third",4:"fourth",5:"fifth",6:"sixth",7:"seventh",8:"eighth",9:"ninth",10:"tenth",
              11:"eleventh",12:"twelfth",13:"thirteenth",14:"fourteenth",15:"fifteenth",20:"twentieth",21:"twenty-first"}
def get_te_replacement_options(rule_id: str, m: re.Match) -> list[str]:
    """Return ordered list of valid replacement choices for editor review dropdown."""
    surface = m.group(0)
    nums = re.findall(r'\d+', surface)
    # Percent style
    if rule_id in ("percent_symbol", "percent_word", "per_cent_word", "percent_percentage"):
        return ["%", "per cent", "percent"]
    # Ellipsis style
    if rule_id in ("ellipsis_symbol", "ellipsis_3dots", "ellipsis_spaced"):
        return ["…", "...", ". . ."]
    # AM/PM
    if "ampm_am" in rule_id:
        return ["AM", "A.M.", "a.m.", "am"]
    if "ampm_pm" in rule_id:
        return ["PM", "P.M.", "p.m.", "pm"]
    # Era (AD/BC)
    if "era_ad" in rule_id:
        return ["AD", "A.D."]
    if "era_bc" in rule_id:
        return ["BC", "B.C."]
    # Number ranges
    if rule_id in ("range_to", "range_endash", "range_hyphen") and len(nums) >= 2:
        n1, n2 = nums[0], nums[1]
        return [f"{n1} to {n2}", f"{n1}–{n2}", f"{n1}-{n2}"]
    # Leading zero
    if rule_id == "leading_zero_missing":
        decimal_m = re.search(r'(\.\d+)', surface)
        if decimal_m:
            return [f"0{decimal_m.group(1)}", decimal_m.group(1)]
    if rule_id == "leading_zero_present":
        decimal_m = re.search(r'(0\.\d+)', surface)
        if decimal_m:
            dec_str = decimal_m.group(1)
            return [dec_str, dec_str.lstrip('0')]
    # Times symbol (×)
    if rule_id in ("times_symbol", "times_letter"):
        return ["×", "x"]
    # Versus
    if rule_id in ("versus_full", "versus_vs_dot", "versus_vs", "versus_v_dot", "versus_v"):
        return ["versus", "vs.", "vs", "v.", "v"]
    # Latin abbreviations
    if "latin_eg" in rule_id:
        return ["e.g.", "e.g.,", "eg.", "eg"]
    if "latin_ie" in rule_id:
        return ["i.e.", "i.e.,", "ie.", "ie"]
    if "latin_etc" in rule_id:
        return ["etc.", "etc"]
    # Numbers (digit/spell)
    if rule_id in ("num_single_numeral", "num_single_spelled", "num_double_numeral", "num_double_spelled",
                   "num_zero_numeral", "num_zero_spelled") and nums:
        n = int(nums[0])
        word = _NUM_WORDS.get(n, str(n))
        if re.match(r'^\d+$', surface):
            return [str(n), word]
        else:
            return [word, str(n)]
    # Ordinals
    if rule_id in ("ordinal_st", "ordinal_nd", "ordinal_rd", "ordinal_th") and nums:
        n = int(nums[0])
        word = _ORD_WORDS.get(n, surface)
        suffix = "st" if n % 10 == 1 and n % 100 != 11 else \
                 "nd" if n % 10 == 2 and n % 100 != 12 else \
                 "rd" if n % 10 == 3 and n % 100 != 13 else "th"
        return [f"{n}{suffix}", word]
    if rule_id == "ordinal_spelled" and nums:
        n = int(nums[0])
        word = _ORD_WORDS.get(n, surface)
        suffix = "st" if n % 10 == 1 and n % 100 != 11 else \
                 "nd" if n % 10 == 2 and n % 100 != 12 else \
                 "rd" if n % 10 == 3 and n % 100 != 13 else "th"
        return [word, f"{n}{suffix}"]
    # Trademark / Register / Copyright
    if rule_id in ("sym_trademark_char", "sym_trademark_text"):
        return ["™", "(TM)"]
    if rule_id in ("sym_registered_char", "sym_registered_text"):
        return ["®", "(R)"]
    if rule_id in ("sym_copyright_char", "sym_copyright_text"):
        return ["©", "(C)"]
    # Quote style
    if rule_id in ("quote_double_curly", "quote_double_straight"):
        return ['"', '"', '"']
    if rule_id in ("quote_single_curly", "quote_single_straight"):
        return ["'", "'", "'"]
    # Fold
    if "fold" in rule_id:
        n_match = re.search(r'(\d+|one|two|three|four|five|six|seven|eight|nine|ten)', surface, re.I)
        if n_match:
            part = n_match.group(1)
            try:
                n = int(part)
                word = _NUM_WORDS.get(n, part)
            except ValueError:
                n = next((k for k, v in _NUM_WORDS.items() if v == part.lower()), None)
                word = part.lower()
                part = str(n) if n else part
            return [f"{part}-fold", f"{part} fold", f"{word}fold"]
    # Times word (X times)
    if "times_numeral" in rule_id or "times_word" in rule_id or "times_twice" in rule_id:
        if nums:
            n = int(nums[0])
            word = _NUM_WORDS.get(n, str(n))
            if n == 2:
                return [f"{n} times", "twice", f"{word} times"]
            return [f"{n} times", f"{word} times"]
        return ["twice", "2 times"]
    # Date
    _MONTHS = {"jan":"January","feb":"February","mar":"March","apr":"April",
               "may":"May","jun":"June","jul":"July","aug":"August",
               "sep":"September","oct":"October","nov":"November","dec":"December",
               "january":"January","february":"February","march":"March",
               "april":"April","june":"June","july":"July","august":"August",
               "september":"September","october":"October","november":"November","december":"December"}
    if "date" in rule_id:
        month_m = re.search(r'([A-Za-z]+)', surface)
        day_m = re.search(r'\b(\d{1,2})\b', surface)
        year_m = re.search(r'\b(\d{4})\b', surface)
        if month_m and day_m and year_m:
            month = _MONTHS.get(month_m.group(1).lower(), month_m.group(1))
            d, y = day_m.group(1), year_m.group(1)
            return [f"{d} {month} {y}", f"{month} {d}, {y}"]
    # Century
    if "century" in rule_id and nums:
        n = int(nums[0])
        word = _ORD_WORDS.get(n, f"{n}th")
        suffix = "st" if n % 10 == 1 and n % 100 != 11 else \
                 "nd" if n % 10 == 2 and n % 100 != 12 else \
                 "rd" if n % 10 == 3 and n % 100 != 13 else "th"
        plural = "centuries" if "centur" in surface.lower() and \
                 ("centuri" in surface.lower() or "centuries" in surface.lower()) else "century"
        return [f"{n}{suffix} {plural}", f"{word} {plural}"]
    # Fractions
    _FRAC_MAP = {
        ("1","2"): ("½", "1/2", "one-half"),
        ("1","3"): ("⅓", "1/3", "one-third"),
        ("2","3"): ("⅔", "2/3", "two-thirds"),
        ("1","4"): ("¼", "1/4", "one-quarter"),
        ("3","4"): ("¾", "3/4", "three-quarters"),
    }
    if "fraction" in rule_id and len(nums) >= 2:
        key = (nums[0], nums[1])
        if key in _FRAC_MAP:
            return list(_FRAC_MAP[key])
        return [f"{nums[0]}/{nums[1]}"]
    # Temperature
    if "celsius" in rule_id or "deg_c" in rule_id:
        n = nums[0] if nums else "#"
        return [f"{n}°C", f"{n} degrees Celsius"]
    if "fahrenheit" in rule_id or "deg_f" in rule_id:
        n = nums[0] if nums else "#"
        return [f"{n}°F", f"{n} degrees Fahrenheit"]
    if rule_id == "degree_symbol_alone":
        return ["°C", "°F"]
    # Comparison symbols
    if rule_id in ("cmp_gte_sym", "cmp_gte_words"):
        return ["≥", "greater than or equal to"]
    if rule_id in ("cmp_lte_sym", "cmp_lte_words"):
        return ["≤", "less than or equal to"]
    if rule_id in ("cmp_gt_sym", "cmp_gt_words"):
        return [">", "greater than"]
    if rule_id in ("cmp_lt_sym", "cmp_lt_words"):
        return ["<", "less than"]
    if rule_id in ("cmp_approx_words", "cmp_tilde"):
        return ["≈", "~", "approximately"]
    # Thousands separator
    if "thousands" in rule_id and nums:
        raw = re.sub(r'[,\s ]', '', surface)
        try:
            n = int(raw)
            return [f"{n:,}", f"{n:}"]
        except ValueError:
            pass
    # Virgule (per)
    if rule_id in ("per_word", "virgule_slash"):
        slash_m = re.search(r'(\S+)/(\S+)', surface)
        if slash_m:
            return [surface, f"{slash_m.group(1)} per {slash_m.group(2)}"]
    # Inline list marker
    if "list" in rule_id:
        return ["(a)", "(1)", "(i)", "a)", "1)", "i)", "A)", "I)"]
    # SI units
    _SI = {"mg":("mg","milligram"), "kg":("kg","kilogram"),
           "ml":("ml","mL","millilitre"), "mcg":("mcg","microgram"),
           "mm":("mm","millimetre"), "cm":("cm","centimetre"),
           "km":("km","kilometre")}
    for abbr, opts in _SI.items():
        if abbr in rule_id:
            return list(opts)
    # Time units
    if "unit_hour" in rule_id:
        n = nums[0] if nums else ""
        pre = f"{n} " if n else ""
        return [f"{pre}h", f"{pre}hr", f"{pre}hour"]
    if "unit_min" in rule_id:
        n = nums[0] if nums else ""
        pre = f"{n} " if n else ""
        return [f"{pre}min", f"{pre}minute"]
    if "unit_sec" in rule_id:
        n = nums[0] if nums else ""
        pre = f"{n} " if n else ""
        return [f"{pre}s", f"{pre}sec", f"{pre}second"]
    # Gas abbreviations (inline vs subscript)
    _GAS = {"paco2":("PaCO₂","PaCO2"), "pco2":("PCO₂","PCO2"),
            "pao2":("PaO₂","PaO2"), "fio2":("FiO₂","FIO2"),
            "sao2":("SaO₂","SaO2"), "spo2":("SpO₂","SpO2")}
    for gas, opts in _GAS.items():
        if gas in rule_id:
            return list(opts)
    return []
