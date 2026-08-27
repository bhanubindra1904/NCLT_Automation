import glob
import os
import re


PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_ROOT_DIR = os.path.join(PROJECT_DIR, "outputs")
RUN_FOLDER_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$")


def configured_output_directory(default=None):
    configured = os.getenv("NCLT_OUTPUT_DIR", "").strip()
    return os.path.abspath(configured or default or os.getcwd())


def latest_output_directory(required_pattern=None):
    if not os.path.isdir(OUTPUT_ROOT_DIR):
        return None

    directories = [
        os.path.join(OUTPUT_ROOT_DIR, name)
        for name in os.listdir(OUTPUT_ROOT_DIR)
        if RUN_FOLDER_PATTERN.fullmatch(name)
        and os.path.isdir(os.path.join(OUTPUT_ROOT_DIR, name))
    ]
    directories.sort(key=os.path.getmtime, reverse=True)

    for directory in directories:
        if not required_pattern or glob.glob(os.path.join(directory, required_pattern)):
            return directory
    return None


def find_latest_file(pattern, preferred_directory=None):
    preferred = preferred_directory or os.getenv("NCLT_OUTPUT_DIR", "").strip()
    if preferred:
        preferred_files = glob.glob(os.path.join(os.path.abspath(preferred), pattern))
        if preferred_files:
            return max(preferred_files, key=os.path.getmtime)

    candidates = glob.glob(os.path.join(os.getcwd(), pattern))
    if os.path.isdir(OUTPUT_ROOT_DIR):
        for name in os.listdir(OUTPUT_ROOT_DIR):
            directory = os.path.join(OUTPUT_ROOT_DIR, name)
            if RUN_FOLDER_PATTERN.fullmatch(name) and os.path.isdir(directory):
                candidates.extend(glob.glob(os.path.join(directory, pattern)))

    return max(candidates, key=os.path.getmtime) if candidates else None


PARTY_SEPARATOR_PATTERN = re.compile(
    r"(?<![A-Z0-9])(?:(?:V\s*/\s*S|V\s*S|VERSUS)(?:\.)?|V\.)(?![A-Z0-9])",
    re.IGNORECASE,
)
BARE_V_PATTERN = re.compile(r"(?<![A-Z0-9])V(?![A-Z0-9])", re.IGNORECASE)
MAIN_MATTER_PATTERN = re.compile(
    r"\bIN\s+THE\s+MATTER\s+OF\b",
    re.IGNORECASE,
)
INTERIM_APPLICATION_PATTERN = re.compile(
    r"^(?:IA|I\.A\.|IVN|IVNP|CA|MA|CONT\.?A)\s*[/.(\-]",
    re.IGNORECASE,
)

LEGAL_ENTITY_WORDS = {
    "COMPANY",
    "CORPORATION",
    "INCORPORATED",
    "LIMITED",
    "LLP",
    "LTD",
    "PLC",
    "PRIVATE",
    "PVT",
}
ROLE_PHRASES = {
    "BOARD OF DIRECTORS",
    "DIRECTOR",
    "IRP",
    "LIQUIDATOR",
    "PERSONAL GUARANTOR",
    "RESOLUTION PROFESSIONAL",
    "RP",
    "SUSPENDED BOARD",
}
COMPANY_TOKEN_STOPWORDS = {
    "AND",
    "COMPANY",
    "CORPORATION",
    "INDIA",
    "LIMITED",
    "LLP",
    "LTD",
    "PRIVATE",
    "PVT",
    "THE",
}

CASE_REFERENCE_PATTERNS = (
    re.compile(
        r"""
        (?<![A-Z0-9])
        (?:R\s*)?C\s*\.?\s*P\s*\.?
        \s*(?:NO\.?\s*)?
        \(\s*I\s*B\s*C?\s*\)
        \s*(?:NO\.?\s*)?
        (?:[/:\-]\s*)?
        \d+[A-Z]?
        (?:
            \s*/\s*[A-Z0-9.\-]+
            |
            \s*\(\s*[A-Z0-9.\-]+\s*\)
        ){0,5}
        \s*(?:/\s*)?(?:OF\s*)?(?:19|20)\d{2}\b
        """,
        re.IGNORECASE | re.VERBOSE,
    ),
    re.compile(
        r"""
        (?<![A-Z0-9(/])
        (?:IBA|IB|TCP|TP)(?![A-Z])
        \s*(?:NO\.?\s*)?
        (?:[/:\-]\s*)?
        \d+[A-Z]?
        (?:
            \s*/\s*[A-Z0-9.\-]+
            |
            \s*\(\s*[A-Z0-9.\-]+\s*\)
        ){0,5}
        \s*(?:/\s*)?(?:OF\s*)?(?:19|20)\d{2}\b
        """,
        re.IGNORECASE | re.VERBOSE,
    ),
    re.compile(
        r"""
        (?<![A-Z0-9])
        \(\s*IBC?\s*\)
        \s*[/:\-]\s*\d+[A-Z]?
        (?:
            \s*/\s*[A-Z0-9.\-]+
            |
            \s*\(\s*[A-Z0-9.\-]+\s*\)
        ){0,4}
        \s*(?:/\s*)?(?:19|20)\d{2}\b
        """,
        re.IGNORECASE | re.VERBOSE,
    ),
    re.compile(
        r"""
        (?<![A-Z0-9])
        (?:R\s*)?C\s*\.?\s*P\s*\.?
        \s*/\s*IBC?\s*/\s*\d+[A-Z]?
        (?:
            \s*/\s*[A-Z0-9.\-]+
            |
            \s*\(\s*[A-Z0-9.\-]+\s*\)
        ){0,5}
        \s*(?:/\s*)?(?:19|20)\d{2}\b
        """,
        re.IGNORECASE | re.VERBOSE,
    ),
    re.compile(
        r"""
        (?<![A-Z0-9])
        (?:R\s*)?C\s*\.?\s*P\s*\.?
        \s*(?:NO\.?\s*)?
        [/:\-]\s*\d+[A-Z]?
        (?:
            \s*/\s*[A-Z0-9.\-]+
            |
            \s*\(\s*[A-Z0-9.\-]+\s*\)
        ){0,5}
        \s*(?:/\s*)?(?:19|20)\d{2}\b
        """,
        re.IGNORECASE | re.VERBOSE,
    ),
)

CIN_PATTERN = re.compile(r"\b([LU][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})\b", re.I)
LLPIN_PATTERN = re.compile(r"\b([A-Z]{2,4}-\d{4})\b", re.I)


def repair_doubled_ocr(value):
    """Collapse tokens whose PDF text layer duplicates every character."""
    source = "" if value is None else str(value)

    def collapse_token(match):
        token = match.group(0)
        if len(token) < 8 or len(token) % 2 or token.isdigit():
            return token
        if all(token[index] == token[index + 1] for index in range(0, len(token), 2)):
            return token[::2]
        return token

    return re.sub(r"\S+", collapse_token, source)


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", repair_doubled_ocr(value)).strip()


def normalize_text(value):
    return re.sub(r"[^A-Z0-9]+", " ", clean_text(value).upper()).strip()


def normalize_company_key(value):
    normalized = normalize_text(value)
    normalized = re.sub(r"\bPRIVATE\s*LIMITED\b", "PRIVATE LIMITED", normalized)
    normalized = re.sub(r"\bPVT\s*LTD\b", "PRIVATE LIMITED", normalized)
    normalized = re.sub(r"\bPVT\b", "PRIVATE", normalized)
    normalized = re.sub(r"\bLTD\b", "LIMITED", normalized)
    normalized = re.sub(r"\bM S\b", "", normalized)
    return clean_text(normalized)


def _join_split_years(value):
    value = re.sub(r"\b(19|20)\s+(\d{2})\b", r"\1\2", value)
    return re.sub(r"\b((?:19|20)\d)\s+(\d)\b", r"\1\2", value)


def _is_interim_reference(source, start):
    prefix = normalize_text(source[max(0, start - 24) : start])
    return bool(
        re.search(
            r"(?:^|\s)(?:IA|IVN|IVN P|IVNP|CA|MA|CONT A|APPLICATION|"
            r"INTERVENTION PETITION|INTERLOCUTORY APPLICATION)\s*$",
            prefix,
        )
    )


def extract_case_references(value):
    source = _join_split_years(repair_doubled_ocr(value))
    matches = []
    for pattern_index, pattern in enumerate(CASE_REFERENCE_PATTERNS):
        for match in pattern.finditer(source):
            matches.append((match.start(), match.end(), pattern_index, match.group(0)))

    accepted_ranges = []
    references = []
    seen = set()
    for start, end, pattern_index, matched_text in sorted(
        matches,
        key=lambda item: (item[0], -(item[1] - item[0]), item[2]),
    ):
        if _is_interim_reference(source, start):
            continue
        if any(start < used_end and end > used_start for used_start, used_end in accepted_ranges):
            continue
        display = clean_text(matched_text).strip(" ,;:-")
        canonical_source = re.sub(r"\b(?:OF|NO)\b", "/", display.upper())
        key = re.sub(r"[^A-Z0-9]+", "", canonical_source)
        if not key or key in seen:
            continue
        references.append({"display": display, "key": key, "pattern": pattern_index})
        accepted_ranges.append((start, end))
        seen.add(key)
    return references


def extract_case_id_from_cells(cells):
    for cell in cells:
        references = extract_case_references(cell)
        if references:
            return references[0]["display"]
    return ""


def recover_case_id_from_row(page, row_bbox, table_bbox=None):
    """Recover case text omitted by a vertically merged PDF table cell."""
    _, top, _, bottom = row_bbox
    for right_fraction in (0.22, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5):
        for vertical_padding in (1, 5, 10):
            crop_bbox = (
                0,
                max(0, top - vertical_padding),
                min(page.width, page.width * right_fraction),
                min(page.height, bottom + vertical_padding),
            )
            text = page.crop(crop_bbox).extract_text(layout=True) or ""
            references = extract_case_references(text)
            if references:
                return references[0]["display"]

    if table_bbox and table_bbox[3] > bottom:
        for right_fraction in (0.22, 0.25, 0.3, 0.35):
            crop_bbox = (
                0,
                max(0, top - 1),
                min(page.width, page.width * right_fraction),
                min(page.height, table_bbox[3] + 2),
            )
            text = page.crop(crop_bbox).extract_text(layout=True) or ""
            references = extract_case_references(text)
            if references:
                return references[0]["display"]
    return ""


def _clean_party_side(value, debtor=False):
    result = clean_text(value)
    result = re.sub(r"^(?:M\s*/\s*S|MR|MRS|MS)\.?\s+", "", result, flags=re.I)
    if debtor:
        result = re.split(
            r"\s+(?:&\s*(?:ORS?|ANR)|AND\s+OTHERS|THROUGH\s+(?:RP|IRP|LIQUIDATOR)|"
            r"REPRESENTED\s+BY|ADV\.)\b",
            result,
            maxsplit=1,
            flags=re.I,
        )[0]
    result = re.sub(r"\bPRIVATE\s*LIMITED\b", "PRIVATE LIMITED", result, flags=re.I)
    result = re.sub(r"\bPVT\s*LTD\b\.?", "PRIVATE LIMITED", result, flags=re.I)
    result = re.sub(r"\bPVT\b\.?", "PRIVATE", result, flags=re.I)
    result = re.sub(r"\bLTD\b\.?", "LIMITED", result, flags=re.I)
    if debtor:
        entity_end = re.search(
            r"\b(?:PRIVATE\s*LIMITED|LIMITED|LLP|PLC|INCORPORATED|CORPORATION)\b",
            result,
            flags=re.I,
        )
        if entity_end:
            result = result[: entity_end.end()]
    result = re.sub(r"\s+", " ", result).strip(" ,.;:-")
    return result


def has_legal_entity_word(value):
    normalized = normalize_company_key(value)
    return any(re.search(rf"\b{re.escape(word)}\b", normalized) for word in LEGAL_ENTITY_WORDS)


def has_role_phrase(value):
    normalized = normalize_text(value)
    return any(re.search(rf"\b{re.escape(phrase)}\b", normalized) for phrase in ROLE_PHRASES)


def extract_main_parties(value):
    """Extract the main creditor/debtor pair, preferring the final main-matter clause."""
    original = clean_text(value)
    main_matches = list(MAIN_MATTER_PATTERN.finditer(original))
    source = original[main_matches[-1].end() :] if main_matches else original
    separator_pattern = PARTY_SEPARATOR_PATTERN
    separators = list(separator_pattern.finditer(source))
    all_separators = list(separator_pattern.finditer(original))
    if not separators:
        separator_pattern = BARE_V_PATTERN
        separators = list(separator_pattern.finditer(source))
        all_separators = list(separator_pattern.finditer(original))
    if not separators:
        return {
            "creditor": "",
            "debtor": "",
            "confident": False,
            "reason": "No creditor/debtor separator",
            "used_main_matter": bool(main_matches),
        }

    separator = separators[0]
    creditor = _clean_party_side(source[: separator.start()])
    debtor = _clean_party_side(source[separator.end() :], debtor=True)

    reason = ""
    confident = bool(creditor and debtor)
    if has_role_phrase(debtor):
        confident = False
        reason = "Debtor text is a person or insolvency role"
    elif not main_matches and INTERIM_APPLICATION_PATTERN.search(source):
        confident = False
        reason = "Interim-application parties detected; main matter is missing"
    elif not main_matches and len(all_separators) > 1:
        confident = False
        reason = "Multiple party pairs without a main-matter label"
    elif not main_matches and has_role_phrase(creditor) and not has_legal_entity_word(debtor):
        confident = False
        reason = "Interim-application parties detected"
    elif len(normalize_text(debtor)) < 4:
        confident = False
        reason = "Debtor text is incomplete"

    return {
        "creditor": creditor,
        "debtor": debtor,
        "confident": confident,
        "reason": reason,
        "used_main_matter": bool(main_matches),
    }


def extract_party_from_cells(cells):
    candidates = []
    candidate_index = 0
    for cell in cells:
        for segment in re.split(r"[\r\n]+", str(cell or "")):
            source_cell = clean_text(segment)
            if not source_cell:
                continue
            parsed = extract_main_parties(source_cell)
            if not parsed["debtor"]:
                continue
            quality = (
                int(parsed["used_main_matter"]),
                int(parsed["confident"]),
                int(has_legal_entity_word(parsed["debtor"])),
                -candidate_index,
                len(normalize_text(parsed["debtor"])),
            )
            candidates.append((quality, parsed, source_cell))
            candidate_index += 1
    if not candidates:
        return None
    _, parsed, source_cell = max(candidates, key=lambda item: item[0])
    return {**parsed, "source_cell": source_cell}


def company_tokens(value):
    tokens = {
        token
        for token in normalize_company_key(value).split()
        if len(token) >= 3 and token not in COMPANY_TOKEN_STOPWORDS
    }
    return tokens


def extract_verified_id(page_text, company_name):
    """Return an ID only when its search-result context matches the company name."""
    expected_tokens = company_tokens(company_name)
    if not expected_tokens:
        return None

    candidates = []
    for pattern in (CIN_PATTERN, LLPIN_PATTERN):
        for match in pattern.finditer(page_text or ""):
            context = (page_text or "")[max(0, match.start() - 350) : match.end() + 350]
            overlap = len(expected_tokens & company_tokens(context))
            candidates.append((overlap, match.group(1).upper()))

    if not candidates:
        return None
    overlap, identifier = max(candidates, key=lambda item: item[0])
    required_overlap = 1 if len(expected_tokens) == 1 else min(2, len(expected_tokens))
    return identifier if overlap >= required_overlap else None
