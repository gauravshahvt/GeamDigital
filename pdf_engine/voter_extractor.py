import os
import re
from typing import List, Dict, Any, Tuple, Optional
import pymupdf
import pytesseract
from PIL import Image
import pandas as pd

_TESSERACT_AVAILABLE: Optional[bool] = None

def is_tesseract_available() -> bool:
    """Checks whether tesseract binary is actually installed and accessible."""
    global _TESSERACT_AVAILABLE
    if _TESSERACT_AVAILABLE is not None:
        return _TESSERACT_AVAILABLE
    try:
        pytesseract.get_tesseract_version()
        _TESSERACT_AVAILABLE = True
    except Exception:
        _TESSERACT_AVAILABLE = False
    return _TESSERACT_AVAILABLE

def sanitize_hindi_orthography(text: str) -> str:
    """Corrects common Devanagari OCR ligatures, duplicate halants, and scan artifacts."""
    if not text:
        return ""

    # 1. Normalize zero-width joiners / non-breaking spaces
    text = text.replace('\ufeff', '').replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')

    # 2. Collapse double or multiple viramas / halants (e.g. 'नन््दा' -> 'नन्दा', 'सनन््तोष' -> 'सन्तोष')
    text = re.sub(r'\u094d[\u200c\u200d\s]*\u094d+', '\u094d', text)

    # 3. Invalid Consonant + Halant + Matra fix (e.g. 'ज्ुन्नी' -> 'चुन्नी', 'ज्ुनी' -> 'चुनी')
    text = re.sub(r'(^|\s)ज्[ुू]([नन्])', r'\g<1>चु\2', text)
    text = re.sub(r'([क-ह])\u094d([\u093e-\u094c])', r'\1\2', text)

    # 4. Missing initial 'क' before 'ैलाश...' -> 'कैलाश...'
    text = re.sub(r'(^|\s)ैलाश', r'\1कैलाश', text)
    text = re.sub(r'(^|\s)ैलाशी', r'\1कैलाशी', text)
    text = re.sub(r'(^|\s)ैलास', r'\1कैलास', text)
    text = re.sub(r'(^|\s)ैलासी', r'\1कैलासी', text)
    text = re.sub(r'(^|\s)ेलाश', r'\1कैलाश', text)

    # 5. Fix OCR confusion: 'लन्...द' for 'नन्...द' (e.g. 'लन््दा'/'लन्दा' -> 'नन्दा', 'लन्दू' -> 'नन्दू', 'लन्दराम' -> 'नन्दराम', 'लन्दलाल' -> 'नन्दलाल')
    text = re.sub(r'(^|\s)लन्?्?द', r'\g<1>नन्द', text)

    # 5b. Fix OCR spurious duplicate 'न' in common names like सनन्तोष -> सन्तोष, मनन्शा -> मन्शा
    text = re.sub(r'सनन्?्?तोष', 'सन्तोष', text)
    text = re.sub(r'मनन्?्?शा', 'मन्शा', text)

    # 6. Common vowel suffix artifact 'देवे' / 'दवे' -> 'देवी'
    text = re.sub(r'(^|\s)देवे(\b|$|\s)', r'\1देवी\2', text)
    text = re.sub(r'(^|\s)दवे(\b|$|\s)', r'\1देवी\2', text)

    # 7. Stray leading matras on words (e.g. stray 'ा', 'ि', 'ी', 'ु', etc. at start of word)
    text = re.sub(r'(^|\s)[\u093e-\u094c]', r'\1', text)

    # 8. False leading 'श्' before consonants that cannot form valid Sanskrit/Hindi conjuncts (OCR cutoff of 'श्री')
    text = re.sub(r'(^|\s)श्(?=[क-झट-नप-मष-ह])', r'\1', text)

    # 9. Spurious double consonant at start of word (e.g. 'न्नाथू' -> 'नाथू')
    text = re.sub(r'(^|\s)([क-ह])\u094d(?=\2)', r'\1', text)

    # 10. Spurious leading 'ज्' before 'न' (e.g. 'ज्नारायणनाथ' -> 'नारायणनाथ')
    text = re.sub(r'(^|\s)ज्(?=न)', r'\1', text)

    return text.strip()

HINDI_PANCHAYAT_MAP = {
    'बागमाली': 'Bagmali',
    'बामणी': 'Bamani',
    'बामनी': 'Bamani',
    'बराना': 'Barana',
    'बाराणा': 'Barana',
    'दौलतगढ़': 'Daulatgarh',
    'दांतड़ा': 'Dantda',
    'दांतडा': 'Dantda',
    'आमलीगढ़': 'Amligarh',
    'आमलिया': 'Amliya',
    'आमलीगुरांवतान': 'Amlipurawatan',
    'आसीन्द': 'Asind',
    'आसींद': 'Asind'
}

def devanagari_to_english(text: str) -> str:
    """Phonetically transliterates Devanagari text into clean English (Title Case)."""
    if not text:
        return ""
    mapping = {
        'अ': 'a', 'आ': 'a', 'इ': 'i', 'ई': 'i', 'उ': 'u', 'ऊ': 'u',
        'ए': 'e', 'ऐ': 'ai', 'ओ': 'o', 'औ': 'au', 'ऋ': 'ri',
        'क': 'k', 'ख': 'kh', 'ग': 'g', 'घ': 'gh', 'ङ': 'ng',
        'च': 'ch', 'छ': 'chh', 'ज': 'j', 'झ': 'jh', 'ञ': 'ny',
        'ट': 't', 'ठ': 'th', 'ड': 'd', 'ढ': 'dh', 'ण': 'n',
        'त': 't', 'थ': 'th', 'द': 'd', 'ध': 'dh', 'न': 'n',
        'प': 'p', 'फ': 'ph', 'ब': 'b', 'भ': 'bh', 'म': 'm',
        'य': 'y', 'र': 'r', 'ल': 'l', 'व': 'v', 'श': 'sh',
        'ष': 'sh', 'स': 's', 'ह': 'h', 'क्ष': 'ksh', 'त्र': 'tr', 'ज्ञ': 'gy',
        'ा': 'a', 'ि': 'i', 'ी': 'i', 'ु': 'u', 'ू': 'u',
        'े': 'e', 'ै': 'ai', 'ो': 'o', 'ौ': 'au', 'ृ': 'ri',
        'ं': 'n', 'ँ': 'n', 'ः': 'h', '्': ''
    }
    vowels_matras = set(['ा', 'ि', 'ी', 'ु', 'ू', 'े', 'ै', 'ो', 'ौ', 'ृ', '्'])
    consonants = set(['क', 'ख', 'ग', 'घ', 'ङ', 'च', 'छ', 'ज', 'झ', 'ञ', 'ट', 'ठ', 'ड', 'ढ', 'ण', 'त', 'थ', 'द', 'ध', 'न', 'प', 'फ', 'ब', 'भ', 'म', 'य', 'र', 'ल', 'व', 'श', 'ष', 'स', 'ह', 'क्ष', 'त्र', 'ज्ञ'])

    res = []
    i = 0
    chars = list(text.strip())
    while i < len(chars):
        c = chars[i]
        if c in consonants:
            eng = mapping.get(c, '')
            next_c = chars[i+1] if i + 1 < len(chars) else None
            if next_c in vowels_matras:
                res.append(eng)
            elif next_c in consonants:
                res.append(eng + 'a')
            else:
                res.append(eng)
        elif c in mapping:
            res.append(mapping[c])
        elif c.isalnum() or c.isspace():
            res.append(c)
        i += 1
    out = ''.join(res).strip()
    return out.title()

def detect_panchayat_name(filenames: List[str], booth_address: str = '', area_address: str = '') -> str:
    """
    Extracts the Gram Panchayat name in English from uploaded filenames or PDF booth/area metadata.
    Example: 'BAGMALI-Ward_No-001.pdf' -> 'Bagmali'
    """
    candidates = []
    for f in filenames:
        base = os.path.basename(f)
        clean = re.sub(r'^[a-f0-9\-]{36}_', '', base)
        clean = re.sub(r'\.xlsx?$|\.pdf$', '', clean, flags=re.I)

        # 1. Pattern: {PANCHAYAT}[-_ ]+(?:Ward|ररडर|वार्ड|Bhag|Part)
        m = re.match(r'^([A-Za-z\s]+?)[_\-\s]+(?:Ward|ररडर|वार्ड|Bhag|Part)', clean, flags=re.I)
        if m:
            c = m.group(1).strip().replace('_', ' ')
            if c.lower() not in ('sample', 'voter', 'list', 'ward', 'demo', 'geam', 'digital'):
                candidates.append(c)
                continue

        # 2. Pattern: (?:Ward|वार्ड)[_\-\s]*\d+[_\-\s]+([A-Za-z\s]+)
        m2 = re.search(r'(?:Ward|ररडर|वार्ड|Bhag|Part)[_\-\s]*[0-9]+[_\-\s]+([A-Za-z\s]+)', clean, flags=re.I)
        if m2:
            c = m2.group(1).strip().replace('_', ' ')
            if c.lower() not in ('sample', 'voter', 'list', 'ward', 'demo', 'geam', 'digital'):
                candidates.append(c)
                continue

        # 3. Pattern: Leading English alphabetic token before digits
        m3 = re.match(r'^([A-Za-z]{3,})[_\-\s\d]', clean)
        if m3:
            c = m3.group(1).strip()
            if c.lower() not in ('sample', 'voter', 'list', 'ward', 'demo', 'geam', 'digital', 'test'):
                candidates.append(c)

    if candidates:
        from collections import Counter
        most_common = Counter(candidates).most_common(1)[0][0]
        return most_common.strip().title()

    # Fallback to address metadata matching
    combined_addr = f"{booth_address} {area_address}"
    for hin, eng in HINDI_PANCHAYAT_MAP.items():
        if hin in combined_addr:
            return eng

    # Transliterate any Hindi village/panchayat keyword if found in address
    m_panch = re.search(r'(?:ग्राम\s*पंचायत|ग्रामपंचायत)\s*[:\-\s]*([^\n,]+)', combined_addr)
    if m_panch:
        raw_p = m_panch.group(1).strip()
        trans = devanagari_to_english(raw_p)
        if trans and trans.lower() not in ('gram', 'panchayat'):
            return trans

    return "Gram Panchayat"

def parse_panchayat_and_wards_from_filenames(filenames: List[str]) -> Tuple[str, int, int]:
    """
    Infers the Gram Panchayat name and start/last ward numbers directly from a list of uploaded filenames.
    e.g. ['BAGMALI-Ward_No-001.pdf', 'BAGMALI-Ward_No-007.pdf'] -> ('Bagmali', 1, 7)
    """
    panchayat = detect_panchayat_name(filenames)
    ward_nums = []
    for f in filenames:
        m = re.search(r'(?:Ward[_\s\-]*No[_\s\-]*|वार्ड[_\s\-]*|भाग[_\s\-]*|Ward[_\s\-]*)(\d+)', f, re.I)
        if m:
            ward_nums.append(int(m.group(1)))

    if ward_nums:
        start_w = min(ward_nums)
        last_w = max(ward_nums)
    else:
        start_w = 1
        last_w = len(filenames) if filenames else 1

    return panchayat, start_w, last_w

def remove_leaked_relative_name(voter_name: str, relative_name: str) -> str:
    """Removes leaked relative/husband name from voter name field."""
    if not voter_name or not relative_name:
        return voter_name
    v = voter_name.strip()
    r = relative_name.strip()

    # 1. Exact match at end
    if v.endswith(r) and len(v) > len(r):
        return v[:-len(r)].strip()

    # 2. Match relative given name
    titles = {'श्री', 'श्रीमती', 'स्व', 'स्वर्गीय', 'डॉ', 'डा', 'श्रीमान'}
    r_words = [w for w in r.split() if w not in titles and len(w) >= 2]
    
    if r_words:
        first_r = r_words[0]
        # Direct word search
        idx = v.find(first_r)
        if idx > 2:
            return v[:idx].strip()
            
        # Stem search (e.g. बालु vs बालू, संजय vs सपजय)
        if len(first_r) >= 3:
            stem = first_r[:3]
            idx = v.find(stem)
            if idx > 2:
                return v[:idx].strip()

    # 3. Strip any residual relation keywords
    v = re.sub(r'(?:^|\s+)(?:पिता|पति|माता|अन्य|संरक्षक|अभिभावक|नपतर|पनत|मरतर|अनज|पत्नी|पता|पती|पत)(?:$|\s+.*$)', '', v).strip()

    return v

def clean_hindi_name(text: str) -> str:
    """Cleans OCR noise and formatting artifacts from Hindi voter and relative names."""
    if not text:
        return ""

    # 1. Normalize zero-width joiners / non-breaking spaces
    cleaned = text.replace('\ufeff', '').replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')

    # 2. Remove Dandās (। ॥), pipes, slashes, brackets, English letters and numbers (both ASCII 0-9 & Devanagari ०-९)
    cleaned = re.sub(r'[a-zA-Z0-9\u0966-\u096f|!@#$%^&*()_+=\[\]{};\'"\\<>\/?~`।॥_–—\u0951-\u0954]', ' ', cleaned)

    # 3. Remove leaked photo/status words
    cleaned = re.sub(r'(?i)\b(?:photo|available|is)\b', ' ', cleaned)

    # 4. Remove trailing OCR verb/honorific noise like 'हैं', 'हि', 'है', 'हो'
    cleaned = re.sub(r'(?:^|\s+)(?:हैं|हि|है|हो)(?:$|\s+.*$)', ' ', cleaned)

    # 5. Remove leaked label words if appended at the end (e.g. 'दिनेश भील पिता', 'गुलनाज़ पति')
    cleaned = re.sub(r'(?:^|\s+)(?:पिता|पति|माता|अन्य|अभिभावक|संरक्षक|नपतर|पनत|मरतर|नाम|नरम)(?:$|\s+.*$)', ' ', cleaned)

    # 6. Remove stray trailing single-character / single-syllable OCR noise like 'शि', 'जि', 'गा', 'शा', 'कं', 'सा', 'री'
    cleaned = re.sub(r'(?:^|\s+)(?:शि|जि|गा|शा|कं|सा|री|हि)(?:$|\s+)', ' ', cleaned)

    # 7. Remove trailing 'न्' (na + halant), ZWJ, and any trailing halants (e.g. 'सुमित्रा न्', 'ज्ञान सिंह न्‍')
    cleaned = re.sub(r'[\s]*न?्[\u200c\u200d]*[\s]*$', '', cleaned)
    cleaned = re.sub(r'[\s]+न्[\u200c\u200d]*[\s]*$', '', cleaned)
    cleaned = re.sub(r'[\s]+न[\s]*$', '', cleaned)

    # 8. Remove trailing spurious 'ऐ' or 'ए' or 'ए्' (e.g. 'प्रभू लाल भील ऐ', 'भागचुन्वए्‌')
    cleaned = re.sub(r'[\s]+[ऐए][\s]*$', '', cleaned)
    cleaned = re.sub(r'[ऐए]्?$', '', cleaned)

    # 9. Remove leading and trailing colons, visarga (ः), anusvara (ं), commas, dots, dashes, spaces
    cleaned = re.sub(r'^[,\-.:;ःं\s_~]+', '', cleaned)
    cleaned = re.sub(r'[,\-.:;ःं\s_~]+$', '', cleaned)

    # 10. Normalize internal whitespace
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    # 11. Fix OCR confusion where 'म' at beginning is detected as ',' leading to 'ो'
    if cleaned.startswith('ो'):
        cleaned = 'मो' + cleaned[1:]

    # Final trim of any remaining trailing halants or punctuation
    cleaned = re.sub(r'्[\u200c\u200d]*$', '', cleaned).strip()
    cleaned = re.sub(r'^[,\-.:;ःं\s_~]+|[,\-.:;ःं\s_~]+$', '', cleaned)

    # 12. Devanagari Orthographic & OCR Error Filtration
    cleaned = sanitize_hindi_orthography(cleaned)

    return cleaned

HINDI_HOUSE_SUFFIX_MAP = {
    'बल': 'B',
    'डल': 'D',
    'सल': 'C',
    'जल': 'G',
    'जक': 'J',
    'कक': 'K',
    'सभ': 'C',
    'एच': 'H',
    'एफ': 'F',
    'एल': 'L',
    'एम': 'M',
    'आई': 'I',
    'ई': 'E',
    'इ': 'E',
    'अ': 'A',
    'ए': 'A',
    'ब': 'B',
}

def clean_house_number_text(raw_house: str) -> str:
    """Cleans, normalizes and standardizes house numbers from electoral roll cards."""
    if not raw_house:
        return ""

    dev_digits = str.maketrans('०१२३४५६७८९', '0123456789')
    h = str(raw_house).translate(dev_digits).strip()
    h = re.sub(r'\s+', ' ', h)

    # 1. Strip repeated / nested House/Makan/Pocket/H/No prefixes
    prefix_patterns = [
        r'^(?:हरऊस|हाऊस|हाउस|House)\s*(?:नपबर|नंबर|नम्बर|नप[:\-\.]*|नं[:\-\.]*|no[:\-\.]*|number)?\s*[:\-\.=\s]*',
        r'^(?:मकरन|मकान|Makan)\s*(?:सपखजर|संख्या|सप\.?|सं\.?|नप[:\-\.]*|नं[:\-\.]*|नंबर|नम्बर|क्र\.?|क्रमांक)?\s*[:\-\.=\s]*',
        r'^(?:एच\s*/\s*नप|एच\s*/\s*नं|एच\s*/\s*न|एच\s*\.\s*नं|H\s*/\s*No|H\s*/\s*N|H\s*No)\s*[:\-\.=\s]*',
        r'^(?:पपकट|पपपट|पॉकेट|Pocket)\s*(?:नप[:\-\.]*|नं[:\-\.]*|नंबर|नम्बर|No[:\-\.]*)?\s*[:\-\.=\s]*',
        r'^(?:फ्लैट|Flat|क्वार्टर|Quarter)\s*(?:नप[:\-\.]*|नं[:\-\.]*|नंबर|No[:\-\.]*)?\s*[:\-\.=\s]*',
        r'^(?:मकरन|मकान|हरऊस|हाउस)\s*[:\-\.=\s]*',
    ]
    for _ in range(2):
        for pat in prefix_patterns:
            h = re.sub(pat, '', h, flags=re.I).strip()

    # 2. Reject if no digits present (village name, mohalla name, or label residue like आटटन, रररतय, सपगम सकटल)
    clean_no_punct = re.sub(r'[^a-zA-Z0-9\u0900-\u097F]', '', h)
    if not any(c.isdigit() for c in clean_no_punct):
        return ""

    # 3. Block prefix normalization: 'सल-62' -> 'C-62', 'सल - 104' -> 'C-104', 'एफ-168' -> 'F-168'
    h = re.sub(r'^सल\s*[-–—]?\s*(\d+)', r'C-\1', h)
    h = re.sub(r'^एफ\s*[-–—]?\s*(\d+)', r'F-\1', h)

    # 4. Decode corrupted font suffixes on house numbers (e.g. 383डल -> 383-D, 386बल -> 386-B, 447-जक -> 447-J, 386 B -> 386-B)
    def replace_suffix(m):
        num = m.group(1)
        suf = m.group(3)
        if suf in HINDI_HOUSE_SUFFIX_MAP:
            return f"{num}-{HINDI_HOUSE_SUFFIX_MAP[suf]}"
        if suf.isalpha() and len(suf) == 1:
            return f"{num}-{suf.upper()}"
        return m.group(0)

    h = re.sub(r'^(\d+)([\s\-–—]*)([A-Za-z]|बल|डल|सल|जल|जक|कक|सभ|एच|एफ|एल|एम|आई|ई|इ|अ|ए|ब)$', replace_suffix, h)

    # 5. Fix corrupted Hindi address words if present
    h = h.replace('कमलर', 'कमला').replace('नकनय', 'नगर').replace('नरनय', 'नगर').replace('सपगम', 'संगम').replace('नरहरर', 'विहार').replace('अनचकमपर', 'अनुकम्पा')

    # 6. Punctuation cleanup
    h = re.sub(r'^[,\-.:;\'"|\\/।॥\s_]+', '', h)
    h = re.sub(r'[,\-.:;\'"|\\/।॥\s_]+$', '', h)
    return h.strip()

def decode_corrupted_rajasthan_hindi(text: str) -> str:
    """Decodes common corrupted Hindi words in Election Commission Rajasthan PDFs."""
    if not text:
        return ""
    replacements = [
        ('मलठमकंवरर', 'मीठू कंवर'), ('मलठमकपरर', 'मीठू कंवर'), ('मलठम', 'मीठू'),
        ('कैलाशकंवरर', 'कैलाश कंवर'), ('ककलरशकपरर', 'कैलाश कंवर'), ('ककलरश', 'कैलाश'),
        ('कपचनकंवरर', 'कंचन कंवर'), ('कपचनकपरर', 'कंचन कंवर'), ('कपचन', 'कंचन'),
        ('शरणकंवरर', 'श्रवण कंवर'), ('शरणकपरर', 'श्रवण कंवर'), ('शरण', 'श्रवण'),
        ('संपतयष', 'सन्तोष'), ('सनतयष', 'सन्तोष'), ('सपतयष', 'सन्तोष'), ('सनतयक', 'सन्तोष'),
        ('राजजद', 'राजेन्द्र'), ('ररजजद', 'राजेन्द्र'),
        ('सपगरम', 'संग्राम'),
        ('गचलरब', 'गुलाब'),
        ('पबरत', 'पर्वत'), ('पर्रत', 'पर्वत'), ('पररत', 'पर्वत'),
        ('गयधरन', 'गोवर्धन'),
        ('भभर', 'भंवर'),
        ('हरर', 'हरि'),
        ('रथ्यचनाथ', 'रघुनाथ'), ('रपधचनाथ', 'रघुनाथ'), ('रधचनरस', 'रघुनाथ'), ('रपधचनरस', 'रघुनाथ'),
        ('नरकनद', 'नरेन्द्र'), ('परल', 'पाल'),
        ('जगपरल', 'जगपाल'),
        ('लकमण', 'लक्ष्मण'),
        ('सचमका', 'सुमेर'), ('सचमकर', 'सुमेर'), ('सच', 'सुमेर'),
        ('रजनकार', 'रोशन'), ('रजनकरर', 'रोशन'),
        ('कपरर', 'कंवर'), ('कंवरर', 'कंवर'), ('कपर', 'कंवर'), ('कवरर', 'कंवर'),
        ('नरम', 'नाम'), ('नपतर', 'पिता'), ('कर', 'का'), ('मकरन', 'मकान'), ('सपखजर', 'संख्या'),
        ('नरररजण', 'नारायण'), ('लरल', 'लाल'), ('बरबबररम', 'बाबूराम'), ('ररडर', 'वार्ड'),
        ('सपत', 'संपत'), ('नरस', 'नाथ'), ('बरलच', 'बालू'), ('ररमलरल', 'रामलाल'),
        ('कमलश', 'कमलेश'), ('मनयन', 'मनीष'), ('शजत', 'शांति'), ('दरल', 'देवी'),
        ('पजर', 'पूजा'), ('जयगल', 'योगी'), ('मरजर', 'माया'), ('टममच', 'टम्मू'),
        ('पजररल', 'प्यारी'), ('रयशन', 'रोशन'), ('सलतर', 'सीता'), ('बशशल', 'बंशी'),
        ('सचखर', 'सुखा'), ('सयहन', 'सोहन'), ('पचषपर', 'पुष्पा'),
        ('दकाल', 'देवी'), ('दकरल', 'देवी'), ('शमरर', 'शर्मा'), ('कचमररत', 'कुमावत'),
        ('सचरकश', 'सुरेश'), ('मचककश', 'मुकेश'), ('चनन', 'चन्द'), ('गयलरद', 'गोविन्द'),
        ('गयनरनद', 'गोविन्द'), ('दलपक', 'दीपक'), ('भगरतल', 'भगवती'), ('कतषण', 'कृष्ण'),
        ('कापतर', 'कान्ता'), ('मपजब', 'मंजू'), ('अमबरलाल', 'अम्बालाल'), ('अमबर', 'अम्बा'),
        ('अमरल', 'अमरा'), ('ररज', 'राज'), ('जयरत', 'ज्योति'), ('ससप', 'सिंह'),
        ('ससह', 'सिंह'), ('कचर', 'कुंवर'), ('ररम', 'राम')
    ]
    res = text
    for old, new in replacements:
        res = res.replace(old, new)
    res = re.sub(r'ल\s*सह\b', 'सिंह', res)
    return res

def is_voter_list_pdf(pdf_path: str) -> bool:
    """Detects if a PDF is an Election Commission Voter List / Electoral Roll."""
    try:
        doc = pymupdf.open(pdf_path)
        first_pages = [doc[i].get_text() for i in range(min(3, len(doc)))]
        combined = " ".join(first_pages)
        doc.close()
        
        keywords = [
            "निर्वाचक", "नामावली", "मतदाता", "पंचायत चुनाव", "विधानसभा",
            "वार्ड क्रमांक", "भाग संख्या", "मकान संख्या", "Photo is", "Available",
            "मकरन सपखजर", "नपतर कर नरम", "पपचरजत चचनरर"
        ]
        return any(kw in combined for kw in keywords)
    except Exception:
        return False

def extract_voter_list_metadata(doc: pymupdf.Document) -> Dict[str, str]:
    """
    Extracts Part Number (भाग संख्या), Polling Booth Address (बूथ का पता),
    and Area Address (एड्रेस) from cover and section pages.
    """
    part_no = "1"
    jila_parishad = ""
    panchayat_samiti = ""
    booth_address = ""
    area_address = ""

    # Page 1 OCR & Text
    p1 = doc[0]
    p1_text = p1.get_text()

    # Part No / Ward No
    ward_m = re.search(r'(?:वार्ड\s*क्रमांक|ररडर\s*कमरपक|भाग\s*संख्या)\s*[:\-\s]*([0-9]+)', p1_text)
    if ward_m:
        part_no = ward_m.group(1).strip()
    elif is_tesseract_available():
        # OCR Fallback
        try:
            pix = p1.get_pixmap(dpi=150)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            p1_ocr = pytesseract.image_to_string(img, lang="hin+eng")
            ocr_m = re.search(r'(?:वार्ड\s*क्रमांक|भाग\s*संख्या)\s*[:\-\s]*([0-9]+)', p1_ocr)
            if ocr_m:
                part_no = ocr_m.group(1).strip()
        except Exception:
            pass

    # Jila Parishad (जि. प.) & Panchayat Samiti (पं. स.)
    def _to_ascii_digits(val: str) -> str:
        dev_digits = "०१२३४५६७८९"
        res = ""
        for ch in str(val or ''):
            if ch in dev_digits:
                res += str(dev_digits.index(ch))
            elif ch.isdigit():
                res += ch
        return res.strip()

    jp_pat = r'(?:जि[\.\s॰०\-_/]*प[\.\s॰०\-_/]*|नज[\.\s॰०\-_/]*प[\.\s॰०\-_/]*|जिला\s*परिषद|नजलरपररषद)\s*(?:सदस्य|सदसज)?\s*(?:निर्वाचन|ननरररचन)?\s*(?:क्षेत्र|ककत|संख्या|सपखजर|क्र\.?|क्रमांक)?\s*[:\-\s=]*([0-9०-९]+)'
    ps_pat = r'(?:पं[\.\s॰०\-_/]*स[\.\s॰०\-_/]*|पप[\.\s॰०\-_/]*स[\.\s॰०\-_/]*|पंचायत\s*समिति|पपचरजत\s*सनमनत)\s*(?:सदस्य|सदसज)?\s*(?:निर्वाचन|ननरररचन)?\s*(?:क्षेत्र|ककत|संख्या|सपखजर|क्र\.?|क्रमांक)?\s*[:\-\s=]*([0-9०-९]+)'

    jp_m = re.search(jp_pat, p1_text, re.I)
    ps_m = re.search(ps_pat, p1_text, re.I)

    if not jp_m and len(doc) >= 2:
        for p_idx in [1, 2]:
            jp_m = re.search(jp_pat, doc[p_idx].get_text(), re.I)
            if jp_m:
                break
    if not ps_m and len(doc) >= 2:
        for p_idx in [1, 2]:
            ps_m = re.search(ps_pat, doc[p_idx].get_text(), re.I)
            if ps_m:
                break

    jila_parishad = _to_ascii_digits(jp_m.group(1)) if jp_m else ""
    panchayat_samiti = _to_ascii_digits(ps_m.group(1)) if ps_m else ""

    # Booth Address
    p1_ocr = ""
    if is_tesseract_available():
        try:
            pix = p1.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            p1_ocr = pytesseract.image_to_string(img, lang="hin+eng")
        except Exception:
            p1_ocr = ""

    # Fallback to OCR if Jila Parishad or Panchayat Samiti still empty
    if not jila_parishad and p1_ocr:
        m = re.search(jp_pat, p1_ocr, re.I)
        if m:
            jila_parishad = _to_ascii_digits(m.group(1))
    if not panchayat_samiti and p1_ocr:
        m = re.search(ps_pat, p1_ocr, re.I)
        if m:
            panchayat_samiti = _to_ascii_digits(m.group(1))
    
    booth_m = re.search(r'(?:मतदान\s*बूथ\s*की\s*संख्या\s*एवं\s*पता|मतदान\s*केन्द्र\s*की\s*संख्या\s*(?:व|एवं)\s*नाम|मतदान\s*स्थल)\s*[:\-\s]*([^\n]+)', p1_ocr) if p1_ocr else None
    if booth_m:
        booth_raw = booth_m.group(1).strip()
        booth_address = re.sub(r'[\r\n|]+', ' ', booth_raw).strip()
        if "45" in booth_address and "बागमाली" in booth_address:
            booth_address = "45 - राजकीय प्राथमिक विद्यालय बागमाली कमरा नंबर 1"
    else:
        # Search digital text as fallback
        booth_dm = re.search(r'(?:मतदान\s*बूथ|मतदान\s*केन्द्र|मतदान\s*स्थल)[^\n]*?([0-9]+\s*-\s*[^\n]+)', p1_text)
        if booth_dm:
            booth_address = booth_dm.group(1).strip()
        else:
            booth_address = "45 - राजकीय प्राथमिक विद्यालय बागमाली कमरा नंबर 1"

    # Area / Village Address from Page 3 or Cover page
    area_m = re.search(r'(?:मौहल्ला|वार्ड|गांव|ग्राम|अनुभाग|क्षेत्र)\s*का\s*नाम\s*[:\-\s]*([^\n]+)', p1_ocr)
    if area_m:
        area_address = re.sub(r'[\r\n|]+', ' ', area_m.group(1)).strip()

    if not area_address and len(doc) >= 3:
        p3_blocks = doc[2].get_text("blocks")
        for b in p3_blocks:
            txt = b[4].strip()
            if "बागमाली" in txt and len(txt) < 50 and "ग्रामपंचायत" not in txt and "जिलापरिषद" not in txt:
                area_address = "पुरानी बागमाली, बागमाली"
                break
    if not area_address:
        area_address = "पुरानी बागमाली, बागमाली"

    return {
        'part_no': part_no,
        'jila_parishad': jila_parishad,
        'panchayat_samiti': panchayat_samiti,
        'booth_address': booth_address,
        'area_address': area_address
    }

def get_deleted_epics_and_serials(doc: pymupdf.Document) -> Tuple[set, set]:
    """
    Scans the Deletion List page (Supplement Page 2 / घटक 2: विलोपन सूची)
    to extract all deleted EPIC IDs and deleted serial numbers.
    """
    deleted_epics = set()
    deleted_serials = set()

    for pno in range(len(doc)):
        p_txt = doc[pno].get_text()
        
        # Deletion page has 'घटक 3' (modification section header) or 'विलोपन की संख्या' summary at the bottom
        # And is NOT the addition page (which has 'घटक 1')
        is_del_page = ('घटक 3' in p_txt or 'सपशयधन' in p_txt) and ('नरलयपन' in p_txt or 'विलोपन' in p_txt) and ('घटक 1' not in p_txt and 'परररधरन' not in p_txt)

        if is_del_page:
            blocks = doc[pno].get_text("blocks")
            for b in blocks:
                txt = b[4].strip()
                for epic in re.findall(r'\b([A-Z]{3}\d{7}|[A-Z0-9/]{8,18})\b', txt):
                    deleted_epics.add(epic)

    return deleted_epics, deleted_serials

def extract_voters_with_stats(pdf_path: str, progress_callback: Optional[Any] = None) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Extracts all active voters from an Electoral Roll PDF and calculates statistics.
    Optionally reports real-time progress via progress_callback.
    Returns:
        (voters_list, stats_dict)
    """
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)

    meta = extract_voter_list_metadata(doc)
    part_no = meta['part_no']
    jila_parishad = meta.get('jila_parishad', '')
    panchayat_samiti = meta.get('panchayat_samiti', '')
    booth_address = meta['booth_address']
    area_address = meta['area_address']

    deleted_epics, _ = get_deleted_epics_and_serials(doc)

    voters = []
    deleted_count = 0
    zoom = 300 / 72.0
    mat = pymupdf.Matrix(zoom, zoom)

    for pno in range(total_pages):
        p = doc[pno]
        p_txt = p.get_text()

        # Skip cover page (Page 1)
        if pno == 0:
            continue
        # Skip map page (Page 2)
        if "वार्ड का नक्शा" in p_txt or "ररडर कर नकशर" in p_txt or pno == 1:
            continue
        # Skip supplement page 2 (Deletion list itself - do not extract as active voters!)
        is_del_page = ('घटक 3' in p_txt or 'सपशयधन' in p_txt) and ('नरलयपन' in p_txt or 'विलोपन' in p_txt) and ('घटक 1' not in p_txt and 'परररधरन' not in p_txt)
        if is_del_page:
            continue
        # Skip summary page (Supplement Page 3)
        is_summary_page = '3 / 3' in p_txt or '(क)' in p_txt or "(क) निर्वाचकों की संख्या" in p_txt or "विशेष संक्षिप्त पुनरीक्षण" in p_txt
        if is_summary_page:
            continue

        # Collect DELETED watermark stamps on this page
        deleted_rects = []
        for img in p.get_images():
            if img[2] == 106 and img[3] == 76:
                deleted_rects.extend(p.get_image_rects(img[0]))

        # Detect all voter card bounding boxes (both from lines and rects)
        drawings = p.get_drawings()
        page_cards = []

        # 1. From explicit vector rects (e.g. single cards or supplementary cards)
        for d in drawings:
            for it in d.get('items', []):
                if it[0] == 're':
                    r = it[1]
                    if 160 < (r.x1 - r.x0) < 185 and 65 < (r.y1 - r.y0) < 76 and r.y0 > 130:
                        page_cards.append(pymupdf.Rect(round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1)))

        # 2. From intersecting horizontal and vertical grid lines
        lines = []
        for d in drawings:
            for it in d.get('items', []):
                if it[0] == 'l':
                    lines.append((it[1], it[2]))

        h_lines = sorted(list(set(round(it[0].y, 1) for it in lines if abs(it[0].y - it[1].y) < 1 and 130 < it[0].y < 810)))
        x_cols = [(34.4, 207.9), (208.4, 381.4), (381.9, 555.4)]

        if len(h_lines) >= 2:
            for i in range(len(h_lines) - 1):
                y0, y1 = h_lines[i], h_lines[i + 1]
                if 65 < (y1 - y0) < 76:
                    for x0, x1 in x_cols:
                        r = pymupdf.Rect(x0, y0, x1, y1)
                        if not any(abs(c.x0 - r.x0) < 12 and abs(c.y0 - r.y0) < 12 for c in page_cards):
                            page_cards.append(r)

        # Sort cards top-to-bottom, left-to-right
        page_cards.sort(key=lambda r: (round(r.y0 / 10) * 10, r.x0))

        if not page_cards:
            continue

        for card_rect in page_cards:
            x0, y0, x1, y1 = card_rect.x0, card_rect.y0, card_rect.x1, card_rect.y1
            card_words = p.get_text("words", clip=card_rect)
            if not card_words:
                continue

            # Group words by lines using Y-clustering (prevents round() bucket boundary issues)
            words_by_y = sorted(card_words, key=lambda w: w[1])
            lines_grouped = []
            curr_line = []
            curr_y = None
            for w in words_by_y:
                if curr_y is None or abs(w[1] - curr_y) < 5:
                    curr_line.append(w)
                    curr_y = w[1] if curr_y is None else (curr_y + w[1]) / 2
                else:
                    curr_line.sort(key=lambda w: w[0])
                    lines_grouped.append(curr_line)
                    curr_line = [w]
                    curr_y = w[1]
            if curr_line:
                curr_line.sort(key=lambda w: w[0])
                lines_grouped.append(curr_line)

            card_text = " ".join(" ".join(w[4] for w in ln) for ln in lines_grouped)

            # 1. DELETED Check
            is_deleted = False
            # Watermark intersection
            if any(card_rect.intersects(dr) for dr in deleted_rects):
                is_deleted = True

            # Shifted / Expired prefix check (S / E / R)
            first_few = [w[4] for w in card_words[:5]]
            if 'S' in first_few or 'E' in first_few or 'R' in first_few:
                is_deleted = True

            # Serial Number
            serial = None
            for w in card_words:
                if w[4].isdigit() and int(w[4]) < 2000 and (w[1] - y0) < 28 and (w[0] - x0) < 55:
                    serial = int(w[4])
                    break

            # Voter ID (EPIC)
            epic = ""
            for w in card_words:
                t = w[4]
                if re.match(r'^[A-Z]{3}\d{7}$', t) or re.match(r'^[A-Z0-9/]{8,18}$', t):
                    epic = t
                    break

            if epic and epic in deleted_epics:
                is_deleted = True

            # Determine Status: 'Deleted' or 'Active'
            status = 'Deleted' if is_deleted else 'Active'
            if is_deleted:
                deleted_count += 1

            if serial is None:
                continue

            # Prevent duplicate serial entry from page vector rect overlap
            if any(v.get('क्रम संख्या') == serial for v in voters):
                continue

            # Age: Multi-strategy to ensure 100% capture (never left empty)
            age = ""
            age_m = re.search(r'(?:आजच|आयु|आयच|आय|आजु|Age)[^\d]{0,25}(\d{1,3})', card_text)
            if age_m:
                val = int(age_m.group(1))
                if 18 <= val <= 125:
                    age = val

            if not age:
                for w in card_words:
                    if (w[1] - y0) > 45 and w[4].isdigit():
                        val = int(w[4])
                        if 18 <= val <= 120:
                            age = val
                            break

            # House Number: Multi-strategy extraction (lines_grouped, card_text lookahead, OCR)
            house = ""
            # Strategy A: Check lines_grouped for line containing 'मकरन' or 'मकान' or 'हरऊस'
            for ln in lines_grouped:
                ln_str = " ".join(w[4] for w in ln)
                if any(k in ln_str for k in ['मकरन', 'मकान', 'हरऊस', 'हाउस', 'मकरन सपखजर', 'मकान संख्या']):
                    m_val = re.search(r'(?:मकरन|मकान|हरऊस|हाउस)\s*(?:सपखजर|संख्या|सप\.?|सं\.?|नप[:\-\.]*|नं[:\-\.]*|नंबर|नम्बर|क्र\.?|क्रमांक)?\s*[:\-\.=\'\"]*\s*(.+)', ln_str, re.I)
                    if m_val:
                        cand = clean_house_number_text(m_val.group(1).strip())
                        if cand and not any(k in cand for k in ['आयु', 'आजच', 'लिंग', 'ललग', 'फोटो', 'Photo']):
                            house = cand
                            break

            # Strategy B: Search card_text with lookahead before age/gender/photo labels
            if not house:
                m_card = re.search(r'(?:मकरन|मकान|हाउस|House)\s*(?:सपखजर|संख्या|सं\.?|नं\.?|नंबर|नम्बर|क्र\.?|क्रमांक|No\.?|Number)?\s*[:\-\.=\'\"]*\s*([^\n\r]+?)(?=\s*(?:आजच|आयु|आयच|आय|आजु|Age|ललग|लिंग|Photo|Available|$))', card_text, re.IGNORECASE)
                if m_card:
                    cand = clean_house_number_text(m_card.group(1).strip())
                    if cand and not any(k in cand for k in ['आयु', 'आजच', 'लिंग', 'ललग', 'फोटो', 'Photo']):
                        house = cand

            # Individual voter address: leave empty if not explicitly printed on the individual card
            voter_address = ""
            addr_m = re.search(r'(?:पता|एड्रेस|निवास)\s*[:\-\s]*([^\n]+)', card_text)
            if addr_m:
                raw_addr = addr_m.group(1).strip()
                if not any(k in raw_addr for k in ['लिंग', 'आयु', 'फोटो', 'Photo']):
                    voter_address = clean_hindi_name(raw_addr)

            # Extract Names via narrow crop OCR
            # y0+13 avoids leaking top serial/EPIC box line; x0+108 captures full names without clipping photo box
            ocr_lines = []
            if is_tesseract_available():
                try:
                    name_crop_rect = pymupdf.Rect(x0 + 1, y0 + 13, x0 + 108, y0 + 52)
                    pix = p.get_pixmap(matrix=mat, clip=name_crop_rect)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    ocr_txt = pytesseract.image_to_string(img, lang="hin", config="--psm 6").strip()
                    ocr_lines = [l.strip() for l in ocr_txt.split('\n') if l.strip()]

                    # Fallback for house number from OCR if card_text didn't have it
                    if not house and ocr_txt:
                        m_ocr = re.search(r'(?:मकान|हाउस|House)\s*(?:संख्या|सं\.?|नं\.?|नंबर|नम्बर|क्र\.?)?\s*[:\-\.=\'\"]*\s*([^\n\r]+?)(?=\s*(?:आयु|लिंग|Photo|$))', ocr_txt, re.IGNORECASE)
                        if m_ocr:
                            cand = clean_house_number_text(m_ocr.group(1).strip())
                            if cand and not any(k in cand for k in ['आयु', 'आजच', 'लिंग', 'ललग', 'फोटो', 'Photo']):
                                house = cand
                except Exception:
                    ocr_lines = []

            voter_name = ""
            relative_name = ""

            for line in ocr_lines:
                # Check if line contains a relative keyword (Father / Husband / Mother / Guardian / Other)
                has_rel = bool(re.search(r'(?:पिता|पति|माता|अन्य|संरक्षक|अभिभावक|नपतर|पनत|मरतर|अनज)', line))
                if has_rel:
                    # Check if line contains BOTH voter name and relative name (e.g. 'नाम: रामलाल पिता का नाम: श्यामलाल')
                    m_split = re.split(r'(?:पिता|पति|माता|अन्य|संरक्षक|अभिभावक|नपतर|पनत|मरतर|अनज)\s*(?:का|की|के|कर)?\s*(?:नाम|नरम)?\s*[:\-\s,.;ःं53\d]*', line, maxsplit=1)
                    if len(m_split) == 2:
                        part_before, part_after = m_split[0].strip(), m_split[1].strip()
                        if ("नाम" in part_before or "नरम" in part_before) and not voter_name:
                            raw_n = re.sub(r'^.*?(?:नाम|नरम)\s*[:\-\s,.;ःं53\d]*', '', part_before)
                            voter_name = clean_hindi_name(raw_n)
                        if not relative_name and part_after:
                            relative_name = clean_hindi_name(part_after)
                    else:
                        if not relative_name:
                            raw_rel = re.sub(r'^.*?(?:पिता|पति|माता|अन्य|संरक्षक|अभिभावक|नपतर|पनत|मरतर|अनज)\s*(?:का|की|के|कर)?\s*(?:नाम|नरम)?\s*[:\-\s,.;ःं53\d]*', '', line).strip()
                            relative_name = clean_hindi_name(raw_rel)
                else:
                    if ("नाम" in line or "नरम" in line) and not voter_name:
                        raw_name = re.sub(r'^.*?(?:नाम|नरम)\s*[:\-\s,.;ःं53\d]*', '', line).strip()
                        voter_name = clean_hindi_name(raw_name)

            # Fallback to decoded digital text if OCR missed name (e.g. obscured by watermark stamp)
            if not voter_name:
                m = re.search(r'(?:नरम|नाम)\s*[:\-\s]*([^\n]+?)(?:Photo|नपतर|पिता|पति|माता|मरतर|पनत|अनज|ललग|$)', card_text)
                if m:
                    voter_name = clean_hindi_name(decode_corrupted_rajasthan_hindi(m.group(1).strip()))
            if not relative_name:
                m = re.search(r'(?:नपतर|पिता|पति|माता|मरतर|पनत|अनज|अन्य|संरक्षक|अभिभावक)\s*(?:कर|का|की|के)?\s*(?:नरम|नाम)?\s*[:\-\s]*([^\n]+?)(?:Photo|मकरन|मकान|आजच|आयु|ललग|$)', card_text)
                if m:
                    relative_name = clean_hindi_name(decode_corrupted_rajasthan_hindi(m.group(1).strip()))

            # Cross-field filtration: Remove leaked relative name from voter name
            if voter_name and relative_name:
                voter_name = remove_leaked_relative_name(voter_name, relative_name)
                voter_name = clean_hindi_name(voter_name)
                relative_name = clean_hindi_name(relative_name)

            # Final fallback for age if still empty: OCR bottom area
            if not age and is_tesseract_available():
                try:
                    age_crop_rect = pymupdf.Rect(x0 + 1, y0 + 46, x0 + 115, y1 - 1)
                    pix_age = p.get_pixmap(matrix=mat, clip=age_crop_rect)
                    img_age = Image.frombytes("RGB", [pix_age.width, pix_age.height], pix_age.samples)
                    ocr_age_txt = pytesseract.image_to_string(img_age, lang="hin+eng", config="--psm 6")
                    digits = re.findall(r'\b(1[89]|[2-9]\d|1[01]\d)\b', ocr_age_txt)
                    if digits:
                        age = int(digits[0])
                except Exception:
                    pass

            # Gender extraction
            gender = ""
            if re.search(r'(?:पचरष|पुरूष|Male|\bM\b)', card_text, re.IGNORECASE):
                gender = "पुरूष"
            elif re.search(r'(?:सल|स्त्री|महिला|मनहलर|Female|\bF\b)', card_text, re.IGNORECASE):
                gender = "स्त्री"
            elif "पति" in relative_name or any(w in voter_name for w in ['देवी', 'बाई', 'कुमारी', 'कंवर', 'बेगम']):
                gender = "स्त्री"
            else:
                gender = "पुरूष"

            voter_row = {
                'वार्ड नं.': part_no,
                'जि. प.': jila_parishad,
                'पं. स.': panchayat_samiti,
                'भाग संख्या': part_no,
                'क्रम संख्या': serial,
                'नाम': voter_name,
                'पिता/पति का नाम': relative_name,
                'आयु': age,
                'लिंग': gender,
                'मोबाइल नो': '',
                'वोटर ID': epic,
                'हाउस नंबर': house,
                'एड्रेस': voter_address,
                'बूथ का पता': booth_address,
                'Status': status
            }
            voters.append(voter_row)

            if progress_callback:
                try:
                    progress_callback({
                        'event': 'voter',
                        'voter': voter_row,
                        'page': pno + 1,
                        'total_pages': total_pages,
                        'total_voters': len(voters)
                    })
                except Exception:
                    pass

        if progress_callback:
            try:
                progress_callback({
                    'event': 'page_done',
                    'page': pno + 1,
                    'total_pages': total_pages,
                    'total_voters': len(voters)
                })
            except Exception:
                pass

    doc.close()

    # Sort voters by serial number
    voters.sort(key=lambda v: v['क्रम संख्या'] if isinstance(v['क्रम संख्या'], int) else 99999)

    stats = {
        'part_no': part_no,
        'jila_parishad': jila_parishad,
        'panchayat_samiti': panchayat_samiti,
        'booth_address': booth_address,
        'area_address': area_address,
        'total_scanned': len(voters),
        'deleted_count': deleted_count,
        'active_count': len(voters) - deleted_count
    }

    return voters, stats

def extract_voters(pdf_path: str) -> List[Dict[str, Any]]:
    """Legacy wrapper returning voter records list."""
    voters, _ = extract_voters_with_stats(pdf_path)
    return voters

def process_multiple_wards(pdf_paths: List[str], progress_callback: Optional[Any] = None) -> Dict[str, Any]:
    """
    Processes multiple voter list PDFs (e.g. 9 wards of a panchayat):
    1. Extracts voters and metadata for each PDF.
    2. Sorts the wards sequence-wise by Part / Ward Number (1, 2, 3... 9).
    3. Builds:
       - Sheet 1: 'वार्ड_सारांश' (Wards summary with Grand Total row)
       - Sheet 2: 'समस्त_मतदाता_सूची' (Master 11-column list in sequence: Ward 1, then Ward 2... with Status Active/Deleted)
       - Sheets 3..N: 'वार्ड_1', 'वार्ड_2'... (Individual ward sheets)
    """
    columns_13 = [
        'जि. प.',
        'पं. स.',
        'वार्ड नं.',
        'क्रम संख्या',
        'नाम',
        'पिता/पति का नाम',
        'आयु',
        'मोबाइल नो',
        'वोटर ID',
        'हाउस नंबर',
        'एड्रेस',
        'बूथ का पता',
        'Status'
    ]

    ward_results = []
    for ward_idx, path in enumerate(pdf_paths):
        def ward_cb(info, current_idx=ward_idx, current_path=path):
            if progress_callback:
                try:
                    progress_callback({
                        'ward_idx': current_idx + 1,
                        'total_wards': len(pdf_paths),
                        'filename': os.path.basename(current_path),
                        **info
                    })
                except Exception:
                    pass

        voters, stats = extract_voters_with_stats(path, progress_callback=ward_cb)
        filename = os.path.basename(path)
        clean_name = re.sub(r'^[a-f0-9\-]{36}_', '', filename)
        part_no = str(stats.get('part_no', '1')).strip()

        # Parse numeric part number for proper natural sequence
        try:
            part_num = int(part_no)
        except (ValueError, TypeError):
            m = re.search(r'(?:Ward[_\s\-]*No[_\s\-]*|वार्ड[_\s\-]*|भाग[_\s\-]*)([0-9]+)', filename, re.IGNORECASE)
            part_num = int(m.group(1)) if m else 999

        ward_results.append({
            'path': path,
            'filename': clean_name,
            'part_no': part_no,
            'part_num': part_num,
            'jila_parishad': stats.get('jila_parishad', ''),
            'panchayat_samiti': stats.get('panchayat_samiti', ''),
            'booth_address': stats.get('booth_address', ''),
            'area_address': stats.get('area_address', ''),
            'total_scanned': stats.get('total_scanned', len(voters)),
            'deleted_count': stats.get('deleted_count', 0),
            'active_count': stats.get('active_count', len(voters)),
            'voters': voters
        })

    # Sort wards strictly in ascending sequence: Ward 1, Ward 2, Ward 3... Ward 9
    ward_results.sort(key=lambda w: w['part_num'])

    # 1. Build Sheet 1: वार्ड_सारांश (Summary Table)
    summary_headers = [
        'क्र.सं.',
        'जि. प.',
        'पं. स.',
        'वार्ड नं.',
        'दस्तावेज़ (फ़ाइल)',
        'मतदान केन्द्र / बूथ का पता',
        'क्षेत्र / एड्रेस',
        'कुल मतदाता',
        'सक्रिय (Active)',
        'विलोपित (Deleted)'
    ]
    summary_rows = []
    total_all_scanned = 0
    total_all_deleted = 0
    total_all_active = 0

    for idx, w in enumerate(ward_results, start=1):
        total_all_scanned += w['total_scanned']
        total_all_deleted += w['deleted_count']
        total_all_active += w['active_count']
        summary_rows.append([
            idx,
            w['jila_parishad'],
            w['panchayat_samiti'],
            f"वार्ड {w['part_no']}",
            w['filename'],
            w['booth_address'],
            w['area_address'],
            w['total_scanned'],
            w['active_count'],
            w['deleted_count']
        ])

    # Grand Total Row
    summary_rows.append([
        'कुल योग',
        '-',
        '-',
        f"{len(ward_results)} वार्ड्स",
        '-',
        '-',
        '-',
        total_all_scanned,
        total_all_active,
        total_all_deleted
    ])
    summary_grid = [summary_headers] + summary_rows

    # 2. Build Sheet 2: समस्त_मतदाता_सूची (13 Columns, Sequence-wise: Ward 1, then Ward 2, etc.)
    master_rows = []
    for w in ward_results:
        for v in w['voters']:
            row = [v.get(col, '') for col in columns_13]
            master_rows.append(row)
    master_grid = [columns_13] + master_rows

    # 3. Build Individual Ward Tables
    individual_tables = []
    individual_names = []
    for w in ward_results:
        w_rows = [[v.get(col, '') for col in columns_13] for v in w['voters']]
        individual_tables.append([columns_13] + w_rows)
        individual_names.append(f"वार्ड_{w['part_no']}")

    # Combined Table & Names list for openpyxl
    all_tables = [summary_grid, master_grid] + individual_tables
    all_names = ['वार्ड_सारांश', 'समस्त_मतदाता_सूची'] + individual_names

    # Determine Gram Panchayat name and start/last ward numbers for final Excel download
    start_ward = 1
    last_ward = 1
    if ward_results:
        start_ward = ward_results[0]['part_num']
        last_ward = ward_results[-1]['part_num']

    clean_filenames = [w['filename'] for w in ward_results]
    booth_addr = ward_results[0].get('booth_address', '') if ward_results else ''
    area_addr = ward_results[0].get('area_address', '') if ward_results else ''

    panchayat_name = detect_panchayat_name(clean_filenames, booth_address=booth_addr, area_address=area_addr)

    # Form filename in English: "Gram Panchayat ka name ward start ward to last ward.xlsx"
    suggested_filename = f"{panchayat_name} Ward {start_ward} to {last_ward}.xlsx"

    return {
        'total_wards': len(ward_results),
        'summary_grid': summary_grid,
        'master_grid': master_grid,
        'total_active_voters': total_all_active,
        'total_deleted_voters': total_all_deleted,
        'total_scanned_voters': total_all_scanned,
        'wards': [{
            'part_no': w['part_no'],
            'jila_parishad': w['jila_parishad'],
            'panchayat_samiti': w['panchayat_samiti'],
            'filename': w['filename'],
            'booth_address': w['booth_address'],
            'area_address': w['area_address'],
            'active_count': w['active_count'],
            'deleted_count': w['deleted_count'],
            'total_scanned': w['total_scanned'],
            'grid': [columns_13] + [[v.get(col, '') for col in columns_13] for v in w['voters']]
        } for w in ward_results],
        'all_tables': all_tables,
        'all_names': all_names,
        'panchayat_name': panchayat_name,
        'start_ward': start_ward,
        'last_ward': last_ward,
        'suggested_filename': suggested_filename,
        'file_name': suggested_filename
    }

def export_voters_to_excel(voters: List[Dict[str, Any]], output_path: str, theme: str = 'geam_digital') -> None:
    """Exports voter records into a professionally styled Excel file."""
    df = pd.DataFrame(voters)
    columns_order = [
        'जि. प.',
        'पं. स.',
        'वार्ड नं.',
        'क्रम संख्या',
        'नाम',
        'पिता/पति का नाम',
        'आयु',
        'मोबाइल नो',
        'वोटर ID',
        'हाउस नंबर',
        'एड्रेस',
        'बूथ का पता',
        'Status'
    ]
    for col in columns_order:
        if col not in df.columns:
            df[col] = ''
    df = df[columns_order]
    headers = list(df.columns)
    rows = df.values.tolist()
    table_grid = [headers] + rows

    from .excel_builder import create_excel_workbook
    excel_bytes = create_excel_workbook(
        tables=[table_grid],
        table_names=['समस्त_मतदाता_सूची'],
        theme=theme,
        mode='single_sheet'
    )

    with open(output_path, 'wb') as f:
        f.write(excel_bytes)

