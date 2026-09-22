
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import math, re
import openpyxl


def num(x, default=0.0):
    try:
        if x is None or x == "":
            return default
        return float(str(x).replace(",", "."))
    except Exception:
        return default


def round_half(x):
    return round(x * 2) / 2


def roundup(x, n=1):
    return math.ceil(x / n) * n


def mround(x, multiple):
    if multiple == 0:
        return 0
    return round(x / multiple) * multiple


def mround_excel(x, multiple):
    """Excel-compatible MROUND for the positive values used here."""
    if multiple == 0:
        return 0.0
    return math.floor((float(x) / multiple) + 0.5) * multiple


def parse_size(s):
    s = str(s or "").lower().replace("×", "x").replace("х", "x").replace("*", "").replace(" ", "")
    nums = re.findall(r"\d+(?:[.,]\d+)?", s)
    if len(nums) < 2:
        raise ValueError(f"Невалиден формат: {s}")
    return float(nums[0].replace(",", ".")), float(nums[1].replace(",", "."))


@dataclass(frozen=True)
class FormatOption:
    source: str
    print_format: str
    sheets_per_source: int


class FormatRepository:
    def __init__(self, workbook):
        self.workbook = Path(workbook)
        wb = openpyxl.load_workbook(self.workbook, data_only=False, keep_vba=True)
        ws = wb["формати"]
        self.options = []
        source = None
        for r in range(2, ws.max_row + 1):
            a, b, c = ws.cell(r,1).value, ws.cell(r,2).value, ws.cell(r,3).value
            if a not in (None, ""):
                source = str(a).strip()
            if b not in (None, "") and c not in (None, ""):
                try:
                    n = int(float(c))
                except Exception:
                    continue
                self.options.append(FormatOption(source, str(b).strip(), n))
        # The workbook is only a data source here; do not leave its ZIP package
        # alive after the formats have been copied into Python objects.
        wb.close()

    def sources(self):
        return list(dict.fromkeys(o.source for o in self.options))

    def options_for(self, source):
        norm = lambda x: str(x).strip().lower().replace('х','x').replace('×','x')
        opts = [o for o in self.options if norm(o.source) == norm(source)]
        # User-confirmed Himiya rule: 43×61 -> 43×30.5 -> 30.5×21.5.
        # The supplied workbook contains only the latter row, so add the first
        # print stage explicitly rather than losing it in the Python model.
        if str(source).strip().lower() == 'himiya':
            if not any(o.print_format.replace('×','x').lower() == '43x30.5' for o in opts):
                opts = [FormatOption('Himiya','43x30.5',2)] + opts
        return opts

    def all_options(self):
        return list(self.options)


def grip_mm(useful: bool) -> float:
    return 3.0 if useful else 10.0


def get_layout(sheet_w, sheet_h, prod_w, prod_h, grip):
    # The gripper occupies the sheet's feed-edge dimension (height here).
    # This is decisive for A4 on Himiya: 305-3 >= 297, while 305-10 < 297.
    usable_h = sheet_h - grip
    cols = math.floor(sheet_w / prod_w) if prod_w else 0
    rows = math.floor(usable_h / prod_h) if prod_h else 0
    if cols <= 0 or rows <= 0:
        return 0, 0
    return cols * rows, (cols, rows)



def solve_guillotine(sheet_w_mm, sheet_h_mm, prod_w_mm, prod_h_mm):
    """Recursive guillotine optimizer copied from the supplied macro-based Python app."""
    memo = {}

    def helper(w, h):
        if w < min(prod_w_mm, prod_h_mm) or h < min(prod_w_mm, prod_h_mm):
            return 0

        key = (round(w, 1), round(h, 1))
        if key in memo:
            return memo[key]

        best_cnt = 0

        # Single item in either orientation.
        if w >= prod_w_mm and h >= prod_h_mm:
            best_cnt = 1
        if w >= prod_h_mm and h >= prod_w_mm:
            best_cnt = max(best_cnt, 1)

        # Vertical cuts.
        cut_x = set()
        x = prod_w_mm
        while x < w:
            cut_x.add(x)
            x += prod_w_mm
        x = prod_h_mm
        while x < w:
            cut_x.add(x)
            x += prod_h_mm

        for cx in cut_x:
            best_cnt = max(best_cnt, helper(cx, h) + helper(w - cx, h))

        # Horizontal cuts.
        cut_y = set()
        y = prod_h_mm
        while y < h:
            cut_y.add(y)
            y += prod_h_mm
        y = prod_w_mm
        while y < h:
            cut_y.add(y)
            y += prod_w_mm

        for cy in cut_y:
            best_cnt = max(best_cnt, helper(w, cy) + helper(w, h - cy))

        memo[key] = best_cnt
        return best_cnt

    return helper(sheet_w_mm, sheet_h_mm)


def smart_calculate(print_format, prod_w_mm, prod_h_mm, turnover="не", useful_grip=False):
    sw_cm, sh_cm = parse_size(print_format)
    psw_mm, psh_mm = sw_cm * 10, sh_cm * 10

    # Exact workfield from the supplied v4.5 calculator.
    grip = 3 if useful_grip else 10
    lim_w = psw_mm - 5
    lim_h = psh_mm - (grip + 3)

    if str(turnover).strip().lower() == "да":
        half_w = lim_w / 2
        cnt = solve_guillotine(half_w, lim_h, prod_w_mm, prod_h_mm)
        return cnt * 2

    return solve_guillotine(lim_w, lim_h, prod_w_mm, prod_h_mm)


def choose_best(repo, source, prod_w, prod_h, turnover, useful_grip):
    opts = repo.options_for(source)
    scored = []
    for o in opts:
        reps = smart_calculate(o.print_format, prod_w, prod_h, turnover, useful_grip)
        if reps <= 0:
            continue
        # Primary objective: minimum source sheets, secondary: higher repetitions.
        source_sheets = math.inf
        # Clean sheets required is computed later; here use the option's yield.
        score = (-o.sheets_per_source / reps, reps)
        scored.append((score, o, reps))
    if not scored:
        return None
    # Prefer the highest number of source-sheet-effective reproductions.
    scored.sort(key=lambda t: (t[0][0], t[0][1]), reverse=True)
    return scored[0][1], scored[0][2]


def find_best_sheet_logic(repo, source, prod_w, prod_h, turnover, useful_grip):
    """Excel-style FindBestSheet: return MIN and MAX (Optim.) choices.

    MIN: physically smallest available print format whose waste is below 60%.
    MAX (Optim.): available format with the lowest waste; ties prefer the
    smaller sheet. Both results include the calculated repetitions per sheet.
    """
    # Excel's FindBestSheet does not restrict the recommendation to the
    # currently selected source sheet.  It searches the available print
    # formats from the Formats table and ignores the special large "плаки"
    # entry.  This is why, for example, a 64x90 source can recommend
    # 23x17.5 or 45x20.
    candidates = []
    for opt in repo.all_options():
        if str(opt.source).strip().lower() == 'плаки':
            continue
        reps = smart_calculate(opt.print_format, prod_w, prod_h, turnover, useful_grip)
        if reps <= 0:
            continue
        w, h = parse_size(opt.print_format)
        sheet_area = w * h
        product_area = prod_w * prod_h
        waste_pct = max(0.0, (1.0 - (product_area * reps) / (sheet_area * 100.0)) * 100.0)
        candidates.append((opt, reps, waste_pct, sheet_area))

    if not candidates:
        return None, None

    # MIN = physically smallest usable print format.  If two formats have
    # the same area, prefer the one with the better yield/waste.
    min_choice = min(candidates, key=lambda c: (c[3], c[2], -c[1]))

    # MAX (Optim.) = the format with the lowest waste.  If waste is tied,
    # prefer the smaller physical format and then the higher yield.
    max_choice = min(candidates, key=lambda c: (c[2], c[3], -c[1]))

    def pack(c):
        opt, reps, waste_pct, _ = c
        return opt, reps, waste_pct

    return pack(min_choice), pack(max_choice)


@dataclass
class Inputs:
    source_format: str
    product_w: float
    product_h: float
    quantity: int  # единични бройки на готовото изделие
    front_colors: int
    back_colors: int
    turnover: str
    useful_grip: bool
    duplication: str = ""
    duplication_count: int = 0
    paper_price_per_sheet: float = 0.15
    vat: str = "без"
    plates: bool = True
    plate_supply: str | None = None
    cutting: str = "стандарт"
    numbering: str = "без"
    perforation: str = "без"
    prepress_type: str = "флаери"
    prepress_price: float = 0.0
    uv: str = "без"
    lamination: str = "без"
    calender: str = "без к"
    film_price: float = 0.0
    gluing: str = "без"
    folding: int = 0
    folding_type: str = "Гатеви"
    bigoving: int = 0
    bigoving_type: str = "без"
    die_cut: str = "без"
    round_punch: str = ""
    breaking: float = 40.0
    electric_montage: str = "без"
    typesetting: str = ""
    sewing: str = "без"
    counting: str = "да"
    # Ръчно въведена цена за „Операции, които не са изброени“.
    # Участва в общата калкулация, но не е довършителна операция.
    other_price: float = 0.0
    separators: int = 0
    separator_material: str = "без"
    transport: str = "не"
    transport_km: float = 15
    fuel_price: float = 1.15
    surcharge_pct: float = 40.0
    print_type: str = "цветно"
    press: str = "автоматичен"
    prices: dict | None = None


def waste_sheets(clean_sheets, qty, front, back, turnover):
    if clean_sheets <= 0:
        return 0
    colors = front + back
    if clean_sheets <= 9000:
        if (clean_sheets <= 2000 and colors <= 4) or (clean_sheets <= 2500 and str(turnover).lower() == "да"):
            return 25 * front
        return colors * 30
    return colors * 40


DEFAULT_PRICES = {
    "paper_per_source_sheet": 0.15,
    "print_color": 10.3,
    "print_bw": 5.2,
    "plate": 2.8,
    "duplication": 2.7,
    "prepress_бошура/покана": 2.56,
    "prepress_етикети/визитки": 1.02,
    "prepress_корици": 1.53,
    "prepress_листовки/стикери": 1.28,
    "prepress_минимално": 0.51,
    "prepress_плакат": 2.56,
    "prepress_страниране": 0.25,
    "prepress_флаери": 1.28,
    "electric_монтаж_бошура/покана": 2.56,
    "electric_монтаж_етикети/визитки": 1.02,
    "electric_монтаж_корици": 1.53,
    "electric_монтаж_листовки/стикери": 1.28,
    "electric_монтаж_минимално": 0.51,
    "electric_монтаж_плакат": 2.56,
    "electric_монтаж_страниране": 0.25,
    "electric_монтаж_флаери": 1.28,
    "uv_гланц": 0.028, "uv_гланц_подготовка": 5.624,
    "uv_мат": 0.056, "uv_мат_подготовка": 5.624,
    "uv_частичен": 0.071, "uv_частичен_подготовка": 41.42,
    "uv_част. дв": 0.142, "uv_част. дв_подготовка": 41.42,
    "uv_част.обем": 0.075, "uv_част.обем_подготовка": 43.46,
    "lamination_гланц": 0.051, "lamination_гланц_подготовка": 5.624,
    "lamination_кадифе": 0.2, "lamination_кадифе_подготовка": 7.67,
    "lamination_мат": 0.064, "lamination_мат_подготовка": 5.624,
    "calender_per_sheet": 0.025,
    "film_manual": 0.0,
    "gluing_джоб": 0.06, "gluing_знаменца": 0.02, "gluing_кутии": 0.012, "gluing_дв. лепяща": 0.041,
    "separator_вестник": 0.028, "separator_друг": 0.3,
    "fuel_per_km": 1.15,
}


def excel_labor_markup(costs, percent):
    """
    Exact worksheet G28 logic:
    ROUNDUP(SUM(G4:G7,G9:G12,G16,G18:G20,G22:G26),0) * H28%
    G27 Transport is excluded.
    """
    keys = (
        "print_face_back", "turnover_cost", "over1000_cost", "duplication",
        "cutting", "numbering", "perforation", "prepress",
        "bigoving", "breaking_cost", "gluing", "folding",
        "electric_montage", "packaging", "other_operations",
        "separator_cost", "round_punch",
    )
    base = sum(float(costs.get(k, 0) or 0) for k in keys)
    rounded_base = math.ceil(base)
    markup = rounded_base * float(percent or 0) / 100.0
    return markup, base, rounded_base


@dataclass
class HimiyaInputs:
    source_format: str
    print_format: str
    product_w: float
    product_h: float
    front_colors: int
    back_colors: int
    turnover: str
    quantity: int
    sheets_per_block: int
    paper_price: float
    paper_colors: int = 1
    plate_price: float = 2.80
    vat: str = "без"
    plates: str = ""
    cutting: str = "не"
    numbering: str = "без"
    perforation: str = "без"
    prepress_price: float = 0.0
    bigoving_count: int = 0
    bigoving_type: str = "без"
    gluing: str = "без"
    sewing: str = "без"
    sewing_type: str = "телчета"
    typesetting: str = ""
    counting: str = "не"
    other_price: float = 0.0
    electric_montage: str = "без"
    separators_material: str = "без"
    separators: int = 0
    transport: str = "не"
    surcharge_pct: float = 40.0
    useful_grip: str = "не"


def _himiya_norm(s):
    return str(s or '').strip().lower().replace('f.','').replace('х','x').replace('×','x').replace(' ','')


def _excel_mround(x, multiple):
    if not multiple:
        return 0.0
    return math.floor(float(x) / multiple + 0.5) * multiple


def calc_himiya(inp: HimiyaInputs, repo: FormatRepository, forced_repetitions=None):
    # The Himiya worksheet uses the same SmartCalculate/format logic as the
    # first calculator for the actual trimmed-size yield.
    reps = smart_calculate(inp.print_format, inp.product_w, inp.product_h, inp.turnover, inp.useful_grip == 'да')
    if forced_repetitions is not None:
        reps = int(forced_repetitions)
        if reps <= 0:
            raise ValueError('Ръчното размножение трябва да е по-голямо от 0.')
    if reps <= 0:
        raise ValueError('Няма подходящо размножение за зададения печатен формат и размер.')

    unit_pieces = inp.quantity * inp.sheets_per_block
    clean = math.ceil(unit_pieces / reps)
    colors = inp.front_colors + inp.back_colors

    if inp.quantity <= 10:
        waste = colors * 5 + math.ceil(clean * 0.1)
    elif inp.quantity <= 500 and colors < 4:
        waste = colors * 25 + _excel_mround(clean * 0.05, 5)
    else:
        waste = colors * 30 + _excel_mround(clean * 0.08, 5)
    total_turnover = clean + waste

    # Source-sheet yield comes from the workbook's Formats table.
    source_norm = _himiya_norm(inp.source_format)
    if source_norm in ('hymiya','himiya'): source_norm = 'himiya'
    yield_per_source = None
    for o in repo.all_options():
        if _himiya_norm(o.source) == source_norm and _himiya_norm(o.print_format) == _himiya_norm(inp.print_format):
            yield_per_source = o.sheets_per_source
            break
    # Himiya is also stored as a separate source in the supplied workbook.
    if yield_per_source is None and source_norm == 'himiya':
        yield_per_source = 2 if _himiya_norm(inp.print_format) == '43x30.5' else 4
    if yield_per_source is None:
        raise ValueError('Не е намерена връзка между формата на хартията и формата за печат.')

    whole_sheets = _excel_mround(total_turnover / yield_per_source, 5) if inp.front_colors > 0 else 0
    # Excel C17 uses MROUND, so preserve that rather than replacing it with CEILING.
    if whole_sheets <= 0:
        whole_sheets = math.ceil(total_turnover / yield_per_source / 5) * 5

    package_color = whole_sheets / 500

    # Prices are read from the same shared price sheet in the workbook.
    # Current workbook values are intentionally kept here as the exact worksheet
    # defaults; the main app can later expose them through the Prices tab.
    price_print = {'черно': 5.2, 'черно с обр.': 8.0, 'не': 8.0, 'да': 8.0}.get(inp.turnover, 8.0)
    print_face_back = colors * price_print if inp.turnover in ('не','черно') else inp.front_colors * price_print

    if inp.turnover == 'да':
        turnover_cost = 2.7 * inp.back_colors * max(1, math.ceil(clean / 1000)) if inp.back_colors else 0
    elif inp.turnover == 'черно с обр.':
        turnover_cost = _excel_mround(2.7 * inp.back_colors * max(1, clean / 1000), 5) if inp.back_colors else 0
    else:
        turnover_cost = 0.0

    over1000_times = math.floor(clean / 1001)
    over1000_units = colors if inp.turnover in ('не','черно') else inp.front_colors
    over1000 = over1000_units * 2.7 * over1000_times

    # Общата цена на печата е една сума: лице/гръб + обръщане + над 1000.
    # Отделните компоненти остават налични за показване и диагностика.
    print_cost = print_face_back + turnover_cost + over1000

    plate_cost = 0.0
    if inp.plates == 'да':
        plate_count = inp.front_colors
        if inp.turnover in ('не','черно'):
            plate_count += inp.back_colors
        plate_cost = float(inp.plate_price) * plate_count
    elif inp.plates == 'всеки цвят отд.':
        plate_cost = float(inp.plate_price) * colors

    if inp.cutting == 'стандарт':
        packs = math.ceil(inp.quantity / 20)
        cutting = max(0.75, min(15, packs * 0.5))
    elif inp.cutting == 'други':
        packs = math.ceil(inp.quantity / 20)
        cutting = max(2.5, min(30, packs * 1.5))
    else:
        cutting = 0.0

    if inp.numbering == 'да':
        numbering = 0.01 * clean * max(0, inp.paper_colors) + 10.23
    elif inp.numbering == '+':
        numbering = 0.007 * clean + 10.23
    else:
        numbering = 0.0

    if inp.perforation == '+':
        perforation = 10.23 + (2.7 * over1000_times) * 1.8
    elif inp.perforation == 'да':
        perforation = 10.23 + (over1000_times * inp.front_colors) + 2 if clean > 1000 else 10.23
    else:
        perforation = 0.0

    bigoving = 0.0
    if inp.bigoving_count > 0 and inp.bigoving_type != 'без':
        if inp.quantity * inp.bigoving_count <= 1000:
            bigoving = (inp.quantity + 40) * 0.005 * inp.bigoving_count
        else:
            bigoving = _excel_mround((inp.quantity + 40) * 0.09 * inp.bigoving_count, 0.5)
        if inp.bigoving_type == 'машинно':
            bigoving = _excel_mround(bigoving / 2, 0.5)

    if inp.gluing == 'без':
        gluing = 0.0
    elif inp.gluing == 'кубчета':
        gluing = 0.1 * (clean / 100) if clean > 1000 else 1.5
    elif inp.gluing == 'каширане':
        gluing = unit_pieces * 0.064 + 7.67
    else:
        glue_prices = {'джоб':0.06, 'знаменца':0.02, 'кутии':0.012, 'дв. лепяща':0.041, 'а':0.0}
        gluing = unit_pieces * glue_prices.get(inp.gluing, 0.0)

    if inp.sewing not in ('','без'):
        try: sewing_factor = float(str(inp.sewing).replace(',','.'))
        except Exception: sewing_factor = 1.0
        sewing = math.ceil(inp.quantity * 0.005 * sewing_factor)
    else:
        sewing = 0.0

    if inp.typesetting == 'ръчно':
        typesetting = clean / 1000 * max(0, inp.paper_colors) * 1.53
    elif inp.typesetting == 'машинно':
        typesetting = clean * 0.0035 * max(0, inp.paper_colors)
    else:
        typesetting = 0.0

    if inp.counting in ('да','<1000'):
        prep = _excel_mround(clean,1000)/1000*2 if inp.counting in ('да','<1000') else 0
        if inp.counting == 'да':
            package = (_excel_mround(clean,1000)/1000*1.5 if clean >= 500 and reps > 2 else _excel_mround(clean,100)/2500*2.56)
        else:
            package = _excel_mround(clean,100)/100*0.5 + math.floor(inp.quantity/100)*0.4
        counting = max(2, prep + min(package,2))
    else:
        counting = 0.0

    separator_sheets = 0
    separator_cost = 0.0
    if inp.separators_material != 'без':
        i19 = math.floor(clean / max(1, inp.sheets_per_block))
        separator_sheets = math.ceil(i19 / max(1, yield_per_source))
        unit = 0.028 if inp.separators_material == 'вестник' else 0.3 if inp.separators_material == 'картон' else 0
        separator_cost = math.ceil(separator_sheets * unit * 2) / 2

    if inp.transport == 'да':
        delivery = 2.7 if total_turnover <= 1000 else 1.53 + (total_turnover/1500)*1.53
    elif inp.transport == 'доставка+':
        delivery = (2.7 if total_turnover <= 1000 else 1.53 + (total_turnover/1500)*1.53) + 5
    else:
        delivery = 0.0

    paper = whole_sheets * inp.paper_price * max(0, int(inp.front_colors + inp.back_colors if False else 1))
    # In the worksheet G3, C12 is the number of paper colours, not the print colour count.
    # It is supplied through the UI as paper_colors.
    paper = whole_sheets * inp.paper_price * max(1, getattr(inp, 'paper_colors', 1))
    if inp.vat == 'със': paper *= 1.2
    # Excel G3 uses ROUNDUP(...,0.5). Excel treats the fractional num_digits
    # argument as 0 here, so 5.4 becomes 6 (not 5.5).
    if inp.vat == 'без': paper = math.ceil(paper)

    costs = {
        'paper':paper, 'print':print_cost, 'turnover':turnover_cost, 'over1000':over1000,
        'plates':plate_cost, 'cutting':cutting, 'numbering':numbering, 'perforation':perforation,
        'prepress':max(0,inp.prepress_price), 'bigoving':bigoving, 'gluing':gluing, 'sewing':sewing,
        'typesetting':typesetting, 'counting':counting, 'other':max(0,inp.other_price),
        'electric_montage': (0.0 if inp.electric_montage in ('','без') else {'бошура/покана':2.56,'етикети/визитки':1.02,'корици':1.53,'листовки/стикери':1.28,'минимално':0.51,'плакат':2.56,'страниране':0.25,'флаери':1.28}.get(inp.electric_montage,0.0) * (2 if inp.back_colors > 0 and inp.turnover == 'не' else 1)),
        'separators':separator_cost, 'transport':delivery,
    }
    # За оскъпяването печатът участва само веднъж.
    # print_cost вече съдържа turnover + over1000, затова компонентите
    # turnover/over1000 не се добавят втори път към базата.
    labor_base = print_cost + sum(costs[k] for k in ('cutting','numbering','perforation','prepress','bigoving','gluing','sewing','typesetting','counting','other','electric_montage','separators'))
    surcharge = math.ceil(labor_base) * inp.surcharge_pct / 100 if inp.surcharge_pct > 0 else 0

    # По същия принцип и общата сума използва агрегирания печат само веднъж.
    total_base = paper + print_cost + plate_cost + cutting + numbering + perforation + max(0, inp.prepress_price) + bigoving + gluing + sewing + typesetting + counting + max(0, inp.other_price) + costs['electric_montage'] + separator_cost + delivery
    total = math.ceil((total_base + surcharge) * 10) / 10
    return {
        'source_format':inp.source_format,'print_format':inp.print_format,'product_w':inp.product_w,'product_h':inp.product_h,
        'unit_pieces':unit_pieces,'quantity':inp.quantity,'sheets_per_block':inp.sheets_per_block,'front_colors':inp.front_colors,'back_colors':inp.back_colors,
        'repetitions':reps,'clean_sheets':clean,'waste_sheets':waste,'total_turnover':total_turnover,'whole_sheets':whole_sheets,'package_color':package_color,
        'paper_colors':inp.paper_colors, 'separator_sheets':separator_sheets, 'plate_count': (int(plate_count) if 'plate_count' in locals() else 0),
        'print_rate':price_print, 'print_face_back': print_face_back,
        'turnover_cost':turnover_cost, 'over1000_cost':over1000,
        **costs,'surcharge':surcharge,'labor_base':labor_base,'total':total,'unit':total/inp.quantity if inp.quantity else 0,'profit':total-(paper+other if False else 0),
    }

def calc(inputs: Inputs, repo: FormatRepository, forced_print=None, forced_repetitions=None):
    prices = inputs.prices or {}
    def p(key, default):
        try: return float(prices.get(key, default))
        except Exception: return default
    choice = None
    if forced_print:
        normf = lambda x: str(x).strip().lower().replace('х','x').replace('×','x')
        opt = next((o for o in repo.options_for(inputs.source_format) if normf(o.print_format) == normf(forced_print)), None)
        if opt:
            reps = smart_calculate(opt.print_format, inputs.product_w, inputs.product_h, inputs.turnover, inputs.useful_grip)
            choice = (opt, reps)
    if choice is None:
        choice = choose_best(repo, inputs.source_format, inputs.product_w, inputs.product_h, inputs.turnover, inputs.useful_grip)
    if not choice:
        raise ValueError("Няма подходящ печатен формат за зададените размери.")
    opt, reps = choice
    if forced_repetitions is not None:
        reps = int(forced_repetitions)
        if reps <= 0:
            raise ValueError("Ръчното размножение трябва да е по-голямо от 0.")

    # Excel terminology: B8 = единични бройки, B11 = размножения, B12 = чист тираж.
    clean = math.ceil(inputs.quantity / reps)
    waste = waste_sheets(clean, inputs.quantity, inputs.front_colors, inputs.back_colors, inputs.turnover)
    total_turnover = clean + waste
    source_sheets_raw = math.ceil(total_turnover / opt.sheets_per_source)
    source_sheets = math.ceil(source_sheets_raw / 5) * 5

    paper = source_sheets * inputs.paper_price_per_sheet
    if inputs.vat.strip().lower() == "със":
        paper *= 1.2
    paper = round_half(paper)

    colors = inputs.front_colors + inputs.back_colors
    # Exact workbook pricing model for the current calculator:
    # G4 is the single print price (8 EUR in the supplied workbook).
    # The price is multiplied by the relevant colour count; there is no
    # separate user-facing "print type".
    print_rate = p('print_g4', 8.0)
    turnover_rate = p('turnover', 2.7)
    if inputs.turnover in ("не", "черно"):
        print_units = colors
    elif inputs.turnover in ("да", "черно с обр."):
        print_units = inputs.front_colors
    else:
        print_units = 0
    print_face_back = print_units * print_rate

    # Excel G6: FLOOR(clean/1001,1). Thus 834 => zero and no surcharge.
    over1000_times = math.floor(clean / 1001)
    over1000_rate = p('print_over1000', 2.7)
    over1000_units = (colors if inputs.turnover in ("не", "черно") else inputs.front_colors)
    over1000_cost = over1000_units * over1000_rate * over1000_times

    # Excel G5: only a real turnover mode adds a separate turnover charge.
    if inputs.turnover == "да":
        turnover_cost = turnover_rate * inputs.back_colors * max(1, math.ceil(clean / 1000)) if inputs.back_colors > 0 else 0
    elif inputs.turnover == "черно с обр.":
        turnover_cost = mround(turnover_rate * inputs.back_colors * max(1, clean / 1000), 5) if inputs.back_colors > 0 else 0
    else:
        turnover_cost = 0
    print_cost = print_face_back + turnover_cost + over1000_cost

    # Excel I8: B6 + B7 only when B9 is не/пантон/черно.
    # Excel G8: 2.8 * I8 only when H8 == "да".
    plate_count = inputs.front_colors
    if str(inputs.turnover).strip().lower() in ("не", "пантон", "черно"):
        plate_count += inputs.back_colors
    plate_supply = inputs.plate_supply if inputs.plate_supply is not None else ("да" if inputs.plates else "не")
    plates_cost = p('plate', 2.8) * plate_count if str(plate_supply).strip().lower() == "да" else 0.0

    if inputs.cutting == "без":
        cutting = 0
    elif inputs.cutting == "форматиране":
        cutting = max(1.5, round_half(math.floor(source_sheets/501) * opt.sheets_per_source * 0.9))
        if inputs.quantity > 150: cutting *= 1.2
    elif inputs.cutting == "март. Ани":
        cutting = mround(inputs.quantity * 0.00051, 1)
    elif inputs.cutting == "други":
        cutting = max(2.5, round_half(inputs.quantity * 0.0006 * (1 + reps/10)))
    else:
        cutting = max(1.5, round_half(clean * 0.0005))

    # Празно поле/неизбрана стойност означава „без“ и не трябва да начислява цена.
    numbering_option = str(inputs.numbering or '').strip().lower()
    if numbering_option in ('', 'без', 'не'):
        numbering = 0.0
    elif numbering_option == 'да':
        numbering = 0.005 * clean + 10.23
    elif numbering_option == '+':
        numbering = 0.01 * clean + 13
    else:
        numbering = 0.0

    perforation_option = str(inputs.perforation or '').strip().lower()
    if perforation_option in ('', 'без', 'не'):
        perforation = 0.0
    elif perforation_option == '+':
        perforation = 10.23 + (2.56 * math.floor(clean / 1001)) * 1.8
    elif perforation_option == 'да':
        perforation = 10.23 + (2.56 * math.floor(clean / 1001))
    else:
        perforation = 0.0

    # Prepress: explicit user price, otherwise workbook price for selected type × reproductions.
    prepress_prices = {"без":0, "бошура/покана":2.56, "етикети/визитки":1.02, "корици":1.53,
                       "листовки/стикери":1.28, "минимално":0.51, "плакат":2.56,
                       "страниране":0.25, "флаери":1.28}
    # Manual prepress is a total price entered by the user.
    # 0/blank means no prepress charge. It must NOT be multiplied by repetitions.
    prepress = float(inputs.prepress_price or 0) if float(inputs.prepress_price or 0) > 0 else 0.0

    uv_setup = {"гланц":5.624,"гланц дв.":5.624,"кадифе":7.67,"кадифе дв.":7.67,"мат":5.624,"мат дв.":5.624,
                "надпечат.":10.23,"частичен":41.42,"част. дв":41.42,"част.обем":43.46}.get(inputs.uv,0)
    uv_unit = {"гланц":p("uv_gloss",0.028),"гланц дв.":p("uv_gloss",0.028)*2,"кадифе":0,"кадифе дв.":0,"мат":p("uv_matt",0.056),"мат дв.":p("uv_matt",0.056)*2,
               "надпечат.":-1,"частичен":p("uv_partial",0.071),"част. дв":p("uv_partial",0.071)*2,"част.обем":p("uv_volume",0.075)}.get(inputs.uv,0)
    if uv_unit > 0: uv = (clean + 70) * uv_unit + uv_setup
    elif uv_unit < 0: uv = 5 * max(1, math.floor(clean/1001)) + uv_setup
    else: uv = 0

    lam_unit = {"гланц":p("lam_gloss",0.051),"гланц дв.":p("lam_gloss",0.051)*2,"кадифе":p("lam_velvet",0.2),"кадифе дв.":p("lam_velvet",0.2)*2,"мат":p("lam_matt",0.064),"мат дв.":p("lam_matt",0.064)*2}.get(inputs.lamination,0)
    lam_setup = {"гланц":5.624,"гланц дв.":5.624,"кадифе":7.67,"кадифе дв.":7.67,"мат":5.624,"мат дв.":5.624}.get(inputs.lamination,0)
    lamination = (clean + 70) * lam_unit + lam_setup if lam_unit else 0

    calender = 0
    if inputs.calender == "каландър": calender = math.ceil((clean+70)*p("calender",0.025))
    elif inputs.calender == "каландър дв.": calender = math.ceil((clean+70)*p("calender",0.025)*2)

    # Film is intentionally manual: the workbook has a blank film cost cell.
    film = max(0.0, inputs.film_price)

    if inputs.gluing == "каширане": gluing = inputs.quantity*p("glue_lamination",0.064) + 7.67
    elif inputs.gluing == "кубчета": gluing = 1.5 if inputs.quantity <= 1000 else mround(p("glue_cubes_over1000",0.1)*clean/100,1)
    elif inputs.gluing != "без": gluing = (inputs.quantity+20)*{"а":0,"джоб":p("glue_pocket",0.06),"знаменца":p("glue_flags",0.02),"кутии":p("glue_boxes",0.012),"дв. лепяща":p("glue_double",0.041)}.get(inputs.gluing,0)
    else: gluing=0

    # Exact Excel F15: набор.
    if str(inputs.typesetting).strip().lower() == "ръчно":
        typesetting = clean / 1000.0 * inputs.quantity * 1.53
    elif str(inputs.typesetting).strip().lower() == "машинно":
        typesetting = clean * 0.0035 * inputs.quantity
    else:
        typesetting = 0.0

    # Exact Excel formula for шиене.
    sv = str(inputs.sewing).strip().lower()
    if sv and sv != "без":
        try:
            sewing_factor = float(sv.replace(",", "."))
        except Exception:
            sewing_factor = 1.0
        sewing = math.ceil(inputs.quantity * 0.005 * sewing_factor)
    else:
        sewing = 0.0

    # Exact Excel formula for G16 (биговане):
    # LET(бигове;H16; избор;I16; тир_фира;B8+40;
    #     удари;B8*бигове;
    #     хиляди_над;IFERROR(FLOOR.PRECISE(удари/1001;1);0);
    #     ръчен_резултат;IF(OR(бигове="";избор="без");"";
    #         IF(удари<=1000;тир_фира*0,005*бигове;
    #            MROUND(тир_фира*0,012*бигове+2,56*хиляди_над;0,5)));
    #     IF(OR(бигове="";избор="без");"";
    #        IF(избор="машинно";MROUND(ръчен_резултат/2;0,5);ръчен_резултат)))
    # В Python празното/0 бигове се третира като липса на операция и цена 0.
    bigoving = 0.0
    bigoving_count = inputs.bigoving
    bigoving_choice = str(inputs.bigoving_type or '').strip().lower()
    if bigoving_count > 0 and bigoving_choice != 'без':
        strikes = inputs.quantity * bigoving_count          # B8 * H16
        tir_fira = inputs.quantity + 40                     # B8 + 40
        over_thousand = math.floor(strikes / 1001)

        if strikes <= 1000:
            manual_result = tir_fira * 0.005 * bigoving_count
        else:
            manual_result = mround_excel(
                tir_fira * 0.012 * bigoving_count + 2.56 * over_thousand,
                0.5
            )

        if bigoving_choice == 'машинно':
            bigoving = mround_excel(manual_result / 2, 0.5)
        else:
            # „ръчно“ и всяка валидна стойност, различна от „без“,
            # следват ръчния резултат от Excel.
            bigoving = manual_result
    if inputs.folding_type == "ръчно": folding = (inputs.quantity+30)*p("fold_manual",0.005)*inputs.folding
    elif inputs.folding_type == "Гатеви": folding = inputs.quantity*p("fold_gatevi",0.0015)*inputs.folding
    elif inputs.folding_type == "Гад ново": folding = inputs.quantity*p("fold_gad",0.002)*inputs.folding
    else: folding=0

    die = 0
    if inputs.die_cut == "щанцов": die=0.031*(clean+50)+8
    elif inputs.die_cut == "½ щанцов": die=(0.033*(clean+50))*2+8
    elif inputs.die_cut == "преге": die=mround(0.07*23*12*1.28,5)
    elif inputs.die_cut == "преге+": die=mround(mround(0.07*23*12*1.28,5)*2,5.5)

    # Excel G18 = G17 * H18%. The die-cut cost G17 itself is NOT in the
    # labor-markup base, but the associated breaking/processing G18 is.
    # Очупването е зависимо от щанцоването:
    # ако има цена за щанцоване, очупването е процент от нея;
    # ако няма щанцоване, очупване няма, независимо от процента.
    breaking_pct = max(0.0, float(inputs.breaking or 0))
    breaking_cost = die * breaking_pct / 100.0 if die > 0 else 0.0

    electric_prices = {"без":0, "флаери":p("em_flayers",1.28), "листовки/стикери":p("em_leaflets",1.28), "етикети/визитки":p("em_labels",1.02),
                       "бошура/покана":p("em_brochure",2.56), "корици":p("em_covers",1.53), "плакат":p("em_poster",2.56), "страниране":p("em_pagination",0.25), "минимално":p("em_min",0.51)}
    em_unit = electric_prices.get(inputs.electric_montage, 0)
    electric_montage = em_unit * reps
    if inputs.back_colors > 0 and inputs.turnover == "не":
        electric_montage *= 2

    if inputs.counting == "да":
        prep=max(0.5, math.ceil(clean/1000)*0.5)
        packing = (mround(clean,1000)/1000*0.5 if clean>=500 and reps>2 else mround(clean,100)/2500*2.5)
        counting=min(30, prep+packing)
    elif inputs.counting == "<1000":
        counting=min(1, mround(clean,100)/100*0.5 + math.floor(inputs.quantity/100)*0.4)
    else: counting=0

    # Separator material + count: workbook formula returns whole source sheets.
    separator_sheets = 0
    separator_cost = 0
    if inputs.separator_material != "без" and inputs.separators > 0 and opt.sheets_per_source > 0:
        separator_sheets = math.ceil((clean / inputs.separators) / opt.sheets_per_source)
        separator_unit = p('separator_вестник',0.028) if inputs.separator_material=='вестник' else (p('separator_друг',0.3) if inputs.separator_material=='друг' else 0)
        separator_cost = separator_sheets * separator_unit

    transport=0
    if inputs.transport=="да": transport=7.67*inputs.transport_km/100*p('fuel_per_km', inputs.fuel_price)+1
    elif inputs.transport=="доставка+": transport=7.67*inputs.transport_km/100*p('fuel_per_km', inputs.fuel_price)+2.3

    # Excel G7:
    # =IF(H7="да";2,7*I7+H6*I7*7;"")
    # H7 = дали има дублаж, I7 = свободно въведен брой дублажи,
    # H6 = брой пъти над 1000.
    duplication_count = max(0, int(inputs.duplication_count or 0))
    duplication_cost = (
        p('duplication', 2.7) * duplication_count
        + over1000_times * duplication_count * 7
        if str(inputs.duplication).strip().lower() == "да"
        else 0.0
    )

    # Exact Excel G28 markup base:
    # G4:G7 + G9:G12 + G16 + G18:G20 + G22:G26.
    labor_costs = {
        "print_face_back": print_face_back,
        "turnover_cost": turnover_cost,
        "over1000_cost": over1000_cost,
        "duplication": duplication_cost,  # G7
        "cutting": cutting,
        "numbering": numbering,
        "perforation": perforation,
        "prepress": prepress,
        "bigoving": bigoving,
        "breaking_cost": breaking_cost,
        "gluing": gluing,
        "folding": folding,
        "electric_montage": electric_montage,
        "packaging": counting,  # G23: броене/разделяне/.../пакетиране
        "other_operations": max(0.0, float(inputs.other_price or 0.0)),  # Ръчно въведена цена за „Операции, които не са изброени“
        "separator_cost": separator_cost,
        "round_punch": 0.0,  # G26: calculated below
    }    # Exact Excel F26.
    if str(inputs.round_punch).strip().lower() == "заобл.":
        round_punch = (inputs.quantity + 40) / 25.56 * 0.01 + 2.56
    elif str(inputs.round_punch).strip().lower() == "замба":
        round_punch = (inputs.quantity + 40) * 0.0021 + 2.56
    else:
        round_punch = 0.0
    labor_costs["round_punch"] = round_punch

    surcharge, labor_base, labor_base_rounded = excel_labor_markup(
        labor_costs, inputs.surcharge_pct
    )

    # Total cost still includes ALL applicable costs, including materials and transport.
    base=sum([
        paper, print_cost, plates_cost, duplication_cost, cutting, numbering, perforation, prepress,
        uv, lamination, calender, film, gluing, bigoving, folding, die,
        breaking_cost, electric_montage, typesetting, sewing, counting, max(0.0, float(inputs.other_price or 0.0)), separator_cost, round_punch,
        transport
    ])
    total=math.ceil((base+surcharge)*10)/10

    return {
        "source_format": inputs.source_format, "print_format": opt.print_format,
        "product_w": inputs.product_w, "product_h": inputs.product_h,
        "unit_pieces": inputs.quantity, "repetitions": reps,
        "clean_sheets": clean, "waste_sheets": waste, "total_print_sheets": total_turnover,
        "source_sheets_raw": source_sheets_raw, "source_sheets": source_sheets,
        "paper":paper, "print":print_cost, "print_rate":print_rate,
        "print_face_back":print_face_back, "turnover_cost":turnover_cost,
        "over1000_rate":over1000_rate, "over1000_times":over1000_times, "over1000_cost":over1000_cost,
        "plates":plates_cost, "plate_count":plate_count, "duplication":duplication_cost, "duplication_count":duplication_count, "cutting":cutting,
        "numbering":numbering, "perforation":perforation, "prepress":prepress, "typesetting":typesetting, "sewing":sewing,
        "uv":uv, "lamination":lamination, "calender":calender, "film":film,
        "gluing":gluing, "bigoving":bigoving, "folding":folding, "die_cut":die, "electric_montage":electric_montage,
        "counting":counting, "other_operations":max(0.0, float(inputs.other_price or 0.0)), "separator_sheets":separator_sheets, "separators":separator_cost,
        "transport":transport, "surcharge":surcharge, "labor_base":labor_base, "labor_base_rounded":labor_base_rounded, "breaking_cost":breaking_cost, "round_punch":round_punch,
        "total":total, "unit": total/inputs.quantity if inputs.quantity else 0,
    }
