import warnings
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed')
warnings.filterwarnings('ignore', message='Unknown extension is not supported and will be removed')
warnings.filterwarnings('ignore', message='Conditional Formatting extension is not supported and will be removed')
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os, datetime, math, textwrap

# Малки вектороподобни иконки: рисуват се с висока резолюция и се
# намаляват до Tkinter PhotoImage, така че да са остри и без външни SVG файлове.
try:
    from PIL import Image, ImageDraw, ImageTk
    _PIL_ICONS = True
except Exception:
    _PIL_ICONS = False
import openpyxl
from pathlib import Path
from engine_himiya_v15 import FormatRepository, Inputs, HimiyaInputs, calc, calc_himiya, parse_size

# Химия комплект: използваме изрично engine_himiya_v15.py,
# за да не се зарежда случайна стара версия на engine.py от друга папка.
try:
    import inspect
    if 'forced_repetitions' not in inspect.signature(calc_himiya).parameters:
        raise RuntimeError('Зареден е грешен engine_himiya_v15.py: липсва forced_repetitions.')
except Exception:
    raise

ROOT = Path(__file__).resolve().parent
WORKBOOK = next((p for p in [ROOT/'calculator_X_Euro_All.xlsm', ROOT/'calculator_X_Euro_Black(1).xlsm'] if p.exists()), ROOT/'calculator_X_Euro_Black(1).xlsm')
repo = FormatRepository(WORKBOOK)

COST_KEYS = [
    ('paper','Хартия'),('print','Печат'),('duplication','Дублаж'),('plates','Плаки'),('cutting','Рязане'),
    ('numbering','Номерация'),('perforation','Перфорация'),('prepress','Предпечат'),
    ('uv','УВ лак'),('lamination','Ламиниране'),('calender','Каландър'),('film','Филм'),
    ('gluing','Лепене'),('bigoving','Биговане'),('folding','Сгъване'),('die_cut','Щанцоване / преге'),('round_punch','Заоб. / замба'),
    ('breaking_cost','Очупване'),('electric_montage','Ел. монтаж'),('counting','Броене/пакетиране'),('typesetting','Набор'),('separators','Разделители'),
    ('transport','Транспорт'),('surcharge','Оскъпяване на труда')]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Print Calculator — Python v15')
        self.geometry('1380x860'); self.minsize(1160,760)
        self.vars={}; self.var('surcharge', '40'); self.var('prepress_price', '')
        self.request_vars={
            'item': tk.StringVar(value=''),
            'client': tk.StringVar(value=''),
            'date': tk.StringVar(value=datetime.date.today().strftime('%d.%m.%Y')),
            'paper': tk.StringVar(value=''),
            'paper_type': tk.StringVar(value=''),
        }
        # Полетата в таб „Заявка“ са самостоятелни ръчни полета за офертата.
        # Те не се подават обратно към калкулацията — тя продължава да използва
        # съществуващите self.vars / book_vars / himiya_vars / calendar_vars.
        self.request_manual_vars={k: tk.StringVar(value='') for k in (
            'format','pages','repetitions','turnover','print_sheets','qty','color','finishing','bigoving'
        )}
        self.result={}; self.himiya_result={}; self.book_result={}; self.spiral_result={}; self.calendar_result={}; self.last_result_mode='order'; self.request_source_tab='Основен'; self._last_valid=False
        self.active_tab_name = 'Основен'
        self._style(); self._build()

    def _fmt_count(self, value):
        """Показва цели бройки/листа с интервал за хилядите, без да променя числовия резултат."""
        try:
            n = int(float(str(value).replace(',', '.')))
            return f'{n:,}'.replace(',', ' ')
        except Exception:
            return str(value)

    def _style(self):
        # Визуален слой — само оформление. НЕ променя изчисленията,
        # формулите, променливите или логиката на калкулаторите.
        s = ttk.Style(self)
        try:
            s.theme_use('clam')
        except Exception:
            pass

        bg = '#F5F7FA'
        card = '#FFFFFF'
        border = '#E3E7EE'
        border_hover = '#CBD2DE'
        text = '#202633'
        muted = '#737C8C'
        accent = '#4F46E5'
        accent_dark = '#3730A3'
        input_bg = '#FFFFFF'
        input_focus = '#FFFFFF'

        self.configure(bg=bg)
        try:
            self.option_add('*Frame.background', bg)
            self.option_add('*Labelframe.background', card)
            self.option_add('*Label.background', bg)
        except Exception:
            pass

        # ---------------------------------------------------------
        # ОСНОВНА ВИЗУАЛНА СИСТЕМА
        # ---------------------------------------------------------
        s.configure('.', background=bg, foreground=text, font=('Segoe UI', 10))
        s.configure('TFrame', background=bg)
        s.configure('TLabel', background=card, foreground=text, font=('Segoe UI', 10))
        s.configure('Sub.TLabel', background=card, foreground=muted, font=('Segoe UI', 9))
        s.configure('Title.TLabel', background=bg, foreground=text, font=('Segoe UI', 20, 'bold'))

        # Светли карти с много дискретна рамка.
        s.configure('Section.TLabelframe',
                    background=card, foreground=text,
                    bordercolor=border, relief='solid', borderwidth=1,
                    padding=8)
        s.configure('Section.TLabelframe.Label',
                    background=card, foreground=text,
                    font=('Segoe UI', 11, 'bold'), padding=(4, 1))

        # ---------------------------------------------------------
        # ПОЛЕТА — чисти, леко „въздушни“, с ясен focus
        # ---------------------------------------------------------
        s.configure('Input.TEntry',
                    fieldbackground=input_bg, foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(7, 2), font=('Segoe UI', 10))
        s.map('Input.TEntry',
              bordercolor=[('focus', accent), ('active', border_hover)],
              lightcolor=[('focus', accent), ('active', border_hover)],
              darkcolor=[('focus', accent), ('active', border_hover)],
              fieldbackground=[('focus', input_focus), ('disabled', '#FFFFFF')],
              foreground=[('disabled', '#9AA1AE')])

        # Цени — отделен стил: при отключване полетата стават бледосини.
        s.configure('PriceEdit.TEntry', fieldbackground='#EAF4FF', foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(7, 2), font=('Segoe UI', 10))
        s.map('PriceEdit.TEntry',
              bordercolor=[('focus', accent), ('active', border_hover)],
              lightcolor=[('focus', accent), ('active', border_hover)],
              darkcolor=[('focus', accent), ('active', border_hover)],
              fieldbackground=[('readonly', '#FFFFFF'), ('focus', '#DDEEFF')],
              foreground=[('readonly', '#4B5563')])

        s.configure('Input.TCombobox',
                    fieldbackground=input_bg, foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(7, 2), font=('Segoe UI', 10))
        s.map('Input.TCombobox',
              bordercolor=[('focus', accent), ('active', border_hover)],
              lightcolor=[('focus', accent), ('active', border_hover)],
              darkcolor=[('focus', accent), ('active', border_hover)],
              fieldbackground=[('focus', input_focus), ('disabled', '#FFFFFF')],
              foreground=[('disabled', '#9AA1AE')])

        # Заявка: запазва се отделният стил/функционалност, но е в новата визия.
        s.configure('RequestOrange.TEntry',
                    fieldbackground=input_bg, foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(7, 2), font=('Segoe UI', 10, 'bold'))
        s.map('RequestOrange.TEntry',
              bordercolor=[('focus', accent), ('active', border_hover)],
              lightcolor=[('focus', accent), ('active', border_hover)],
              darkcolor=[('focus', accent), ('active', border_hover)],
              fieldbackground=[('focus', input_focus)])

        # Книжки — леко синьо поле, за да остане визуално различимо.
        s.configure('BookRepetition.TEntry',
                    fieldbackground='#FFFFFF', foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(7, 2), font=('Segoe UI', 10))
        s.configure('BookRepetition.TCombobox',
                    fieldbackground='#FFFFFF', foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(7, 2), font=('Segoe UI', 10))

        # ---------------------------------------------------------
        # РЕЗУЛТАТИ / ДИАГНОСТИКА
        # ---------------------------------------------------------
        s.configure('Calc.TLabel', background=card, foreground='#4B5563', padding=(0, 2))
        s.configure('Input.TLabel', background=card, foreground=text)
        s.configure('Value.TLabel', background=card, foreground=text, font=('Segoe UI', 10, 'bold'))
        # Кочани → Ценообразуване: цялото каре е единен бял блок.
        s.configure('HimiyaWhite.TLabelframe',
                    background='#FFFFFF', foreground=text,
                    bordercolor='#D9DEE7', relief='solid', borderwidth=1, padding=8)
        s.configure('HimiyaWhite.TLabelframe.Label',
                    background='#FFFFFF', foreground=text,
                    font=('Segoe UI', 11, 'bold'), padding=(4, 1))
        s.configure('HimiyaPrice.TFrame', background='#FFFFFF')
        s.configure('HimiyaPrice.TLabel', background='#FFFFFF', foreground=text, font=('Segoe UI', 10))
        s.configure('HimiyaPriceBold.TLabel', background='#FFFFFF', foreground=text, font=('Segoe UI', 10, 'bold'))
        # Етикети вътре в карета: фонът е идентичен с фона на самото каре.
        s.configure('Card.TLabel', background=card, foreground=text, font=('Segoe UI', 10))
        s.configure('Card.TFrame', background=card)
        # Специален стил за текста в „Книжки → Ценообразуване на печата“.
        # Използваме същия фон като самото каре и не оставяме TLabel да
        # наследява отделен фон от общия TLabel стил.
        s.configure('BookPrice.TLabel', background=card, foreground=text,
                    font=('Segoe UI', 10))
        s.configure('BookPriceBold.TLabel', background=card, foreground=text,
                    font=('Segoe UI', 10, 'bold'))
        s.configure('DiagHighlight.TLabel', background=card, foreground=accent_dark,
                    font=('Segoe UI', 10, 'bold'))

        # ---------------------------------------------------------
        # БУТОНИ — само „ИЗЧИСЛИ“ е силният акцент.
        # ---------------------------------------------------------
        s.configure('TButton',
                    background='#FFFFFF', foreground=text,
                    bordercolor=border, lightcolor=border, darkcolor=border,
                    padding=(12, 6), font=('Segoe UI', 10))
        s.map('TButton',
              background=[('active', '#F5F6F8'), ('pressed', '#ECEEF2'), ('disabled', '#F3F4F6')],
              foreground=[('disabled', '#A0A6B1')],
              bordercolor=[('focus', accent), ('active', border_hover)])

        s.configure('Big.TButton',
                    background=accent, foreground='#FFFFFF',
                    bordercolor=accent, lightcolor=accent, darkcolor=accent,
                    padding=(16, 8), font=('Segoe UI', 10, 'bold'))
        s.map('Big.TButton',
              background=[('active', accent_dark), ('pressed', accent_dark), ('disabled', '#C7CBD4')],
              foreground=[('disabled', '#FFFFFF')],
              bordercolor=[('focus', accent_dark)])

        # Крайната цена е най-силният текстов акцент.
        s.configure('Total.TLabel', background=bg, foreground=accent_dark,
                    font=('Segoe UI', 20, 'bold'))
        s.configure('TopTitle.TLabel', background=bg, foreground=text,
                    font=('Segoe UI', 16, 'bold'))
        s.configure('TopSubtitle.TLabel', background=bg, foreground=muted,
                    font=('Segoe UI', 9))
        s.configure('PriceCaption.TLabel', background=bg, foreground=muted,
                    font=('Segoe UI', 8, 'bold'))
        s.configure('SummaryValue.TLabel', background=card, foreground=accent_dark,
                    font=('Segoe UI', 10, 'bold'))
        s.configure('SummaryLabel.TLabel', background=card, foreground=muted,
                    font=('Segoe UI', 9))
        s.configure('PrimaryInfo.TLabel', background=card, foreground=text,
                    font=('Segoe UI', 11, 'bold'))

        # ---------------------------------------------------------
        # ТАБОВЕ — по-ясен активен раздел, без промяна на поведението.
        # ---------------------------------------------------------
        s.configure('TNotebook', background=bg, borderwidth=0, tabmargins=(0, 0, 0, 0),
                    tabposition='n')
        s.configure('TNotebook.Tab',
                    background='#E9ECF2', foreground=muted,
                    padding=(18, 9), font=('Segoe UI', 10, 'bold'),
                    borderwidth=0)
        s.map('TNotebook.Tab',
              background=[('selected', '#FFFFFF'), ('active', '#F3F4F8')],
              foreground=[('selected', accent_dark), ('active', text)],
              padding=[('selected', (20, 9)), ('active', (19, 9))],
              font=[('selected', ('Segoe UI', 10, 'bold'))])

    def card(self,parent,title):
        # Общ контейнер за картите. Визуално е светъл, с дискретна рамка;
        # вътрешната логика и grid подредбата на съществуващите табове не се пипат.
        return ttk.Labelframe(parent,text=title,style='Section.TLabelframe',padding=12)

    def var(self,key,default=''):
        # Reuse an existing StringVar when the same field appears in more than
        # one UI section (e.g. prepress_price is shown in Materials and Prices).
        # Creating a second variable would disconnect the first Entry from
        # the value read by calculate().
        if key in self.vars:
            return self.vars[key]
        self.vars[key]=tk.StringVar(value=default)
        return self.vars[key]
    def combo(self,parent,key,values,default=None,width=20):
        v=self.var(key,default if default is not None else (values[0] if values else ''))
        return ttk.Combobox(parent,textvariable=v,values=values,state='readonly',width=width,style='Input.TCombobox')

    def _make_dropdown_wide(self, combo, values):
        """Разширява само падащия списък на Combobox-а.

        Самото поле остава със същата компактна ширина. Разширяваме
        popup-прозореца след отварянето му, така че дългите опции да се
        виждат изцяло, без да променяме размера на основния прозорец.
        """
        try:
            import tkinter.font as tkfont
            font = tkfont.nametofont('TkDefaultFont')
            # Ширина в пиксели според реалния текст + малък запас.
            text_px = max([font.measure(str(v)) for v in values] + [0])
            dropdown_px = min(max(220, text_px + 34), 420)
        except Exception:
            dropdown_px = 260

        def find_listboxes(widget):
            found = []
            try:
                if self.tk.call('winfo', 'class', widget) == 'Listbox':
                    found.append(widget)
                for child in self.tk.call('winfo', 'children', widget).split():
                    found.extend(find_listboxes(child))
            except Exception:
                pass
            return found

        def widen(event=None):
            def apply_width():
                try:
                    popdown = self.tk.call('ttk::combobox::PopdownWindow', str(combo))
                    # При ttk popdown-ът е Toplevel. Трябва да се използва wm
                    # geometry, а не widget geometry, иначе при част от темите
                    # ширината реално не се променя.
                    # Запазваме текущата височина и позиция на popup-а.
                    # Не задаваме височина 1px — това свива списъка до
                    # хоризонтална черта и скрива опциите.
                    geom = self.tk.call('wm', 'geometry', popdown)
                    import re
                    m = re.match(r'\d+x(\d+)([+-]\d+)([+-]\d+)$', str(geom))
                    if m:
                        height = int(m.group(1))
                        x, y = m.group(2), m.group(3)
                        self.tk.call('wm', 'geometry', popdown,
                                     f'{dropdown_px}x{height}{x}{y}')
                    else:
                        self.tk.call('wm', 'geometry', popdown, f'{dropdown_px}x200')
                    chars = max(20, int(dropdown_px / 7))
                    for lb in find_listboxes(popdown):
                        self.tk.call(lb, 'configure', '-width', str(chars))
                        # Изключваме хоризонталното изрязване при теми, които го
                        # налагат чрез ширината на вътрешния frame.
                        try:
                            self.tk.call(lb, 'configure', '-height', str(min(12, max(4, len(values)))))
                        except Exception:
                            pass
                except Exception:
                    pass
            # При някои Tk теми listbox-ът се създава непосредствено след
            # postcommand, затова прилагаме настройката и след idle.
            try:
                self.after_idle(apply_width)
                self.after(20, apply_width)
            except Exception:
                apply_width()

        combo.configure(postcommand=widen)

    def entry(self,parent,key,default='',width=16):
        return ttk.Entry(parent,textvariable=self.var(key,default),width=width,style='Input.TEntry')
    def card(self,parent,title):
        return ttk.Labelframe(parent, text=title, style='Section.TLabelframe', padding=8)


    def _sync_repetitions_manual_state(self, event=None):
        # Manual sheet count is writable only when the user selects "Ръчно".
        mode = self.vars.get('repetition_mode', tk.StringVar()).get().strip().lower()
        state = 'normal' if mode == 'ръчно' else 'disabled'
        if hasattr(self, 'repetitions_manual_widget'):
            self.repetitions_manual_widget.configure(state=state)

    def _bvar(self, key, default=''):
        if not hasattr(self, 'book_vars'):
            self.book_vars = {}
        if key not in self.book_vars:
            self.book_vars[key] = tk.StringVar(value=default)
        return self.book_vars[key]

    def _bcombo(self, parent, key, values, default='', width=16):
        return ttk.Combobox(parent, textvariable=self._bvar(key, default), values=values,
                            state='readonly', width=width, style='Input.TCombobox')

    def _bentry(self, parent, key, default='', width=12):
        return ttk.Entry(parent, textvariable=self._bvar(key, default), width=width, style='Input.TEntry')

    def _refresh_book_prints(self, event=None):
        if not hasattr(self, 'book_print_widget'):
            return
        source = self._bvar('source', '64х90').get()
        source = source.replace('f.','').replace('F.','')
        try:
            values = [str(o.print_format) for o in repo.options_for(source)]
        except Exception:
            values = []
        self.book_print_widget.configure(values=values)
        current = self._bvar('print').get()
        if current not in values:
            self._bvar('print').set(values[0] if values else '')

    def _sync_book_repetitions(self, event=None):
        if not hasattr(self, 'book_vars'):
            return
        mode = self.book_vars.get('repetition_mode')
        manual = self.book_vars.get('repetitions_manual')
        if not mode or not manual:
            return
        is_manual = mode.get() == 'Ръчно'
        try:
            self._book_repetitions_manual.configure(state='normal' if is_manual else 'disabled')
        except Exception:
            pass

    def _book_mround(self, value, multiple):
        try:
            value=float(value); multiple=float(multiple)
            if multiple == 0: return 0.0
            return math.floor(value/multiple + 0.5) * multiple
        except Exception:
            return 0.0

    def _calendar_mround(self, value, multiple):
        """Excel MROUND equivalent used by the Calendars calculator."""
        try:
            value=float(value); multiple=float(multiple)
            if multiple == 0:
                return 0.0
            return math.floor(value / multiple + 0.5) * multiple
        except Exception:
            return 0.0

    def _book_floor(self, value, multiple):
        try:
            value=float(value); multiple=float(multiple)
            if multiple == 0: return value
            return math.floor(value/multiple) * multiple
        except Exception:
            return 0.0

    def _book_layout_repetitions(self):
        """Автоматично размножение: брой готови страници върху едната страна на печатния лист.

        Тиражният/печатният формат в таблицата е в сантиметри (напр. 45x32),
        а обрязаният размер в "Книжки" се въвежда в милиметри (напр. 297x210).
        Затова първо преобразуваме формата в mm.
        """
        try:
            pw, ph = parse_size(self._bvar('print').get())
            tw = float(self._bvar('trim_w').get().replace(',', '.'))
            th = float(self._bvar('trim_h').get().replace(',', '.'))
        except Exception:
            return 0
        # Форматите от repo са в cm; обрязан размер е в mm.
        pw *= 10.0
        ph *= 10.0
        best = 0
        for sw, sh in ((pw, ph), (ph, pw)):
            if tw > 0 and th > 0:
                best = max(best, int(sw // tw) * int(sh // th))
        return max(0, best)

    def _update_book_print_logic(self, *args):
        try:
            pages = float(self._bvar('pages').get().strip().replace(',', '.'))
        except Exception:
            pages = 0
        mode = self._bvar('repetition_mode').get().strip()
        if mode == 'Ръчно':
            try:
                reps = float(self._bvar('repetitions_manual').get().strip().replace(',', '.'))
            except Exception:
                reps = 0
        else:
            reps = self._book_layout_repetitions()
        cols = (pages / reps / 2) if pages > 0 and reps > 0 else 0
        self._bvar('book_cols').set(str(cols) if cols else '')
        turnover = '—'
        if cols:
            frac = round(cols % 1, 2)
            turnover = {0:'не', .25:'¼ с обръщ.', .5:'½ с обръщ.', .75:'¾'}.get(frac, 'индивидуално')
        self._bvar('turnover_calc').set('' if turnover == '—' else turnover)
        if hasattr(self, 'book_cols_label'):
            self.book_cols_label.configure(text=(f'{cols:g}' if cols else '—'))
        if hasattr(self, 'book_turnover_label'):
            self.book_turnover_label.configure(text=turnover)

    def _calculate_booklets(self):
        """Книжки — директна реализация на формулите от Excel листа „Книжки".
        Интерфейсът използва същите полета, но всички стойности се изчисляват тук,
        без да се променя engine_himiya_v15.py.
        """
        try:
            def fnum(key, default=0.0):
                try:
                    raw=self._bvar(key, str(default)).get().strip().replace(' ','').replace(',','.')
                    return float(raw) if raw else default
                except Exception:
                    return default
            def sval(key, default=''):
                return self._bvar(key, default).get().strip()
            def roundup(v, digits=0):
                p=10**digits
                return math.ceil(float(v)*p)/p
            def mround(v, multiple):
                return self._book_mround(v, multiple)
            def floor(v, multiple):
                return self._book_floor(v, multiple)

            source=sval('source').replace('f.','').replace('F.','')
            print_fmt=sval('print')
            qty=int(round(fnum('qty')))
            pages=fnum('pages')
            front=int(round(fnum('front'))); back=int(round(fnum('back')))
            tw=fnum('trim_w'); th=fnum('trim_h')
            if not source: raise ValueError('Моля, изберете изходен формат.')
            if not print_fmt: raise ValueError('Моля, изберете Печатен формат.')
            if qty<=0: raise ValueError('Ед. бройки трябва да е по-голямо от 0.')
            if pages<=0: raise ValueError('Брой страници трябва да е по-голям от 0.')
            if front<0 or back<0: raise ValueError('Цветността не може да е отрицателна.')
            if tw<=0 or th<=0: raise ValueError('Въведете валиден обрязан размер.')

            mode=sval('repetition_mode')
            if mode == 'Ръчно':
                reps=fnum('repetitions_manual')
                if reps<=0: raise ValueError('Ръчното размножение трябва да е по-голямо от 0.')
            else:
                reps=self._book_layout_repetitions()
                if reps<=0: raise ValueError('Избраният Печатен формат не побира зададения обрязан размер.')

            # Excel C11 — печатни коли. C10 е размножението на тиражния лист/лице.
            cols=pages/reps/2
            if cols < 1:
                cols=1
            # Excel C12 — обръщане от дробната част на C11.
            frac=round(cols%1,2)
            turnover={0:'не',.25:'¼ с обръщ.',.5:'½ с обръщ.',.75:'¾'}.get(frac,'индивидуално')
            clean=qty/2 if turnover=='да' else qty
            colors=front+back

            opt=next((o for o in repo.options_for(source) if str(o.print_format).strip().lower().replace('х','x').replace('×','x')==print_fmt.strip().lower().replace('х','x').replace('×','x')),None)
            if not opt: raise ValueError('Избраният Печатен формат не е намерен в таблицата с формати.')
            source_count=opt.sheets_per_source
            waste=clean+40 if front==1 else (clean+40*front if front>1 else '')
            if waste=='': waste=clean
            whole=mround(cols,0.5)*mround(float(waste)/source_count,10)

            # G3 — хартия.
            paper_price=fnum('paper',0.085)
            paper=whole*paper_price*(1.2 if sval('vat')=='със' else 1)

            # G4 — печат лице/гръб; цената е Цени -> Печат лице/гръб.
            try: print_rate=float(self.var('price_print_g4','8.00').get().replace(',','.'))
            except Exception: print_rate=8.0
            rounded_cols=mround(cols,0.25)
            if turnover=='да': print_cost=print_rate*front
            elif turnover=='не': print_cost=print_rate*cols*colors
            elif turnover in ('½ с обръщ.','¼ с обръщ.'): print_cost=rounded_cols*print_rate*colors
            elif turnover=='¾': print_cost=print_rate*math.ceil(cols)*colors
            else: print_cost=0

            # G5 — обръщане.
            base_turn=fnum('price_turnover',2.70)
            back_color_cost=0 if turnover=='не' else base_turn*back
            tirazh_zakr=floor(clean,100)
            corrected={'½ с обръщ.':tirazh_zakr/2,'¼ с обръщ.':tirazh_zakr/2,'да':tirazh_zakr,'¾':tirazh_zakr}.get(turnover,0)
            val_L=max(1,corrected/1000)*back_color_cost
            val_M=max(1,tirazh_zakr/1000)*back_color_cost
            if turnover=='да': turn_cost=mround(val_M,base_turn)
            elif turnover in ('½ с обръщ.','¼ с обръщ.'): turn_cost=mround(val_L,base_turn)
            else: turn_cost=0

            # H6/G6 — брой плаки и стойност. Цената за продажба идва от таб Цени.
            if turnover=='да': plate_count=front
            elif turnover=='не': plate_count=colors*cols
            elif turnover in ('½ с обръщ.','¼ с обръщ.'): plate_count=colors*mround(cols,0.5)
            elif turnover=='¾': plate_count=colors*mround(cols,0.5)
            else: plate_count=0
            plate_price=fnum('price_plate',2.80)
            plates_cost=plate_price*plate_count

            # G7 — рязане.
            cutting=sval('cutting','без')
            cutting_cost=0
            if cutting=='стандарт':
                cutting_cost=1.53*cols if clean<=1000 else 0.151*(clean/100)*cols
            elif cutting=='обряз.3стр.':
                prep=9
                per_1000=5
                thickness=1.5 if pages>100 else 1
                cutting_cost=mround(prep+(clean/1000)*per_1000*thickness,0.5)
            elif cutting=='без':
                cutting_cost=0

            # G8/H8 — печат над 1000.
            over1000_times=math.floor(clean/1001)
            over_cost=cols*front*2.7*over1000_times

            # G9/G10 — ръчни стойности.
            prepress=fnum('prepress',0)
            cover=fnum('cover_price',0)

            # G11 — шиене + ръчно зададен брой телчета H11.
            sewing_opt=sval('sewing')
            staples=fnum('sewing_count',2)
            if sewing_opt=='без': sewing=0
            elif qty<=1000: sewing=roundup((qty+7.67)*0.0051*staples,0)
            else: sewing=mround((qty+7.67)*0.0051*staples + 2.7*math.floor(qty*staples/1001),1)

            # G12 — набиране; I12 е действителните коли за набор.
            ts_opt=sval('typesetting')
            ts_cols=fnum('typesetting_cols',0)
            if ts_opt=='не': typesetting=0
            elif ts_opt=='машинно': typesetting=qty*0.0036*ts_cols
            elif ts_opt=='ръчно': typesetting=roundup(qty*0.0051*ts_cols,0)
            elif ts_opt=='календари': typesetting=qty*0.0018*ts_cols
            else: typesetting=0

            # G13 — сгъване; H13=брой гънки, I13=начин.
            fold_opt=sval('folding'); folds=fnum('fold_count',0)
            if fold_opt=='без': folding=0
            elif fold_opt=='машинно': folding=(clean+30)*0.0051*folds*cols
            elif fold_opt=='ръчно': folding=roundup((clean+40)*0.0013*ts_cols*folds + (5.11 if folds>0 else 0),0)
            else: folding=0

            # G14 — влагане.
            inserting=(clean+10)*0.00511 if sval('inserting')=='да' else 0

            # G15 — ел. монтаж: C9 * VLOOKUP(H15,montaj[#All],2), когато C12 <> "без".
            montage_prices={'без':0,'бошура/покана':2.56,'етикети/визитки':1.02,'корици':1.53,'листовки/стикери':1.28,'минимално':0.51,'плакат':2.56,'страниране':0.25,'флаери':1.28}
            em=pages*montage_prices.get(sval('electric_montage'),0) if turnover!='без' else 0

            # G16 — номерация.
            num_opt=sval('numbering')
            if num_opt=='да': numbering=(0.005*clean*pages)+10.23
            elif num_opt=='+': numbering=(0.01*clean*pages)+15
            else: numbering=0

            # G17 — броене/разделяне/лепене/пакетиране.
            cnt=sval('counting')
            layers=reps
            preparation=mround(clean,1000)/1000*2 if cnt in ('да','<1000') else 0
            if cnt=='да':
                package = (mround(clean,1000)/1000*1 if clean>=500 and layers>2 else mround(clean,100)/2500*1)
            elif cnt=='<1000':
                package=(mround(clean,100)/100*0.5)+(math.floor(qty/100)*0.4)
            else:
                package=0
            counting=0 if cnt=='не' else preparation+min(package,30)

            # G18 — биговане.
            big_n=fnum('bigoving',0)
            big=0 if big_n<=0 else (qty+40)*0.0051*big_n*(pages/2)

            # G19 — лепене.
            gluing=(qty+40)*0.156 if sval('gluing')=='да' else 0

            # G20 — УВ лак/надпечат.
            uv_opt=sval('uv')
            uv_data={
                'без':(0,0), 'гланц':(0.028,5.624), 'гланц дв.':(0.056,5.624),
                'кадифе':(0,7.67), 'кадифе дв.':(0,7.67), 'мат':(0.056,5.624),
                'мат дв.':(0.112,5.624), 'над дв.':(-1,23.008), 'надпечат.':(-1,10.23),
                'частичен':(0.071,41.42), 'част. дв':(0.142,41.42), 'част.обем':(0.075,43.46)
            }
            uv_unit,uv_setup=uv_data.get(uv_opt,(0,0))
            if uv_unit>0: uv=(clean+70)*uv_unit+uv_setup
            elif uv_unit<0: uv=5*math.floor(clean/1001)+uv_setup
            else: uv=0

            # G21 — ламиниране.
            lam_opt=sval('lamination')
            lam_data={
                'без':(0,0), 'гланц':(0.051,5.624), 'гланц дв.':(0.102,5.624),
                'кадифе':(0.2,7.67), 'кадифе дв.':(0.4,7.67), 'мат':(0.064,5.624), 'мат дв.':(0.128,5.624)
            }
            lam_unit,lam_setup=lam_data.get(lam_opt,(0,0))
            lamination=(clean+50)*lam_unit+lam_setup if lam_unit>0 else 0

            # G22 — каландър.
            calender=(clean+70)*0.025 if sval('calender')=='каландър' else 0

            # G23 — пътни разходи.
            transport_pct=fnum('transport',0)
            base_for_transport=paper+print_cost+turn_cost+plates_cost+cutting_cost+over_cost+prepress+cover+sewing+typesetting+folding+inserting+em+numbering+counting+big+gluing+uv+lamination+calender
            transport=roundup(base_for_transport,0)*transport_pct/100 if transport_pct>0 else 0

            # G24 — оскъпяване.
            surcharge_pct=fnum('surcharge',40)
            surcharge_base=print_cost+turn_cost+cutting_cost+over_cost+prepress+sewing+typesetting+folding+inserting+em+numbering+counting+big+gluing
            surcharge=roundup(surcharge_base,0)*surcharge_pct/100 if surcharge_pct>0 else 0

            # G26:G29 — общо, ед. бройка, печалба, разходи.
            total=roundup(paper+print_cost+turn_cost+plates_cost+cutting_cost+over_cost+prepress+cover+sewing+typesetting+folding+inserting+em+numbering+counting+big+gluing+uv+lamination+calender+transport+surcharge,1)
            unit=total/qty if qty else 0
            expenses=paper+plate_count*2.45+transport
            profit=total-expenses

            self.request_source_tab='Книжки'
            self.book_result={
                'total':total,'unit':unit,'profit':profit,'expenses':expenses,
                'paper':paper,'paper_price':paper_price,'print':print_cost,'turnover_cost':turn_cost,'plates':plates_cost,
                'plate_count':plate_count,'cutting':cutting_cost,'over1000':over_cost,
                'over1000_times':over1000_times,'prepress':prepress,'cover':cover,'sewing':sewing,
                'sewing_count':staples,'typesetting':typesetting,'typesetting_cols':ts_cols,
                'folding':folding,'fold_count':folds,'inserting':inserting,'electric_montage':em,
                'numbering':numbering,'counting':counting,'bigoving':big,'gluing':gluing,'uv':uv,
                'lamination':lamination,'calender':calender,'transport':transport,'surcharge':surcharge,
                'cols':cols,'turnover':turnover,'clean':clean,'waste':waste,'source_sheets':source_count,
                'whole_sheets':whole,'print_format':print_fmt,'trim':f'{tw:g}×{th:g}','repetitions':reps,
                'qty':qty,'unit_pieces':qty,'pages':pages,'source':source,'source_format':source
            }
            # Данни за графичната схема. Форматите в таблицата са в cm, размерът е в mm.
            try:
                pw_cm, ph_cm = parse_size(print_fmt)
                self._book_scheme_data=(pw_cm, ph_cm, reps, tw/10.0, th/10.0, turnover)
                self._draw_book_scheme()
            except Exception:
                self._book_scheme_data=None
                self._draw_book_scheme()

            # Основни данни — единствените места за печатни коли и обръщане.
            self.book_cols_label.configure(text=f'{cols:g}')
            self.book_turnover_label.configure(text=turnover)
            d=self.book_diag_labels
            d['book_clean'].configure(text=self._fmt_count(clean))
            d['book_waste'].configure(text=self._fmt_count(waste))
            d['book_repetitions_diag'].configure(text=self._fmt_count(reps))
            d['book_whole_sheets'].configure(text=self._fmt_count(whole))
            d['book_print_diag'].configure(text=print_fmt)
            d['book_trim_diag'].configure(text=f'{tw:g} × {th:g} mm')

            # Материали.
            self.book_material_labels['paper_total'].configure(text=f'{paper:.2f} €')
            self.book_material_labels['plate_count'].configure(text=self._fmt_count(plate_count))
            self.book_material_labels['plate_total'].configure(text=f'{plates_cost:.2f} €')

            # Ценообразуване — само печатни показатели и общ финансов резултат.
            surcharge_base_display=roundup(surcharge_base,0)
            for key,val in {
                'print':print_cost,'surcharge_base':surcharge_base_display,'surcharge':surcharge,
                'total':total,'profit':profit,'expenses':expenses
            }.items():
                if key in self.book_price_labels:
                    self.book_price_labels[key].configure(text=f'{val:.2f} €')
            if 'turn' in self.book_price_labels:
                self.book_price_labels['turn'].configure(text='НЕ' if turn_cost <= 0 else f'{turn_cost:.2f} €')
            if 'over1000' in self.book_price_labels:
                self.book_price_labels['over1000'].configure(text=f'{over_cost:.2f} € / {self._fmt_count(over1000_times)}' if over1000_times > 0 else '0,00 € / 0')

            # Цената на всяка довършителна операция се показва само до самата операция.
            finish_values={
                'cover':cover,'sewing':sewing,'typesetting':typesetting,'folding':folding,
                'inserting':inserting,'electric_montage':em,'numbering':numbering,'counting':counting,
                'bigoving':big,'gluing':gluing,'uv':uv,'lamination':lamination,'calender':calender,
                'cutting':cutting_cost,'prepress':prepress,'transport':transport,'surcharge':surcharge
            }
            for key,label in getattr(self,'book_finish_price_labels',{}).items():
                value=finish_values.get(key,0)
                label.configure(text=f'{value:.2f} €' if value else '—')
            self._last_valid=True
            self._update_top_bar()
        except Exception as e:
            self._last_valid=False
            messagebox.showerror('Грешка при изчислението',str(e))

    def _booklets(self, f):
        """Книжки — интерфейс, организиран като Основен/Кочани:
        материалните цени са само в „Материали“, довършителните цени са до
        самите операции, а „Ценообразуване“ съдържа само печатните финансови
        показатели и общия финансов резултат.
        """
        self.book_vars = {}
        self.book_result = {}
        outer = ttk.Frame(f); outer.pack(fill='both', expand=True)
        outer.columnconfigure(0, weight=9); outer.columnconfigure(1, weight=10); outer.columnconfigure(2, weight=13)
        outer.rowconfigure(0, weight=0); outer.rowconfigure(1, weight=0); outer.rowconfigure(2, weight=1)

        left = self.card(outer, 'Основни данни / Печат')
        left.grid(row=0,column=0,sticky='nsew',padx=(0,4),pady=(0,6)); left.columnconfigure(1,weight=1)
        ttk.Label(left,text='Изходен формат').grid(row=0,column=0,padx=(8,8),pady=3,sticky='w')
        cb=self._bcombo(left,'source',repo.sources(),'64х90',14); cb.grid(row=0,column=1,padx=(0,8),pady=3,sticky='w'); cb.bind('<<ComboboxSelected>>', self._refresh_book_prints)
        ttk.Label(left,text='Печатен формат').grid(row=1,column=0,padx=(8,8),pady=3,sticky='w')
        self.book_print_widget=self._bcombo(left,'print',[], '',14); self.book_print_widget.grid(row=1,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(left,text='Размер, мм').grid(row=2,column=0,padx=(8,8),pady=3,sticky='w')
        trimf=ttk.Frame(left); trimf.grid(row=2,column=1,padx=(0,8),pady=3,sticky='w')
        self._bentry(trimf,'trim_w','297',7).pack(side='left'); ttk.Label(trimf,text=' × ',padding=(2,0)).pack(side='left'); self._bentry(trimf,'trim_h','210',7).pack(side='left')
        ttk.Label(left,text='Цветност').grid(row=3,column=0,padx=(8,8),pady=3,sticky='w')
        cf=ttk.Frame(left); cf.grid(row=3,column=1,padx=(0,8),pady=3,sticky='w')
        self._bcombo(cf,'front',[str(i) for i in range(11)],'4',4).pack(side='left'); ttk.Label(cf,text=' + ',padding=(2,0)).pack(side='left'); self._bcombo(cf,'back',[str(i) for i in range(11)],'4',4).pack(side='left')
        ttk.Label(left,text='Ед. бройки').grid(row=4,column=0,padx=(8,8),pady=3,sticky='w'); self._bentry(left,'qty','1500',10).grid(row=4,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(left,text='Брой страници').grid(row=5,column=0,padx=(8,8),pady=3,sticky='w'); self._bentry(left,'pages','16',10).grid(row=5,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(left,text='Размножения').grid(row=6,column=0,padx=(8,8),pady=3,sticky='w')
        rf=ttk.Frame(left); rf.grid(row=6,column=1,padx=(0,8),pady=3,sticky='w')
        self._book_repetition_mode=ttk.Combobox(rf, textvariable=self._bvar('repetition_mode','Автоматично'), values=['','Автоматично','Ръчно'], state='readonly', width=11, style='BookRepetition.TCombobox'); self._book_repetition_mode.pack(side='left')
        self._book_repetitions_manual=ttk.Entry(rf, textvariable=self._bvar('repetitions_manual','2'), width=7, style='BookRepetition.TEntry'); self._book_repetitions_manual.pack(side='left',padx=(6,0)); self._book_repetition_mode.bind('<<ComboboxSelected>>', self._sync_book_repetitions); self._sync_book_repetitions()
        ttk.Label(left,text='Брой печатни коли').grid(row=7,column=0,padx=(8,8),pady=3,sticky='w'); self.book_cols_label=ttk.Label(left,text='—',style='Value.TLabel'); self.book_cols_label.grid(row=7,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(left,text='Обръщане').grid(row=8,column=0,padx=(8,8),pady=3,sticky='w'); self.book_turnover_label=ttk.Label(left,text='—',style='Value.TLabel'); self.book_turnover_label.grid(row=8,column=1,padx=(0,8),pady=3,sticky='w')
        for key in ('pages','repetition_mode','repetitions_manual','source','print','trim_w','trim_h'):
            try: self._bvar(key).trace_add('write', self._update_book_print_logic)
            except Exception: pass
        self._refresh_book_prints(); self._update_book_print_logic()

        diag=self.card(outer,'Диагностика — печат и тираж'); diag.grid(row=0,column=1,sticky='nsew',padx=4,pady=(0,6)); diag.columnconfigure(1,weight=1)
        # Няма дублиране на печатни коли/обръщане — те остават само в Основни данни.
        diag_rows=[('Тираж','book_clean'),('Тираж + макулатура','book_waste'),('Размножения','book_repetitions_diag'),('Цели листа','book_whole_sheets'),('Печатен формат','book_print_diag'),('Размер, мм','book_trim_diag')]
        self.book_diag_labels={}
        for i,(lab,key) in enumerate(diag_rows):
            ttk.Label(diag,text=lab).grid(row=i,column=0,sticky='w',padx=8,pady=4); value_style='DiagHighlight.TLabel' if key in ('book_clean','book_whole_sheets','book_print_diag','book_repetitions_diag') else 'Value.TLabel'; w=ttk.Label(diag,text='—',style=value_style); w.grid(row=i,column=1,sticky='w',padx=8,pady=4); self.book_diag_labels[key]=w

        finish=self.card(outer,'Довършителни операции')
        finish.grid(row=0,column=2,sticky='nsew',padx=(4,0),pady=(0,6)); finish.columnconfigure(1,weight=1); finish.columnconfigure(3,weight=1)
        self.book_finish_price_labels={}
        book_finish=[('Корица','cover'),('Шиене','sewing'),('Набиране','typesetting'),('Сгъване','folding'),('Влагане','inserting'),('Ел. монтаж','electric_montage'),('Номерация','numbering'),('Броене/пакетиране','counting'),('Биговане','bigoving'),('Лепене','gluing'),('УВ лак / надпечат','uv'),('Ламиниране','lamination'),('Каландър','calender'),('Обрязване','cutting'),('Предпечат','prepress'),('Пътни разходи','transport'),('Оскъпяване, %','surcharge')]
        choices={'sewing':['без','да'],'typesetting':['не','машинно','ръчно','календари'],'folding':['без','машинно','ръчно'],'inserting':['не','да'],'electric_montage':['без','бошура/покана','етикети/визитки','корици','листовки/стикери','минимално','плакат','страниране','флаери'],'numbering':['без','да','+'],'counting':['не','да','<1000'],'bigoving':['без','1','2','3'],'gluing':['без','да'],'uv':['без','гланц','гланц дв.','кадифе','кадифе дв.','мат','мат дв.','над дв.','надпечат.','частичен','част. дв','част.обем'],'lamination':['без','гланц','гланц дв.','кадифе','кадифе дв.','мат','мат дв.'],'calender':['без к','каландър'],'cutting':['стандарт','обряз.3стр.','без']}
        defaults={'sewing':'да','typesetting':'ръчно','folding':'ръчно','inserting':'да','electric_montage':'без','numbering':'без','counting':'да','bigoving':'без','gluing':'без','uv':'без','lamination':'без','calender':'без к','cutting':'без'}
        for i,(lab,key) in enumerate(book_finish):
            col=(i%2)*2; row=i//2
            ttk.Label(finish,text=lab).grid(row=row,column=col,sticky='w',padx=(8,4),pady=3)
            holder=ttk.Frame(finish); holder.grid(row=row,column=col+1,sticky='w',padx=(0,8),pady=3)
            if key=='cover':
                self._bentry(holder,'cover_price','15',7).pack(side='left')
            elif key=='sewing':
                self._bcombo(holder,key,choices[key],defaults[key],7).pack(side='left'); ttk.Label(holder,text=' телчета').pack(side='left',padx=(4,2)); self._bentry(holder,'sewing_count','2',5).pack(side='left')
            elif key=='prepress':
                self._bentry(holder,key,'55',7).pack(side='left')
            elif key=='surcharge':
                self._bentry(holder,key,'40',6).pack(side='left'); ttk.Label(holder,text='%').pack(side='left',padx=(3,0))
            elif key=='transport':
                self._bentry(holder,key,'0.3',7).pack(side='left'); ttk.Label(holder,text='%').pack(side='left',padx=(3,0))
            elif key=='typesetting':
                self._bcombo(holder,key,choices[key],defaults[key],9).pack(side='left'); ttk.Label(holder,text=' коли').pack(side='left',padx=(4,2)); self._bentry(holder,'typesetting_cols','3',5).pack(side='left')
            elif key=='folding':
                self._bcombo(holder,key,choices[key],defaults[key],8).pack(side='left'); ttk.Label(holder,text=' гънки').pack(side='left',padx=(4,2)); self._bentry(holder,'fold_count','1',5).pack(side='left')
            else:
                vals=choices.get(key,['без','да']); self._bcombo(holder,key,vals,defaults.get(key,vals[0]),10).pack(side='left')
            pl=ttk.Label(holder,text='—',style='Value.TLabel'); pl.pack(side='left',padx=(7,0)); self.book_finish_price_labels[key]=pl

        materials=self.card(outer,'Материали')
        materials.grid(row=1,column=0,columnspan=2,sticky='nsew',padx=(0,4),pady=(0,6)); self.book_material_labels={}
        ttk.Label(materials,text='Цена хартия / лист, €').grid(row=0,column=0,sticky='w',padx=8,pady=4); self._bentry(materials,'paper','0.085',10).grid(row=0,column=1,sticky='w',padx=8,pady=4)
        ttk.Label(materials,text='Хартия').grid(row=0,column=2,sticky='w',padx=8,pady=4); self.book_material_labels['paper_total']=ttk.Label(materials,text='—',style='Value.TLabel'); self.book_material_labels['paper_total'].grid(row=0,column=3,sticky='w',padx=8,pady=4)
        ttk.Label(materials,text='ДДС').grid(row=0,column=4,sticky='w',padx=8,pady=4); self._bcombo(materials,'vat',['без','със'],'без',8).grid(row=0,column=5,sticky='w',padx=8,pady=4)
        # В материалите не показваме единичната цена на плака — само брой и обща сума.
        ttk.Label(materials,text='Брой плаки').grid(row=1,column=0,sticky='w',padx=8,pady=4); self.book_material_labels['plate_count']=ttk.Label(materials,text='—',style='Value.TLabel'); self.book_material_labels['plate_count'].grid(row=1,column=1,sticky='w',padx=8,pady=4)
        ttk.Label(materials,text='Плаки').grid(row=1,column=2,sticky='w',padx=8,pady=4); self.book_material_labels['plate_total']=ttk.Label(materials,text='—',style='Value.TLabel'); self.book_material_labels['plate_total'].grid(row=1,column=3,sticky='w',padx=8,pady=4)

        pricing=self.card(outer,'Ценообразуване на печата')
        pricing.grid(row=2,column=0,columnspan=2,sticky='nsew',padx=(0,4),pady=(0,0)); self.book_price_labels={}
        price_rows=[
            (0,'Печат лице/гръб','print','Обръщане','turn'),
            (1,'Печат над 1000','over1000',None,None),
            (2,'База за оскъпяване','surcharge_base','Оскъпяване','surcharge'),
            (3,'Печалба','profit','Разходи','expenses'),
        ]
        for r, c1, k1, c2, k2 in price_rows:
            label_style = 'BookPriceBold.TLabel' if k1 in ('total', 'profit') else 'BookPrice.TLabel'
            ttk.Label(pricing, text=c1, style=label_style).grid(row=r, column=0, sticky='w', padx=8, pady=3)
            value_style = 'BookPriceBold.TLabel' if k1 in ('total', 'profit') else 'Value.TLabel'
            self.book_price_labels[k1] = ttk.Label(pricing, text='—', style=value_style)
            self.book_price_labels[k1].grid(row=r, column=1, sticky='w', padx=8, pady=3)
            if c2:
                label_style2 = 'BookPriceBold.TLabel' if k2 == 'expenses' else 'BookPrice.TLabel'
                ttk.Label(pricing, text=c2, style=label_style2).grid(row=r, column=2, sticky='w', padx=8, pady=3)
                value_style2 = 'BookPriceBold.TLabel' if k2 == 'expenses' else 'Value.TLabel'
                self.book_price_labels[k2] = ttk.Label(pricing, text='—', style=value_style2)
                self.book_price_labels[k2].grid(row=r, column=3, sticky='w', padx=8, pady=3)
        scheme=self.card(outer,'Схема на разполагане'); scheme.grid(row=1,column=2,rowspan=2,sticky='nsew',padx=(4,0),pady=(0,0)); scheme.columnconfigure(0,weight=1); scheme.rowconfigure(0,weight=1)
        self.book_scheme_canvas=tk.Canvas(scheme,bg='white',highlightthickness=1,highlightbackground='#D0D0D0')
        self.book_scheme_canvas.grid(row=0,column=0,sticky='nsew',padx=4,pady=4)
        self.book_scheme_canvas.bind('<Configure>', lambda e: self._draw_book_scheme())
        self._book_scheme_data=None

    def _draw_book_scheme(self):
        """Визуална схема на разполагането върху тиражния лист."""
        c=getattr(self,'book_scheme_canvas',None)
        data=getattr(self,'_book_scheme_data',None)
        if c is None:
            return
        c.delete('all')
        if not data:
            c.create_text(20,20,anchor='nw',text='Изчислете, за да се покаже схемата.',fill='#777777',font=('Segoe UI',10))
            return
        try:
            pw,ph,reps,tw,th,turnover=data
            cw=max(c.winfo_width(),180); ch=max(c.winfo_height(),160)
            margin_x=26
            margin_y=34
            scale=min((cw-2*margin_x)/pw,(ch-2*margin_y)/ph)
            if scale<=0: return
            sw,sh=pw*scale,ph*scale
            x0=(cw-sw)/2; y0=margin_y+(ch-2*margin_y-sh)/2
            c.create_rectangle(x0,y0,x0+sw,y0+sh,outline='#444444',width=2)
            # Намираме най-доброто разполагане на готовите страници.
            best=None
            for rw,rh in ((tw,th),(th,tw)):
                nx=int(pw//rw) if rw else 0; ny=int(ph//rh) if rh else 0
                n=nx*ny
                if n>=reps and (best is None or n>best[0]): best=(n,nx,ny,rw,rh)
            if best:
                _,nx,ny,rw,rh=best
                for iy in range(ny):
                    for ix in range(nx):
                        if iy*nx+ix>=reps: break
                        x=x0+ix*rw*scale; y=y0+iy*rh*scale
                        c.create_rectangle(x,y,x+rw*scale,y+rh*scale,fill='#DDEBF7',outline='#2B579A',width=1)
                        if rw*scale>24 and rh*scale>16:
                            c.create_text(x+rw*scale/2,y+rh*scale/2,text=str(iy*nx+ix+1),fill='#1F1F1F',font=('Segoe UI',8,'bold'))
            c.create_text(cw/2,8,anchor='n',text=f'{pw:g} × {ph:g} cm',fill='#444444',font=('Segoe UI',9,'bold'))
            c.create_text(cw/2,ch-8,anchor='s',text=f'{reps:g} бр.  •  обръщане: {turnover.upper()}',fill='#555555',font=('Segoe UI',9))
        except Exception:
            c.create_text(20,20,anchor='nw',text='Няма достатъчно данни за схема.',fill='#777777',font=('Segoe UI',10))

    def _reset_booklets(self):
        # Изчистваме всички полета на „Книжки“, включително ръчните цени
        # в „Довършителни“. Ценовите настройки в таб „Цени“ не се пипат.
        for v in getattr(self,'book_vars',{}).values():
            try: v.set('')
            except Exception: pass
        # Изрично изчистване на ръчните ценови полета, за да не останат
        # визуално старите стойности при следващо задание.
        for key in ('cover_price','prepress','surcharge','transport','paper',
                    'sewing_count','typesetting_cols','fold_count'):
            if key in getattr(self,'book_vars',{}):
                try: self.book_vars[key].set('')
                except Exception: pass
        for w in getattr(self,'book_diag_labels',{}).values():
            try: w.configure(text='—')
            except Exception: pass
        for w in getattr(self,'book_material_labels',{}).values():
            try: w.configure(text='—')
            except Exception: pass
        self.book_result={}
        self._book_scheme_data=None
        self._draw_book_scheme()
        for w in getattr(self,'book_price_labels',{}).values():
            try: w.configure(text='')
            except Exception: pass
        for w in getattr(self,'book_finish_price_labels',{}).values():
            try: w.configure(text='')
            except Exception: pass
        self._last_valid=False
        self._update_top_bar()

    def _make_ui_icon(self, kind, size=20, color='#4F46E5'):
        """Създава малка line-icon графика за Tkinter.

        Иконите са чисто визуални: не участват в състоянието или калкулациите.
        Рисуват се на 4x резолюция за по-гладки диагонали/заобляния.
        """
        if not _PIL_ICONS:
            return tk.PhotoImage(width=1, height=1)
        scale = 4
        px = max(12, int(size))
        im = Image.new('RGBA', (px*scale, px*scale), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        c = color
        w = max(2, scale)
        def line(points, width=w):
            d.line([(int(x*scale), int(y*scale)) for x,y in points], fill=c, width=width*scale, joint='curve')
        def rect(box, radius=0, width=w):
            b=tuple(int(v*scale) for v in box)
            d.rounded_rectangle(b, radius=radius*scale, outline=c, width=width*scale)
        def ellipse(box, width=w, fill=None):
            b=tuple(int(v*scale) for v in box)
            d.ellipse(b, outline=c, width=width*scale, fill=fill)
        def poly(points, width=w):
            d.line([(int(x*scale), int(y*scale)) for x,y in points+[points[0]]], fill=c, width=width*scale, joint='curve')

        m=2
        if kind == 'home':
            poly([(m+2,9),(px/2,3),(px-4,9)], 2)
            line([(4,8),(4,px-3),(px-4,px-3),(px-4,8)],2)
            line([(px/2-2,px-3),(px/2-2,px-8),(px/2+2,px-8),(px/2+2,px-3)],2)
        elif kind == 'stack':
            poly([(px/2,3),(px-3,7),(px/2,11),(3,7)],2)
            line([(3,10),(px/2,14),(px-3,10)],2)
            line([(3,13),(px/2,17),(px-3,13)],2)
        elif kind == 'book':
            line([(3,4),(px/2,6),(px/2,px-3),(3,px-5),(3,4)],2)
            line([(px-3,4),(px/2,6),(px/2,px-3),(px-3,px-5),(px-3,4)],2)
            line([(5,8),(px/2-2,9)],1); line([(px/2+2,9),(px-5,8)],1)
        elif kind == 'spiral':
            # Три отворени елиптични намотки, близки до иконата за спирала.
            for off in (0,4,8):
                ellipse((3,3+off/2,px-3,10+off/2),1)
            line([(5,6),(5,px-3),(px-5,px-3)],2)
        elif kind == 'calendar':
            rect((3,4,px-3,px-3),2,2)
            line([(3,8),(px-3,8)],2)
            line([(6,2),(6,6)],2); line([(px-6,2),(px-6,6)],2)
            for x in (7,11,15):
                if x < px-3: ellipse((x,11,x+1.5,12.5),1)
            for x in (7,11,15):
                if x < px-3: ellipse((x,15,x+1.5,16.5),1)
        elif kind == 'document':
            line([(5,2),(13,2),(px-3,6),(px-3,px-3),(5,px-3),(5,2)],2)
            line([(13,2),(13,6),(px-3,6)],2)
            line([(8,10),(px-6,10)],1); line([(8,13),(px-6,13)],1); line([(8,16),(px-7,16)],1)
        elif kind == 'prices':
            ellipse((3,4,11,9),2); ellipse((9,7,17,12),2); ellipse((5,11,13,16),2)
            line([(11,5),(11,8)],1); line([(14,10),(14,13)],1)
        elif kind == 'settings':
            # Семпла gear-подобна иконка, за да остане четима на 18–20 px.
            ellipse((5,5,px-5,px-5),2)
            ellipse((8,8,px-8,px-8),2)
            for a,b in [((10,2),(10,5)),((10,px-5),(10,px-2)),((2,10),(5,10)),((px-5,10),(px-2,10))]: line([a,b],2)
        elif kind == 'materials':
            poly([(px/2,3),(px-3,6),(px/2,9),(3,6)],2)
            line([(3,9),(px/2,12),(px-3,9)],2); line([(3,12),(px/2,15),(px-3,12)],2)
        elif kind == 'costs':
            ellipse((4,4,px-4,10),2); line([(4,7),(4,14)],2); line([(px-4,7),(px-4,14)],2)
            line([(4,14),(px/2,17),(px-4,14)],2)
        elif kind == 'results':
            for x,h in ((4,6),(9,11),(14,8)):
                d.rounded_rectangle((x*scale,(px-h)*scale,(x+3)*scale,px*scale),radius=1*scale,fill=c)
        elif kind == 'total':
            rect((3,3,px-3,px-3),2,2); line([(6,8),(px-6,8)],1); line([(6,12),(px-6,12)],1); line([(6,16),(px-9,16)],1)
        elif kind == 'finish':
            line([(3,14),(8,9),(12,13),(px-3,4)],2); ellipse((px-5,2,px-2,5),1)
        else:
            ellipse((5,5,px-5,px-5),2)

        im = im.resize((px, px), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(im)

    def _icon_for_tab(self, name, size=18):
        kinds = {
            'Основен':'home', 'Кочани':'stack', 'Книжки':'book',
            'Спирали':'spiral', 'Календари':'calendar', 'Заявка':'document', 'Цени':'prices'
        }
        return self._make_ui_icon(kinds.get(name, 'document'), size, '#4F46E5')

    def _build(self):
        # Общата горна лента винаги принадлежи на активния таб.
        # Цената, Изчисли и Изчисти никога не показват/управляват друг таб.
        top = ttk.Frame(self, padding=(22, 12, 22, 10))
        top.pack(fill='x')

        left_top = ttk.Frame(top)
        left_top.pack(side='left', fill='x', expand=True)
        self.top_title_var = tk.StringVar(value='ОСНОВЕН')
        self.top_subtitle_var = tk.StringVar(value='Активен раздел')
        ttk.Label(left_top, textvariable=self.top_title_var,
                  style='TopTitle.TLabel').pack(anchor='w')
        ttk.Label(left_top, textvariable=self.top_subtitle_var,
                  style='TopSubtitle.TLabel').pack(anchor='w', pady=(2,0))

        controls = ttk.Frame(top)
        controls.pack(side='right')

        price_box = ttk.Frame(controls)
        price_box.pack(side='left', padx=(0,18))
        ttk.Label(price_box, text='КРАЙНА ЦЕНА', style='PriceCaption.TLabel').pack(anchor='e')
        self.top_total_var = tk.StringVar(value='—')
        self.top_unit_var = tk.StringVar(value='')
        ttk.Label(price_box, textvariable=self.top_total_var,
                  style='Total.TLabel').pack(anchor='e', pady=(0,0))
        ttk.Label(price_box, textvariable=self.top_unit_var,
                  style='TopSubtitle.TLabel').pack(anchor='e', pady=(0,0))

        self.top_clear_btn = ttk.Button(controls, text='ИЗЧИСТИ', command=self._active_reset)
        self.top_clear_btn.pack(side='left')
        self.top_calc_btn = ttk.Button(controls, text='ИЗЧИСЛИ', style='Big.TButton', command=self._active_calculate)
        self.top_calc_btn.pack(side='left', padx=(8,0))

        self.top_offer_btn = ttk.Button(controls, text='Клиентска оферта', command=self._save_client_offer)
        self.top_offer_btn.pack(side='left', padx=(8,0))

        # Тънка разделителна линия под контролната лента.
        ttk.Separator(self, orient='horizontal').pack(fill='x', padx=22, pady=(0,8))

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=16, pady=(0, 12))
        self.tabs = {}
        self.main_notebook = nb
        nb.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
        # ТУК РАЗМЕСТВАМЕ РЕДА НА ТАБОВЕТЕ:
        # Иконите са само визуален слой към табовете; имената остават
        # непроменени, за да не се засяга логиката на калкулаторите.
        self.tab_icons = {}
        for name in ['Основен', 'Кочани', 'Книжки', 'Спирали', 'Календари', 'Заявка', 'Цени']:
            f = ttk.Frame(nb, padding=14)
            # Реална line-icon графика вместо Unicode символ; само визуално.
            icon = self._icon_for_tab(name, 18)
            self.tab_icons[name] = icon
            nb.add(f, text=f'  {name}', image=icon, compound='left')
            self.tabs[name] = f
            
        self._order_print(self.tabs['Основен'])
        self._prices(self.tabs['Цени'])
        self._himiya(self.tabs['Кочани'])
        self._booklets(self.tabs['Книжки'])
        self._spirals(self.tabs['Спирали'])
        self._calendars(self.tabs['Календари'])
        self._result(self.tabs['Заявка'])
        self._update_top_bar()

    def _on_tab_changed(self, event=None):
        try:
            display = self.main_notebook.tab(self.main_notebook.select(), 'text')
            name = display.strip().split()[-1] if display else 'Основен'
            # „Кочани“ и „Календари“ са една дума; при останалите също
            # последният токен е стабилното вътрешно име на таба.
            if name not in self.tabs:
                name = next((n for n in self.tabs if display.strip().endswith(n)), 'Основен')
        except Exception:
            return
        self.active_tab_name = name
        if name == 'Заявка':
            self._refresh_request_offer()
        self._update_top_bar()

    def _active_calculate(self):
        if self.active_tab_name == 'Кочани':
            self.calculate_himiya()
        elif self.active_tab_name == 'Основен':
            self.calculate()
        elif self.active_tab_name == 'Книжки':
            self._calculate_booklets()
        elif self.active_tab_name == 'Спирали':
            self.calculate_spirals()
        elif self.active_tab_name == 'Календари':
            self.calculate_calendars()

    def _active_reset(self):
        if self.active_tab_name == 'Кочани':
            self.reset_himiya()
        elif self.active_tab_name == 'Основен':
            self.reset()
        elif self.active_tab_name == 'Книжки':
            self._reset_booklets()
        elif self.active_tab_name == 'Спирали':
            self._reset_spirals()
        elif self.active_tab_name == 'Календари':
            self._reset_calendars()

    def _update_top_bar(self):
        name = getattr(self, 'active_tab_name', 'Основен')
        self.top_title_var.set(name.upper())
        self.top_subtitle_var.set('Активен раздел — изчислението се пази отделно за този таб')

        result = None
        can_calculate = False
        if name == 'Основен':
            result = self.result if self.result else None
            can_calculate = True
        elif name == 'Кочани':
            result = self.himiya_result if self.himiya_result else None
            can_calculate = True
        elif name == 'Книжки':
            result = self.book_result if getattr(self, 'book_result', None) else None
            can_calculate = True
        elif name == 'Спирали':
            result = self.spiral_result if getattr(self, 'spiral_result', None) else None
            can_calculate = True
        elif name == 'Календари':
            result = self.calendar_result if getattr(self, 'calendar_result', None) else None
            can_calculate = True
        elif name == 'Заявка':
            source = getattr(self, 'request_source_tab', 'Основен')
            if source == 'Кочани' and self.himiya_result:
                result = self.himiya_result
            elif source == 'Книжки' and getattr(self, 'book_result', None):
                result = self.book_result
            elif source == 'Спирали' and getattr(self, 'spiral_result', None):
                result = self.spiral_result
            elif source == 'Календари' and getattr(self, 'calendar_result', None):
                result = self.calendar_result
            elif source == 'Основен' and self.result:
                result = self.result
        
        if result and result.get('total') is not None:
            self.top_total_var.set(f"{float(result['total']):.2f} €")
            self.top_unit_var.set(f"{float(result.get('unit',0)):.4f} €/бр.")
        else:
            self.top_total_var.set('—')
            self.top_unit_var.set('')

        state = 'normal' if can_calculate else 'disabled'
        self.top_calc_btn.configure(state=state)
        self.top_clear_btn.configure(state=state)
        # Клиентската оферта е общо действие за двата изчислителни таба.
        self.top_offer_btn.configure(state=state)

    def _hvar(self, key, default=''):
        if not hasattr(self, 'himiya_vars'):
            self.himiya_vars = {}
        if key not in self.himiya_vars:
            self.himiya_vars[key] = tk.StringVar(value=default)
        return self.himiya_vars[key]

    def _hcombo(self, parent, key, values, default='', width=18):
        cb = ttk.Combobox(parent, textvariable=self._hvar(key, default), values=values,
                          state='readonly', width=width, style='Input.TCombobox')
        return cb

    def _hentry(self, parent, key, default='', width=14):
        return ttk.Entry(parent, textvariable=self._hvar(key, default), width=width,
                         style='Input.TEntry')

    def _hlabel(self, parent, text, row, col=0, **kw):
        ttk.Label(parent, text=text, **kw).grid(row=row, column=col, sticky='w', padx=6, pady=3)

    def _himiya(self, f):
        """Химия/кочани — оформление по новата схема от превюто.
        Лявата зона е леко стеснена, а дясната е по-широка, за да има
        повече място за довършителните операции и техните полета.
        """
        self.himiya_result = {}
        outer = ttk.Frame(f)
        outer.pack(fill='both', expand=True)

        # Компактно подреждане: без отделен ред „КОЧАНИ“.
        # Горният ред съдържа основни данни/диагностика/довършителни;
        # под него са материалите и ценообразуването вляво, а схемата е
        # непосредствено под довършителните операции вдясно.
        outer.columnconfigure(0, weight=9)
        outer.columnconfigure(1, weight=10)
        outer.columnconfigure(2, weight=13)
        outer.rowconfigure(0, weight=0)
        outer.rowconfigure(1, weight=0)
        outer.rowconfigure(2, weight=1)

        # ---------------------------------------------------------
        # ЛЯВО — ОСНОВНИ ДАННИ / ПЕЧАТ
        # ---------------------------------------------------------
        left = self.card(outer, 'Основни данни / Печат')
        left.grid(row=0, column=0, sticky='nsew', padx=(0,4), pady=(0,3))
        left.columnconfigure(1, weight=1)

        ttk.Label(left,text='Изходен формат').grid(row=0,column=0,padx=(8,8),pady=1,sticky='w')
        cb=self._hcombo(left,'source',
            ['f.64x90','f.70x100','Hymiya','f.60x90','f.64x94','f.64x88','f.60x84','f.61x86','f.43x61'],
            'f.64x90',14)
        cb.grid(row=0,column=1,padx=(0,8),pady=1,sticky='w')
        cb.bind('<<ComboboxSelected>>', self._refresh_himiya_prints)

        ttk.Label(left,text='Печатен формат').grid(row=1,column=0,padx=(8,8),pady=1,sticky='w')
        self.h_print_widget=self._hcombo(left,'print',[], '',14)
        self.h_print_widget.configure(postcommand=self._refresh_himiya_prints)
        self.h_print_widget.grid(row=1,column=1,padx=(0,8),pady=1,sticky='w')

        ttk.Label(left,text='Размер, мм').grid(row=2,column=0,padx=(8,8),pady=1,sticky='w')
        sizef=ttk.Frame(left)
        sizef.grid(row=2,column=1,padx=(0,8),pady=1,sticky='w')
        self._hentry(sizef,'w','214',7).pack(side='left')
        ttk.Label(sizef,text=' × ',padding=(2,0)).pack(side='left')
        self._hentry(sizef,'h','152',7).pack(side='left')

        ttk.Label(left,text='Размножения').grid(row=3,column=0,padx=(8,8),pady=1,sticky='w')
        rf=ttk.Frame(left)
        rf.grid(row=3,column=1,padx=(0,8),pady=1,sticky='w')
        self._h_repetition_mode=self._hcombo(rf,'repetition_mode',['','Автоматично','Ръчно'],'Автоматично',11)
        self._h_repetition_mode.pack(side='left')
        self._h_repetitions_manual=self._hentry(rf,'repetitions_manual','',7)
        self._h_repetitions_manual.pack(side='left',padx=(6,0))
        self._h_repetition_mode.bind('<<ComboboxSelected>>', self._sync_himiya_repetitions)
        self._sync_himiya_repetitions()

        ttk.Label(left,text='Кочани').grid(row=4,column=0,padx=(8,8),pady=1,sticky='w')
        self._hentry(left,'qty','35',10).grid(row=4,column=1,padx=(0,8),pady=1,sticky='w')

        ttk.Label(left,text='Листа в кочан / от цвят').grid(row=5,column=0,padx=(8,8),pady=1,sticky='w')
        sheets_color=ttk.Frame(left)
        sheets_color.grid(row=5,column=1,padx=(0,8),pady=1,sticky='w')
        self._hentry(sheets_color,'sheets','30',8).pack(side='left')
        ttk.Label(sheets_color,text='   Цвят листа').pack(side='left',padx=(8,3))
        self._hentry(sheets_color,'paper_colors','1',5).pack(side='left')

        ttk.Label(left,text='Цветност').grid(row=6,column=0,padx=(8,8),pady=1,sticky='w')
        cf=ttk.Frame(left)
        cf.grid(row=6,column=1,padx=(0,8),pady=1,sticky='w')
        self._hcombo(cf,'front',[str(i) for i in range(11)],'4',4).pack(side='left')
        ttk.Label(cf,text=' + ',padding=(2,0)).pack(side='left')
        self._hcombo(cf,'back',[str(i) for i in range(11)],'0',4).pack(side='left')

        ttk.Label(left,text='Обръщане').grid(row=7,column=0,padx=(8,8),pady=1,sticky='w')
        self._hcombo(left,'turnover',['','не','да'],'не',8).grid(row=7,column=1,padx=(0,8),pady=1,sticky='w')

        ttk.Label(left,text='Полезен грайфер').grid(row=8,column=0,padx=(8,8),pady=1,sticky='w')
        self._hcombo(left,'grip',['','не','да'],'не',8).grid(row=8,column=1,padx=(0,8),pady=1,sticky='w')

        ttk.Label(left,text='Оскъпяване, %').grid(row=9,column=0,padx=(8,8),pady=1,sticky='w')
        self._hentry(left,'surcharge','40',8).grid(row=9,column=1,padx=(0,8),pady=1,sticky='w')

        self._refresh_himiya_prints()

        # ---------------------------------------------------------
        # СРЕДА — ДИАГНОСТИКА / ПЕЧАТ И ТИРАЖ
        # ---------------------------------------------------------
        diag_box=self.card(outer,'Диагностика — печат и тираж')
        diag_box.grid(row=0,column=1,sticky='nsew',padx=4,pady=(0,3))
        diag_box.columnconfigure(1,weight=1)
        diag_rows=[
            ('Размер на печатното изделие','h_product_size'),
            ('Изходен формат','h_source_format'),
            ('Печатен формат','h_print_format'),
            ('Размножения','h_reps'),
            ('Брой кочани','h_blocks_summary'),
            ('цвята листа','h_paper_colors'),
            ('Единични бройки','h_unit'),
            ('Тираж - чист от цвят','h_clean'),
            ('Макулатура','h_waste'),
            ('Тираж + макулатура','h_total_turnover'),
            ('цели листа/ цвят','h_whole'),
            ('пакет от цвят','h_package_color'),
        ]
        blue_attrs={'h_print_format','h_clean','h_whole'}
        for i,(label,attr) in enumerate(diag_rows):
            ttk.Label(diag_box,text=label,style='Card.TLabel').grid(row=i,column=0,padx=(8,8),pady=2,sticky='w')
            value_style='DiagHighlight.TLabel' if attr in blue_attrs else 'Value.TLabel'
            lab=ttk.Label(diag_box,text='—',style=value_style)
            lab.grid(row=i,column=1,padx=(0,8),pady=2,sticky='w')
            setattr(self,attr,lab)

        # ---------------------------------------------------------
        # ВТОРИ / ТРЕТИ РЕД — МАТЕРИАЛИ | ЦЕНООБРАЗУВАНЕ | СХЕМА
        # Материалите и ценообразуването са вляво под Основни данни/Диагностика,
        # а схемата е изцяло под Довършителни операции вдясно.
        materials = self.card(outer, 'Материали')
        materials.grid(row=1,column=0,columnspan=2,sticky='nsew',padx=(0,4),pady=(0,4))
        for c in (1,3,5): materials.columnconfigure(c,weight=1)
        ttk.Label(materials,text='Цена хартия / лист, €').grid(row=0,column=0,padx=7,pady=4,sticky='w')
        self._hentry(materials,'paper','0.054',8).grid(row=0,column=1,padx=(0,8),pady=4,sticky='w')
        self.h_material_paper=ttk.Label(materials,text='—',style='Value.TLabel'); self.h_material_paper.grid(row=0,column=2,padx=(0,12),pady=4,sticky='w')
        ttk.Label(materials,text='ДДС').grid(row=0,column=3,padx=7,pady=4,sticky='e')
        self._hcombo(materials,'vat',['','без','със'],'',7).grid(row=0,column=4,padx=(0,8),pady=4,sticky='w')
        ttk.Label(materials,text='Плаки').grid(row=1,column=0,padx=7,pady=4,sticky='w')
        self._hcombo(materials,'plates',['','да','всеки цвят отд.','не'],'',13).grid(row=1,column=1,padx=(0,8),pady=4,sticky='w')
        self.h_material_plates=ttk.Label(materials,text='—',style='Value.TLabel'); self.h_material_plates.grid(row=1,column=2,padx=(0,12),pady=4,sticky='w')
        ttk.Label(materials,text='Цена предпечат, €').grid(row=1,column=3,padx=7,pady=4,sticky='e')
        self._hentry(materials,'prepress','',8).grid(row=1,column=4,padx=(0,8),pady=4,sticky='w')

        print_box=self.card(outer,'Ценообразуване на печата')
        # Принудително единен бял фон за рамката, заглавието и вътрешността.
        print_box.configure(style='HimiyaWhite.TLabelframe')
        print_box.grid(row=2,column=0,columnspan=2,sticky='nsew',padx=(0,4),pady=(0,4))
        self.himiya_body=ttk.Frame(print_box, style='HimiyaPrice.TFrame')
        self.himiya_body.pack(fill='both',expand=True)

        # Довършителните са само в горната дясна колона, с малко повече ширина/въздух.
        right = self.card(outer, 'Довършителни операции')
        right.grid(row=0,column=2,sticky='nsew',padx=(4,0),pady=(0,4))
        right.columnconfigure(1,weight=1); right.columnconfigure(4,weight=1)
        finish = [
            (('Рязане','cutting',['не','стандарт','форматиране','март. Ани','други'],'стандарт','cutting'),('Лепене','gluing',['без','кубчета','каширане','а','джоб','знаменца','кутии','дв. лепящ'],'без','gluing')),
            (('Номерация','numbering',['без','да','+'],'без','numbering'),('Перфорация','perforation',['без','да','+'],'без','perforation')),
            (('Биговане — вид','big_type',['без','ръчно','машинно'],'без','bigoving'),('Биговане — брой','big_count',None,'0',None)),
            (('Набор','typesetting',['без','машинно','ръчно'],'без','typesetting'),('Шиене/телчета','sewing',['без','1','2','3','4'],'без','sewing')),
            (('Ел. монтаж','em',['без','бошура/покана','етикети/визитки','корици','листовки/стикери','минимално','плакат','страниране','флаери'],'без','electric_montage'),('Пакетиране','counting',['да','<1000','не'],'да','counting')),
            (('Разделители','sep_mat',['без','вестник','картон'],'без','separators'),(None,None,None,None,None)),
            (('Операции, друго','other',None,'','other'),(None,None,None,None,None)),
            (('Транспорт','transport',['не','да','доставка+'],'не','transport'),  ('Километри','km',None,'15',None)),
        ]
        self.h_finish_price_labels={}; self.h_sep_sheets_label=None
        for r_idx,(left_item,right_item) in enumerate(finish):
            for c_idx,item in enumerate((left_item,right_item)):
                if not item[0]: continue
                label,key,values,default,price_key=item
                base_col=c_idx*3
                ttk.Label(right,text=label,style='Card.TLabel').grid(row=r_idx,column=base_col,padx=(7,5),pady=1,sticky='w')
                widget=self._hentry(right,key,default,8) if values is None else self._hcombo(right,key,values,default,13 if key in ('em','gluing') else 10)
                widget.grid(row=r_idx,column=base_col+1,padx=(0,4),pady=1,sticky='w')
                price_label=ttk.Label(right,text='',style='Value.TLabel'); price_label.grid(row=r_idx,column=base_col+2,padx=(0,7),pady=1,sticky='w')
                if key == 'sep_mat':
                    self.h_sep_sheets_label=price_label
                elif price_key: self.h_finish_price_labels[price_key]=price_label
        self.himiya_finish_diag=ttk.Frame(f)

        # Схемата е непосредствено под довършителните операции — без
        # отделно каре „Оптимален формат за печат“, за да се освободи
        # вертикално място в прозореца.
        # Реален етикет за оптималния формат; не оставяме атрибута None,
        # защото _update_himiya_optimal() го обновява след изчисление.
        self.h_optimal_format_suggestion = ttk.Label(
            f, text='MIN: —    |    MAX (Оптим.): —',
            font=('Segoe UI', 9, 'bold'), foreground='#2B579A'
        )

        graphic_box=self.card(outer,'Схема на разполагане')
        graphic_box.grid(row=1,column=2,rowspan=2,sticky='nsew',padx=(4,0),pady=(0,0))
        graphic_box.columnconfigure(0,weight=1)
        self.h_graphic_canvas=tk.Canvas(graphic_box,height=250,bg='white',highlightthickness=1,highlightbackground='#D0D0D0')
        self.h_graphic_canvas.pack(fill='both',expand=True,padx=4,pady=(0,3))


    def _sync_himiya_repetitions(self, event=None):
        mode=self._hvar('repetition_mode').get().strip().lower()
        self._h_repetitions_manual.configure(state='normal' if mode == 'ръчно' else 'disabled')

    def _update_himiya_graphic(self):
        if not hasattr(self,'h_graphic_canvas'): return
        c=self.h_graphic_canvas; c.delete('all'); r=getattr(self,'himiya_result',{}) or {}
        if not r: return
        try:
            sw_cm,sh_cm=parse_size(r['print_format']); sw,sh=sw_cm*10,sh_cm*10
            pw,ph=float(r['product_w']),float(r['product_h']); reps=int(r.get('repetitions',0) or 0)
            grip=3 if self._hvar('grip').get().strip().lower()=='да' else 10; lim_w=sw-5; lim_h=sh-(grip+3)
            turnover=self._hvar('turnover').get().strip().lower()=='да'
            if turnover:
                half=lim_w/2; _,a=self._layout_for_graphic(half,lim_h,pw,ph); rects=a+[(x+half,y,w,h) for x,y,w,h in a]
            else: _,rects=self._layout_for_graphic(lim_w,lim_h,pw,ph)
            rects=rects[:reps]; self.update_idletasks(); W=max(300,c.winfo_width()-8); H=max(190,c.winfo_height()-8)
            scale=min((W-30)/sw,(H-20)/sh) if sw and sh else 1; ox=(W-sw*scale)/2; oy=8
            c.create_rectangle(ox,oy,ox+sw*scale,oy+sh*scale,fill='#F2F2F2',outline='#333333')
            c.create_rectangle(ox+2.5*scale,oy+3*scale,ox+(sw-2.5)*scale,oy+(sh-grip)*scale,outline='#C0392B',dash=(3,3))
            if turnover: c.create_line(ox+(sw/2)*scale,oy,ox+(sw/2)*scale,oy+sh*scale,fill='#C0392B',dash=(4,3))
            for i,(x,y,w,h) in enumerate(rects,1):
                x1=ox+(2.5+x)*scale; y1=oy+(3+y)*scale; x2=x1+w*scale; y2=y1+h*scale
                c.create_rectangle(x1,y1,x2,y2,fill='#DDEBF7',outline='#2F5597')
                if x2-x1>24 and y2-y1>16: c.create_text((x1+x2)/2,(y1+y2)/2,text=str(i),fill='#1F1F1F',font=('Segoe UI',8,'bold'))
        except Exception: pass

    def _update_himiya_optimal(self):
        if not hasattr(self,'h_optimal_format_suggestion'): return
        from engine_himiya_v15 import find_best_sheet_logic
        try:
            source=self._hvar('source').get(); raw_w=self._hvar('w').get().strip().replace(',','.'); raw_h=self._hvar('h').get().strip().replace(',','.')
            if not raw_w or not raw_h: self.h_optimal_format_suggestion.configure(text='MIN: —    |    MAX (Оптим.): —'); return
            mn,mx=find_best_sheet_logic(repo,source,float(raw_w),float(raw_h),self._hvar('turnover').get(),self._hvar('grip').get()=='да')
            def show(title,choice):
                if not choice: return f'{title}: няма подходящ формат'
                opt,reps,waste=choice; return f'{title}: {opt.print_format} • {reps} бр./лист • фира {waste:.1f}%'
            self.h_optimal_format_suggestion.configure(text=show('MIN',mn)+'    |    '+show('MAX (Оптим.)',mx))
        except Exception: self.h_optimal_format_suggestion.configure(text='—')

    def _refresh_himiya_prints(self, event=None):
        if not hasattr(self,'himiya_vars'): return
        src=self._hvar('source').get()
        
        # Опитваме динамично да заредим всички формати от репозиторито (като в Основен таб)
        try:
            # Преобразуваме името, за да съвпадне с ключовете в Excel (напр. 'f.64x90' -> '64х90')
            repo_src = src.replace('f.', '').replace('x', 'х')
            vals = [str(o.print_format) for o in repo.options_for(repo_src)]
        except Exception:
            vals = []
            
        # Ако репозиторито не върне нищо или за даден формат няма опции, 
        # използваме старите твърдо кодирани стойности като резервен вариант (Fallback)
        if not vals:
            mapping={
                'f.70x100':['50x35','35x25','25x23.3','25x17.5'],
                'f.64x90':['45x32','32x30','45x21.3','22.5x21'],
                'Hymiya':['43x30.5','30.5x21.5'],
                'f.60x90':['45x30','45x20','30x20','30x30'],
                'f.64x94':['47x32','31x21','32x21'], 
                'f.64x88':['44x32','44x21','29x21','32x29'],
                'f.60x84':['42x30','30x28'], 
                'f.61x86':['43x30.5','30.5x28.6','30.5x21.5'], 
                'f.43x61':['43x30.5','30.5x21.5']
            }
            vals=mapping.get(src,[])
            
        if hasattr(self,'h_print_widget'):
            self.h_print_widget.configure(values=vals)
            if vals and self._hvar('print').get() not in vals: 
                self._hvar('print').set(vals[0])
        else:
            # locate the widget created above
            pass

    def calculate_himiya(self):
        try:
            def n(k,d=0.0):
                raw=self._hvar(k,str(d)).get().strip().replace(' ','')
                if not raw:return d
                return float(raw.replace(',','.'))
            # Плаките са задължителен избор за Химия.
            plates_choice = self._hvar('plates').get().strip()
            if not plates_choice:
                raise ValueError('Моля, изберете стойност за „Плаки“.')

            inp=HimiyaInputs(
                source_format=self._hvar('source').get(), print_format=self._hvar('print').get(),
                product_w=n('w'), product_h=n('h'), front_colors=int(n('front')), back_colors=int(n('back')),
                turnover=self._hvar('turnover').get(), quantity=int(n('qty')), sheets_per_block=int(n('sheets')),
                paper_price=n('paper'), paper_colors=int(n('paper_colors',1)), plate_price=n('price_plate',2.80), vat=self._hvar('vat').get(), plates=self._hvar('plates').get(),
                cutting=self._hvar('cutting').get(), numbering=self._hvar('numbering').get(), perforation=self._hvar('perforation').get(),
                prepress_price=n('prepress'), bigoving_count=int(n('big_count')), bigoving_type=self._hvar('big_type').get(),
                gluing=self._hvar('gluing').get(), sewing=self._hvar('sewing').get(), sewing_type='телчета',
                typesetting=self._hvar('typesetting').get(), counting=self._hvar('counting').get(), other_price=n('other'),
                electric_montage=self._hvar('em').get(), separators_material=self._hvar('sep_mat').get(),
                # Броят разделители не се въвежда ръчно.
                # Той се изчислява автоматично по формулата от Excel.
                # Вътрешно подаваме листата в кочан като делител за формулата на Excel;
                # това не е ръчно поле и не се показва на потребителя.
                separators=max(1, int(n('sheets'))),
                transport=self._hvar('transport').get(),
                transport_km=n('km', 15),
                surcharge_pct=n('surcharge',40),
                useful_grip=self._hvar('grip').get())
            forced_reps=None
            if self._hvar('repetition_mode').get().strip().lower() == 'ръчно':
                raw=self._hvar('repetitions_manual').get().strip()
                if not raw: raise ValueError('Избрано е ръчно размножение, но не е въведен брой бройки на печатен лист.')
                forced_reps=int(float(raw.replace(',','.')))
                if forced_reps <= 0: raise ValueError('Ръчното размножение трябва да е по-голямо от 0.')
            r=calc_himiya(inp,repo,forced_repetitions=forced_reps); self.himiya_result=r; self.last_result_mode='himiya'; self.request_source_tab='Кочани'
            self._update_top_bar()
            self._refresh_request_offer()
            # Сумите на материалите се показват в карето „Материали“, както в първия таб.
            self.h_material_paper.configure(text=f"{float(r.get('paper',0) or 0):.2f} €")
            self.h_material_plates.configure(text=f"{float(r.get('plates',0) or 0):.2f} €")
            # Цените на довършителните операции са непосредствено до полетата.
            for key, label_widget in getattr(self, 'h_finish_price_labels', {}).items():
                value = float(r.get(key, 0) or 0)
                label_widget.configure(text=f"{value:.2f} €" if value > 0 else '')
            diag_values={
                'h_product_size': f"{r['product_w']:g} × {r['product_h']:g} мм",
                'h_source_format': str(r.get('source_format','')),
                'h_print_format': str(r.get('print_format','')),
                'h_reps': str(r.get('repetitions','')),
                'h_blocks_summary': f"{r.get('quantity','')}х{r.get('sheets_per_block','')} л.",
                'h_paper_colors': str(r.get('paper_colors', self._hvar('paper_colors','1').get())),
                'h_unit': str(r.get('unit_pieces','')),
                'h_clean': str(r.get('clean_sheets','')),
                'h_waste': str(r.get('waste_sheets','')),
                'h_total_turnover': str(r.get('total_turnover','')),
                'h_whole': str(r.get('whole_sheets','')),
                'h_package_color': str(r.get('package_color','')),
            }
            for attr,value in diag_values.items():
                getattr(self,attr).configure(text=value)
            for w in self.himiya_body.winfo_children(): w.destroy()
            grid=ttk.Frame(self.himiya_body, style='HimiyaPrice.TFrame'); grid.pack(fill='x',padx=8,pady=2)
            for c in range(6):
                grid.columnconfigure(c, weight=1 if c in (1,3,5) else 0)
            # Производствената калкулация в Химия следва същата структура
            # като „Ценообразуване на печата“ в първия таб. Не повтаряме
            # една и съща сума два пъти, а показваме цена/единица, дублаж
            # и отделно сумата на печата и общия печат.
            total_print = float(r.get('print',0) or 0)
            rate = float(r.get('print_rate',0) or 0)
            if rate <= 0:
                rate = 5.2 if str(self._hvar('turnover').get()).strip().lower() == 'черно' else 8.0
            print_sum = float(r.get('print_face_back',0) or 0)
            if print_sum <= 0:
                front = int(float(self._hvar('front').get() or 0))
                back = int(float(self._hvar('back').get() or 0))
                mode = str(self._hvar('turnover').get() or '').strip().lower()
                colors = front if mode in ('да','черно с обр.') else front + back
                print_sum = colors * rate

            # Ценообразуването на печата оформено по същия начин като в първия таб.
            # Преди изчисление карето е празно; след изчисление се попълват редовете.
            turnover_cost = float(r.get('turnover',0) or 0)
            duplication_cost = float(r.get('duplication',0) or 0)
            over1000_cost = float(r.get('over1000',0) or 0)
            labor_base = float(r.get('labor_base',0) or 0)
            labor_surcharge = float(r.get('surcharge',0) or 0)
            turnover_txt = 'НЕ' if turnover_cost <= 0 else f'{turnover_cost:.2f} €'
            duplication_txt = 'НЕ' if duplication_cost <= 0 else f'{duplication_cost:.2f} € / да'
            base_txt = f'{labor_base:.2f} € | закръглена: {math.ceil(labor_base):.0f} €' if labor_base > 0 else ''

            # Разходи/печалба — точно по Excel листа „Химия“:
            # G27 = хартия + друго + разделители + транспорт + брой плаки × 2.45;
            # G26 = общо - разходи.
            plate_expense = float(r.get('plate_count', 0) or 0) * 2.45
            expenses = (
                float(r.get('paper', 0) or 0)
                + float(r.get('other', 0) or 0)
                + float(r.get('separators', 0) or 0)
                + float(r.get('transport', 0) or 0)
                + plate_expense
            )
            total_value = float(r.get('total', 0) or 0)
            profit = total_value - expenses

            rows = [
                ('Печат лице/гръб', f'{print_sum:.2f} €', 'Обръщане', turnover_txt),
                ('Обръщане', turnover_txt, 'Дублаж', duplication_txt),
                ('Печат над 1000', f'{over1000_cost:.2f} € / 1 пъти' if over1000_cost > 0 else '', 'Общо печат', f'{total_print:.2f} €'),
                ('База за оскъпяване', base_txt, 'Оскъпяване на труда', f'{labor_surcharge:.2f} €' if labor_surcharge > 0 else ''),
                ('Печалба', f'{profit:.2f} €', 'Разходи', f'{expenses:.2f} €'),
            ]
            for rr,(l1,v1,l2,v2) in enumerate(rows):
                finance1 = l1 in ('Печалба','Разходи')
                finance2 = l2 in ('Печалба','Разходи')
                # Описанията са болд само на „Общо печат“, „Печалба“ и „Разходи“.
                # Всички стойности/цени остават болднати.
                label_style1 = 'HimiyaPriceBold.TLabel' if (l1 == 'Общо печат' or finance1) else 'HimiyaPrice.TLabel'
                label_style2 = 'HimiyaPriceBold.TLabel' if (l2 == 'Общо печат' or finance2) else 'HimiyaPrice.TLabel'
                value_style = 'HimiyaPriceBold.TLabel'
                ttk.Label(grid,text=l1,style=label_style1).grid(row=rr,column=0,sticky='w',padx=6,pady=3)
                ttk.Label(grid,text=v1,style=value_style).grid(row=rr,column=1,sticky='w',padx=4,pady=3)
                ttk.Label(grid,text=l2,style=label_style2).grid(row=rr,column=2,sticky='w',padx=6,pady=3)
                ttk.Label(grid,text=v2,style=value_style).grid(row=rr,column=3,sticky='w',padx=4,pady=3)

            # Диагностика на довършителните операции.
            # За разделителите Excel използва:
            # I19 = FLOOR(чист тираж / листа в кочан; 1)
            # J19 = ROUNDUP(I19 / цели листа от изходния формат; 0)
            # Броят цели листа за разделители е производствен резултат,
            # а не ръчно въведено поле. Първо използваме стойността от engine-а,
            # а ако стара версия на engine-а не я връща — изчисляваме я по Excel.
            sep_sheets = int(r.get('separator_sheets', 0) or 0)
            if self._hvar('sep_mat').get().strip().lower() != 'без' and sep_sheets <= 0:
                try:
                    clean_for_sep = float(r.get('clean_sheets', 0) or 0)
                    sheets_block = max(1, int(r.get('sheets_per_block', self._hvar('sheets').get() or 1)))
                    source = str(r.get('source_format', self._hvar('source').get()))
                    print_fmt = str(r.get('print_format', self._hvar('print').get()))
                    norm = lambda x: str(x).strip().lower().replace('х','x').replace('×','x').replace(' ','')
                    opt = next((o for o in repo.options_for(source) if norm(o.print_format) == norm(print_fmt)), None)
                    if opt:
                        i19 = math.floor(clean_for_sep / sheets_block)
                        sep_sheets = math.ceil(i19 / max(1, int(opt.sheets_per_source)))
                except Exception:
                    sep_sheets = 0

            fin_diag = []
            for lab,key in [
                ('Рязане','cutting'),('Номерация','numbering'),('Перфорация','perforation'),
                ('Предпечат','prepress'),('Биговане','bigoving'),('Лепене','gluing'),
                ('Шиене/телчета','sewing'),('Набор','typesetting'),('Пакетиране','counting'),
                ('Разделители','separators'),('Ел. монтаж','electric_montage'),
                ('Транспорт','transport'),('Друго','other')
            ]:
                if float(r.get(key,0) or 0) > 0:
                    fin_diag.append((lab, f"{r[key]:.2f} €"))
            r['separator_sheets'] = sep_sheets
            if hasattr(self, 'h_sep_sheets_label'):
                self.h_sep_sheets_label.configure(text=(f"{sep_sheets}л. / {float(r.get('separators', 0) or 0):.2f} €" if sep_sheets and float(r.get('separators', 0) or 0) > 0 else (f"{sep_sheets}л." if sep_sheets else '—')))
            if sep_sheets:
                fin_diag.append(('Цели листа за разделители', str(sep_sheets)))
            if float(r.get('labor_base',0) or 0) > 0:
                fin_diag.append(('База за оскъпяване', f"{r['labor_base']:.2f} € | закръглена: {math.ceil(float(r['labor_base'])):.0f} €"))
            if float(r.get('surcharge',0) or 0) > 0:
                fin_diag.append(('Оскъпяване на труда', f"{r['surcharge']:.2f} €"))
            self._set_diag(self.himiya_finish_diag, fin_diag, columns=2)

            # Крайната цена вече е само в общата горна лента.
            # Не я повтаряме в края на производствената калкулация на Химия.
            # Важно: след изчислението обновяваме и общия Result таб.
            # Това липсваше във v13 и затова се виждаше само старото резюме.
            self._update_himiya_graphic()
            self._update_himiya_optimal()
            self._render_result()
        except Exception as e:
            messagebox.showerror('Химия', str(e))

    def _order_print(self,f):
        outer=ttk.Frame(f)
        outer.pack(fill='both',expand=True,anchor='nw')

        # Трите основни вертикални зони следват картинката:
        # 1) Основни данни / Печат
        # 2) Диагностика — печат и тираж
        # 3) Материали отгоре + Довършителни операции под тях.
        work=ttk.Frame(outer)
        work.grid(row=0,column=0,columnspan=2,sticky='nw')

        # Основният таб: Довършителни е разширено за сметка на Диагностика,
        # за да се виждат по-добре дългите стойности в списъчните полета.
        outer.columnconfigure(0,weight=0,minsize=270)
        outer.columnconfigure(1,weight=0,minsize=125)
        outer.columnconfigure(2,weight=1,minsize=0)

        work.columnconfigure(0,weight=0,minsize=270)
        work.columnconfigure(1,weight=0,minsize=125)
        # Долният ред е еднакво висок за ценообразуването и схемата.
        work.rowconfigure(2,minsize=255)

        right_panel=ttk.Frame(outer)
        right_panel.grid(row=0,column=2,sticky='nsew',padx=(8,0))
        right_panel.columnconfigure(0,weight=1)
        right_panel.rowconfigure(2,weight=1)

        # ---------------------------------------------------------
        # ГОРНА ЧАСТ — ЕДНО ОБЩО КАРЕ "ОСНОВНИ ДАННИ / ПЕЧАТ"
        # Подредбата вътре е точно последователна, а не по две
        # отделни секции:
        # Изходен формат
        # Печатен формат
        # Размножения
        # Единични бройки
        # Размер
        # Цветност
        # Обръщане
        # Полезен грайфер
        # Дублаж
        # Оскъпяване
        # ---------------------------------------------------------
        top_card=self.card(work,'Основни данни / Печат')
        top_card.grid(row=0,column=0,sticky='nsew',padx=(0,4),pady=(0,4))

        top_card.columnconfigure(1,weight=1)

        # 1. Изходен формат
        ttk.Label(top_card,text='Изходен формат').grid(
            row=0,column=0,padx=(8,8),pady=2,sticky='w'
        )
        cb=self.combo(top_card,'source',repo.sources(),' ',16)
        cb.grid(row=0,column=1,padx=(0,8),pady=2,sticky='w')
        cb.bind('<<ComboboxSelected>>',self._refresh_print_formats)

        # 2. Печатен формат
        ttk.Label(top_card,text='Печатен формат').grid(
            row=1,column=0,padx=(8,8),pady=2,sticky='w'
        )
        self.print_format_widget=self.combo(top_card,'print_format',[],'Автоматичен',16)
        self.print_format_widget.grid(row=1,column=1,padx=(0,8),pady=2,sticky='w')
        self._refresh_print_formats()

        # 3. Размножения
        ttk.Label(top_card,text='Размножения').grid(
            row=2,column=0,padx=(8,8),pady=2,sticky='w'
        )
        rf=ttk.Frame(top_card)
        rf.grid(row=2,column=1,padx=(0,8),pady=2,sticky='w')
        self.repetition_mode_widget=self.combo(
            rf,'repetition_mode',['','Автоматично','Ръчно'],'',9
        )
        self.repetition_mode_widget.pack(side='left')
        self.repetitions_manual_widget=self.entry(
            rf,'repetitions_manual','',6
        )
        self.repetitions_manual_widget.pack(side='left',padx=(6,0))
        self.repetition_mode_widget.bind(
            '<<ComboboxSelected>>',self._sync_repetitions_manual_state
        )
        self._sync_repetitions_manual_state()

        # 4. Единични бройки
        ttk.Label(top_card,text='Единични бройки').grid(
            row=3,column=0,padx=(8,8),pady=2,sticky='w'
        )
        self.entry(top_card,'qty','',10).grid(
            row=3,column=1,padx=(0,8),pady=2,sticky='w'
        )

        # 5. Размер, мм
        ttk.Label(top_card,text='Размер, мм').grid(
            row=4,column=0,padx=(8,8),pady=2,sticky='w'
        )
        sizef=ttk.Frame(top_card)
        sizef.grid(row=4,column=1,padx=(0,8),pady=2,sticky='w')
        self.entry(sizef,'w','',7).pack(side='left')
        ttk.Label(sizef,text=' × ',padding=(2,0)).pack(side='left')
        self.entry(sizef,'h','',7).pack(side='left')

        # 6. Цветност
        ttk.Label(top_card,text='Цветност').grid(
            row=5,column=0,padx=(8,8),pady=2,sticky='w'
        )
        cf=ttk.Frame(top_card)
        cf.grid(row=5,column=1,padx=(0,8),pady=2,sticky='w')
        self.combo(cf,'front',[str(i) for i in range(11)],'1',4).pack(side='left')
        ttk.Label(cf,text=' + ',padding=(2,0)).pack(side='left')
        self.combo(cf,'back',[str(i) for i in range(11)],'1',4).pack(side='left')

        # 7. Обръщане
        ttk.Label(top_card,text='Обръщане').grid(
            row=6,column=0,padx=(8,8),pady=2,sticky='w'
        )
        self.combo(
            top_card,'turnover',['','не','да','черно','черно с обр.'],'',12
        ).grid(row=6,column=1,padx=(0,8),pady=2,sticky='w')

        # 8. Полезен грайфер
        ttk.Label(top_card,text='Полезен грайфер').grid(
            row=7,column=0,padx=(8,8),pady=2,sticky='w'
        )
        self.combo(top_card,'grip',['','не','да'],'',11).grid(
            row=7,column=1,padx=(0,8),pady=2,sticky='w'
        )

        # 9. Дублаж + брой
        ttk.Label(top_card,text='Дублаж').grid(
            row=8,column=0,padx=(8,8),pady=2,sticky='w'
        )
        df=ttk.Frame(top_card)
        df.grid(row=8,column=1,padx=(0,8),pady=2,sticky='w')
        self.combo(df,'duplication',['','не','да'],'',8).pack(side='left')
        ttk.Label(df,text=' бр.').pack(side='left',padx=(3,1))
        self.entry(df,'duplication_count','',6).pack(side='left')

        # 10. Оскъпяване
        ttk.Label(top_card,text='Оскъпяване на труда, %').grid(
            row=9,column=0,padx=(8,8),pady=2,sticky='w'
        )
        self.surcharge_entry=self.entry(top_card,'surcharge','40',7)
        self.surcharge_entry.grid(row=9,column=1,padx=(0,8),pady=2,sticky='w')
        self.surcharge_warning=ttk.Label(top_card,text='')
        self.surcharge_warning.grid(
            row=10,column=0,columnspan=2,padx=8,pady=(0,2),sticky='w'
        )

        # ---------------------------------------------------------
        # ДИАГНОСТИКА — отделно каре вдясно, в същия горен ред.
        # ---------------------------------------------------------
        diag_box=self.card(work,'Диагностика — печат и тираж')
        diag_box.grid(row=0,column=1,sticky='nsew',padx=(4,0),pady=(0,4))
        diag_frame=ttk.Frame(diag_box)
        diag_frame.pack(fill='both',expand=True,padx=4,pady=4)
        self.diag=diag_frame
        self._set_diag(self.diag,[])

        # Проследяване на оскъпяването.
        def surcharge_watch(*_):
            try:
                v=self.vars['surcharge'].get().strip()
            except Exception:
                return
            if not v:
                self.surcharge_entry.configure(style='Input.TEntry')
                self.surcharge_warning.config(text='')
                return
            if v.replace(',','.')=='40':
                self.surcharge_entry.configure(style='Input.TEntry')
                self.surcharge_warning.config(text='')
            else:
                self.surcharge_entry.configure(style='Warning.TEntry')
                self.surcharge_warning.config(
                    text=f'⚠ Основният процент е 40%. Въведен е {v}%.'
                )
        self.vars['surcharge'].trace_add('write',surcharge_watch)
        surcharge_watch()

        # ---------------------------------------------------------
        # ДОЛНА ЧАСТ — ЦЕНООБРАЗУВАНЕ + ГРАФИКА ВЛЯВО;
        # ОПТИМАЛЕН ФОРМАТ В ДЯСНАТА КОЛОНА.
        # ---------------------------------------------------------

        # Ценообразуване на печата — под Основни данни / Печат.
        print_box=self.card(work,'Ценообразуване на печата')
        print_box.grid(row=2,column=0,sticky='nsew',pady=(0,0))
        self.print_price_diag=ttk.Frame(print_box)
        self.print_price_diag.pack(fill='x',padx=3,pady=2)

        # Оптимален формат — вдясно, непосредствено под довършителните.
        optimal = tk.Frame(right_panel,bg='#FFFFFF',relief='groove',borderwidth=1)
        optimal.grid(row=1,column=0,sticky='ew',pady=(0,4))
        ttk.Label(
            optimal,text='Оптимален формат за печат',background='#FFFFFF',
            foreground='#0078D7',font=('Segoe UI',9,'italic')
        ).pack(anchor='w',padx=8,pady=(3,0))
        self.optimal_format_suggestion=ttk.Label(
            optimal,text='MIN: —    |    MAX (Оптим.): —',
            background='#FFFFFF', font=('Segoe UI',9,'bold'),foreground='#2B579A',wraplength=570
        )
        self.optimal_format_suggestion.pack(anchor='w',padx=8,pady=(0,4))

        # Схема — под Диагностика, непосредствено до Ценообразуване.
        graphic_box=self.card(work,'Схема на разполагане')
        graphic_box.grid(row=2,column=1,sticky='nsew',padx=(4,0),pady=(0,0))
        graphic_box.columnconfigure(0,weight=1)
        self.graphic_canvas=tk.Canvas(
            graphic_box,height=250,bg='white',
            highlightthickness=1,highlightbackground='#D0D0D0'
        )
        self.graphic_canvas.pack(fill='both',expand=True,padx=4,pady=(0,3))

        # ---------------------------------------------------------
        # ДЯСНО — ДОВЪРШИТЕЛНИ ОПЕРАЦИИ
        # ---------------------------------------------------------
        finish_card=self.card(right_panel,'Довършителни операции')
        finish_card.grid(row=0,column=0,sticky='nsew',pady=(0,4))
        finish_card.columnconfigure(0,weight=1)

        mat = ttk.Frame(finish_card, style='Section.TLabelframe')
        mat.grid(row=0, column=0, sticky='ew', padx=2, pady=(2, 3))
        mat.columnconfigure(1, weight=0)
        mat.columnconfigure(2, weight=0)
        mat.columnconfigure(4, weight=1)

        ttk.Label(mat, text='Цена хартия / лист, €', background='#FFFFFF').grid(row=0, column=0, padx=(7, 5), pady=3, sticky='w')
        self.entry(mat, 'paper', '', 8).grid(row=0, column=1, padx=(0, 5), pady=3, sticky='w')
        self.material_paper_price_label = ttk.Label(mat, text='—', style='Value.TLabel', background='#FFFFFF')
        self.material_paper_price_label.grid(row=0, column=2, padx=(2, 14), pady=3, sticky='w')

        ttk.Label(mat, text='ДДС', background='#FFFFFF').grid(row=0, column=3, padx=(4, 5), pady=3, sticky='e')
        self.combo(mat, 'vat', ['', 'без', 'със'], '', 6).grid(row=0, column=4, padx=(0, 8), pady=3, sticky='w')

        ttk.Label(mat, text='Плаки', background='#FFFFFF').grid(row=1, column=0, padx=(7, 5), pady=3, sticky='w')
        self.combo(mat, 'plates', ['', 'да', 'не'], '', 8).grid(row=1, column=1, padx=(0, 5), pady=3, sticky='w')
        self.material_plates_price_label = ttk.Label(mat, text='—', style='Value.TLabel', background='#FFFFFF')
        self.material_plates_price_label.grid(row=1, column=2, padx=(2, 14), pady=3, sticky='w')

        ttk.Label(mat, text='Цена предпечат, €', background='#FFFFFF').grid(row=1, column=3, padx=(4, 5), pady=3, sticky='e')
        self.entry(mat, 'prepress_price', '', 8).grid(row=1, column=4, padx=(0, 8), pady=3, sticky='w')
        # Съвместимост със стария код: material_diag вече не е визуален блок.
        self.material_diag=ttk.Frame(mat)

        ttk.Separator(finish_card,orient='horizontal').grid(
            row=1,column=0,sticky='ew',padx=6,pady=(9,9))
        finish_panel=ttk.Frame(finish_card)
        finish_panel.grid(row=2,column=0,sticky='nsew',padx=2,pady=(0,2))
        finish_card.rowconfigure(2,weight=1)
        self._finish_first(finish_panel)

        # Проследяване на оскъпяването.
        def surcharge_watch(*_):
            try:
                v=self.vars['surcharge'].get().strip()
            except Exception:
                return
            if not v:
                self.surcharge_entry.configure(style='Input.TEntry')
                self.surcharge_warning.config(text='')
                return
            if v.replace(',','.') == '40':
                self.surcharge_entry.configure(style='Input.TEntry')
                self.surcharge_warning.config(text='')
            else:
                self.surcharge_entry.configure(style='Warning.TEntry')
                self.surcharge_warning.config(
                    text=f'⚠ Основният процент е 40%. Въведен е {v}%.'
                )
        self.vars['surcharge'].trace_add('write', surcharge_watch)
        surcharge_watch()


    def _refresh_print_formats(self,event=None):
        source=self.vars.get('source',tk.StringVar(value='Himiya')).get()
        values=['Автоматичен']+[o.print_format for o in repo.options_for(source)]
        if hasattr(self,'print_format_widget'):
            self.print_format_widget['values']=values
            if self.vars['print_format'].get() not in values:self.vars['print_format'].set('Автоматичен')

    def _update_optimal_suggestion(self, *args):
        from engine_himiya_v15 import find_best_sheet_logic
        try:
            source=self.vars['source'].get()
            raw_w=self.vars['w'].get().strip().replace(',','.')
            raw_h=self.vars['h'].get().strip().replace(',','.')
            if not raw_w or not raw_h:
                self.optimal_format_suggestion.config(text='MIN: —    |    MAX (Оптим.): —')
                return
            mn,mx=find_best_sheet_logic(repo,source,float(raw_w),float(raw_h),self.vars['turnover'].get(),self.vars['grip'].get()=='да')
            def show(title,choice):
                if not choice:
                    return f'{title}: няма подходящ формат'
                opt,reps,waste=choice
                return f'{title}: {opt.print_format} • {reps} бр./лист • фира {waste:.1f}%'
            self.optimal_format_suggestion.config(text=show('MIN',mn)+'    |    '+show('MAX (Оптим.)',mx))
        except Exception:
            self.optimal_format_suggestion.config(text='-')

    def _spiral_var(self,key,default=''):
        if not hasattr(self,'spiral_vars'): self.spiral_vars={}
        if key not in self.spiral_vars:
            self.spiral_vars[key]=tk.StringVar(value=default)
        return self.spiral_vars[key]

    def _sentry(self,parent,key,default='',width=12):
        return ttk.Entry(parent,textvariable=self._spiral_var(key,default),width=width,style='Input.TEntry')

    def _scombo(self,parent,key,values,default='',width=14):
        return ttk.Combobox(parent,textvariable=self._spiral_var(key,default),values=values,state='readonly',width=width,style='Input.TCombobox')

    def _spiral_price(self,key,default=0.0):
        try: return float(self.var(key,str(default)).get().replace(',','.'))
        except Exception: return float(default)

    def _spiral_size_info(self,size):
        data={
            '3/16':(4.8,2),'1/4':(6.4,3),'5/16':(7.9,4),'3/8':(9.5,4),
            '7/16':(11.1,5),'1/2':(12.7,5),'9/16':(14.3,6)
        }
        return data.get(str(size).strip(),('', ''))

    def _update_spiral_tooth_helper(self,*args):
        try:
            raw=self._spiral_var('length','').get().strip().replace(',','.')
            length=float(raw) if raw else 0
            teeth=max(0,math.floor((length-6)/8.467)) if length else 0
        except Exception:
            teeth=0
        size=self._spiral_var('size','1/2').get()
        mm,norm=self._spiral_size_info(size)
        self.spiral_diag_labels.get('teeth_calc',tk.Label()).configure(text=str(teeth) if teeth else '—')
        self.spiral_diag_labels.get('mm',tk.Label()).configure(text=f'{mm:g}' if mm else '—')
        self.spiral_diag_labels.get('sheets_per_body',tk.Label()).configure(text=str(norm) if norm else '—')
        self.spiral_diag_labels.get('sheets_per_hit',tk.Label()).configure(text=str(norm+1) if norm else '—')

    def _spirals(self, f):
        self.spiral_vars = {}; self.spiral_result = {}; self.spiral_diag_labels = {}
        # Използваме Frame вместо card, за да премахнем заглавния ред най-отгоре
        outer = ttk.Frame(f)
        outer.pack(fill='both', expand=True, padx=4, pady=4)
        outer.columnconfigure(0, weight=1); outer.columnconfigure(1, weight=1); outer.columnconfigure(2, weight=1)

        # 1. Колона 0: Основни данни
        main = self.card(outer, 'Основни данни')
        main.grid(row=0, column=0, sticky='nsew', padx=(0, 4), pady=(0, 6)); main.columnconfigure(1, weight=1)
        rows = [
            ('Общ брой изделия', 'qty', '1500'),
            ('Размер на спиралата', 'size', '1/2'),
            ('Брой зъби', 'teeth', ''),
            ('Тяло — грамаж', 'body_gsm', '115'),
            ('Брой листа', 'body_sheets', '20'),
            ('Корица — грамаж', 'cover_gsm', '250'),
            ('Брой листа', 'cover_sheets', '1'),
        ]
        sizes = ['3/16', '1/4', '5/16', '3/8', '7/16', '1/2', '9/16']
        for i, (lab, key, d) in enumerate(rows):
            ttk.Label(main, text=lab).grid(row=i, column=0, sticky='w', padx=8, pady=4)
            if key == 'size':
                w = self._scombo(main, key, sizes, d, 10)
            else:
                w = self._sentry(main, key, d, 10)
            w.grid(row=i, column=1, sticky='w', padx=8, pady=4)
            if key == 'size':
                w.bind('<<ComboboxSelected>>', self._update_spiral_tooth_helper)

        # 2. Колона 1: Довършителни / ръчни операции
        fin = self.card(outer, 'Довършителни / ръчни операции')
        fin.grid(row=0, column=1, sticky='nsew', padx=4, pady=(0, 6)); fin.columnconfigure(1, weight=1)
        self.spiral_price_labels = {}

        ops = [
            ('Спирала и перфо ×2', 'spiral_perfo_x2', ['да', 'не'], 'да', 'spiral_perfo'),
            ('За нарязване/зъб', 'cut_per_tooth', ['да', 'не'], 'не', 'cut_cost'),
            ('Кукички', 'hooks', ['без', 'до 80мм', 'до 150мм'], 'без', 'hooks_cost'),
            ('Куриер, €', 'courier', '', '', 'courier_cost'),
            ('Такса извън стандарт, €', 'outside_standard', '', '', 'outside_cost'),
            ('Увеличение %', 'increase_pct', '', '40', None),
        ]
        for i, (lab, key, vals, d, price_key) in enumerate(ops):
            ttk.Label(fin, text=lab).grid(row=i, column=0, sticky='w', padx=8, pady=4)
            w = self._scombo(fin, key, vals, d, 12) if isinstance(vals, list) else self._sentry(fin, key, d, 10)
            w.grid(row=i, column=1, sticky='w', padx=8, pady=4)
            if price_key:
                pl = ttk.Label(fin, text='—', style='Value.TLabel')
                pl.grid(row=i, column=2, sticky='w', padx=(4, 8), pady=4)
                self.spiral_price_labels[price_key] = pl

        # 3. Колона 2: Изчисляване на зъбите / норматив
        diag = self.card(outer, 'Изчисляване на зъбите / норматив')
        diag.grid(row=0, column=2, sticky='nsew', padx=(4, 0), pady=(0, 6)); diag.columnconfigure(1, weight=1)
        drows = [
            ('Дължина на изделието, мм', 'length', ''),
            ('Изчислени зъби', 'teeth_calc', '—'),
            ('Размер, мм', 'mm', '—'),
            ('Листа за 1 тяло', 'sheets_per_body', '—'),
            ('Листа за 1 перфориране', 'sheets_per_hit', '—')
        ]
        for i, (lab, key, d) in enumerate(drows):
            ttk.Label(diag, text=lab, style='Card.TLabel').grid(row=i, column=0, sticky='w', padx=8, pady=4)
            if key == 'length':
                w = self._sentry(diag, key, d, 10); w.grid(row=i, column=1, sticky='w', padx=8, pady=4); w.bind('<KeyRelease>', self._update_spiral_tooth_helper)
            else:
                w = ttk.Label(diag, text=d, style='DiagHighlight.TLabel'); w.grid(row=i, column=1, sticky='w', padx=8, pady=4); self.spiral_diag_labels[key] = w

        # 4. Долен ред: Ценообразуване
        pricing = self.card(outer, 'Ценообразуване')
        pricing.grid(row=1, column=0, columnspan=3, sticky='nsew', pady=(0, 0)); pricing.columnconfigure(1, weight=1)
        prows = [
            ('Спирала', 'spiral'),
            ('Затваряне', 'closing'),
            ('Перфо', 'perfo'),
            ('Увеличение', 'increase'),
        ]
        for r, (label_text, key) in enumerate(prows):
            ttk.Label(pricing, text=label_text, style='Card.TLabel').grid(row=r, column=0, sticky='w', padx=8, pady=3)
            w = ttk.Label(pricing, text='—', style='Value.TLabel')
            w.grid(row=r, column=1, sticky='w', padx=8, pady=3)
            self.spiral_price_labels[key] = w

        self._update_spiral_tooth_helper()
        # Пояснение за полето „Брой зъби“ при 2 отделни спирали (в долния ляв край)
        hint_label = ttk.Label(
            outer, 
            text="* Ако в „Спирала и перфо х 2“ е избрано „да“, то в „Брой зъби“ не се задава общия брой (когато се прави календар с отвор за кукичка).",
            font=('Segoe UI', 9, 'italic'),
            foreground='#667085'
        )
        # Позиционираме го в долния ляв край (под основните данни/ценообразуването)
        hint_label.grid(row=2, column=0, columnspan=2, sticky='w', padx=4, pady=(6, 2))
    def calculate_spirals(self):
        try:
            v = lambda k: self._spiral_var(k, '').get().strip()
            num = lambda k: float(v(k).replace(',', '.')) if v(k) else 0.0
            qty = num('qty'); teeth = num('teeth'); body_sheets = num('body_sheets'); cover_sheets = num('cover_sheets')
            size = v('size') or '1/2'
            tooth_price = self._spiral_price('price_spiral_tooth_' + size.replace('/', '_'), 0.0)
            spiral = tooth_price * teeth * qty * 1.8 * 1.15
            closing = math.ceil((qty + 20) * 0.04)
            _, norm = self._spiral_size_info(size); sheets_per_hit = (norm + 1) if norm else 1
            total_sheets = body_sheets + cover_sheets
            hits = (qty * total_sheets) / sheets_per_hit if sheets_per_hit else 0
            perfo = round((hits * 0.011 * 1.2) * 10) / 10 if qty else 0
            spiral_perfo = (spiral + perfo) if v('spiral_perfo_x2').lower() == 'да' else 0
            hooks_type = v('hooks'); hook_rate = {'до 80мм': 0.04, 'до 150мм': 0.051}.get(hooks_type, 0)
            hook_inc = 0.35
            hooks_cost = hook_rate * qty * hook_inc if hook_rate else 0
            courier = num('courier'); cut_cost = 0.0031 * qty * teeth if v('cut_per_tooth').lower() == 'да' else 0
            outside = num('outside_standard')
            increase_pct = num('increase_pct') or 40
            increase = math.ceil((closing + cut_cost) * increase_pct / 100)
            total = math.ceil(sum([spiral, closing, perfo, spiral_perfo, hooks_cost, courier, cut_cost, increase, outside]))
            unit = total / qty if qty else 0

            self.request_source_tab='Спирали'
            self.spiral_result = {
                'qty': qty, 'size': size, 'teeth': teeth, 'total': total, 'unit': unit,
                'spiral': spiral, 'closing': closing, 'perfo': perfo, 'spiral_perfo': spiral_perfo,
                'hooks_cost': hooks_cost, 'courier_cost': courier, 'cut_cost': cut_cost,
                'increase': increase, 'outside_cost': outside
            }
            self._update_top_bar()
            self._refresh_request_offer()

            # Обновяване на етикетите: при 0 слагаме "—", а при сума > 0 показваме "X.XX €"
            for k, val in self.spiral_result.items():
                if k in self.spiral_price_labels:
                    val_float = float(val or 0)
                    display_text = f'{val_float:.2f} €' if val_float > 0 else '—'
                    self.spiral_price_labels[k].configure(text=display_text)

            self.top_total_var.set(f'Крайна цена: {total:.2f} €')
            self.top_unit_var.set(f'{unit:.4f} € / бр.' if qty else '')
            self._last_valid = True; self._update_top_bar()
        except Exception as e:
            self._last_valid = False; messagebox.showerror('Грешка при изчислението', str(e))
    def _reset_spirals(self):
        defaults = {
            'qty': '1500', 'size': '1/2', 'teeth': '', 'body_gsm': '115',
            'body_sheets': '20', 'cover_gsm': '250', 'cover_sheets': '1',
            'length': '', 'spiral_perfo_x2': 'да', 'cut_per_tooth': 'не',
            'hooks': 'без', 'courier': '', 'outside_standard': '', 'increase_pct': '40'
        }
        for k, d in defaults.items():
            self._spiral_var(k, d).set(d)
        self.spiral_result = {}
        for w in getattr(self, 'spiral_price_labels', {}).values():
            w.configure(text='—')
        self._update_spiral_tooth_helper()
        self._update_top_bar()
    def _cvar(self, key, default=''):
        if not hasattr(self, 'calendar_vars'):
            self.calendar_vars = {}
        if key not in self.calendar_vars:
            self.calendar_vars[key] = tk.StringVar(value=default)
        return self.calendar_vars[key]

    def _centry(self, parent, key, default='', width=12):
        return ttk.Entry(parent, textvariable=self._cvar(key, default), width=width, style='Input.TEntry')

    def _ccombo(self, parent, key, values, default='', width=15):
        return ttk.Combobox(parent, textvariable=self._cvar(key, default), values=values,
                            state='readonly', width=width, style='Input.TCombobox')

    def _calendar_sheet_count(self, fmt):
        """Lookup of source-sheet count from workbook table таб_Цял_Лист."""
        try:
            wb = openpyxl.load_workbook(WORKBOOK, data_only=True, keep_vba=True)
            # The table is on sheet формати in the current workbook. Find matching
            # print-format and return the corresponding whole-sheet count.
            ws = wb['формати']
            target = str(fmt).strip().lower().replace('х','x').replace('×','x')
            for row in ws.iter_rows(min_row=1, values_only=True):
                vals=[str(v).strip() if v is not None else '' for v in row]
                if any(v.lower().replace('х','x').replace('×','x') == target for v in vals):
                    for v in row:
                        try:
                            n=float(v)
                            if n>0 and n.is_integer():
                                # avoid returning the format dimension itself
                                if n not in (60,64,70,84,90,100,120,140):
                                    return int(n)
                        except Exception: pass
            wb.close()
        except Exception:
            pass
        # Known workbook whole-sheet counts used by the print formats.
        known={'a3':8,'a4':16,'a5':32,'a6':64,'f.60x84':1,'f.64x90':1,'f.70x100':1}
        return known.get(str(fmt).strip().lower(), 1)

    def _calendar_layout_repetitions(self):
        """Автоматичен брой размножения според необрязания размер и печатния формат.

        Печатният формат от таблицата е в cm, а необрязаният размер се въвежда в mm.
        Проверяваме и двете ориентации и вземаме максималния брой изделия,
        които се побират върху един печатен лист — същата логика като в „Книжки“.
        """
        try:
            pf=str(self._cvar('print_format','').get()).strip()
            import re
            m=re.search(r'(\d+(?:[.,]\d+)?)\s*[xх×]\s*(\d+(?:[.,]\d+)?)',pf.lower())
            if not m:
                return 0
            pw=float(m.group(1).replace(',','.'))*10.0
            ph=float(m.group(2).replace(',','.'))*10.0
            tw=float(self._cvar('trim_w','').get().strip().replace(',','.'))
            th=float(self._cvar('trim_h','').get().strip().replace(',','.'))
            if pw<=0 or ph<=0 or tw<=0 or th<=0:
                return 0
            best=0
            for sw,sh in ((pw,ph),(ph,pw)):
                best=max(best,int(sw//tw)*int(sh//th))
            return max(0,best)
        except Exception:
            return 0

    def _sync_calendar_repetitions(self, event=None):
        if not hasattr(self,'calendar_vars'):
            return
        mode=self._cvar('repetition_mode','Автоматично').get() or 'Автоматично'
        manual=self._cvar('repetitions_manual','2')
        if mode=='Ръчно':
            try:
                self._calendar_repetitions_manual.configure(state='normal')
            except Exception:
                pass
            try:
                reps=float(manual.get().strip().replace(',','.'))
            except Exception:
                reps=0
        else:
            try:
                self._calendar_repetitions_manual.configure(state='disabled')
            except Exception:
                pass
            reps=self._calendar_layout_repetitions()
            if reps>0:
                manual.set(str(reps).rstrip('0').rstrip('.') if isinstance(reps,float) else str(reps))
        self._cvar('repetitions','').set(str(reps) if reps else '')

    def _calendar_recalc_diag(self, *args):
        try:
            # При автоматичен режим размноженията винаги следват размера и печатния формат.
            mode=self._cvar('repetition_mode','Автоматично').get() or 'Автоматично'
            if mode=='Автоматично' and not getattr(self,'_calendar_syncing_repetitions',False):
                self._calendar_syncing_repetitions=True
                try:
                    self._sync_calendar_repetitions()
                finally:
                    self._calendar_syncing_repetitions=False
            pages=float(self._cvar('pages','').get().replace(',','.')) if self._cvar('pages','').get() else 0
            reps=float(self._cvar('repetitions','').get().replace(',','.')) if self._cvar('repetitions','').get() else 1
            qty=float(self._cvar('qty','').get().replace(',','.')) if self._cvar('qty','').get() else 0
            cf=float(self._cvar('color_front','').get() or 0)
            cb=float(self._cvar('color_back','').get() or 0)
            cols=max(1, pages/reps) if reps else 1
            clean=qty
            self.calendar_diag_labels.get('cols',tk.Label()).configure(text=self._fmt_count(cols) if cols else '—')
            self.calendar_diag_labels.get('clean',tk.Label()).configure(text=self._fmt_count(clean) if clean else '—')
            self.calendar_diag_labels.get('colors',tk.Label()).configure(text=f'{int(cf) if cf.is_integer() else cf:g}+{int(cb) if cb.is_integer() else cb:g}')
            if hasattr(self,'calendar_scheme_canvas'):
                self._draw_calendar_scheme()
        except Exception:
            pass

    def _calendars(self, f):
        self.calendar_vars={}; self.calendar_result={}; self.calendar_diag_labels={}; self.calendar_price_labels={}
        outer=ttk.Frame(f); outer.pack(fill='both',expand=True,padx=4,pady=4)

        # Подредба по референтния екран:
        # ред 0: Основни | Диагностика | Довършителни
        # ред 1: Материали | Материали | Схема
        # ред 2: Ценообразуване на печата | Ценообразуване на печата | Схема
        for c in range(3):
            outer.columnconfigure(c, weight=1)
        outer.rowconfigure(0, weight=0)
        outer.rowconfigure(1, weight=0)
        outer.rowconfigure(2, weight=1)

        main=self.card(outer,'Основни данни / Печат')
        main.grid(row=0,column=0,sticky='nsew',padx=(0,4),pady=(0,6))
        main.columnconfigure(1,weight=1)

        ttk.Label(main,text='Формат на хартията').grid(row=0,column=0,padx=(8,8),pady=3,sticky='w')
        try:
            source_values=[str(x) for x in repo.sources()]
        except Exception:
            source_values=['a4','a5','a6']
        source_values=[x for x in source_values if str(x).strip().lower() != 'a3']
        source_values=list(dict.fromkeys(source_values))
        source_default='a4' if 'a4' in source_values else (source_values[0] if source_values else '')
        cb=self._ccombo(main,'source_format',source_values,source_default,14)
        cb.grid(row=0,column=1,padx=(0,8),pady=3,sticky='w')
        cb.bind('<<ComboboxSelected>>',self._calendar_update_print_formats)

        ttk.Label(main,text='Формат за печат').grid(row=1,column=0,padx=(8,8),pady=3,sticky='w')
        self.calendar_print_combo=self._ccombo(main,'print_format',[],'',14)
        self.calendar_print_combo.grid(row=1,column=1,padx=(0,8),pady=3,sticky='w')

        ttk.Label(main,text='Необязан размер').grid(row=2,column=0,padx=(8,8),pady=3,sticky='w')
        trimf=ttk.Frame(main)
        trimf.grid(row=2,column=1,padx=(0,8),pady=3,sticky='w')
        self._centry(trimf,'trim_h','210',7).pack(side='left')
        ttk.Label(trimf,text=' x ',padding=(2,0)).pack(side='left')
        self._centry(trimf,'trim_w','297',7).pack(side='left')
        ttk.Label(trimf,text=' мм',padding=(4,0)).pack(side='left')

        ttk.Label(main,text='Цветност').grid(row=3,column=0,padx=(8,8),pady=3,sticky='w')
        cf_frame=tk.Frame(main,bg='white'); cf_frame.grid(row=3,column=1,padx=(0,8),pady=3,sticky='w')
        self._ccombo(cf_frame,'color_front',[str(i) for i in range(1,9)],'1',4).pack(side='left')
        ttk.Label(cf_frame,text=' + ',padding=(2,0)).pack(side='left')
        self._ccombo(cf_frame,'color_back',[str(i) for i in range(0,9)],'0',4).pack(side='left')

        ttk.Label(main,text='Единични бройки').grid(row=4,column=0,padx=(8,8),pady=3,sticky='w')
        self._centry(main,'qty','1000',10).grid(row=4,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(main,text='Брой страници').grid(row=5,column=0,padx=(8,8),pady=3,sticky='w')
        self._centry(main,'pages','1',10).grid(row=5,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(main,text='Обръщане').grid(row=6,column=0,padx=(8,8),pady=3,sticky='w')
        self._ccombo(main,'turnover',['не','да','½ с обръщ.','¼ с обръщ.'],'не',14).grid(row=6,column=1,padx=(0,8),pady=3,sticky='w')
        ttk.Label(main,text='Размножения').grid(row=7,column=0,padx=(8,8),pady=3,sticky='w')
        rf=tk.Frame(main,bg='white')
        rf.grid(row=7,column=1,padx=(0,8),pady=3,sticky='w')
        self._calendar_repetition_mode=ttk.Combobox(
            rf,textvariable=self._cvar('repetition_mode','Автоматично'),
            values=['Автоматично','Ръчно'],state='readonly',width=11,
            style='BookRepetition.TCombobox')
        self._calendar_repetition_mode.pack(side='left')
        self._calendar_repetitions_manual=self._centry(rf,'repetitions_manual','2',7)
        self._calendar_repetitions_manual.configure(style='BookRepetition.TEntry')
        self._calendar_repetitions_manual.pack(side='left',padx=(6,0))
        self._calendar_repetition_mode.bind('<<ComboboxSelected>>',self._sync_calendar_repetitions)

        ttk.Label(main,text='Смяна на цвят').grid(row=8,column=0,padx=(8,8),pady=3,sticky='w')
        cc=tk.Frame(main,bg='white'); cc.grid(row=8,column=1,padx=(0,8),pady=3,sticky='w')
        self._ccombo(cc,'color_change_opt',['не','да'],'не',7).pack(side='left')
        ttk.Label(cc,text=' брой').pack(side='left',padx=(5,3))
        self._centry(cc,'color_change_count','2',7).pack(side='left')
        self._calendar_update_print_formats()
        self._sync_calendar_repetitions()

        diag=self.card(outer,'Диагностика — печат и тираж')
        diag.grid(row=0,column=1,sticky='nsew',padx=4,pady=(0,6))
        diag.columnconfigure(1,weight=1)
        drows=[
            ('Печатни коли','cols'),('Чист тираж','clean'),
            ('Тираж + макулатура','waste'),('Цели листа','whole'),
            ('Печатен формат','print'),('Цветност','colors')
        ]
        for r,(lab,key) in enumerate(drows):
            ttk.Label(diag,text=lab,style='Card.TLabel').grid(row=r,column=0,sticky='w',padx=7,pady=3)
            st='DiagHighlight.TLabel' if key in ('print','clean','whole') else 'Value.TLabel'
            w=ttk.Label(diag,text='—',style=st)
            w.grid(row=r,column=1,sticky='w',padx=7,pady=3)
            self.calendar_diag_labels[key]=w

        # Довършителни са в две вътрешни колони, както на референтния екран.
        finish=self.card(outer,'Довършителни')
        finish.grid(row=0,column=2,rowspan=2,sticky='nsew',padx=(4,0),pady=(0,6))
        for c in range(6):
            finish.columnconfigure(c,weight=1 if c in (1,4) else 0)

        left_finish=[
            ('3 стр. обрязване/рязане','cut_opt',['стандарт','обряз.3стр.','без'],'стандарт','cutting'),
            ('Шиене','sewing_opt',['без','1','2','3','4'],'без','sewing'),
            ('Биговане','bigoving_opt',['без','1','2','3','4','5','6'],'без','bigoving'),
            ('Набор','typesetting_opt',['машинно','ръчно','не','календ.'],'не','typesetting'),
            ('Пакетиране','packaging_opt',['не','да','<1000'],'не','packaging'),
        ]
        right_finish=[
            ('Лепене','gluing_opt',['без','да'],'без','gluing'),
            ('Разрязване','cut2_opt',['без','1','2','3','4','0'],'без','cutting2'),
            ('Леп. гръб','spine_opt',['без','1','2','3','4','0'],'без','spine'),
            ('Прозорци','windows_opt',['без','да'],'без','windows'),
            ('Ел. монтаж','electric_opt',['без','бошура/покана','етикети/визитки','корици','листовки/стикери','минимално','плакат','страниране','флаери'],'без','electric'),
        ]
        self.calendar_finish_price_labels={}
        for r,(lab,key,vals,d,price_key) in enumerate(left_finish):
            ttk.Label(finish,text=lab).grid(row=r,column=0,sticky='w',padx=(7,4),pady=3)
            self._ccombo(finish,key,vals,d,14).grid(row=r,column=1,sticky='w',padx=(0,4),pady=3)
            price_lbl=ttk.Label(finish,text='0,00 €',style='Value.TLabel',width=10,anchor='w')
            price_lbl.grid(row=r,column=2,sticky='w',padx=(0,7),pady=3)
            self.calendar_finish_price_labels[price_key]=price_lbl
        for r,(lab,key,vals,d,price_key) in enumerate(right_finish):
            ttk.Label(finish,text=lab).grid(row=r,column=3,sticky='w',padx=(7,4),pady=3)
            self._ccombo(finish,key,vals,d,14).grid(row=r,column=4,sticky='w',padx=(0,4),pady=3)
            price_lbl=ttk.Label(finish,text='0,00 €',style='Value.TLabel',width=10,anchor='w')
            price_lbl.grid(row=r,column=5,sticky='w',padx=(0,7),pady=3)
            self.calendar_finish_price_labels[price_key]=price_lbl

        sep_row=5
        ttk.Label(finish,text='Разделители').grid(row=sep_row,column=0,sticky='w',padx=(7,4),pady=3)
        self._ccombo(finish,'sep_material',['без','вестник','друг'],'без',10).grid(row=sep_row,column=1,sticky='w',padx=(0,4),pady=3)
        sep_price=ttk.Label(finish,text='0л',style='Value.TLabel',width=12,anchor='w')
        sep_price.grid(row=sep_row,column=2,sticky='w',padx=(0,7),pady=3)
        self.calendar_finish_price_labels['separators']=sep_price

        ttk.Label(finish,text='на листа').grid(row=sep_row,column=3,sticky='w',padx=(7,4),pady=3)
        self._centry(finish,'sep_count','100',7).grid(row=sep_row,column=4,sticky='w',padx=(0,4),pady=3)
        sep_right_price=ttk.Label(finish,text='0,00 €',style='Value.TLabel',width=10,anchor='w')
        sep_right_price.grid(row=sep_row,column=5,sticky='w',padx=(0,7),pady=3)
        self.calendar_separator_total_label=sep_right_price

        # РЕД 6: Други разходи (вляво) и Транспорт (вдясно, точно под "на листа")
        other_row=6
        ttk.Label(finish,text='Други разходи, €').grid(row=other_row,column=0,sticky='w',padx=(7,4),pady=3)
        self._centry(finish,'other_price','',10).grid(row=other_row,column=1,sticky='w',padx=(0,4),pady=3)

        ttk.Label(finish,text='Транспорт, €').grid(row=other_row,column=3,sticky='w',padx=(7,4),pady=3)
        self._centry(finish,'transport_price','',10).grid(row=other_row,column=4,sticky='w',padx=(0,4),pady=3)

        # РЕД 7: Оскъпяване остава само на последния ред
        surcharge_row=7
        ttk.Label(finish,text='Оскъпяване %').grid(row=surcharge_row,column=0,sticky='w',padx=(7,4),pady=3)
        self._centry(finish,'surcharge','40',10).grid(row=surcharge_row,column=1,sticky='w',padx=(0,4),pady=3)
        self.calendar_finish_price_labels['surcharge']=ttk.Label(finish,text='0,00 €',style='Value.TLabel',width=10,anchor='w')
        self.calendar_finish_price_labels['surcharge'].grid(row=surcharge_row,column=2,sticky='w',padx=(0,7),pady=3)

        # Материали — напълно възстановени и перфектно подравнени в две колони
        materials=self.card(outer,'Материали')
        materials.grid(row=1,column=0,columnspan=2,sticky='nsew',padx=(0,4),pady=(0,6))
        
        # Настройваме 3 колони с фиксирани тегла за безупречно вертикално подравняване
        materials.columnconfigure(0, weight=0, minsize=180)  # За текстовите етикети вляво
        materials.columnconfigure(1, weight=0, minsize=160)  # За първата колона с елементи (полета и бройки)
        materials.columnconfigure(2, weight=1)              # За втората колона (ДДС менюто и сумата за плаки)

        self.calendar_material_labels={}

        # РЕД 0: Хартия (Етикет, Поле за писане, Изчислен Резултат) и ДДС (Етикет, Меню за избор)
        ttk.Label(materials,text='Хартия — цена / лист, €').grid(row=0,column=0,sticky='w',padx=7,pady=2)
        
        paper_frame = ttk.Frame(materials)
        paper_frame.grid(row=0,column=1,sticky='w',padx=7,pady=2)
        self._centry(paper_frame,'paper_price','',8).pack(side='left')
        self.calendar_material_labels['paper_total']=ttk.Label(paper_frame,text='—',style='Value.TLabel')
        self.calendar_material_labels['paper_total'].pack(side='left',padx=(8,0))
        
        vat_frame = ttk.Frame(materials)
        vat_frame.grid(row=0,column=2,sticky='w',padx=7,pady=2)
        ttk.Label(vat_frame,text='ДДС:').pack(side='left',padx=(0,4))
        self._ccombo(vat_frame,'vat',['без','със'],'без',8).pack(side='left')

        # РЕД 1: Плаки (Бройката е в кол. 1, а Общата Сума отива в кол. 2)
        ttk.Label(materials,text='Плаки').grid(row=1,column=0,sticky='w',padx=7,pady=2)
        
        plate_count_frame = ttk.Frame(materials)
        plate_count_frame.grid(row=1,column=1,sticky='w',padx=7,pady=2)
        self.calendar_material_labels['plate_count']=ttk.Label(plate_count_frame,text='—',style='Value.TLabel')
        self.calendar_material_labels['plate_count'].pack(side='left')
        ttk.Label(plate_count_frame,text=' бр.').pack(side='left')
        
        # Общата сума за плаки застава абсолютно самостоятелно в колона 2 - точно под ДДС менюто
        self.calendar_material_labels['plate_total']=ttk.Label(materials,text='—',style='Value.TLabel')
        self.calendar_material_labels['plate_total'].grid(row=1,column=2,sticky='w',padx=45,pady=2)

        # РЕД 2: Предпечат (Полето за въвеждане е подравнено под това за хартията)
        ttk.Label(materials,text='Предпечат, €').grid(row=2,column=0,sticky='w',padx=7,pady=2)
        self._centry(materials,'prepress_price','',10).grid(row=2,column=1,sticky='w',padx=7,pady=2)

        self._cvar('plates_opt','да')

        # Производствената калкулация е долу вляво; схемата заема дясната колона.
        prod=self.card(outer,'Ценообразуване на печата')
        prod.grid(row=2,column=0,columnspan=2,sticky='nsew',padx=(0,4),pady=(0,0))
        for c in range(6):
            prod.columnconfigure(c,weight=1 if c in (1,3,5) else 0)

        prows=[
            ('Печат лице/гръб','print'),('Обръщане','turn'),
            ('Печат над 1000','over1000'),('Смяна на цвят','color_change'),
            ('Оскъпяване','surcharge'),('Печалба','profit'),('Разходи','expenses')
        ]
        # В референтната подредба основните пера са вертикално вляво,
        # а Печалба и Разходи са на последния ред.
        for i,(lab,key) in enumerate(prows[:5]):
            ttk.Label(prod,text=lab,font=('Segoe UI',10,'bold') if key=='surcharge' else None).grid(row=i,column=0,sticky='w',padx=(16,6),pady=2)
            w=ttk.Label(prod,text='—',style='Value.TLabel',font=('Segoe UI',10,'bold') if key=='surcharge' else None)
            w.grid(row=i,column=1,sticky='w',padx=6,pady=2)
            self.calendar_price_labels[key]=w
        ttk.Label(prod,text='Печалба',font=('Segoe UI',10,'bold')).grid(row=5,column=0,sticky='w',padx=(16,6),pady=2)
        self.calendar_price_labels['profit']=ttk.Label(prod,text='—',style='Value.TLabel',font=('Segoe UI',10,'bold'))
        self.calendar_price_labels['profit'].grid(row=5,column=1,sticky='w',padx=6,pady=2)
        ttk.Label(prod,text='Разходи',font=('Segoe UI',10,'bold')).grid(row=5,column=2,sticky='w',padx=(16,6),pady=2)
        self.calendar_price_labels['expenses']=ttk.Label(prod,text='—',style='Value.TLabel',font=('Segoe UI',10,'bold'))
        self.calendar_price_labels['expenses'].grid(row=5,column=3,sticky='w',padx=6,pady=2)

        scheme=self.card(outer,'Схема на разполагане')
        scheme.grid(row=1,column=2,rowspan=2,sticky='nsew',padx=(4,0),pady=(0,0))
        scheme.columnconfigure(0,weight=1); scheme.rowconfigure(0,weight=1)
        self.calendar_scheme_canvas=tk.Canvas(scheme,bg='white',highlightthickness=1,highlightbackground='#D0D0D0')
        self.calendar_scheme_canvas.grid(row=0,column=0,sticky='nsew',padx=4,pady=4)
        self.calendar_scheme_canvas.bind('<Configure>', lambda e: self._draw_calendar_scheme())
        self._calendar_scheme_data=None

        for v in self.calendar_vars.values():
            try: v.trace_add('write',self._calendar_recalc_diag)
            except Exception: pass
        self._calendar_recalc_diag()

    def _draw_calendar_scheme(self):
        """Схема: избраният печатен формат е работната площ, а необрязаният размер
        се размножава върху нея. Около печатния формат показваме технологичен лист.
        """
        c=getattr(self,'calendar_scheme_canvas',None)
        if c is None:
            return
        c.delete('all')
        try:
            import re
            pf=str(self._cvar('print_format','').get() or '').strip()
            m=re.search(r'(\d+(?:[.,]\d+)?)\s*[xх×]\s*(\d+(?:[.,]\d+)?)',pf.lower())
            if not m:
                raise ValueError
            pw=float(m.group(1).replace(',','.'))
            ph=float(m.group(2).replace(',','.'))
            tw=float(self._cvar('trim_w','').get().strip().replace(',','.'))/10.0
            th=float(self._cvar('trim_h','').get().strip().replace(',','.'))/10.0
            mode=self._cvar('repetition_mode','Автоматично').get() or 'Автоматично'
            if mode=='Автоматично':
                reps=self._calendar_layout_repetitions()
            else:
                reps=float(self._cvar('repetitions_manual','').get().strip().replace(',','.'))
            reps=max(0,reps)
            turnover=str(self._cvar('turnover','не').get() or 'не')
            if pw<=0 or ph<=0 or tw<=0 or th<=0:
                raise ValueError

            # Технологичният лист е с 3 cm допълнително по ширина и 2 cm по височина,
            # както в референтната схема. Самият избран печатен формат е вътре.
            sheet_w=pw+3.0
            sheet_h=ph+2.0
            cw=max(c.winfo_width(),180); ch=max(c.winfo_height(),160)
            margin_x=26; margin_y=34
            scale=min((cw-2*margin_x)/sheet_w,(ch-2*margin_y)/sheet_h)
            if scale<=0:
                return
            sw,sh=sheet_w*scale,sheet_h*scale
            x0=(cw-sw)/2; y0=margin_y+(ch-2*margin_y-sh)/2
            c.create_rectangle(x0,y0,x0+sw,y0+sh,outline='#444444',width=2)

            # Работната площ на печатния формат е центрирана в технологичния лист.
            px=x0+1.5*scale; py=y0+1.0*scale
            c.create_rectangle(px,py,px+pw*scale,py+ph*scale,outline='#777777',width=1)

            best=None
            for rw,rh in ((tw,th),(th,tw)):
                nx=int(pw//rw) if rw else 0
                ny=int(ph//rh) if rh else 0
                n=nx*ny
                if n>0 and (best is None or n>best[0]):
                    best=(n,nx,ny,rw,rh)
            shown=0
            if best and reps>0:
                _,nx,ny,rw,rh=best
                for iy in range(ny):
                    for ix in range(nx):
                        if shown>=reps:
                            break
                        x=px+ix*rw*scale; y=py+iy*rh*scale
                        c.create_rectangle(x,y,x+rw*scale,y+rh*scale,fill='#DDEBF7',outline='#2B579A',width=1)
                        shown+=1
                        if rw*scale>24 and rh*scale>16:
                            c.create_text(x+rw*scale/2,y+rh*scale/2,text=str(shown),fill='#1F1F1F',font=('Segoe UI',8,'bold'))
                    if shown>=reps:
                        break

            c.create_text(cw/2,8,anchor='n',text=f'Печатен формат: {pw:g} × {ph:g} cm',fill='#444444',font=('Segoe UI',9,'bold'))
            c.create_text(cw/2,ch-8,anchor='s',text=f'{reps:g} бр.  •  обръщане: {turnover.upper()}',fill='#555555',font=('Segoe UI',9))
        except Exception:
            c.create_text(20,20,anchor='nw',text='Въведете необрязан размер и изберете формат за печат.',fill='#777777',font=('Segoe UI',10))

    def _calendar_update_print_formats(self,*args):
        source=self._cvar('source_format','a4').get()
        opts=[]
        try:
            opts=[o.print_format for o in repo.options_for(source)]
        except Exception: pass
        if not opts:
            opts=['a4','a5','a6'] if source in ('a3','a4','a5','a6') else [source]
        cb=getattr(self,'calendar_print_combo',None)
        # Find the widget by stored reference or create it in the main card is not practical;
        # the source-format change is therefore handled by the dedicated stored combo.
        if cb is not None:
            cb.configure(values=list(dict.fromkeys(opts)))
            if self._cvar('print_format','').get() not in opts: self._cvar('print_format',opts[0] if opts else '').set(opts[0] if opts else '')
        else:
            # locate through calendar widgets created above
            pass

    def calculate_calendars(self):
        """Календари — изчисления по формулите от Excel, дадени за „Диагностика — печат и тираж“.
        Мапингът към интерфейса е: C6=цвета лице, C7=цвета гръб, C8=тираж,
        C9=страници/печатни коли, C10=обръщане, C11=размножения,
        C12=изчислени печатни коли, C13=тираж.
        """
        try:
            v=lambda k:self._cvar(k,'').get().strip()
            num=lambda k:float(v(k).replace(',','.')) if v(k) else 0.0
            source=v('source_format') or 'a4'
            pf=v('print_format')
            try:
                trim_w=num('trim_w'); trim_h=num('trim_h')
                trim=f'{trim_w:g}×{trim_h:g} мм'
            except Exception:
                trim=''
            cf=int(num('color_front')); cb=int(num('color_back'))
            qty=num('qty')          # Excel C8/C13
            pages=num('pages')       # Excel C9
            turnover=v('turnover') or 'не'
            reps=num('repetitions')  # Excel C11
            if reps <= 0: reps=1

            # C12 = IFERROR(IF(A12<1,1,C9/C11),"") ; A12=C9/C11/2
            raw_cols=pages/reps
            cols=1 if raw_cols < 1 else raw_cols
            colors=cf+cb

            # C14 — тиражни листа от изходния формат.
            sheet_count=self._calendar_sheet_count(source)
            # C15 = IF(AND(C6<=4,C8<=1000),C13+25*C6,C13+40*(C6+C7))
            waste=qty+25*cf if (cf<=4 and qty<=1000) else qty+40*(cf+cb)
            # C16 = MROUND(C15/C14*C12,10)
            whole=self._calendar_mround(waste/sheet_count*cols,10) if sheet_count else 0

            # G3 — хартия. Цената за лист е задължително текущо въведена.
            # Не използваме никаква предишна стойност: четем директно полето
            # paper_price от текущия tab „Календари“.
            paper_price_text=v('paper_price')
            if not paper_price_text:
                self._last_valid=False
                self.calendar_result={}
                try:
                    self.calendar_material_labels['paper_total'].configure(text='—')
                except Exception:
                    pass
                messagebox.showwarning('Календари', 'Моля, въведете цена за лист на хартия.')
                return
            try:
                paper_price=float(paper_price_text.replace(',','.'))
            except Exception:
                self._last_valid=False
                self.calendar_result={}
                messagebox.showwarning('Календари', 'Моля, въведете валидна цена за лист на хартия.')
                return
            if v('vat')=='със':
                paper=whole*paper_price*1.2
            elif v('vat')=='без':
                paper=whole*(paper_price+0.06)
            else:
                paper=0

            # G4 — печат лице/гръб.
            print_price=self._spiral_price('price_print_g4',8.0)
            rounded_cols=self._calendar_mround(cols,0.25)
            if turnover=='да':
                printing=print_price*cols
            elif turnover=='не':
                printing=print_price*cols*colors
            elif turnover in ('½ с обръщ.','¼ с обръщ.'):
                printing=rounded_cols*print_price*colors
            elif turnover=='¾':
                printing=print_price*math.ceil(cols)*colors
            else:
                printing=0

            # I8 — брой плаки/печатни единици според дадената формула.
            if turnover=='да':
                plate_units=pages/reps
            elif turnover in ('не','½ с обръщ.','пантон'):
                plate_units=colors*cols
            elif turnover=='¼ с обръщ.':
                plate_units=math.ceil(pages/reps)
            else:
                plate_units=0
            # G5 — обръщане.
            plate_base=2.7*cb if turnover in ('да','½ с обръщ.','¼ с обръщ.') else 0
            clean_turn=math.floor(qty/100)*100
            coef={'да':1,'½ с обръщ.':0.5,'¼ с обръщ.':0.25}.get(turnover,0)
            corrected=clean_turn*coef
            turn_calc=plate_base if qty<=1000 else (corrected/1000)*plate_base
            turn=0 if turnover=='не' else self._calendar_mround(turn_calc,2.7)

            # G6 — печат над 1000.
            over=0 if qty<=1000 else math.ceil((qty-1000)/1000)*2.7*cols*cf

            # G7 — смяна на цвят.
            color_change=20.45*num('color_change_count') if v('color_change_opt')=='да' else 0

            # G8 — плаки.
            plates=2.8*plate_units if plate_units>0 else 0
            plate_count=int(math.ceil(plate_units)) if plate_units>0 else 0

            # G9 — рязане/3-стр. обрязване.
            cut_type=v('cut_opt')
            base_labour=1.02*cols if qty<=1000 else 0.1*(qty/100)*cols
            cut_3=self._calendar_mround((7/1000*1.53*pages)+base_labour,1)
            if cut_type=='стандарт': cut=base_labour
            elif cut_type=='обряз.3стр.': cut=cut_3
            else: cut=0

            # Биговане — броят бигове идва от H/полето за довършителна операция.
            # Не използваме несъществуващата променлива `big`, преди да е изчислена.
            try:
                big_n=float(v('bigoving_opt').replace(',','.'))
            except Exception:
                big_n=0
            big=(qty+40)*0.0051*big_n*(pages/2) if big_n>0 else 0

            # G10 няма отделен ред в интерфейса.
            prepress=num('prepress_price')
            other=num('other_price')
            transport=num('transport_price')

            # G11 — шиене/телчета: (C8+15)*H11*0.005.
            sew_type=v('sewing_opt')
            try: sew_factor=float(sew_type.replace(',','.'))
            except Exception: sew_factor=0
            sewing=math.ceil((qty+15)*sew_factor*0.005) if sew_factor>0 else 0

            # G12 — набор. Интерфейсът няма отделно числово H12; използваме броя коли
            # от полето за набор, когато е числово, иначе 0.
            try: ts_factor=float(v('typesetting_opt').replace(',','.'))
            except Exception: ts_factor=0
            typeset=(qty+40)*cols*0.0051*ts_factor if ts_factor>0 else 0

            # G13 — сгъване/операция според текущия избор.
            ts=v('typesetting_opt')
            if ts in ('не','без',''):
                fold=0
            elif ts=='машинно':
                fold=math.ceil(pages/1000*3*cols)
            elif ts=='ръчно':
                fold=math.ceil(0.00356*pages*cols)
            elif ts in ('календ.','календари'):
                fold=math.ceil(0.00175*pages*cols*reps)
            else:
                fold=0

            # Останалите текущи полета.
            glue=qty*0.0075 if v('gluing_opt')=='да' else 0
            cut2=qty*float(v('cut2_opt'))*0.0052 if v('cut2_opt') not in ('без','0','') else 0
            spine=qty*float(v('spine_opt'))*0.0125 if v('spine_opt') not in ('без','0','') else 0
            windows=qty*0.0125 if v('windows_opt')=='да' else 0
            packaging=(math.floor(pages*cols/100)*0.77 if pages*cols>=500 else math.floor(qty/1000)*2.56) if v('packaging_opt')=='да' else (math.floor(pages/100)*2.56 if v('packaging_opt')=='<1000' else 0)
            electric=0

            # G19 — монтаж/други. G19/I19/J19 остава по текущата структура.
            other_cost=other
            # „Транспорт“ е отделен от „Други“ и е самостоятелен директен разход.
            transport_cost=transport

            # J21/G21 — разделители.
            sep_mat=v('sep_material')
            sep_every=max(0,int(num('sep_count')))
            sep_sheets=0; sep_unit=0.0; sep=0.0
            if sep_mat!='без' and sep_every>0 and qty>0:
                sep_sheets=math.ceil(((qty+100)/sep_every)*reps)
                sep_unit=self._spiral_price('price_sep_news' if sep_mat=='вестник' else 'price_sep_other',0.028 if sep_mat=='вестник' else 0.30)
                sep=math.ceil(sep_sheets*sep_unit*2)/2

            # G22 — оскъпяване; G21 (разделители) не влиза в базата.
            surcharge_pct=num('surcharge') or 40
            surcharge_base=sum([printing,turn,over,color_change,cut,prepress,other_cost,sewing,typeset,fold,glue,cut2,spine,windows,packaging,electric])
            surcharge=math.ceil(surcharge_base)*surcharge_pct/100 if surcharge_pct>0 else 0

            # G24/G25/G26/G27.
            total=math.ceil((paper+printing+turn+over+color_change+plates+cut+prepress+other_cost+transport_cost+sewing+typeset+fold+glue+cut2+spine+windows+packaging+electric+sep+surcharge)*10)/10
            unit=total/qty if qty else 0
            expenses=paper+plate_units*2.45+sep+transport_cost
            profit=total-expenses

            self.request_source_tab='Календари'
            self.calendar_result={'qty':qty,'total':total,'unit':unit,'source_format':source,'print_format':pf,'trim_size':trim,'color_front':cf,'color_back':cb,'turnover':turnover,'pages':pages,'repetitions':reps,
                # Допълнителни полета само за представянето в „Заявка“ — не участват
                # в изчисленията и запазват вече изчислените стойности.
                'cols':cols,'waste':waste,'whole_sheets':whole,'paper_price':paper_price,
                'color_change_count':num('color_change_count'),'paper_g':self.request_vars['paper'].get().strip(),
                'separator_material':sep_mat,'separator_sheets':sep_sheets,
                'finish_costs': {'3 стр. обрязване/рязане':cut,'Шиене':sewing,'Биговане':big,'Набор':typeset,'Пакетиране':packaging,'Лепене':glue,'Разрязване':cut2,'Леп. гръб':spine,'Прозорци':windows}}
            self._update_top_bar()
            self._refresh_request_offer()
            vals={'paper':paper,'print':printing,'turn':turn,'over1000':over,'color_change':color_change,'plates':plates,'cutting':cut,'prepress':prepress,'sewing':sewing,'bigoving':0,'typesetting':fold,'gluing':glue,'cutting2':cut2,'spine':spine,'windows':windows,'packaging':packaging,'electric':electric,'other':other_cost,'transport':transport_cost,'separators':sep,'surcharge':surcharge,'total':total,'unit':unit,'profit':profit,'expenses':expenses}

            # Данни за схемата: текущият печатен формат + необрязаният размер.
            # Самата схема чете директно полетата, за да се обновява веднага при промяна.
            self._calendar_scheme_data=True
            self._draw_calendar_scheme()

            if hasattr(self,'calendar_material_labels'):
                self.calendar_material_labels['paper_total'].configure(text=f'{paper:.2f} €' if paper>0 else '—')
                self.calendar_material_labels['plate_count'].configure(text=self._fmt_count(plate_count) if plate_count>0 else '—')
                self.calendar_material_labels['plate_total'].configure(text=f'{plates:.2f} €' if plates>0 else '—')

            for k,w in self.calendar_price_labels.items():
                val=float(vals.get(k,0) or 0)
                w.configure(text=f'{val:.2f} €' if val>0 else '—')

            for k,w in getattr(self,'calendar_finish_price_labels',{}).items():
                val=float(vals.get(k,0) or 0)
                if k=='separators':
                    # При разделителите първият резултат показва само необходимите листа.
                    # Общата сума е отделно, след „на листа“ и полето за брой листа.
                    w.configure(text=f'{int(sep_sheets)}л' if sep_sheets else '0л')
                    if hasattr(self, 'calendar_separator_total_label'):
                        self.calendar_separator_total_label.configure(
                            text=f'{sep:.2f} €'.replace('.',',') if sep > 0 else '0,00 €'
                        )
                else:
                    w.configure(text=f'{val:.2f} €'.replace('.',',') if val>0 else '0,00 €')

            for k,w in self.calendar_diag_labels.items(): w.configure(text='—')
            self.calendar_diag_labels['cols'].configure(text=self._fmt_count(cols))
            self.calendar_diag_labels['clean'].configure(text=self._fmt_count(qty))
            self.calendar_diag_labels['waste'].configure(text=self._fmt_count(waste))
            self.calendar_diag_labels['whole'].configure(text=self._fmt_count(whole))
            self.calendar_diag_labels['print'].configure(text=pf or '—')
            self.calendar_diag_labels['colors'].configure(text=f'{cf}+{cb}')
            self._last_valid=True; self._update_top_bar()
        except Exception as e:
            self._last_valid=False; messagebox.showerror('Грешка при изчислението — Календари',str(e))

    def _calendar_plate_count(self,qty,pages,cols,cf,cb,turnover):
        if turnover=='не': return 0
        base=2.7*(cf+cb if False else cb)
        if qty<=1000: return 2.7*cb
        coef={'да':1,'½ с обръщ.':0.5,'¼ с обръщ.':0.25}.get(turnover,0)
        return round((math.floor(qty/100)*coef*2.7*cb)/2.7)*2.7

    def _reset_calendars(self):
        defaults={'source_format':'a4','print_format':'','trim_h':'210','trim_w':'297','color_front':'1','color_back':'0','qty':'1000','pages':'1','turnover':'не','repetitions':'2','repetition_mode':'Автоматично','repetitions_manual':'2','paper_price':'','vat':'със','color_change_opt':'не','color_change_count':'2','plates_opt':'да','cut_opt':'стандарт','sewing_opt':'без','bigoving_opt':'без','typesetting_opt':'не','gluing_opt':'без','cut2_opt':'без','spine_opt':'без','windows_opt':'без','packaging_opt':'не','electric_opt':'без','sep_material':'без','sep_count':'100','surcharge':'40','prepress_price':'','other_price':'','transport_price':'','separators_material':'','separators_count':''}
        for k,d in defaults.items(): self._cvar(k,d).set(d)
        self.calendar_result={}
        self._calendar_scheme_data=None
        self._draw_calendar_scheme()
        for w in getattr(self,'calendar_price_labels',{}).values(): w.configure(text='—')
        for w in getattr(self,'calendar_material_labels',{}).values(): w.configure(text='—')
        for w in getattr(self,'calendar_finish_price_labels',{}).values(): w.configure(text='')
        self._calendar_update_print_formats(); self._calendar_recalc_diag(); self._update_top_bar()

    def _prices(self, f):
        """Таб Цени с разделена колона за Спирали (Размер + Цена)."""
        body = ttk.Frame(f)
        body.pack(fill='both', expand=True)

        header_frame = ttk.Frame(body)
        header_frame.pack(fill='x', pady=(0, 4))

        ttk.Label(
            header_frame, 
            text="Цени", 
            style='Section.TLabelframe.Label'
        ).pack(side='left', anchor='w')

        self._prices_locked = tk.BooleanVar(value=True)
        self._price_entries = []

        def toggle_lock():
            locked = self._prices_locked.get()
            new_locked = not locked
            self._prices_locked.set(new_locked)

            btn_text = "🔒 Заключени цени" if new_locked else "🔓 Отключени цени"
            lock_btn.config(text=btn_text)

            target_state = 'readonly' if new_locked else 'normal'
            for entry in self._price_entries:
                entry.config(state=target_state, style='PriceEdit.TEntry')

        lock_btn = ttk.Button(
            header_frame, 
            text="🔒 Заключени цени", 
            command=toggle_lock
        )
        lock_btn.pack(side='right', padx=5)

        grid = ttk.Frame(body)
        grid.pack(fill='both', expand=True)

        # Стандартни колони
        cols = [
            ('Основни', [('Печат лице/гръб','price_print_g4','8.00'),('Печат над 1000 / цвят','price_print_over1000','2.70'),('Обръщане','price_turnover','2.70'),('Дублаж / бр.','price_duplication','2.70'),('Плака / бр.','price_plate','2.80'),('Гориво / км','price_fuel','1.15')]),
            ('Предпечат / монтаж', [('флаери','price_em_flayers','1.28'),('листовки/стикери','price_em_leaflets','1.28'),('етикети/визитки','price_em_labels','1.02'),('корици','price_em_covers','1.53'),('бошура/покана','price_em_brochure','2.56'),('минимално','price_em_min','0.51'),('плакат','price_em_poster','2.56'),('страниране','price_em_pagination','0.25')]),
            ('Довършителни', [('УВ гланц / лист','price_uv_gloss','0.028'),('УВ мат / лист','price_uv_matt','0.056'),('УВ частичен / лист','price_uv_partial','0.071'),('УВ частичен обем / лист','price_uv_volume','0.075'),('Ламиниране гланц / лист','price_lam_gloss','0.051'),('Ламиниране кадифе / лист','price_lam_velvet','0.20'),('Ламиниране мат / лист','price_lam_matt','0.064'),('Каландър / лист','price_calender','0.025'),('Биговане / удар','price_bigoving','0.005'),('Сгъване ръчно / бр.','price_fold_manual','0.005'),('Сгъване Гатеви / бр.','price_fold_gatevi','0.0015'),('Сгъване Гад ново / бр.','price_fold_gad','0.002')]),
            ('Лепене', [
                ('Джоб','price_glue_pocket','0.06'),
                ('Знаменца','price_glue_flags','0.02'),
                ('Кутии','price_glue_boxes','0.012'),
                ('Дв. лепяща','price_glue_double','0.041'),
                ('Разделители', None, None),
                ('Вестник / лист','price_sep_news','0.028'),
                ('Друг / лист','price_sep_other','0.30')
            ])
        ]

        for c, (title, items) in enumerate(cols):
            card = self.card(grid, title)
            card.grid(row=0, column=c, padx=4, pady=0, sticky='nsew')
            for i, item_data in enumerate(items):
                if len(item_data) == 3 and item_data[1] is None:
                    lab, _, _ = item_data
                    lbl = ttk.Label(card, text=lab, font=('Segoe UI', 10, 'bold'))
                    lbl.grid(row=i, column=0, columnspan=2, padx=6, pady=(12, 6), sticky='w')
                else:
                    lab, key, d = item_data
                    ttk.Label(card, text=lab).grid(row=i, column=0, padx=5, pady=2, sticky='w')
                    en = self.entry(card, key, d, 10)
                    en.grid(row=i, column=1, padx=5, pady=2, sticky='e')
                    en.config(state='readonly', style='PriceEdit.TEntry')
                    self._price_entries.append(en)

        # Отделна секция за Спирали (Разделена на 2 подколони)
        spiral_card = self.card(grid, 'Спирали')
        spiral_card.grid(row=1, column=0, columnspan=4, padx=4, pady=(4, 0), sticky='nsew')

        spiral_items = [
            (('Зъб 3/16', '3/16"'), 'price_spiral_tooth_3_16', '0.00156'),
            (('Зъб 1/4', '1/4"'), 'price_spiral_tooth_1_4', '0.00210'),
            (('Зъб 5/16', '5/16"'), 'price_spiral_tooth_5_16', '0.00251'),
            (('Зъб 3/8', '3/8"'), 'price_spiral_tooth_3_8', '0.00220'),
            (('Зъб 7/16', '7/16"'), 'price_spiral_tooth_7_16', '0.00470'),
            (('Зъб 1/2', '1/2"'), 'price_spiral_tooth_1_2', '0.00360'),
            (('Зъб 9/16', '9/16"'), 'price_spiral_tooth_9_16', '0.00520'),
            (('Перфо — коефициент', 'кока'), 'price_spiral_perfo_coeff', '0.011'),
            (('Перфо — буфер труд', 'буфер'), 'price_spiral_perfo_buffer', '1.20'),
            (('Кукички до 80 мм', '80мм'), 'price_spiral_hooks_80', '0.04'),
            (('Кукички до 150 мм', '150мм'), 'price_spiral_hooks_150', '0.051'),
            (('За нарязване/зъб', 'зъб'), 'price_spiral_cut_tooth', '0.0031'),
        ]

        # Подредба на спиралите в 2 подколони (по 6 реда на колона)
        for idx, (label_data, key, d) in enumerate(spiral_items):
            col_offset = (idx // 6) * 3
            row_idx = idx % 6

            name_label, size_sublabel = label_data
            
            # Име / Описание
            ttk.Label(spiral_card, text=name_label, width=20).grid(row=row_idx, column=col_offset, padx=(5, 2), pady=1, sticky='w')
            # Размер / Тип
            ttk.Label(spiral_card, text=f"[{size_sublabel}]", font=('Segoe UI', 8, 'italic'), foreground='#555555').grid(row=row_idx, column=col_offset+1, padx=(0, 5), pady=1, sticky='w')
            # Поле за цена
            en = self.entry(spiral_card, key, d, 9)
            en.grid(row=row_idx, column=col_offset+2, padx=(0, 12), pady=1, sticky='e')
            en.config(state='readonly', style='PriceEdit.TEntry')
            self._price_entries.append(en)

        for c_idx in range(4):
            grid.columnconfigure(c_idx, weight=1)

        ttk.Label(
            body, 
            text='Цените в този раздел се настройват за параметрите на калкулатора.', 
            style='Sub.TLabel'
        ).pack(anchor='w', pady=3)
    def _finish_first(self, f):
        # Довършителните операции остават с ВСИЧКИ полета за избор.
        # Задаваме style='Section.TLabelframe', за да има белият фон на карето.
        operations = ttk.Frame(f, style='Section.TLabelframe')
        operations.pack(fill='x', anchor='n')
        for c in (1, 4):
            operations.columnconfigure(c, weight=1)

        rows = [
            (('Рязане', 'cutting', ['без', 'стандарт', 'форматиране', 'март. Ани', 'други'], 'стандарт', 'cutting'),
             ('Номерация', 'numbering', ['без', 'да', '+'], 'без', 'numbering')),
            (('Перфорация', 'perforation', ['без', 'да', '+'], 'без', 'perforation'),
             ('УВ лак', 'uv', ['без', 'гланц', 'гланц дв.', 'кадифе', 'кадифе дв.', 'мат', 'мат дв.', 'надпечат.', 'частичен', 'част. дв', 'част.обем'], 'без', 'uv')),
            (('Ламиниране', 'lamination', ['без', 'гланц', 'гланц дв.', 'кадифе', 'кадифе дв.', 'мат', 'мат дв.'], 'без', 'lamination'),
             ('Каландър', 'calender', ['без к', 'каландър', 'каландър дв.'], 'без к', 'calender')),
            (('Филм, €', 'film_price', None, '', None),
             ('Лепене', 'gluing', ['без', 'а', 'джоб', 'знаменца', 'каширане', 'кубчета', 'кутии', 'дв. лепящ'], 'без', 'gluing')),
            (('Брой гънки', 'folding', None, '0', 'folding'),
             ('Тип сгъване', 'foldtype', ['без', 'ръчно', 'Гатеви', 'Гад ново'], 'без', 'folding')),
            (('Щанцоване / преге', 'die', ['без', 'щанцов', '½ щанцов', 'преге', 'преге+'], 'без', 'die_cut'),
             ('Очупване, %', 'breaking', None, '40', 'breaking_cost')),
            (('Брой бигове', 'bigoving', ['0', '1', '2', '3', '4', '5', '6'], '0', None),
             ('Ел. монтаж', 'electric_montage', ['без', 'бошура/покана', 'етикети/визитки', 'корици', 'листовки/стикери', 'минимално', 'плакат', 'страниране', 'флаери'], 'без', 'electric_montage')),
            (('Биговане', 'bigoving_type', ['без', 'ръчно', 'машинно'], 'без', 'bigoving'),
             ('Операции, друго', 'other_operations', None, '', 'other')),
            (('Броене/пакетиране', 'counting', ['да', '<1000', 'не'], 'да', 'counting'),
             ('Заоб. / замба', 'round_punch', ['', 'заобл.', 'замба'], '', 'round_punch')),
            (('Разделители', 'sepmat', ['без', 'вестник', 'друг'], 'без', None),
             ('Брой разделители', 'sepn', None, '50', 'separators')),
            (('Транспорт', 'transport', ['не', 'да', 'доставка+'], 'не', 'transport'),
             ('Километри', 'km', None, '15', None)),
        ]

        self.finish_price_labels = {}
        self.separator_sheet_label = None
        for r, (left_item, right_item) in enumerate(rows):
            for c, item in ((0, left_item), (3, right_item)):
                lab, key, vals, d, price_key = item
                # Изрично задаваме bg='#FFFFFF' за текстовите етикети
                tk.Label(operations, text=lab, anchor='e', bg='#FFFFFF', fg='#202633', font=('Segoe UI', 10)).grid(
                    row=r, column=c, padx=(7, 2), pady=5, sticky='e')
                w = self.combo(operations, key, vals, d, 18) if isinstance(vals, list) else self.entry(operations, key, d, 8)
                # Полето остава компактно, но падащият списък се разширява
                # според най-дългата опция, за да не се реже текстът.
                if isinstance(vals, list):
                    self._make_dropdown_wide(w, vals)
                w.grid(row=r, column=c + 1, padx=(0, 4), pady=5, sticky='w')
                price_lbl = ttk.Label(operations, text='', style='Value.TLabel', background='#FFFFFF', width=8, anchor='w')
                price_lbl.grid(row=r, column=c + 2, padx=(3, 7), pady=5, sticky='w')
                if key == 'sepmat':
                    self.separator_sheet_label = price_lbl
                elif price_key:
                    self.finish_price_labels[key] = price_lbl

        self.finish_diag = ttk.Frame(f, style='Section.TLabelframe')

    def _finish(self,f):
        # Химия / кочани — компактна подредба на довършителните операции.
        # Подредбата е по 2 полета на ред:
        # 1) Рязане | Лепене
        # 2) Номерация | Перфорация
        # 3) Биговане — вид | Брой бигове
        # 4) Набор | Шиене/телчета
        # 5) Ел. монтаж | Пакетиране
        # 6) Разделители | Брой
        # 7) Транспорт | Друго
        # 8) Оскъпяване, %

        operations = self.card(f, 'Довършителни операции')
        operations.pack(fill='x', anchor='n')
        operations.columnconfigure(1, weight=1)
        operations.columnconfigure(3, weight=1)

        rows = [
            (('Рязане', 'cutting', ['без','стандарт','форматиране','март. Ани','други'], 'стандарт'),
             ('Лепене', 'gluing', ['без','а','джоб','знаменца','каширане','кубчета','кутии','дв. лепящ'], 'без')),

            (('Номерация', 'numbering', ['без','да','+'], 'без'),
             ('Перфорация', 'perforation', ['без','да','+'], 'без')),

            (('Биговане — вид', 'big_type', ['без','ръчно','машинно'], 'без'),
             ('Биговане — брой', 'big_count', None, '0')),

            (('Набор', 'typesetting', ['','ръчно','машинно'], ''),
             ('Шиене/телчета', 'sewing', ['без','1','2','3','4'], 'без')),

            (('Ел. монтаж', 'em',
              ['без','флаери','листовки/стикери','етикети/визитки','корици',
               'бошура/покана','минимално','плакат','страниране'], 'без'),
             ('Пакетиране', 'counting', ['да','<1000','не'], 'да')),

            (('Разделители', 'sep_mat', ['без','вестник','картон'], 'без'),
             ('Брой', 'sep_n', None, '0')),

            (('Транспорт', 'transport', ['не','да','доставка+'], 'не'),
             ('Километри', 'km', None, '15')),

            (('Оскъпяване, %', 'surcharge', None, '40'),
             (None, None, None, None)),
        ]

        for r, (left_item, right_item) in enumerate(rows):
            for c, item in ((0, left_item), (2, right_item)):
                lab, key, vals, default = item
                if not lab:
                    continue

                ttk.Label(
                    operations, text=lab
                ).grid(row=r, column=c, padx=(7,2), pady=4, sticky='e')

                if isinstance(vals, list):
                    w = self.combo(operations, key, vals, default, 13)
                else:
                    w = self.entry(operations, key, default, 9)

                w.grid(
                    row=r, column=c+1,
                    padx=(0, 8 if c == 2 else 10),
                    pady=4, sticky='w'
                )

        # Диагностиката остава непосредствено под довършителните операции.
        diag_box = self.card(f, 'Диагностика — довършителни')
        diag_box.pack(fill='both', expand=True, pady=(8,0), anchor='n')
        diag_frame = ttk.Frame(diag_box)
        diag_frame.pack(fill='both', expand=True, padx=4, pady=4)
        self.finish_diag = diag_frame
        self._set_diag(self.finish_diag, [], columns=2)

    def _diag_card(self,parent,title):
        lf=self.card(parent,title);lf.pack(side='left',fill='both',expand=True,padx=(8,0),anchor='n')
        box=ttk.Frame(lf);box.pack(fill='both',expand=True);return box

    def _set_diag(self, container, rows, columns=1):
        for c in container.winfo_children():
            c.destroy()

        # Поставяме бял фон на самия контейнер, за да няма разминаване в цвета
        try:
            container.configure(style='Section.TLabelframe')
        except Exception:
            pass

        shown = [(k, v) for k, v in rows if v not in (None, '', 0, '0', 0.0, '0.00 €', '0.00')]
        if not shown:
            ttk.Label(container, text='Няма резултати, различни от 0.', style='Sub.TLabel', background='#FFFFFF').grid(
                row=0, column=0, columnspan=max(1, columns), sticky='w', padx=6, pady=6)
            return

        key_value_highlight = {
            'Тираж', 'Чист тираж', 'Печатен формат', 'Цели листа',
            'Цели листа за тиража'
        }
        columns = max(1, int(columns))
        per_col = (len(shown) + columns - 1) // columns
        for i, (k, v) in enumerate(shown):
            col = i // per_col
            row = i % per_col
            base = col * 2
            container.columnconfigure(base, weight=0)
            container.columnconfigure(base + 1, weight=0)
            
            # Изрично задаваме background='#FFFFFF' (бял фон) за текстовете и стойностите
            ttk.Label(container, text=k, style='Card.TLabel', background='#FFFFFF').grid(row=row, column=base, sticky='w', padx=(2, 4), pady=4)
            value_style = 'DiagHighlight.TLabel' if k in key_value_highlight else 'Value.TLabel'
            ttk.Label(container, text=str(v), style=value_style, background='#FFFFFF').grid(
                row=row, column=base + 1, sticky='w', padx=(0, 2), pady=4)
    def _result(self, f):
        # Основният „Резултат“ таб вече е само „Заявка“.
        # Старите резултатни контейнери се запазват скрити за съвместимост
        # с вътрешното обновяване, но не се показват като отделен интерфейс.
        self._result_hidden = ttk.Frame(f)
        self.result_summary = ttk.Frame(self._result_hidden)
        self.result_finish = ttk.Frame(self._result_hidden)
        self.result_body = ttk.Frame(self._result_hidden)
        self.result_heading = ttk.Label(self._result_hidden, text='', font=('Segoe UI', 15, 'bold'))

        self._request_offer(f)

    def _request_entry(self, parent, key, width=18):
        return ttk.Entry(parent, textvariable=self.request_vars[key],
                         width=width, style='RequestOrange.TEntry')

    def _request_style(self):
        return {
            'grey':'#EEF1F4',
            'orange':'#FCE4D6',
            'white':'#FFFFFF',
            'line':'#808080',
            'text':'#333333',
            'orange_text':'#8B4A12'
        }

    def _request_cell(self, parent, text, row, col, bg='#FFFFFF', bold=False,
                      align='left', colspan=1, rowspan=1, font_size=10, fg='#333333', wraplength=None):
        anchor={'left':'w','center':'center','right':'e'}.get(align,'w')
        # Таблицата остава бяла. Само автоматично изчислените клетки,
        # които подаваме със стария #C6C6C6, получават много светло сиво.
        # Старите извиквания с #E7E6E6 се третират като обикновени бели клетки.
        cell_bg = '#EEF1F4' if bg == '#C6C6C6' else '#FFFFFF'
        lbl=tk.Label(parent, text=str(text), bg=cell_bg, fg=fg,
                     font=('Segoe UI', font_size, 'bold' if bold else 'normal'),
                     anchor=anchor, padx=9, pady=4, relief='solid', borderwidth=1,
                     highlightthickness=0, justify='left',
                     wraplength=(wraplength if wraplength is not None else 0))
        lbl.grid(row=row, column=col, columnspan=colspan, rowspan=rowspan,
                 sticky='nsew', padx=0, pady=0)
        return lbl

    def _request_entry(self, parent, key, width=18):
        # Двете групи ръчни полета са независими: горните данни са в
        # request_vars, а производствените полета от референтната заявка
        # са в request_manual_vars. Нито една от тях не се подава към engine-а.
        manual_keys = ('item','client','date','paper','paper_type')
        if key in getattr(self, 'request_manual_vars', {}):
            variable = self.request_manual_vars[key]
            bg = '#EAF4FF'
        else:
            variable = self.request_vars[key]
            bg = '#E6F0FA' if key in manual_keys else '#FFFFFF'
        e = tk.Entry(parent, textvariable=variable, width=width,
                     bg=bg, fg='#202633', font=('Segoe UI',10),
                     relief='solid', borderwidth=1, highlightthickness=1,
                     highlightbackground='#D9E1EA', highlightcolor='#4F46E5',
                     insertbackground='#202633', justify='left')
        return e
    def _request_sheet_grid(self, sheet):
        # Four fixed columns matching the Excel request sheet proportions.
        widths=[16,26,17,26]
        for i,w in enumerate(widths):
            sheet.grid_columnconfigure(i, minsize=w*8, weight=0)

    def _request_manual(self, sheet, label, key, row, col, label_bold=True):
        self._request_cell(sheet, label.upper(), row, col, bg='#E7E6E6', bold=label_bold, align='right')
        e=self._request_entry(sheet,key)
        e.grid(row=row,column=col+1,sticky='nsew',ipady=3,padx=0,pady=0)
        return e

    def _request_auto_row(self, sheet, row, left_label, left_value, right_label=None, right_value=None,
                          left_label_bold=False, right_label_bold=True,
                          left_value_bg='#E7E6E6', right_value_bg='#E7E6E6',
                          row_bold=False):
        self._request_cell(sheet,left_label,row,0,bg='#E7E6E6',bold=(left_label_bold or row_bold),align='right')
        self._request_cell(sheet,left_value,row,1,bg=left_value_bg,bold=row_bold,align='left')
        if right_label is not None:
            self._request_cell(sheet,right_label,row,2,bg='#E7E6E6',bold=(right_label_bold or row_bold),align='right')
            self._request_cell(sheet,right_value,row,3,bg=right_value_bg,bold=(right_label=='ЦЕНА БЕЗ ДДС:' or row_bold),align='center')

    def _request_finish_text(self):
        selected=[]
        def val(k): return str(self.vars[k].get()).strip() if k in self.vars else ''
        def num(k):
            try: return float(val(k).replace(',','.'))
            except Exception: return 0
        if val('cutting').lower() not in ('','без','не'): selected.append('Рязане')
        if val('numbering').lower() not in ('','без','не'): selected.append('Номерация')
        if val('perforation').lower() not in ('','без','не'): selected.append('Перфорация')
        if val('uv').lower() not in ('','без','не'): selected.append('УВ лак')
        if val('lamination').lower() not in ('','без','не'): selected.append('Ламиниране')
        if val('calender').lower() not in ('','без к','без','не'): selected.append('Каландър')
        if val('film_price') not in ('','0','0.0','0,0'): selected.append('Филм')
        if val('gluing').lower() not in ('','без','не'): selected.append('Лепене')
        if num('bigoving')>0: selected.append(f"Биговане - {int(num('bigoving'))} бига")
        if num('folding')>0: selected.append(f"Сгъване ({int(num('folding'))} бр.)")
        if val('die').lower() not in ('','без','не'): selected.append('Щанцоване')
        if val('round_punch').lower() not in ('','без','не'): selected.append('Заоб. / замба')
        if val('electric_montage').lower() not in ('','без','не'): selected.append('Ел. монтаж')
        if val('counting').lower() not in ('','без','не'): selected.append('Пакетиране')
        if val('sepmat').lower() not in ('','без','не'): selected.append(f"Разделители ({val('sepn') or '50'})")
        return ' | '.join(selected)

    def _request_values(self):
        r = self.result or {}
        front = self.vars.get('front', tk.StringVar(value='')).get()
        back = self.vars.get('back', tk.StringVar(value='')).get()
        turnover = self.vars.get('turnover', tk.StringVar(value='')).get()
        source_fmt = str(r.get('source_format', ''))
        print_fmt = str(r.get('print_format', ''))
        source_sheets = r.get('source_sheets', 0)
        
        # Взимаме реалния брой общи листове от изчислението (total_print_sheets)
        total_sheets_val = r.get('total_print_sheets', r.get('sheets', 0))
        
        paper_per_sheet = ''
        if source_sheets:
            paper_per_sheet = f"{r.get('paper', 0)/source_sheets:.3f} € цена / лист"
            
        return {
            'item': self.request_vars['item'].get().strip(),
            'client': self.request_vars['client'].get().strip(),
            'date': self.request_vars['date'].get().strip(),
            'paper': self.request_vars['paper'].get().strip().upper(),
            'paper_type': self.request_vars['paper_type'].get().strip(),
            'paper_price': paper_per_sheet,
            'size': f"{r.get('product_w', 0):g}x{r.get('product_h', 0):g} мм" if r else '',
            'qty': self._fmt_count(r.get('unit_pieces', '')),
            'color': f"{front} + {back}" if (front or back) else '',
            'turnover': turnover,
            'price': f"{r.get('total', 0):.2f} €" if r else '',
            'unit': f"{r.get('unit', 0):.4f} €" if r else '',
            'source_format': source_fmt,
            'print_format': print_fmt,
            'total_sheets': self._fmt_count(total_sheets_val), # <-- Тук подаваме правилната стойност
            'clean': self._fmt_count(r.get('clean_sheets', '')),
            'reps': str(r.get('repetitions', '')),
            'waste': str(r.get('waste_sheets', '')),
            'source_sheets': self._fmt_count(source_sheets),
            'separator_sheets': self._fmt_count(r.get('separator_sheets', 0)),
            'finish': self._request_finish_text() or '—'
        }
    def _request_seed_manual(self, key, value):
        v=self.request_manual_vars[key]
        if not v.get().strip() and value not in (None, ''):
            v.set(str(value))

    def _request_production_display(self, content, active, r):
        """Показва производствените параметри според активния калкулатор.

        Това са само прочетени резултати/параметри от съответния таб.
        Не са полета за ръчно въвеждане и не участват обратно в калкулацията.
        """
        def txt(value, fallback='—'):
            if value is None:
                return fallback
            value=str(value).strip()
            return value if value else fallback

        def num(value):
            try:
                return float(str(value).replace(',', '.'))
            except Exception:
                return 0.0

        rows=[]
        if active == 'Основен':
            front=txt(self.vars.get('front', tk.StringVar(value='')).get(), '')
            back=txt(self.vars.get('back', tk.StringVar(value='')).get(), '')
            color=f'{front} + {back}' if front or back else '—'
            reps=num(r.get('repetitions', 0))
            qty=num(r.get('unit_pieces', r.get('qty', 0)))
            tirage=(qty / reps) if reps else num(r.get('qty', 0))
            sheets=r.get('total_print_sheets', r.get('sheets', r.get('source_sheets', '')))
            waste=r.get('waste_sheets', '')
            sheets_text=txt(self._fmt_count(sheets))
            if sheets_text != '—':
                sheets_text += ' листа (вкл. макулатура)'
            turnover=txt(r.get('turnover', self.vars.get('turnover', tk.StringVar(value='')).get()))
            source=txt(r.get('source_format', self.vars.get('source', tk.StringVar(value='')).get()))
            pf=txt(r.get('print_format', self.vars.get('print_format', tk.StringVar(value='')).get()))
            product=f"{num(r.get('product_w',0)):g}x{num(r.get('product_h',0)):g} мм" if r.get('product_w') not in (None,'') and r.get('product_h') not in (None,'') else '—'
            rows=[
                ('ЕД. БРОЙКИ', self._fmt_count(qty)),
                ('ТИРАЖ', self._fmt_count(tirage)),
                ('ЦВЕТНОСТ', color),
                ('ПЕЧАТНИ ЛИСТА', sheets_text),
                ('ОБРЪЩАНЕ', turnover),
                ('ФОРМАТ (НА Х-Я)', source),
                ('ФОРМАТ ЗА ПЕЧАТ', pf),
                ('РАЗМНОЖЕНИЯ', self._fmt_count(reps)),
                ('ОБРЯЗАН РАЗМЕР', product),
            ]
        elif active == 'Книжки':
            bv=lambda k: str(self.book_vars.get(k).get()).strip() if k in getattr(self,'book_vars',{}) else ''
            front=txt(bv('front'),''); back=txt(bv('back'),'')
            color=f'{front} + {back}' if front or back else '—'
            rows=[
                ('ЕД. БРОЙКИ', self._fmt_count(r.get('unit_pieces',r.get('qty','')))),
                ('ТИРАЖ', self._fmt_count(r.get('qty',r.get('unit_pieces','')))),
                ('ЦВЕТНОСТ', color),
                ('ОБРЪЩАНЕ', txt(r.get('turnover',bv('turnover')))),
                ('ФОРМАТ (НА Х-Я)', txt(r.get('source',r.get('source_format',bv('source'))))),
                ('ФОРМАТ ЗА ПЕЧАТ', txt(r.get('print_format',''))),
                ('ПЕЧАТНИ КОЛИ', self._fmt_count(r.get('cols',''))),
                ('РАЗМНОЖЕНИЯ', self._fmt_count(r.get('repetitions',''))),
                ('ЛИСТА ВКЛ. МАКУЛАТУРА', f"{self._fmt_count(r.get('whole_sheets',''))} листа"),
                ('ОБРЯЗАН РАЗМЕР', txt(r.get('trim',''))),
                ('СТРАНИЦИ', self._fmt_count(r.get('pages',bv('pages')))),
            ]
        elif active == 'Кочани':
            hv=lambda k: str(self._hvar(k).get()).strip() if k in getattr(self,'himiya_vars',{}) else ''
            front=txt(hv('front'),''); back=txt(hv('back'),'')
            color=f'{front} + {back}' if front or back else txt(r.get('paper_colors'),'—')
            sheets_block = r.get('sheets_per_block','')
            paper_colors = r.get('paper_colors', hv('paper_colors'))
            clean = r.get('clean_sheets','')
            rows=[
                # „Ед. бройки“ при Кочани е броят на кочаните от самия таб.
                ('ЕД. БРОЙКИ', f"{self._fmt_count(r.get('quantity',r.get('unit_pieces','')))} кочана"),
                # Листовете и цветовете се показват като една производствена
                # стойност: напр. „33 листа × 3 цвят/а“.
                ('БР. ЛИСТА В КОЧАН / ОТ ЦВЯТ', f"{self._fmt_count(sheets_block)} листа × {self._fmt_count(paper_colors)} цвят/а"),
                ('ТИРАЖ', self._fmt_count(clean)),
                ('ЦВЕТНОСТ', color),
                # В engine резултатът „turnover“ е цена за обръщането, а в
                # „Производствени параметри“ трябва да показваме избора да/не.
                ('ОБРЪЩАНЕ', txt(hv('turnover')).upper()),
                ('ФОРМАТ (НА Х-Я)', txt(r.get('source_format',''))),
                ('ФОРМАТ ЗА ПЕЧАТ', txt(r.get('print_format',''))),
                ('РАЗМНОЖЕНИЯ', self._fmt_count(r.get('repetitions',''))),
                ('ЦЕЛИ ЛИСТА', f"{self._fmt_count(r.get('whole_sheets',''))} листа/цвят"),
                ('ОБРЯЗАН РАЗМЕР', f"{num(r.get('product_w',0)):g}x{num(r.get('product_h',0)):g} мм"),
            ]
        elif active == 'Календари':
            color=f"{txt(r.get('color_front'),'0')} + {txt(r.get('color_back'),'0')}"
            rows=[
                ('ЕД. БРОЙКИ', self._fmt_count(r.get('qty',''))),
                ('ТИРАЖ', self._fmt_count(r.get('qty',''))),
                ('ЦВЕТНОСТ', color),
                ('ПЕЧАТНИ ЛИСТА', self._fmt_count(r.get('whole_sheets',''))),
                ('ОБРЪЩАНЕ', txt(r.get('turnover',''))),
                ('ФОРМАТ (НА Х-Я)', txt(r.get('source_format',''))),
                ('ФОРМАТ ЗА ПЕЧАТ', txt(r.get('print_format',''))),
                ('РАЗМНОЖЕНИЯ', self._fmt_count(r.get('repetitions',''))),
                ('ОБРЯЗАН РАЗМЕР', txt(r.get('trim_size',''))),
                ('СТРАНИЦИ', self._fmt_count(r.get('pages',''))),
            ]
        elif active == 'Спирали':
            sv=lambda k: str(self._spiral_var(k,'').get()).strip()
            rows=[
                ('ЕД. БРОЙКИ', self._fmt_count(r.get('qty',sv('qty')))),
                ('РАЗМЕР НА СПИРАЛАТА', txt(r.get('size',sv('size')))),
                ('БРОЙ ЗЪБИ', self._fmt_count(r.get('teeth',sv('teeth')))),
                ('ТЯЛО — ЛИСТА', self._fmt_count(sv('body_sheets'))),
                ('КОРИЦА — ЛИСТА', self._fmt_count(sv('cover_sheets'))),
                ('ТЯЛО — ГРАМАЖ', txt(sv('body_gsm'))),
                ('КОРИЦА — ГРАМАЖ', txt(sv('cover_gsm'))),
            ]
        else:
            rows=[('ПАРАМЕТРИ','—')]

        # Две колони за производствените параметри, за да не се
        # разтяга карето надолу и резултатът да остава видим.
        def _is_zero_display(value):
            # Скриваме само реални нулеви стойности; текстови параметри като
            # „Не“ и други описания не се приемат за нула.
            text = str(value).strip().replace(',', '.')
            if not text:
                return False
            try:
                return float(text) == 0
            except Exception:
                parts = [p.strip() for p in text.split('+')]
                if len(parts) > 1:
                    try:
                        return all(float(p) == 0 for p in parts)
                    except Exception:
                        return False
                return text.startswith('0 ') or text.startswith('0 мм') or text.startswith('0x0')

        visible_rows = [(label, value) for label, value in rows if not _is_zero_display(value)]
        for row,(label,value) in enumerate(visible_rows):
            col = row % 2
            grid_row = row // 2
            self._request_field(content, label, value, grid_row, col, value_bold=True)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)

    def _request_manual_production(self, content, active, r):
        # Съвместимост със стари извиквания: производствените параметри вече
        # са само информативни, а ръчно въвеждането е единствено горе в заявката.
        return self._request_production_display(content, active, r)

    def _request_offer(self, f):
        """Вътрешен работен екран „Заявка“.

        Визуалният слой е отделен от калкулациите: всички съществуващи
        request_vars, резултати и TXT запис остават непроменени.
        """
        outer = tk.Frame(f, bg='#F5F7FA')
        outer.pack(fill='both', expand=True)

        # Самата страница започва директно с картата „Данни за заявката“,
        # както в референтния екран. Горната обща лента на приложението
        # вече показва името на активния таб и крайната цена.
        body = tk.Frame(outer, bg='#F5F7FA')
        body.pack(fill='both', expand=True)
        self.request_sheet = body

        footer = tk.Frame(outer, bg='#F5F7FA')
        footer.pack(fill='x', pady=(6, 0))
        ttk.Button(footer, text='Запази заявка (TXT)',
                   command=self._save_offer_text).pack(side='right')
        self._refresh_request_offer()

    def _request_card(self, parent, title, column=0, row=0, colspan=1):
        """Светла секция за „Заявка“, без таблична мрежа."""
        card = tk.Frame(parent, bg='#FFFFFF', highlightbackground='#DCE2EA',
                        highlightcolor='#DCE2EA', highlightthickness=1,
                        bd=0)
        card.grid(row=row, column=column, columnspan=colspan,
                  sticky='nsew', padx=5, pady=5)
        head = tk.Frame(card, bg='#FFFFFF')
        head.pack(fill='x', padx=14, pady=(12, 7))
        icon_kinds = {
            'Данни за заявката': 'document',
            'Производствени параметри': 'settings',
            'Материали': 'materials',
            'Допълнителни разходи': 'costs',
            'Резултати': 'results',
            'Обща цена': 'total',
            'Довършителни работи': 'finish',
        }
        icon = self._make_ui_icon(icon_kinds.get(title, 'document'), 19, '#4F46E5')
        # Държим референция, за да не бъде събрана PhotoImage от garbage collector.
        self._request_icon_refs = getattr(self, '_request_icon_refs', [])
        self._request_icon_refs.append(icon)
        tk.Label(head, image=icon, bg='#FFFFFF').pack(side='left', padx=(0, 7))
        tk.Label(head, text=title, bg='#FFFFFF', fg='#243B5A',
                 font=('Segoe UI', 11, 'bold')).pack(side='left')
        content = tk.Frame(card, bg='#FFFFFF')
        content.pack(fill='both', expand=True, padx=14, pady=(0, 12))
        return card, content

    def _request_field(self, parent, label, value, row, col=0,
                       manual_key=None, width=20, value_bold=False):
        """Поле в светла секция; manual_key прави истинско Entry."""
        box = tk.Frame(parent, bg='#FFFFFF')
        box.grid(row=row, column=col, sticky='ew', padx=5, pady=4)
        tk.Label(box, text=label, bg='#FFFFFF', fg='#667085',
                 font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 3))
        if manual_key is not None:
            e = self._request_entry(box, manual_key, width=width)
            # Ръчните полета в „Данни за заявката“ са леко синкави,
            # за да се различават ясно от автоматично попълнените стойности.
            e.configure(bg='#EAF4FF', relief='solid', borderwidth=1,
                         highlightthickness=1, highlightbackground='#D9E1EA',
                         highlightcolor='#7C5CFF', insertbackground='#202633')
            e.pack(fill='x', ipady=5)
            return e
        tk.Label(box, text=str(value if value not in (None, '') else '—'),
                 bg='#FFFFFF', fg='#202633',
                 font=('Segoe UI', 10, 'bold' if value_bold else 'normal'),
                 anchor='w').pack(fill='x', ipady=5)
        return None

    def _request_stat(self, parent, label, value, row, col=0, accent=False):
        box = tk.Frame(parent, bg='#F8F9FC', highlightbackground='#E5E9F0',
                       highlightthickness=1, bd=0)
        box.grid(row=row, column=col, sticky='ew', padx=5, pady=4)
        tk.Label(box, text=label, bg='#F8F9FC', fg='#667085',
                 font=('Segoe UI', 9)).pack(anchor='w', padx=10, pady=(7, 1))
        tk.Label(box, text=str(value if value not in (None, '') else '—'),
                 bg='#F8F9FC', fg='#5B3FD1' if accent else '#202633',
                 font=('Segoe UI', 11, 'bold')).pack(anchor='w', padx=10, pady=(0, 7))

    def _request_result_stats(self, parent, result, source='Основен'):
        """Показва цената, единичната цена и информационната бележка в секция Резултати."""
        total = float(result.get('total', 0) or 0)
        unit = float(result.get('unit', 0) or 0)
        row = 0
        
        if abs(total) > 1e-12:
            self._request_stat(parent, 'Цена без ДДС', f'{total:.2f} €', row, 0, True)
            row += 1
        if abs(unit) > 1e-12:
            self._request_stat(parent, 'Ед. цена', f'{unit:.4f} €', row, 0)
            row += 1

        # Бледолилавата информационна бележка с преливащ текст
        note = tk.Frame(parent, bg='#F2EEFF', bd=0, highlightthickness=0)
        note.grid(row=row, column=0, sticky='nsew', padx=5, pady=(8, 4))
        note.columnconfigure(0, weight=1)

        msg_text = f'Резултатите са от „{source}“ на база на въведените параметри и текущите цени в системата.'
        
        tk.Label(
            note,
            text=f'ℹ  {msg_text}',
            bg='#F2EEFF', fg='#7C6BD8',
            font=('Segoe UI', 9), justify='left', anchor='w',
            wraplength=230
        ).pack(fill='both', expand=True, padx=10, pady=10)

    def _request_empty_state(self, sheet):
        for w in sheet.winfo_children():
            w.destroy()
        sheet.columnconfigure(0, weight=1)
        empty = tk.Frame(sheet, bg='#FFFFFF', highlightbackground='#DCE2EA',
                         highlightthickness=1, bd=0)
        empty.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        tk.Label(empty, text='Няма готов резултат', bg='#FFFFFF', fg='#243B5A',
                 font=('Segoe UI', 12, 'bold')).pack(pady=(28, 4))
        tk.Label(empty, text='Направете изчисление в активния таб, за да се попълни заявката.',
                 bg='#FFFFFF', fg='#667085', font=('Segoe UI', 10)).pack(pady=(0, 28))

    def _refresh_request_offer(self):
        """Обновява визуалното представяне на „Заявка“ според последната калкулация.

        Източникът и всички стойности се взимат от същите резултати и
        textvariable-и като преди; променен е само визуалният слой.
        """
        if not hasattr(self, 'request_sheet'):
            return
        sheet = self.request_sheet
        self._request_icon_refs = []
        for w in sheet.winfo_children():
            w.destroy()
        self._request_icon_refs = []
        for w in sheet.winfo_children():
            w.destroy()

        active = getattr(self, 'active_tab_name', 'Основен')
        if active == 'Заявка':
            active = getattr(self, 'request_source_tab', 'Основен')

        # В „Основен“ имаме 4 визуални зони:
        # Производствени параметри | Материали | Довършителни | Резултати.
        # Производствената колона е леко по-широка, за да не се режат дългите параметри.
        if active == 'Основен':
            sheet.columnconfigure(0, weight=4, uniform='request_main')
            sheet.columnconfigure(1, weight=3, uniform='request_main')
            sheet.columnconfigure(2, weight=3, uniform='request_main')
            sheet.columnconfigure(3, weight=3, uniform='request_main')
        else:
            # При „Кочани“ карето „Производствени параметри“ трябва да има
            # повече хоризонтално място, защото съдържа по-дълги описания.
            if active in ('Кочани', 'Календари'):
                sheet.columnconfigure(0, weight=4, uniform='request_main')
                sheet.columnconfigure(1, weight=3, uniform='request_main')
                sheet.columnconfigure(2, weight=3, uniform='request_main')
                sheet.columnconfigure(3, weight=3, uniform='request_main')
            else:
                for c in range(4):
                    sheet.columnconfigure(c, weight=1, uniform='request')
        sheet.rowconfigure(1, weight=0)
        sheet.rowconfigure(2, weight=0)
        sheet.rowconfigure(3, weight=1)

        # Общата горна секция — ръчно въвеждаемите данни остават същите.
        _, details = self._request_card(sheet, 'Данни за заявката', 0, 0, 4)
        for c in range(4):
            details.columnconfigure(c, weight=1)
        self._request_field(details, 'Изделие', '', 0, 0, 'item')
        self._request_field(details, 'Клиент', '', 0, 1, 'client')
        self._request_field(details, 'Дата', '', 0, 2, 'date')
        self._request_field(details, 'Хартия', '', 0, 3, 'paper')

        # --- Книжки ---
        if active == 'Книжки' and getattr(self, 'book_result', None):
            r = self.book_result
            bv = lambda k: str(self.book_vars.get(k).get()).strip() if k in getattr(self, 'book_vars', {}) else ''
            qty = r.get('unit_pieces', r.get('qty', bv('qty'))) or '—'
            pages = r.get('pages', bv('pages')) or '—'
            reps = r.get('repetitions', '') or '—'
            clean = r.get('clean', '') or '—'
            waste = r.get('waste', '') or '—'
            whole = r.get('whole_sheets', '') or '—'
            source_sheets = r.get('source_sheets', '') or '—'
            color = f"{bv('front')} + {bv('back')}" if (bv('front') or bv('back')) else '—'
            turnover = str(r.get('turnover', '')).strip() or '—'
            size = str(r.get('trim', '—')).strip() or '—'
            source_format = str(r.get('source', bv('source'))).strip() or '—'
            print_format = str(r.get('print_format', '—')).strip() or '—'
            paper_price = float(r.get('paper_price', 0) or 0)
            if not paper_price:
                try: paper_price = float(bv('paper').replace(',', '.')) if bv('paper') else 0.0
                except Exception: paper_price = 0.0
            paper_price_text = f'{paper_price:.3f} € цена/лист' if paper_price else '—'
            # За Книжки в Заявка показваме четирите задължителни операции
            # от самия таб, като запазваме и броя гънки.
            finish = []
            if bv('sewing').lower() not in ('', 'без', 'не'):
                finish.append('Шиене')
            if bv('typesetting').lower() not in ('', 'без', 'не'):
                finish.append('Набор')
            if bv('folding').lower() not in ('', 'без', 'не') and float(r.get('fold_count', 0) or 0) > 0:
                finish.append(f"Сгъване ({self._fmt_count(r.get('fold_count'))} бр.)")
            if bv('inserting').lower() == 'да':
                finish.append('Влагане')
            finish_text = ' • '.join(finish) if finish else '—'

            _, prod = self._request_card(sheet, 'Производствени параметри', 0, 1, 1)
            self._request_production_display(prod, active, r)

            _, finish_card = self._request_card(sheet, 'Довършителни работи', 1, 1, 1)
            tk.Label(finish_card, text=finish_text, bg='#FFFFFF', fg='#344054',
                     font=('Segoe UI',10), justify='left', anchor='w', wraplength=300).pack(
                         fill='x', padx=5, pady=4)

            _, res = self._request_card(sheet, 'Резултати', 2, 1, 1)
            res.columnconfigure(0, weight=1)
            self._request_result_stats(res, r, active)
            return

        # --- Кочани / Химия ---
        if active == 'Кочани' and getattr(self, 'himiya_result', None):
            r = self.himiya_result
            hv = lambda k: str(self._hvar(k).get()).strip() if k in getattr(self, 'himiya_vars', {}) else ''
            def hn(k, default=0):
                try: return float(hv(k).replace(',', '.')) if hv(k) else float(default)
                except Exception: return float(default)
            paper_g = self.request_vars['paper'].get().strip().upper(); paper_type = self.request_vars['paper_type'].get()
            paper_parts = []
            if paper_g and paper_g != '-':
                paper_parts.append(f'{paper_g} г')
            if paper_type and paper_type != '-':
                paper_parts.append(paper_type)
            paper_manual = ' '.join(paper_parts) or '—'
            qty = int(r.get('quantity') or hn('qty')); sheets_block = int(r.get('sheets_per_block') or hn('sheets'))
            paper_colors = int(r.get('paper_colors') or hn('paper_colors', 1)); clean = r.get('clean_sheets',''); whole = r.get('whole_sheets',''); reps = r.get('repetitions','')
            paper_price = hn('paper'); paper_price_text = f'{paper_price:.3f} € цена/лист' if paper_price else '—'
            color = f"{hv('front')} + {hv('back')}" if (hv('front') or hv('back')) else '—'; turnover = hv('turnover') or '—'
            size = f"{r.get('product_w',0):g}×{r.get('product_h',0):g} мм"; total = float(r.get('total',0) or 0)
            finish = [x for x in self._client_offer_finish_list() if x not in ('Ел. монтаж','Транспорт','Пътни разходи') and not x.startswith('Друго')]
            # Разделителите се показват с конкретния резултат от калкулацията.
            sep_sheets = int(r.get('separator_sheets', 0) or 0)
            if hv('sep_mat').lower() not in ('', 'без', 'не') and sep_sheets:
                finish = [f'картон - {sep_sheets} цели листа' if x == 'Разделяне' else x for x in finish]
            finish_text = ' • '.join(finish) if finish else '—'

            # За Кочани данните се взимат изцяло от himiya_result / himiya_vars,
            # а не от self.result на таб „Основен“.
            _, prod = self._request_card(sheet, 'Производствени параметри', 0, 1, 1)
            self._request_production_display(prod, active, r)

            _, mat = self._request_card(sheet, 'Материали', 1, 1, 1)
            mat.columnconfigure(0, weight=1)
            material_rows = [
                ('ХАРТИЯ', f'{paper_price:.3f} € цена/лист' if paper_price else ''),
                ('ХАРТИЯ', paper_manual),
            ]
            for i, (lab, val) in enumerate(material_rows):
                if str(val).strip() and str(val).strip() not in ('0', '0.000', '—'):
                    self._request_field(mat, lab, val, i, 0, value_bold=(i == 0))

            # Довършителни + допълнителни разходи използват една обща колона.
            finish_zone = tk.Frame(sheet, bg='#F7F9FC', bd=0, highlightthickness=0)
            finish_zone.grid(row=1, column=2, rowspan=2, sticky='nsew', padx=6, pady=4)
            finish_zone.grid_rowconfigure(0, weight=1)
            finish_zone.grid_rowconfigure(1, weight=0)
            finish_zone.grid_columnconfigure(0, weight=1)
            finish_card, finish_content = self._request_card(finish_zone, 'Довършителни работи', 0, 0, 1)
            finish_content.columnconfigure(0, weight=1)
            tk.Label(finish_content, text=finish_text, bg='#FFFFFF', fg='#344054',
                     font=('Segoe UI',10), justify='left', anchor='nw', wraplength=300).pack(
                         fill='both', expand=True, padx=5, pady=4)

            _, extra = self._request_card(finish_zone, 'Допълнителни разходи', 0, 1, 1)
            extra.columnconfigure(0, weight=1)
            transport_total = float(r.get('transport', 0) or 0)
            surcharge_total = float(r.get('surcharge', 0) or 0)
            extra_rows=[]
            if abs(transport_total) > 1e-12:
                extra_rows.append(('ТРАНСПОРТ', f'{transport_total:.2f} €'))
            if abs(surcharge_total) > 1e-12:
                extra_rows.append(('ОСКЪПЯВАНЕ НА ТРУДА', f'{surcharge_total:.2f} €'))
            for i,(lab,val) in enumerate(extra_rows):
                self._request_field(extra, lab, val, i, 0, value_bold=True)

            _, res = self._request_card(sheet, 'Резултати', 3, 1, 1)
            res.columnconfigure(0, weight=1)
            self._request_result_stats(res, r, active)
            return

        # --- Спирали ---
        if active == 'Спирали' and getattr(self, 'spiral_result', None):
            r=self.spiral_result; sv=lambda k: str(self._spiral_var(k,'').get()).strip()
            size=str(r.get('size','—')); qty=f"{self._fmt_count(r.get('qty',0))} бр."; total=float(r.get('total',0) or 0); unit=float(r.get('unit',0) or 0)
            finish=[]
            if sv('spiral_perfo_x2').lower()=='да': finish.append('Спирала и перфо ×2')
            if sv('cut_per_tooth').lower()=='да': finish.append('За нарязване/зъб')
            if sv('hooks').lower() not in ('','без'): finish.append('Закачане + кукички')
            finish_text=' • '.join(str(x).upper() for x in finish) if finish else '—'
            _, prod = self._request_card(sheet, 'Производствени параметри', 0, 1, 1)
            self._request_production_display(prod, active, r)

            _, finish_card = self._request_card(sheet, 'Довършителни работи', 1, 1, 1)
            tk.Label(finish_card, text=finish_text, bg='#FFFFFF', fg='#344054',
                     font=('Segoe UI',10), justify='left', anchor='w', wraplength=300).pack(
                         fill='x', padx=5, pady=4)

            _, res = self._request_card(sheet, 'Резултати', 2, 1, 1)
            res.columnconfigure(0, weight=1)
            self._request_result_stats(res, r, active)
            return

        # --- Календари ---
        if active == 'Календари' and getattr(self, 'calendar_result', None):
            r=self.calendar_result
            cv=lambda k: str(self._cvar(k,'').get()).strip() if k in getattr(self,'calendar_vars',{}) else ''
            def cnum(k, default=0):
                try: return float(cv(k).replace(',', '.')) if cv(k) else float(default)
                except Exception: return float(default)

            color=f"{r.get('color_front','0')} + {r.get('color_back','0')}"
            turnover=str(r.get('turnover','—')).strip() or '—'
            paper_price=float(r.get('paper_price',0) or 0)
            paper_price_text=f'{paper_price:.3f} € цена/лист' if paper_price else '—'
            paper_manual=str(self.request_vars['paper'].get() or '').strip() or '—'
            whole=r.get('whole_sheets','')
            waste=r.get('waste','')
            cols=r.get('cols','')
            change_count=r.get('color_change_count',cnum('color_change_count'))
            change_opt=cv('color_change_opt').lower()
            change_text=self._fmt_count(change_count) if change_opt == 'да' and change_count else ('ДА' if change_opt == 'да' else 'НЕ')
            finish=[]
            fc=r.get('finish_costs',{}) or {}
            for name,cost in fc.items():
                if float(cost or 0)>0:
                    finish.append(name)
            # В „Заявка“ разделителите се показват само ако е избран материал.
            # Материалът не се изписва — показва се само изчисленият брой листове.
            sep_material=str(r.get('separator_material','')).strip().lower()
            sep_sheets=int(r.get('separator_sheets',0) or 0)
            if sep_material not in ('','без','не') and sep_sheets>0:
                finish.append(f'Разделители ({sep_sheets} листа)')
            finish_text=' • '.join(finish) if finish else '—'

            _, prod = self._request_card(sheet, 'Производствени параметри', 0, 1, 1)
            calendar_rows=[
                ('ЕД. БРОЙКИ', f"{self._fmt_count(r.get('qty',''))} бр."),
                ('ЦВЕТНОСТ', color),
                ('ФОРМАТ (НА Х-Я)', str(r.get('source_format','—'))),
                ('ФОРМАТ ЗА ПЕЧАТ', str(r.get('print_format','—'))),
                ('БР. СТРАНИЦИ', self._fmt_count(r.get('pages',''))),
                ('ПЕЧАТНИ КОЛИ', self._fmt_count(cols)),
                ('СМЯНА НА ЦВЯТ', change_text),
                ('ОБРЪЩАНЕ', turnover.upper()),
                ('РАЗМНОЖЕНИЯ', self._fmt_count(r.get('repetitions',''))),
                ('ЛИСТА (ВКЛ. МАКУЛАТУРА)', f"{self._fmt_count(waste)} листа (вкл. макулатура)"),
                ('ОБРЯЗАН РАЗМЕР', str(r.get('trim_size','—'))),
            ]
            for idx,(lab,val) in enumerate(calendar_rows):
                col = idx % 2
                grid_row = idx // 2
                # Всички автоматично получени резултати в производствените
                # параметри са удебелени, не само ЕД. БРОЙКИ.
                self._request_field(prod, lab, val, grid_row, col, value_bold=True)
            prod.columnconfigure(0, weight=1)
            prod.columnconfigure(1, weight=1)

            _, mat = self._request_card(sheet, 'Материали', 1, 1, 1)
            material_rows=[('ХАРТИЯ ЦЕНА',paper_price_text),('ХАРТИЯ',paper_manual)]
            for i,(lab,val) in enumerate(material_rows):
                if str(val).strip() and str(val).strip()!='—':
                    self._request_field(mat, lab, val, i, 0, value_bold=(i==0))

            _, finish_card = self._request_card(sheet, 'Довършителни работи', 2, 1, 1)
            tk.Label(finish_card, text=finish_text, bg='#FFFFFF', fg='#344054',
                     font=('Segoe UI',10), justify='left', anchor='nw', wraplength=250).pack(
                         fill='both', expand=True, padx=5, pady=4)

            _, res = self._request_card(sheet, 'Резултати', 3, 1, 1)
            res.columnconfigure(0, weight=1)
            self._request_result_stats(res, r, active)
            return

        # --- Основен таб ---
        if active == 'Основен' and getattr(self, 'result', None):
            r=self.result; v=self._request_values()
            # Производствени параметри — леко разширена колона.
            _, prod = self._request_card(sheet, 'Производствени параметри', 0, 1, 1)
            self._request_production_display(prod, active, r)

            # Материалите вече са самостоятелна колона между производството
            # и довършителните работи. Допълнителните разходи остават под тях,
            # за да не се увеличава излишно височината на екрана.
            def euro(key):
                try: return float(r.get(key, 0) or 0)
                except Exception: return 0.0

            paper_total = euro('paper')
            print_total = euro('print')
            prepress_total = euro('prepress')
            other_total = euro('other')
            transport_total = euro('transport')
            surcharge_total = euro('surcharge')
            excluded = {'paper','print','prepress','transport','surcharge','other'}
            finish_keys = ('duplication','plates','cutting','numbering','perforation','uv','lamination',
                           'calender','film','gluing','bigoving','folding','die_cut','round_punch',
                           'breaking_cost','electric_montage','counting','typesetting','separators')
            finish_total = sum(euro(k) for k in finish_keys)

            _, mat = self._request_card(sheet, 'Материали', 1, 1, 1)
            for c in range(2): mat.columnconfigure(c, weight=1)
            material_rows=[('ХАРТИЯ',paper_total),('ПЛАКИ',euro('plates'))]
            material_rows=[(lab,val) for lab,val in material_rows if abs(val) > 1e-12]
            for i,(lab,val) in enumerate(material_rows):
                self._request_field(mat, lab, f'{val:.2f} €', i, 0, value_bold=(lab=='ХАРТИЯ'))

            finish_text=v['finish'] or '—'

            # Довършителни работи + Допълнителни разходи използват една обща
            # височина в колона 2. Довършителните заемат свободното място,
            # а допълнителните разходи остават компактни в долната част.
            finish_zone = tk.Frame(sheet, bg='#F7F9FC', bd=0, highlightthickness=0)
            finish_zone.grid(row=1, column=2, rowspan=2, sticky='nsew', padx=6, pady=4)
            finish_zone.grid_rowconfigure(0, weight=1)
            finish_zone.grid_rowconfigure(1, weight=0)
            finish_zone.grid_columnconfigure(0, weight=1)

            finish_card, finish_content = self._request_card(finish_zone, 'Довършителни работи', 0, 0, 1)
            finish_content.columnconfigure(0, weight=1)
            tk.Label(finish_content, text=finish_text, bg='#FFFFFF', fg='#344054',
                     font=('Segoe UI',10), justify='left', anchor='nw', wraplength=280).pack(
                         fill='both', expand=True, padx=5, pady=4)

            _, extra = self._request_card(finish_zone, 'Допълнителни разходи', 0, 1, 1)
            extra.columnconfigure(0, weight=1)
            extra_rows=[]
            if abs(transport_total) > 1e-12:
                extra_rows.append(('ТРАНСПОРТ', transport_total))
            if abs(surcharge_total) > 1e-12:
                extra_rows.append(('ОСКЪПЯВАНЕ НА ТРУДА', surcharge_total))
            for i,(lab,val) in enumerate(extra_rows):
                self._request_field(extra, lab, f'{val:.2f} €', i, 0)

            _, res = self._request_card(sheet, 'Резултати', 3, 1, 1)
            res.columnconfigure(0, weight=1)
            self._request_result_stats(res, r, active)
            return

        self._request_empty_state(sheet)

    def _client_offer_finish_list(self):
        """Връща само имената на довършителните операции от горното каре.
        Тази информация е общата клиентска част на резултата и не съдържа цени.
        """
        source = getattr(self, 'request_source_tab', getattr(self, 'active_tab_name', 'Основен'))
        if source == 'Кочани' and self.himiya_result:
            selected=[]
            def hv(key):
                return str(self._hvar(key).get() if key in getattr(self, 'himiya_vars', {}) else '').strip()
            def hn(key, default=0):
                try:
                    return float(hv(key).replace(',', '.')) if hv(key) else float(default)
                except Exception:
                    return float(default)
            if hv('cutting').lower() not in ('', 'без', 'не'): selected.append('Рязане')
            if hn('big_count') > 0: selected.append('Биговане')
            if hv('gluing').lower() not in ('', 'без', 'не'): selected.append('Лепене')
            if hv('numbering').lower() not in ('', 'без', 'не'): selected.append('Номерация')
            if hv('perforation').lower() not in ('', 'без', 'не'): selected.append('Перфорация')
            if hv('typesetting').lower() not in ('', 'без', 'не'): selected.append('Набор')
            if hv('sewing').lower() not in ('', 'без', 'не'): selected.append('Шиене/телчета')
            if hv('em').lower() not in ('', 'без', 'не'): selected.append('Ел. монтаж')
            if hv('counting').lower() not in ('', 'без', 'не'): selected.append('Пакетиране')
            if hv('sep_mat').lower() not in ('', 'без', 'не'): selected.append('Разделяне')
            if hv('transport').lower() not in ('', 'без', 'не'): selected.append('Транспорт')
            if hn('other') > 0: selected.append('Друго')
            return selected

        # Книжки — собственен резултат и собствени полета.
        if source == 'Книжки' and getattr(self, 'book_result', None):
            selected=[]
            def bv(key):
                try:
                    return str(self.book_vars[key].get()).strip() if key in self.book_vars else ''
                except Exception:
                    return ''
            def bn(key):
                try:
                    return float(bv(key).replace(',', '.')) if bv(key) else 0.0
                except Exception:
                    return 0.0
            if bn('cover_price') > 0: selected.append('Корица')
            if bv('sewing').lower() not in ('', 'без', 'не'): selected.append('Шиене')
            if bv('typesetting').lower() not in ('', 'без', 'не'): selected.append('Набиране')
            if bv('folding').lower() not in ('', 'без', 'не') and bn('fold_count') > 0: selected.append('Сгъване')
            if bv('inserting').lower() == 'да': selected.append('Влагане')
            if bv('electric_montage').lower() not in ('', 'без', 'не'): selected.append('Ел. монтаж')
            if bv('numbering').lower() not in ('', 'без', 'не'): selected.append('Номерация')
            if bv('counting').lower() not in ('', 'без', 'не'): selected.append('Броене/пакетиране')
            if bn('bigoving') > 0: selected.append('Биговане')
            if bv('gluing').lower() not in ('', 'без', 'не'): selected.append('Лепене')
            if bv('uv').lower() not in ('', 'без', 'не'): selected.append('УВ лак / надпечат')
            if bv('lamination').lower() not in ('', 'без', 'не'): selected.append('Ламиниране')
            if bv('calender').lower() not in ('', 'без к', 'без', 'не'): selected.append('Каландър')
            if bv('cutting').lower() not in ('', 'без', 'не'): selected.append('Обрязване')
            if bn('prepress') > 0: selected.append('Предпечат')
            if bn('transport') > 0: selected.append('Пътни разходи')
            if bn('surcharge') > 0: selected.append('Оскъпяване')
            return selected

        if source == 'Спирали' and getattr(self, 'spiral_result', None):
            selected=[]
            sv=lambda k: str(self._spiral_var(k,'').get()).strip()
            if sv('spiral_perfo_x2').lower() == 'да': selected.append('Спирала и перфо ×2')
            if sv('cut_per_tooth').lower() == 'да': selected.append('За нарязване/зъб')
            if sv('hooks').lower() not in ('','без'): selected.append('Закачане + кукички')
            if self._spiral_var('courier','').get().strip(): selected.append('Куриер')
            if self._spiral_var('outside_standard','').get().strip(): selected.append('Такса извън стандарт')
            return selected

        selected=[]
        if source != 'Основен':
            return selected
        def val(key):
            return str(self.vars[key].get()).strip() if key in self.vars else ''
        def num(key):
            try: return float(val(key).replace(',', '.')) if val(key) else 0
            except Exception: return 0
        if val('cutting').lower() not in ('', 'без', 'не'): selected.append('Рязане')
        if num('folding') > 0: selected.append('Сгъване')
        if num('bigoving') > 0: selected.append('Биговане')
        if val('counting').lower() not in ('', 'без', 'не'): selected.append('Пакетиране')
        if val('sepmat').lower() not in ('', 'без', 'не'): selected.append('Разделяне')
        if val('gluing').lower() not in ('', 'без', 'не'): selected.append('Лепене')
        if val('uv').lower() not in ('', 'без', 'не'): selected.append('УВ лак')
        if val('lamination').lower() not in ('', 'без', 'не'): selected.append('Ламиниране')
        if val('die').lower() not in ('', 'без', 'не'): selected.append('Щанцоване')
        if val('electric_montage').lower() not in ('', 'без', 'не'): selected.append('Ел. монтаж')
        if val('typesetting').lower() not in ('', 'без', 'не'): selected.append('Набор')
        if val('numbering').lower() not in ('', 'без', 'не'): selected.append('Номерация')
        if val('perforation').lower() not in ('', 'без', 'не'): selected.append('Перфорация')
        if val('calender').lower() not in ('', 'без к', 'без', 'не'): selected.append('Каландър')
        if val('film_price') not in ('', '0', '0.0', '0,0'): selected.append('Филм')
        if val('other_operations'): selected.append('Друго')
        # ТУК ДОБАВЕТЕ ТОЗИ РЕД ЗА ТАБ ОСНОВЕН:
        selected = [x for x in selected if x != 'Ел. монтаж']
        return selected

    def _save_client_offer(self):
        """Записва клиентска оферта от горните две клиентски карета на Резултат.
        Не използва формата/съдържанието на "Заявка" освен за изделие, клиент и дата.
        """
        try:
            # Всеки изчислителен таб има собствен резултат. Не използваме
            # self.result/self.himiya_result за „Книжки“, защото те са отделни.
            is_book = bool(getattr(self, 'book_result', None)) and getattr(self, 'active_tab_name', '') == 'Книжки'
            is_spiral = bool(getattr(self, 'spiral_result', None)) and getattr(self, 'active_tab_name', '') == 'Спирали'
            is_calendar = bool(getattr(self, 'calendar_result', None)) and getattr(self, 'active_tab_name', '') == 'Календари'
            is_h = (not is_book and not is_spiral) and getattr(self, 'last_result_mode', 'order') == 'himiya' and bool(self.himiya_result)
            if is_book:
                r = self.book_result
            elif is_spiral:
                r = self.spiral_result
            elif is_calendar:
                r = self.calendar_result
            elif is_h:
                r = self.himiya_result
            else:
                r = self.result

            if not r:
                raise ValueError('Първо направете изчисление.')

            # Данните, които са общи и за двете клиентски карета.
            item = self.request_vars.get('item').get().strip() if hasattr(self, 'request_vars') and 'item' in self.request_vars else ''
            client = self.request_vars.get('client').get().strip() if hasattr(self, 'request_vars') and 'client' in self.request_vars else ''
            date = self.request_vars.get('date').get().strip() if hasattr(self, 'request_vars') and 'date' in self.request_vars else ''
            paper = self.request_vars.get('paper').get().strip() if hasattr(self, 'request_vars') and 'paper' in self.request_vars else ''

            if is_book:
                front = self.book_vars.get('front', tk.StringVar(value='')).get().strip()
                back = self.book_vars.get('back', tk.StringVar(value='')).get().strip()
                color = f'{front} + {back}' if (front or back) else '—'
                size = f"{r.get('trim','—')}"
                qty = f"{r.get('unit_pieces', self.book_vars.get('qty', tk.StringVar(value='')).get() or '—')} бр."
                turnover = str(r.get('turnover','')).strip() or 'НЕ'
                print_format = str(r.get('print_format','—'))
            elif is_spiral:
                size = str(r.get('size','—'))
                qty = f"{self._fmt_count(r.get('qty',0))} бр."
                color = '—'
                turnover = '—'
                print_format = '—'
                spiral_teeth = str(self._spiral_var('teeth', '').get()).strip() or '—'
                body_gsm = str(self._spiral_var('body_gsm', '').get()).strip() or '—'
                body_sheets = str(self._spiral_var('body_sheets', '').get()).strip() or '—'
                cover_gsm = str(self._spiral_var('cover_gsm', '').get()).strip() or '—'
                cover_sheets = str(self._spiral_var('cover_sheets', '').get()).strip() or '—'
                hooks = str(self._spiral_var('hooks', '').get()).strip()
            elif is_calendar:
                size = str(r.get('trim_size', '—')).strip() or '—'
                qty = f"{self._fmt_count(r.get('qty', 0))} бр."
                cf = str(r.get('color_front', '0')).strip()
                cb = str(r.get('color_back', '0')).strip()
                color = f'{cf} + {cb}' if (cf or cb) else '—'
                turnover = str(r.get('turnover', '—')).strip() or '—'
                print_format = str(r.get('print_format', '—')).strip() or '—'
            elif is_h:
                front = self._hvar('front').get() or ''
                back = self._hvar('back').get() or ''
                color = f'{front} + {back}'.strip(' +')
                size = f"{r.get('product_w',0):g} × {r.get('product_h',0):g} мм"
                qty = (f"{r.get('quantity', '')} кочана х {r.get('sheets_per_block', self._hvar('sheets').get() or '')} л." if r.get('quantity') else f"{r.get('unit_pieces','')} бр.")
                turnover = self._hvar('turnover').get() or '—'
                print_format = str(r.get('print_format','—'))
            else:
                front = self.vars.get('front', tk.StringVar(value='')).get().strip()
                back = self.vars.get('back', tk.StringVar(value='')).get().strip()
                color = f'{front} + {back}' if (front or back) else '—'
                size = f"{r.get('product_w',0):g} × {r.get('product_h',0):g} мм"
                qty = f"{r.get('unit_pieces','')} бр."
                turnover = self.vars.get('turnover', tk.StringVar(value='')).get().strip() or '—'
                print_format = str(r.get('print_format','—'))

            if is_calendar:
                # В календарната оферта показваме само реално начислените довършителни операции.
                # Изрично не включваме Ел. монтаж, Разделители, Транспорт и Оскъпяване.
                excluded = {'Ел. монтаж', 'Разделители', 'Транспорт', 'Оскъпяване'}
                finish_costs = r.get('finish_costs', {})
                # В клиентската оферта се изреждат само операциите, без показване на цени.
                finishes = [name for name, cost in finish_costs.items()
                            if name not in excluded and float(cost or 0) > 0]
                finish_text = ', '.join(finishes) if finishes else '—'
            else:
                finishes = self._client_offer_finish_list()
                if is_book:
                    # В офертата за „Книжки“ не показваме Ел. монтаж и Транспорт
                    # (както и вече изключените Корица, Предпечат и Оскъпяване).
                    finishes = [x for x in finishes if x not in (
                        'Корица', 'Ел. монтаж', 'Предпечат', 'Пътни разходи',
                        'Транспорт', 'Оскъпяване'
                    )]
                elif is_h:
                    # В офертата за „Кочани“ не показваме Ел. монтаж и Транспорт.
                    finishes = [x for x in finishes if x not in ('Ел. монтаж', 'Транспорт')]
                finish_text = ', '.join(finishes) if finishes else '—'

            lines = [
                'КЛИЕНТСКА ОФЕРТА',
                '',
                f'Изделие: {item or "—"}',
                f'Клиент: {client or "—"}',
                f'Дата: {date or "—"}',
                f'Хартия: {paper or "—"}',
                '',
                'ОСНОВНИ ПАРАМЕТРИ',
                *([f'Размер на спирала: {size}',
                   f'Ед. бройки: {self._fmt_count(r.get("qty", 0))} бр.'] if is_spiral else [f'Размер: {size}', f'Тираж: {qty}']),
                *([f'Цветност: {color}',
                   f'Обръщане: {turnover}',
                   f'Печатен формат: {print_format}'] if not is_spiral else []),
                *([f'Страници: {r.get("pages", self.book_vars.get("pages", tk.StringVar(value="")).get() or "—")}',
                   f'Размножения: {r.get("repetitions", "—")}'] if is_book or is_calendar else []),
                *([f'Брой коли: {self._fmt_count(r.get("cols", "—"))}'] if is_book or is_calendar else []),
                *([f'Брой цветове листа: {color}'] if is_h else []),
                *([f'Брой зъби: {spiral_teeth}',
                   f'Тяло: {body_sheets} листа / {body_gsm} грамаж',
                   f'Корица: {cover_sheets} листа / {cover_gsm} грамаж',
                   f'Кукички: {hooks or "без"}'] if is_spiral else []),
                *(['', 'ДОВЪРШИТЕЛНИ ОПЕРАЦИИ', finish_text] if not is_spiral else []),
                '',
                f'Крайна цена: {float(r.get("total",0)):.2f} €',
                f'Цена / бр.: {float(r.get("unit",0)):.4f} €',
            ]

            path = filedialog.asksaveasfilename(
                title='Запази клиентската оферта като текстов файл',
                defaultextension='.txt',
                filetypes=[('Текстов файл', '*.txt'), ('Всички файлове', '*.*')],
                initialfile=f'Клиентска_оферта_{client or ""}.txt'
            )
            if not path:
                return
            Path(path).write_text('\n'.join(lines), encoding='utf-8')
        except Exception as e:
            messagebox.showerror('Грешка при запис на клиентска оферта', str(e))


    def _save_offer_text(self):
        """Записва „Заявка“ като TXT по стария шаблон на заявката.
        Конструкцията е еднаква за Таб 1 и Химия; различават се само данните.
        """
        try:
            # Източникът на „Заявка“ се пази отделно, за да не се записват
            # случайно стойности от „Основен“/„Кочани“, когато последно е
            # изчисляван „Книжки“.
            source = getattr(self, 'request_source_tab', getattr(self, 'active_tab_name', 'Основен'))
            is_book = source == 'Книжки' and bool(getattr(self, 'book_result', None))
            is_h = source == 'Кочани' and bool(self.himiya_result)
            is_calendar = source == 'Календари' and bool(getattr(self, 'calendar_result', None))
            is_spiral = source == 'Спирали' and bool(getattr(self, 'spiral_result', None))
            is_main = source == 'Основен' and bool(self.result)
            if not (is_book or is_h or is_calendar or is_spiral or is_main):
                raise ValueError('Първо направете изчисление в съответния таб.')

            # Старият TXT шаблон („Невена“): две колони по 38 знака.
            CELL_W = 38
            INNER_W = CELL_W * 2

            def fit(value, width=CELL_W, align='left'):
                text = str(value if value is not None else '')
                text = text.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')
                if len(text) > width:
                    text = text[:max(0, width - 1)] + '…'
                if align == 'right':
                    return text.rjust(width)
                if align == 'center':
                    return text.center(width)
                return text.ljust(width)

            def border():
                return '+' + '-' * CELL_W + '+' + '-' * CELL_W + '+'

            def row(left='', right=''):
                # Точно като стария TXT: '| ' + 38 знака + '| ' + 38 знака + '|'
                return f'| {fit(left)}| {fit(right)}|'

            def full(text='', center=False):
                # Старият шаблон използва общо поле между двете колони.
                return f'| {fit(text, INNER_W, "center" if center else "left")} |'

            def full_wrapped(text='', center=False):
                # За дълги списъци (напр. много довършителни операции)
                # текстът продължава на следващ ред вместо да се отрязва.
                raw = str(text if text is not None else '')
                raw = raw.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')
                import textwrap
                chunks = textwrap.wrap(
                    raw, width=INNER_W, break_long_words=False,
                    break_on_hyphens=False, replace_whitespace=True
                ) or ['']
                return [fit(chunk, INNER_W, 'center' if center else 'left') for chunk in chunks]

            lines = [border(), full('ЗАЯВКА', True), border()]

            if is_book:
                r = self.book_result
                bv = lambda k: str(self.book_vars.get(k).get()).strip() if k in getattr(self, 'book_vars', {}) else ''
                item = str(self.request_vars['item'].get() or '').strip()
                client = str(self.request_vars['client'].get() or '').strip()
                date = str(self.request_vars['date'].get() or '').strip()
                qty = r.get('unit_pieces', r.get('qty', bv('qty'))) or '—'
                pages = r.get('pages', bv('pages')) or '—'
                reps = r.get('repetitions', '') or '—'
                clean = r.get('clean', '') or '—'
                waste = r.get('waste', '') or '—'
                whole = r.get('whole_sheets', '') or '—'
                source_sheets = r.get('source_sheets', '') or '—'
                paper_price = float(r.get('paper_price', 0) or 0)
                if not paper_price:
                    try: paper_price = float(bv('paper').replace(',', '.')) if bv('paper') else 0.0
                    except Exception: paper_price = 0.0
                paper_price_text = f'{paper_price:.3f} € цена/ лист' if paper_price else '—'
                paper_g = str(self.request_vars['paper'].get() or '').strip()
                paper_type = str(self.request_vars['paper_type'].get() or '').strip()
                if paper_type:
                    try:
                        float(paper_g.replace(',', '.'))
                        paper_manual = f'{paper_g} г {paper_type}'
                    except Exception:
                        paper_manual = f'{paper_g} {paper_type}'.strip()
                elif paper_g:
                    paper_manual = paper_g
                else:
                    paper_manual = '—'
                front, back = bv('front'), bv('back')
                color = f'{front} + {back}' if (front or back) else '—'
                turnover = str(r.get('turnover', '')).strip() or '—'
                size = str(r.get('trim', '—')).strip() or '—'
                finish = self._client_offer_finish_list()
                finish = [x for x in finish if x not in ('Корица','Ел. монтаж','Предпечат','Пътни разходи','Транспорт','Оскъпяване')]
                finish_line = ' ● '.join(finish) if finish else '—'

                lines += [
                    row(f'КЛИЕНТ: {client or "—"}', f'ДАТА: {date or "—"}'), border(),
                    row(f'ИЗДЕЛИЕ: {item or "—"}', f'ЦЕНА БЕЗ ДДС: {float(r.get("total",0) or 0):.2f}'), border(),
                    row(f'ЕД. БРОЙКИ: {qty} бр.', f'ед. бройка: {float(r.get("unit",0) or 0):.4f}'), border(),
                    row(f'СТРАНИЦИ: {pages}', f'размножения: {reps}'), border(),
                    row(f'ХАРТИЯ: {paper_price_text}', f'формат (на х-я): {r.get("source", bv("source")) or "—"}'), border(),
                    row(paper_manual, f'формат за ПЕЧАТ: {r.get("print_format", "—")}'), border(),
                    row(f'ПЕЧАТНИ КОЛИ: {r.get("cols", "—")}', f'ТИРАЖ: {clean}'), border(),
                    row(f'ТИРАЖ + МАКУЛАТУРА: {waste} листа (вкл. макулатура)', f'ЦЕЛИ ЛИСТА: {whole} листа'), border(),
                    row(f'ЦВЕТНОСТ: {color}', f'ОБРЪЩАНЕ: {turnover.upper() if turnover else "—"}'), border(),
                    row(f'ОБРЯЗАН РАЗМЕР: {size}', ''), border(),
                    full('ДОВЪРШИТЕЛНИ РАБОТИ', True),
                    *[f'| {line} |' for line in full_wrapped(f'{finish_line}')]
                ]
            elif is_calendar:
                r = self.calendar_result
                cv = lambda k: str(self._cvar(k,'').get()).strip() if k in getattr(self,'calendar_vars',{}) else ''
                def cnum(k, default=0):
                    try:
                        return float(cv(k).replace(',', '.')) if cv(k) else float(default)
                    except Exception:
                        return float(default)

                item = str(self.request_vars['item'].get() or '').strip()
                client = str(self.request_vars['client'].get() or '').strip()
                date = str(self.request_vars['date'].get() or '').strip()
                qty = r.get('qty',0)
                total = float(r.get('total',0) or 0)
                unit = float(r.get('unit',0) or 0)
                color = f"{r.get('color_front','0')} + {r.get('color_back','0')}"
                turnover = str(r.get('turnover','—')).strip() or '—'
                paper_price = float(r.get('paper_price',0) or 0)
                paper_price_text = f'{paper_price:.3f} € цена/ лист' if paper_price else '—'
                paper_manual = str(self.request_vars['paper'].get() or '').strip() or '—'
                cols = r.get('cols','—')
                reps = r.get('repetitions','—')
                waste = r.get('waste','')
                change_count = r.get('color_change_count', cnum('color_change_count'))
                change_opt = cv('color_change_opt').lower()
                change_text = self._fmt_count(change_count) if change_opt == 'да' and change_count else ('ДА' if change_opt == 'да' else 'НЕ')
                finish = [name for name,cost in (r.get('finish_costs',{}) or {}).items() if float(cost or 0)>0]
                finish_line = ' ● '.join(str(x).strip() for x in finish if str(x).strip()) or '—'

                lines += [
                    row(f'ИЗДЕЛИЕ: {item or "Календари"}', f'КЛИЕНТ: {client or "—"}'), border(),
                    row(f'ОБРЯЗАН РАЗМЕР: {r.get("trim_size", "—")}', f'ДАТА: {date or "—"}'), border(),
                    row(f'ЕД. БРОЙКИ: {self._fmt_count(qty)} бр.', f'ЦЕНА БЕЗ ДДС: {total:.2f} €'), border(),
                    row(f'ЦВЕТНОСТ: {color}', f'ед. бройка: {unit:.4f} €'), border(),
                    row(f'ХАРТИЯ ЦЕНА: {paper_price_text}', f'формат (на х-я): {r.get("source_format", "—")}'), border(),
                    row(f'ХАРТИЯ: {paper_manual}', f'формат за печат: {r.get("print_format", "—")}'), border(),
                    row(f'бр. страници: {self._fmt_count(r.get("pages", ""))}', f'ПЕЧАТНИ КОЛИ: {self._fmt_count(cols)}'), border(),
                    row(f'СМЯНА НА ЦВЯТ: {change_text}', f'размножения: {self._fmt_count(reps)}'), border(),
                    row(f'ОБРЪЩАНЕ: {turnover.upper()}', f'{self._fmt_count(waste)} листа (вкл. макулатура)'), border(),
                    full('ДОВЪРШИТЕЛНИ РАБОТИ', True),
                    *[f'| {line} |' for line in full_wrapped(f'{finish_line}')]
                ]
            elif is_h:
                r = self.himiya_result
                hv = lambda k: str(self._hvar(k).get()).strip() if k in getattr(self, 'himiya_vars', {}) else ''
                def hn(k, default=0):
                    try:
                        return float(hv(k).replace(',', '.')) if hv(k) else float(default)
                    except Exception:
                        return float(default)

                item = str(self.request_vars['item'].get() or '').strip()
                client = str(self.request_vars['client'].get() or '').strip()
                date = str(self.request_vars['date'].get() or '').strip()
                paper_g = str(self.request_vars['paper'].get() or '').strip()
                paper_type = str(self.request_vars['paper_type'].get() or '').strip()

                qty = int(r.get('quantity') or hn('qty'))
                sheets_block = int(r.get('sheets_per_block') or hn('sheets'))
                paper_colors = int(r.get('paper_colors') or hn('paper_colors', 1))
                clean = r.get('clean_sheets', '')
                whole = r.get('whole_sheets', '')
                reps = r.get('repetitions', '')
                paper_price = hn('paper')
                paper_price_text = f'{paper_price:.3f} € цена/ лист' if paper_price else '—'
                front, back = hv('front'), hv('back')
                color = f'{front} + {back}' if (front or back) else '—'
                turnover = hv('turnover') or '—'
                size = f"{r.get('product_w', 0):g}×{r.get('product_h', 0):g} мм"
                total = float(r.get('total', 0) or 0)
                unit = float(r.get('unit', 0) or 0)

                if paper_g and paper_type:
                    paper_manual = f'{paper_g} г {paper_type}'
                elif paper_g:
                    paper_manual = f'{paper_g} г'
                elif paper_type:
                    paper_manual = paper_type
                else:
                    paper_manual = '—'

                finish = [x for x in self._client_offer_finish_list() if not x.startswith('Друго') and x != 'Транспорт']
                if hv('sep_mat').lower() not in ('', 'без', 'не'):
                    sep_sheets = r.get('separator_sheets')
                    if not sep_sheets:
                        try:
                            sep_sheets = max(0, int(r.get('clean_sheets', 0) or 0) // max(1, sheets_block))
                        except Exception:
                            sep_sheets = 0
                    if sep_sheets:
                        finish = [f'картон - {sep_sheets} цели листа' if x == 'Разделяне' else x for x in finish]
                finish_line = ' ● '.join(str(x).strip() for x in finish if str(x).strip()) or '—'

                # Подредбата е същата като в „Заявка_Невена“.
                lines += [
                    row(f'КЛИЕНТ: {client or "—"}', f'ДАТА: {date or "—"}'), border(),
                    row(f'ИЗДЕЛИЕ: {item or "—"}', f'ЦЕНА БЕЗ ДДС: {total:.2f}'), border(),
                    row(f'ЕД. БРОЙКИ: {qty} кочана', f'ед. бройка: {unit:.4f}'), border(),
                    row(f'бр. листа в кочан/ от цвят: {sheets_block} листа в кочан',
                        f'бр. ЦВЯТА листа: {paper_colors} цвят/а'), border(),
                    row(f'ХАРТИЯ: {paper_price_text}', f'формат (на х-я): {r.get("source_format", "—")}'), border(),
                    row(paper_manual, f'формат за ПЕЧАТ: {r.get("print_format", "—")}'), border(),
                    row(f'TИРАЖ: {clean}', f'размножения: {reps}'), border(),
                    row(f'ЦВЕТНОСТ: {color}', f'ОБРЪЩАНЕ: {turnover.upper() if turnover else "—"}'), border(),
                    row(f'ЦЕЛИ ЛИСТА: {whole} листа/ цвят (вкл. макулатура)', f'ОБРЯЗАН РАЗМЕР: {size}'), border(),
                    full('ДОВЪРШИТЕЛНИ РАБОТИ', True),
                    *[f'| {line} |' for line in full_wrapped(f'{finish_line}')]
                ]
            else:
                v = self._request_values()
                paper_manual = f'{v["paper"]} г {v["paper_type"]}'.strip()
                if not paper_manual:
                    paper_manual = '—'

                finish_items = [x.strip() for x in (v.get('finish') or '').split('|') if x.strip()] or ['—']
                finish_line = ' ● '.join(finish_items)

                # Таб 1 използва същата конструкция и същите ширини като Химия.
                lines += [
                    row(f'КЛИЕНТ: {v["client"] or "—"}', f'ДАТА: {v["date"] or "—"}'), border(),
                    row(f'ИЗДЕЛИЕ: {v["item"] or "—"}', f'ЦЕНА БЕЗ ДДС: {v["price"] or "—"}'), border(),
                    row(f'ЕД. БРОЙКИ: {v["qty"] or "—"}', f'ед. бройка: {v["unit"] or "—"}'), border(),
                    row(f'ЦВЕТНОСТ: {v["color"] or "—"}', f'ОБРЪЩАНЕ: {v["turnover"] or "—"}'), border(),
                    row(f'ХАРТИЯ: {v["paper_price"] or "—"}', f'формат (на х-я): {v["source_format"] or "—"}'), border(),
                    row(paper_manual, f'формат за ПЕЧАТ: {v["print_format"] or "—"}'), border(),
                    row(f'TИРАЖ: {v["clean"] or "—"}', f'размножения: {v["reps"] or "—"}'), border(),
                    row(f'{v["source_sheets"] or "—"} листа (вкл. макулатура)', f'ОБРЯЗАН РАЗМЕР: {v["size"] or "—"}'), border(),
                    full('ДОВЪРШИТЕЛНИ РАБОТИ', True),
                    *[f'| {line} |' for line in full_wrapped(f'{finish_line}')]
                ]

            path = filedialog.asksaveasfilename(
                title='Запази заявката като текстов файл',
                defaultextension='.txt',
                filetypes=[('Текстов файл', '*.txt'), ('Всички файлове', '*.*')],
                initialfile=f'Заявка_{self.request_vars["client"].get().strip() or ""}.txt'
            )
            if not path:
                return
            # Старият файл „Невена“ е UTF-16; запазваме и това поведение.
            Path(path).write_text('\n'.join(lines), encoding='utf-16')
        except Exception as e:
            messagebox.showerror('Грешка при запис на заявка', str(e))

    def _clear(self,f):
        for c in f.winfo_children():c.destroy()
    def _set_result_mode(self, mode):
        # Съвместимост със стари извиквания; резултатът вече е един общ изглед.
        self.result_mode = mode
        self._render_result()

    def _compact_cost_grid(self, parent, rows):
        """Тройна разбивка, при която цената стои непосредствено след описанието."""
        grid = ttk.Frame(parent)
        grid.pack(fill='x', padx=10, pady=(2, 8))

        # Три равни визуални групи. Във всяка група описанието и цената
        # са непосредствено една до друга, вместо цената да се изтласква
        # в крайния десен край на групата.
        for c in range(3):
            grid.columnconfigure(c, weight=1, uniform='cost_group')

        groups = 3
        per_group = (len(rows) + groups - 1) // groups
        for g in range(groups):
            data = rows[g * per_group:(g + 1) * per_group]
            group = ttk.Frame(grid)
            group.grid(row=0, column=g, sticky='nw', padx=(4, 12 if g < 2 else 4))
            group.columnconfigure(0, weight=0)
            group.columnconfigure(1, weight=0)

            for row, (label, value) in enumerate(data):
                ttk.Label(group, text=label).grid(
                    row=row, column=0, sticky='w', padx=(0, 7), pady=3
                )
                ttk.Label(group, text=str(value), style='Value.TLabel').grid(
                    row=row, column=1, sticky='w', padx=0, pady=3
                )

    def _render_result(self):
        # Вече няма отделен визуален таб „Резултат“. Данните се пазят
        # във self.result / self.himiya_result и се използват от „Заявка“
        # и клиентската оферта.
        if hasattr(self, 'request_sheet'):
            self._refresh_request_offer()

    def _himiya_plate_count(self, r):
        value = r.get('plate_count', r.get('plates_count'))
        if value not in (None, '', '—'):
            return value
        try:
            front = int(float(self._hvar('front').get() or 0))
            back = int(float(self._hvar('back').get() or 0))
            turnover = str(self._hvar('turnover').get() or '').strip().lower()
            choice = str(self._hvar('plates').get() or '').strip().lower()
            if choice == 'не' or not choice:
                return 0
            count = front + back if choice == 'всеки цвят отд.' else front
            if choice == 'да' and turnover in ('не','пантон','черно'):
                count = front + back
            return count
        except Exception:
            return 0

    def _render_himiya_result(self, r):
        """Подробен резултат за Химия в същия чист стил като първия таб."""
        # Горен ред: резюме + довършителни
        top = self.result_summary.master
        top.columnconfigure(0, weight=2, uniform='hres')
        top.columnconfigure(1, weight=3, uniform='hres')

        summary = self.card(self.result_summary, 'Ключово резюме')
        summary.pack(fill='both', expand=True)
        summary.columnconfigure(1, weight=1)
        summary.columnconfigure(3, weight=1)
        vals = [
            ('Размер на изделието', f"{r.get('product_w',0):g} × {r.get('product_h',0):g} мм"),
            ('Изходен формат', r.get('source_format','—')),
            ('Печатен формат', r.get('print_format','—')),
            ('Размножения', r.get('repetitions','—')),
            ('Брой кочани', f"{r.get('quantity','')} × {r.get('sheets_per_block','')} л."),
            ('Цветност', f"{self._hvar('front').get()}+{self._hvar('back').get()}"),
            ('Обръщане', self._hvar('turnover').get() or '—'),
            ('Тираж чист от цвят', self._fmt_count(r.get('clean_sheets','—'))),
            ('Цели листа / цвят', self._fmt_count(r.get('whole_sheets','—'))),
            ('Пакет от цвят', r.get('package_color','—')),
        ]
        for i,(label,value) in enumerate(vals):
            row=i//2; col=(i%2)*2
            ttk.Label(summary,text=label,style='SummaryLabel.TLabel').grid(row=row,column=col,sticky='w',padx=(10,6),pady=4)
            ttk.Label(summary,text=str(value),style='SummaryValue.TLabel').grid(row=row,column=col+1,sticky='w',padx=(0,18),pady=4)

        finish = self.card(self.result_finish, 'Довършителни операции')
        finish.pack(fill='both', expand=True)

        # В горното каре няма цени. То е клиентско-информационен списък,
        # еднакъв по логика с първия таб и служи като източник за
        # бъдещата "Клиентска оферта".
        selected = []
        def hv(key):
            return str(self._hvar(key).get() if key in getattr(self, 'himiya_vars', {}) else '').strip()
        def hn(key, default=0):
            try:
                return float(hv(key).replace(',', '.')) if hv(key) else float(default)
            except Exception:
                return float(default)

        if hv('cutting').lower() not in ('', 'без', 'не'): selected.append('Рязане')
        if hn('big_count') > 0: selected.append('Биговане')
        if hv('gluing').lower() not in ('', 'без', 'не'): selected.append('Лепене')
        if hv('numbering').lower() not in ('', 'без', 'не'): selected.append('Номерация')
        if hv('perforation').lower() not in ('', 'без', 'не'): selected.append('Перфорация')
        if hv('typesetting').lower() not in ('', 'без', 'не'): selected.append('Набор')
        if hv('sewing').lower() not in ('', 'без', 'не'): selected.append('Шиене/телчета')
        if hv('em').lower() not in ('', 'без', 'не'): selected.append('Ел. монтаж')
        if hv('counting').lower() not in ('', 'без', 'не'): selected.append('Пакетиране')
        if hv('sep_mat').lower() not in ('', 'без', 'не'): selected.append('Разделяне')
        # Транспортът остава в калкулацията, но не се показва в резултата.
        if hn('other') > 0: selected.append('Друго')

        if selected:
            op_grid = ttk.Frame(finish)
            op_grid.pack(fill='both', expand=True, padx=10, pady=8)
            op_grid.columnconfigure(0, weight=1)
            op_grid.columnconfigure(1, weight=1)
            half = (len(selected) + 1) // 2
            for i, operation in enumerate(selected):
                col = 0 if i < half else 1
                row = i if i < half else i - half
                ttk.Label(op_grid, text='• ' + operation).grid(
                    row=row, column=col, sticky='w', padx=4, pady=3
                )
        else:
            ttk.Label(finish, text='Няма избрани довършителни операции.', style='Sub.TLabel').pack(
                anchor='w', padx=10, pady=10
            )

        # Долен ред: анализ на печата + материали
        body=self.result_body
        body.columnconfigure(0,weight=2,uniform='hbody')
        body.columnconfigure(1,weight=3,uniform='hbody')
        body.rowconfigure(0,weight=1)
        analysis_frame=ttk.Frame(body)
        analysis_frame.grid(row=0,column=0,sticky='nsew',padx=(0,4),pady=(4,0))
        material_frame=ttk.Frame(body)
        material_frame.grid(row=0,column=1,sticky='nsew',padx=(4,0),pady=(4,0))

        analysis=self.card(analysis_frame,'Анализ на печата'); analysis.pack(fill='both',expand=True)
        front=int(float(self._hvar('front').get() or 0)); back=int(float(self._hvar('back').get() or 0))
        # Някои по-стари engine.py версии връщат само общия печат.
        # Затова при липсваща разбивка я възстановяваме от входовете и
        # наличната обща цена, вместо да показваме подвеждащи нули.
        turnover_mode = str(self._hvar('turnover').get() or '').strip().lower()
        rate=float(r.get('print_rate', 0) or 0)
        if rate <= 0:
            rate = 5.2 if turnover_mode == 'черно' else 8.0
        if r.get('print_face_back') is not None and float(r.get('print_face_back',0) or 0) > 0:
            pf=float(r.get('print_face_back',0) or 0)
        else:
            print_colors = front if turnover_mode in ('да','черно с обр.') else front + back
            pf = print_colors * rate
        over1000=float(r.get('over1000_cost',0) or 0)
        turnover_cost=float(r.get('turnover_cost',0) or 0)
        if turnover_cost <= 0 and turnover_mode == 'да' and back > 0:
            clean=float(r.get('clean_sheets',0) or 0)
            turnover_cost=2.7 * back * max(1, math.ceil(clean/1000))
        if over1000 <= 0:
            clean=float(r.get('clean_sheets',0) or 0)
            times=math.floor(clean/1001)
            units=(front+back) if turnover_mode in ('не','черно') else front
            over1000=units * 2.7 * times
        print_total=pf + turnover_cost + over1000
        analysis_rows=[
            ('Печат лице/гръб', ''),
            (f'  лице: {front} цвят × {rate:.2f} €', f'{front*rate:.2f} €'),
            (f'  гръб: {back} цвята × {rate:.2f} €', f'{back*rate:.2f} €'),
            ('Печат над 1000', f"{over1000:.2f} €"),
            ('Обръщане', f"{turnover_cost:.2f} €"),
        ]
        for i,(label,value) in enumerate(analysis_rows):
            font=('Segoe UI',10,'bold') if value=='' else None
            ttk.Label(analysis,text=label,font=font).grid(row=i,column=0,sticky='w',padx=(10,6),pady=3)
            if value: ttk.Label(analysis,text=value,style='Value.TLabel').grid(row=i,column=1,sticky='e',padx=(4,12),pady=3)
        rr=len(analysis_rows)
        ttk.Separator(analysis).grid(row=rr,column=0,columnspan=2,sticky='ew',padx=10,pady=(5,4))
        ttk.Label(analysis,text='Общо печат',font=('Segoe UI',10,'bold')).grid(row=rr+1,column=0,sticky='w',padx=10,pady=3)
        ttk.Label(analysis,text=f'{print_total:.2f} €',font=('Segoe UI',10,'bold')).grid(row=rr+1,column=1,sticky='e',padx=(4,12),pady=3)

        material=self.card(material_frame,'Материали'); material.pack(fill='both',expand=True)
        mat_rows=[
            ('Хартия — цели листа', r.get('whole_sheets','—')),
            ('Цена / лист', f"{float(self._hvar('paper').get() or 0):.3f} €"),
            ('Сума хартия', f"{float(r.get('paper',0) or 0):.2f} €"),
            ('Плаки / брой', self._himiya_plate_count(r)),
            ('Сума плаки', f"{float(r.get('plates',0) or 0):.2f} €"),
            ('Предпечат', f"{float(r.get('prepress',0) or 0):.2f} €"),
        ]
        for i,(label,value) in enumerate(mat_rows):
            ttk.Label(material,text=label).grid(row=i,column=0,sticky='w',padx=(10,10),pady=3)
            ttk.Label(material,text=str(value),style='Value.TLabel').grid(row=i,column=1,sticky='e',padx=(0,12),pady=3)

    def _update_print_price_diag(self, r):
        # Поставяме бял фон на самия контейнер
        try:
            self.print_price_diag.configure(style='Section.TLabelframe')
        except Exception:
            pass

        for child in self.print_price_diag.winfo_children():
            child.destroy()

        turnover_value = float(r.get('turnover_cost', 0) or 0)
        duplication_value = float(r.get('duplication', 0) or 0)
        turnover_display = 'НЕ' if turnover_value <= 0 else f"{turnover_value:.2f} € / да"
        duplication_display = 'НЕ' if duplication_value <= 0 else f"{duplication_value:.2f} € / да"

        rows = [
            ('Печат лице/гръб', f"{r.get('print_face_back',0):.2f} €"),
            ('Обръщане', turnover_display),
            ('Дублаж', duplication_display),
        ]
        if r.get('over1000_times', 0) > 0 or r.get('over1000_cost', 0):
            rows.append((
                'Печат над 1000',
                f"{r.get('over1000_cost',0):.2f} € / {r.get('over1000_times',0)} пъти"
            ))
        rows.append(('Общо печат', f"{r.get('print',0):.2f} €"))
        if r.get('labor_base', 0):
            rows.append((
                'База за оскъпяване',
                f"{r.get('labor_base',0):.2f} € | закръглена: {r.get('labor_base_rounded',0):.0f} €"
            ))
        if r.get('surcharge', 0):
            rows.append(('Оскъпяване на труда', f"{r.get('surcharge',0):.2f} €"))

        plate_expense = float(r.get('plate_count', 0) or 0) * 2.45
        expenses = (
            float(r.get('paper', 0) or 0)
            + float(r.get('uv', 0) or 0)
            + float(r.get('lamination', 0) or 0)
            + float(r.get('calender', 0) or 0)
            + float(r.get('die_cut', 0) or 0)
            + float(r.get('film', 0) or 0)
            + float(r.get('other', 0) or 0)
            + float(r.get('separator_cost', r.get('separators', 0)) or 0)
            + float(r.get('transport', 0) or 0)
            + plate_expense
        )
        total_value = float(r.get('total', 0) or 0)
        profit = total_value - expenses
        rows.append(('Печалба', f"{profit:.2f} €"))
        rows.append(('Разходи', f"{expenses:.2f} €"))

        for i, (label, value) in enumerate(rows):
            is_total = label == 'Общо печат'
            is_finance = label.strip().capitalize() in ('Печалба', 'Разходи')

            # Описанията са обикновен шрифт. Само „Общо печат“,
            # „Печалба“ и „Разходи“ са болд. Стойностите с цени
            # остават болднати както досега.
            ttk.Label(
                self.print_price_diag, text=label,
                background='#FFFFFF',
                font=('Segoe UI', 10, 'bold') if is_total or is_finance else None
            ).grid(row=i, column=0, sticky='w', padx=(10, 4), pady=3)

            ttk.Label(
                self.print_price_diag, text=value,
                style='Value.TLabel',
                background='#FFFFFF',
                font=('Segoe UI', 10, 'bold') if is_total or is_finance or not is_finance else None
            ).grid(row=i, column=1, sticky='w', padx=(4, 10), pady=3)

        self.print_price_diag.columnconfigure(0, weight=0)
        self.print_price_diag.columnconfigure(1, weight=1)
    def _update_diagnostics(self):
        if not self.result:return
        r=self.result
        self._update_print_price_diag(r)
        self._set_diag(self.diag,[('Размер на печатното изделие',f"{r['product_w']:g} × {r['product_h']:g} мм"),('Единични бройки',self._fmt_count(r['unit_pieces'])),('Чист тираж',self._fmt_count(r['clean_sheets'])),('Размножения',r['repetitions']),('Изходен формат',r['source_format']),('Печатен формат',r['print_format']),('Макулатура',self._fmt_count(r['waste_sheets'])),('Тираж + макулатура',self._fmt_count(r['total_print_sheets'])),('Цели листа за тиража',self._fmt_count(r['source_sheets']))])
        # Материалните цени се показват директно в карето „Материали“.
        # За хартията използваме директно резултата от двигателя. Така
        # стойността до полето винаги следва същата логика като общата
        # калкулация, включително избора „със/без ДДС“.
        if hasattr(self, 'material_paper_price_label'):
            paper_total = float(r.get('paper', 0) or 0)
            self.material_paper_price_label.configure(
                text=f"{paper_total:.2f} €" if paper_total else ''
            )
        if hasattr(self, 'material_plates_price_label'):
            self.material_plates_price_label.configure(
                text=f"{r.get('plates',0):.2f} €" if r.get('plates',0) else ''
            )
        if hasattr(self, 'separator_sheet_label') and self.separator_sheet_label is not None:
            sep_sheets = int(r.get('separator_sheets', 0) or 0)
            self.separator_sheet_label.configure(
                text=f"{self._fmt_count(sep_sheets)} л." if sep_sheets else ''
            )

        # Цената на всяка довършителна операция е непосредствено до нейното
        # поле за избор/въвеждане. Това заменя отделното диагностично каре.
        if hasattr(self, 'finish_price_labels'):
            for key, label_widget in self.finish_price_labels.items():
                cost_key = {
                    'cutting':'cutting','numbering':'numbering','perforation':'perforation',
                    'uv':'uv','lamination':'lamination','calender':'calender','gluing':'gluing',
                    'folding':'folding','die':'die_cut','breaking':'breaking_cost',
                    'electric_montage':'electric_montage','bigoving_type':'bigoving',
                    'other_operations':'other','sepn':'separators','transport':'transport',
                    'counting':'counting','round_punch':'round_punch'
                }.get(key)
                value = r.get(cost_key,0) if cost_key else 0
                label_widget.configure(text=f"{value:.2f} €" if value else '')

        fin=[]
        for key,label in COST_KEYS:
            if key in ('paper','plates','prepress','transport','surcharge'):continue
            if r.get(key,0):fin.append((label,f"{r[key]:.2f} €"))
        if r.get('separator_sheets',0):fin.append(('Цели листа за разделители',r['separator_sheets']))
        if r.get('transport',0):fin.append(('Транспорт',f"{r['transport']:.2f} €"))
        # База за оскъпяване и оскъпяване на труда вече са в
        # „Ценообразуване на печата“, затова не се повтарят тук.
        self._set_diag(self.finish_diag,fin,columns=2)

    def _layout_for_graphic(self, sheet_w, sheet_h, prod_w, prod_h):
        """Връща координати на изделията за визуализация, без да променя engine резултата."""
        def solve(w,h):
            memo={}
            def helper(w,h):
                if w < min(prod_w,prod_h) or h < min(prod_w,prod_h):
                    return 0,[]
                key=(round(w,1),round(h,1))
                if key in memo:return memo[key]
                best=0; rects=[]
                if w>=prod_w and h>=prod_h:
                    best=1; rects=[(0,0,prod_w,prod_h)]
                if w>=prod_h and h>=prod_w and best<1:
                    best=1; rects=[(0,0,prod_h,prod_w)]
                cut_x=set(); x=prod_w
                while x<w:cut_x.add(x);x+=prod_w
                x=prod_h
                while x<w:cut_x.add(x);x+=prod_h
                for cx in cut_x:
                    c1,r1=helper(cx,h);c2,r2=helper(w-cx,h)
                    if c1+c2>best:
                        best=c1+c2;rects=r1+[(rx+cx,ry,rw,rh) for rx,ry,rw,rh in r2]
                cut_y=set(); y=prod_h
                while y<h:cut_y.add(y);y+=prod_h
                y=prod_w
                while y<h:cut_y.add(y);y+=prod_w
                for cy in cut_y:
                    c1,r1=helper(w,cy);c2,r2=helper(w,h-cy)
                    if c1+c2>best:
                        best=c1+c2;rects=r1+[(rx,ry+cy,rw,rh) for rx,ry,rw,rh in r2]
                memo[key]=(best,rects);return memo[key]
            return helper(w,h)
        return solve(sheet_w,sheet_h)

    def _update_graphic(self):
        if not hasattr(self,'graphic_canvas'): return
        c=self.graphic_canvas
        c.delete('all')
        r=self.result or {}
        if not r:
            return
        try:
            sw_cm,sh_cm=parse_size(r['print_format'])
            sw,sh=sw_cm*10,sh_cm*10
            pw,ph=float(r['product_w']),float(r['product_h'])
            reps=int(r.get('repetitions',0) or 0)
            useful=self.vars.get('grip',tk.StringVar(value='')).get()=='да'
            grip=3 if useful else 10
            # Това е същото работно поле, което engine smart_calculate използва.
            lim_w=sw-5; lim_h=sh-(grip+3)
            turnover=str(self.vars.get('turnover',tk.StringVar(value='не')).get()).strip().lower()=='да'
            if turnover:
                half=lim_w/2
                _,rects1=self._layout_for_graphic(half,lim_h,pw,ph)
                rects=rects1+[(x+half,y,w,h) for x,y,w,h in rects1]
            else:
                _,rects=self._layout_for_graphic(lim_w,lim_h,pw,ph)
            # Показваме точно толкова изделия, колкото е записал engine резултатът.
            rects=rects[:reps]
            self.update_idletasks()
            # Вземаме реалните динамични размери на платното
            W = c.winfo_width()
            H = c.winfo_height()
            if W < 10: W = 260
            if H < 10: H = 220

            # Намаляваме излишните бели полета от 40/35 на 16 пиксела
            scale = min((W - 16) / sw, (H - 16) / sh) if sw and sh else 1
            
            # Центрираме схемата перфектно – както хоризонтално, така и вертикално
            ox = (W - sw * scale) / 2
            oy = (H - sh * scale) / 2
            c.create_rectangle(ox,oy,ox+sw*scale,oy+sh*scale,fill='#F2F2F2',outline='#333333')
            c.create_rectangle(ox+2.5*scale,oy+3*scale,ox+(sw-2.5)*scale,oy+(sh-grip)*scale,outline='#C0392B',dash=(3,3))
            if turnover:
                c.create_line(ox+(sw/2)*scale,oy,ox+(sw/2)*scale,oy+sh*scale,fill='#C0392B',dash=(4,3))
            for i,(x,y,w,h) in enumerate(rects,1):
                x1=ox+(2.5+x)*scale;y1=oy+(3+y)*scale
                x2=x1+w*scale;y2=y1+h*scale
                c.create_rectangle(x1,y1,x2,y2,fill='#DDEBF7',outline='#2F5597')
                if (x2-x1)>38 and (y2-y1)>22:
                    c.create_text((x1+x2)/2,(y1+y2)/2,text=str(i),fill='#1F1F1F',font=('Segoe UI',8,'bold'))
        except Exception as e:
            pass
    def calculate(self):
        try:
            def fnum(k, default=0.0):
                if k not in self.vars: return default
                raw = str(self.vars[k].get()).strip().replace(' ', '')
                if not raw: return default
                try: return float(raw.replace(',', '.'))
                except ValueError: return default

            def sval(k, default=''):
                # Някои стари довършителни полета (UV, ламиниране,
                # каландър и др.) вече не са част от интерфейса.
                # Изчислението трябва да ги приема като "без",
                # вместо да търси несъществуваща StringVar.
                v = self.vars.get(k)
                return str(v.get()).strip() if v is not None else default

            prices={
                'print_g4':fnum('price_print_g4',8.00),
                'print_over1000':fnum('price_print_over1000',2.70),
                'turnover':fnum('price_turnover',2.70),
                 'duplication':fnum('price_duplication',2.70),
                'plate':fnum('price_plate',2.80),
                'fuel_per_km':fnum('price_fuel',1.15),
                'uv_gloss':fnum('price_uv_gloss',0.028),
                'uv_matt':fnum('price_uv_matt',0.056),
                'uv_partial':fnum('price_uv_partial',0.071),
                'uv_volume':fnum('price_uv_volume',0.075),
                'lam_gloss':fnum('price_lam_gloss',0.051),
                'lam_velvet':fnum('price_lam_velvet',0.20),
                'lam_matt':fnum('price_lam_matt',0.064),
                'calender':fnum('price_calender',0.025),
                'bigoving':fnum('price_bigoving',0.005),
                'bigoving_over1000':0.012,
                'bigoving_setup':2.56,
                'fold_manual':fnum('price_fold_manual',0.005),
                'fold_gatevi':fnum('price_fold_gatevi',0.0015),
                'fold_gad':fnum('price_fold_gad',0.002),
                'glue_pocket':fnum('price_glue_pocket',0.06),
                'glue_flags':fnum('price_glue_flags',0.02),
                'glue_boxes':fnum('price_glue_boxes',0.012),
                'glue_double':fnum('price_glue_double',0.041),
                'glue_lamination':0.064,
                'glue_cubes_over1000':0.1,
                'separator_вестник':fnum('price_sep_news',0.028),
                'separator_друг':fnum('price_sep_other',0.30),
                'em_flayers':fnum('price_em_flayers',1.28),
                'em_leaflets':fnum('price_em_leaflets',1.28),
                'em_labels':fnum('price_em_labels',1.02),
                'em_covers':fnum('price_em_covers',1.53),
                'em_brochure':fnum('price_em_brochure',2.56),
                'em_min':fnum('price_em_min',0.51),
                'em_poster':fnum('price_em_poster',2.56),
                'em_pagination':fnum('price_em_pagination',0.25),
            }
            required=['source','w','h','qty','front','back','turnover','grip','paper','plates']
            for key in required:
                if key in self.vars and self.vars[key].get().strip()=='':
                    raise ValueError(f'Моля, попълнете полето: {key}')
            # Предпазване от Division by zero в engine при празен/нулев тираж
            # или размер. Валидираме ги тук, преди да подадем Inputs към engine.
            if fnum('qty') <= 0:
                raise ValueError('Тиражът трябва да е по-голям от 0.')
            if fnum('w') <= 0 or fnum('h') <= 0:
                raise ValueError('Размерът трябва да е по-голям от 0.')
            if self.vars.get('repetition_mode', tk.StringVar(value='')).get().strip() == 'Ръчно':
                if fnum('repetitions_manual') <= 0:
                    raise ValueError('Ръчното размножение трябва да е по-голямо от 0.')
            
            # Създаване на обекта Inputs с поправената цена за предпечат
            inp=Inputs(
                source_format=self.vars['source'].get(),
                product_w=fnum('w'),
                product_h=fnum('h'),
                quantity=int(fnum('qty')),
                front_colors=int(self.vars['front'].get()),
                back_colors=int(self.vars['back'].get()),
                turnover=self.vars['turnover'].get(),
                useful_grip=self.vars['grip'].get()=='да',
                duplication=self.vars['duplication'].get(),
                duplication_count=max(0,int(fnum('duplication_count',0))),
                paper_price_per_sheet=fnum('paper'),
                vat=self.vars['vat'].get(),
                plates=self.vars['plates'].get()=='да',
                plate_supply=self.vars['plates'].get(),
                cutting=self.vars['cutting'].get(),
                numbering=self.vars['numbering'].get(),
                perforation=self.vars['perforation'].get(),
                prepress_type=self.vars.get('prepress_type', tk.StringVar(value='без')).get(),
                prepress_price=fnum('prepress_price'),                
                uv=sval('uv','без'),
                lamination=sval('lamination','без'),
                calender=sval('calender','без'),
                film_price=fnum('film_price'),
                gluing=sval('gluing','без'),
                folding=int(fnum('folding')),
                folding_type=sval('foldtype','без'),
                bigoving=int(fnum('bigoving')),
                bigoving_type=self.vars.get('bigoving_type', tk.StringVar(value='без')).get(),
                die_cut=sval('die','без'),round_punch=sval('round_punch','без'),
                breaking=fnum('breaking',40),
                electric_montage=sval('electric_montage','без'),typesetting=sval('typesetting','без'),sewing=sval('sewing','без'),
                counting=sval('counting','да'),
                # „Операции, които не са изброени“ е цена, не текстово описание.
                other_price=fnum('other_operations'),
                separators=int(fnum('sepn')),
                separator_material=sval('sepmat','без'),
                transport=sval('transport','не'),
                transport_km=fnum('km',15),
                fuel_price=fnum('fuel',1.15),
                surcharge_pct=fnum('surcharge',40),
                prices=prices
            )

            forced=self.vars['print_format'].get();forced=None if forced=='Автоматичен' else forced
            manual=None
            if self.vars['repetition_mode'].get()=='Ръчно':
                raw=self.vars['repetitions_manual'].get().strip()
                if not raw:raise ValueError('Избрано е ръчно размножение, но не е въведен брой бройки на печатен лист.')
                manual=int(float(raw.replace(',','.')))
                
            self.result=calc(inp,repo,forced,manual); self.last_result_mode='order'; self.request_source_tab='Основен'; self._update_diagnostics(); self._update_graphic(); self._update_optimal_suggestion();self._render_result();self._last_valid=True;self._update_top_bar();self._refresh_request_offer()
        except Exception as e:
            self._last_valid=False;messagebox.showerror('Грешка при изчислението',str(e))

    def reset_himiya(self):
        # Изчистваме само данните и резултата на Химия.
        # Изчислението в „Поръчка и печат“ остава запазено.
        # Основни данни / Печат се изчистват напълно. Не връщаме
        # автоматично старите примерни стойности, за да е ясно, че табът
        # е празен и е готов за ново изчисление.
        defaults = {
            'source':'','print':'','repetition_mode':'','repetitions_manual':'',
            'qty':'','w':'','h':'','front':'','back':'','turnover':'','grip':'',
            'sheets':'','paper_colors':'','paper':'','vat':'','plates':'',
            'prepress':'','em':'',
            # Довършителните операции остават както са били — бутонът
            # „Изчисти“ тук е поискан да изчиства само Основни данни / Печат.
        }
        for key, value in defaults.items():
            self._hvar(key, value).set(value)
        self.himiya_result = {}
        if hasattr(self, 'h_material_paper'):
            self.h_material_paper.configure(text='—')
        if hasattr(self, 'h_material_plates'):
            self.h_material_plates.configure(text='—')
        if hasattr(self, 'h_sep_sheets_label'):
            self.h_sep_sheets_label.configure(text='')
        self._refresh_himiya_prints()
        self._sync_himiya_repetitions()
        self._set_diag(self.himiya_finish_diag, [], columns=2)
        for attr in getattr(self, 'himiya_diag_attrs', []):
            try:
                getattr(self, attr).configure(text='')
            except Exception:
                pass
        if hasattr(self, 'himiya_body'):
            for w in self.himiya_body.winfo_children():
                w.destroy()
            ttk.Label(self.himiya_body, text='Попълнете оранжевите полета и натиснете „ИЗЧИСЛИ“.', foreground='#2B579A').pack(anchor='w', padx=6, pady=6)
        self._update_top_bar()

    def reset(self):
        # Clear all order/material/finishing inputs so a new calculation cannot
        # accidentally reuse values from the previous job. Price settings remain.
        input_keys = [
            'source','w','h','qty','front','back','turnover','grip','print_format',
            'repetition_mode','repetitions_manual','duplication','duplication_count','paper','vat','plates','prepress_price',
            'cutting','numbering','perforation','uv','lamination','calender','film_price',
            'gluing','bigoving','bigoving_type','folding','foldtype','die','round_punch','breaking','electric_montage',
            'other_operations','counting','sepmat','sepn','transport','km','fuel'
        ]
        for key in input_keys:
            if key in self.vars:
                self.vars[key].set('')
        self.result={}; self._last_valid=False
        self._set_diag(self.diag,[])
        if hasattr(self, 'optimal_format_suggestion'):
            self.optimal_format_suggestion.config(text='След изчисление ще се покаже оптималният формат.')
        if hasattr(self, 'material_paper_price_label'):
            self.material_paper_price_label.configure(text='')
        if hasattr(self, 'material_plates_price_label'):
            self.material_plates_price_label.configure(text='')
        if hasattr(self, 'finish_price_labels'):
            for w in self.finish_price_labels.values():
                w.configure(text='')
        self._set_diag(self.finish_diag,[])
        self._set_diag(self.print_price_diag,[])
        self._render_result()
        if hasattr(self, 'optimal_format_suggestion'):
            self.optimal_format_suggestion.config(text='-')
        self._update_top_bar()

if __name__=='__main__':App().mainloop()
