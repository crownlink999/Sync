#!/usr/bin/env python3
"""Well Status -> data/wellstatus.json (вкладка «Скважины» приложения «Основные сводки»).

Запуск:  python3 tools/build_wellstatus.py <Well_Status_*.xlsx> [out=data/wellstatus.json]

Листы файла:
  TOTAL WELLS      - по строке на скважину (признак реальной скважины: колонка A = 1); остальное - итоги по ГЗУ
  Сводная таблица  - суточная история по объектам с 2013 г. (Qж, Qн, обводнённость, Ндин) + отклонение Qн от техрежима
  гориз скв / ЗБС  - списки горизонтальных скважин и зарезок боковых стволов (только пометка)
"""
import sys, re, json, datetime as dt, collections
import openpyxl

SRC = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else 'data/wellstatus.json'

wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
ws = wb['TOTAL WELLS']


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) and v == v else None


def rnd(v, n=1):
    v = num(v)
    if v is None:
        return None
    v = round(v, n)
    return int(v) if v == int(v) else v      # 22.0 -> 22 (короче JSON)


def dte(v):
    """дата или None (в файле пустая дата = время 00:00:00 или 1899)"""
    if isinstance(v, dt.datetime) and v.year >= 2000:
        return v.date()
    return None


def txt(v):
    if v is None:
        return ''
    s = re.sub(r'\s+', ' ', str(v)).strip()
    s = re.sub(r'[一-鿿（）、，：；·]+', '', s).strip()
    return '' if s == '0' else s


# дата среза
rep_date = ws['A1'].value
if not isinstance(rep_date, dt.datetime):
    raise SystemExit('TOTAL WELLS!A1 не дата')
REP = rep_date.date()

# ---- горизонтальные и ЗБС ----
def well_ids(name):
    s = set()
    try:
        for r in wb[name].iter_rows(min_row=2, max_col=3, values_only=True):
            if r[2] not in (None, '', 0):
                s.add(str(r[2]).strip().replace('.0', ''))
    except KeyError:
        pass
    return s
HZ, ZBS = well_ids('гориз скв'), well_ids('ЗБС')

# ---- статусы ----
def categorize(st):
    s = st.lower().strip()
    if not s or s == '0':
        return 'other', None
    flag = None
    if 'проп' in s:
        flag = 'leak'            # работает с пропуском (клапан / НКТ)
    if 'низ' in s and 'эфф' in s:
        flag = 'loweff'          # ВН с низкой эффективностью насоса
    if s.startswith('п.э') or s.startswith('п. э'):
        return 'pe', flag        # периодическая эксплуатация
    if s.startswith('бездейств'):
        return 'idle', flag
    if s.startswith(('опрс', 'прс', 'окрс', 'крс', 'кву', 'ож.', 'птос', 'рпл', 'свабир')):
        return 'rep', flag       # ремонт / ожидание ремонта / исследования
    if s.startswith('обводнен'):
        return 'wet', flag
    if s.startswith(('рег', 'контрол')):
        return 'reg', flag
    if any(k in s for k in ('работает', 'раб.', 'раб ', 'винтовой', 'фонтан')) or s.startswith(('ск ', 'сау', 'лпшн', 'вн ')):
        return 'work', flag
    return 'other', flag


NUMTOK = re.compile(r'^\d+(?:[.,]\d+)?$')


def hist(s, keep):
    """числа из текстовой истории замеров ('2026,А,29,29,68внс60↑70,29...'): только чистые числа, без годов"""
    if not isinstance(s, str):
        return []
    out = []
    for t in s.split(','):
        t = t.strip()
        if NUMTOK.match(t):
            v = float(t.replace(',', '.'))
            if t.isdigit() and 1990 <= v <= 2100:
                continue
            out.append(int(v) if v == int(v) else round(v, 1))
    return out[-keep:]


wells = []
for r in ws.iter_rows(min_row=5, max_col=95, values_only=True):
    if r[0] != 1 or r[3] in (None, ''):
        continue
    wid = str(r[3]).strip().replace('.0', '')
    st = txt(r[4])
    cat, flag = categorize(st)
    eff = num(r[27])
    if eff is not None and eff <= 1.5:
        eff *= 100
    dq, dd, wo = dte(r[30]), dte(r[31]), dte(r[28])
    w = {
        'id': wid, 'a': str(r[1] or ''), 't': txt(r[2]), 'st': st, 'c': cat, 'g': txt(r[5]),
        'q': rnd(r[6]), 'n': rnd(r[7], 2), 'w': rnd(r[8], 0), 'ht': rnd(r[10], 0), 'tp': rnd(r[11]), 'cp': rnd(r[12]),
        'dl': rnd(r[13], 0), 'pd': rnd(r[18], 0), 'zn': txt(r[21]), 'pm': txt(r[22]), 'pj': txt(r[23]),
        'rpm': rnd(r[24], 1), 'sl': rnd(r[25], 1), 'th': rnd(r[26], 0), 'ef': rnd(eff, 0),
        'wo': wo.isoformat() if wo else None,
        'dq': dq.isoformat() if dq else None, 'dd': dd.isoformat() if dd else None,
        'sb': rnd(r[34], 0), 'hr': rnd(r[64], 0),
        'pq': rnd(r[67]), 'pn': rnd(r[68], 2), 'pw': rnd(r[69], 0),
        'inf': txt(r[85]), 'nip': txt(r[65]).replace('NIPI_', '').replace('_Part', ''),
        'rp': int(r[94]) if isinstance(r[94], (int, float)) else 0,
        'hq': hist(r[41], 24), 'hd': hist(r[35], 10), 'hw': hist(r[38], 10),
    }
    if flag:
        w['fl'] = flag
    if wid in HZ:
        w['hz'] = 1
    if wid in ZBS:
        w['zbs'] = 1
    # возраст замеров
    w['lq'] = (REP - dq).days if dq else None
    w['ld'] = (REP - dd).days if dd else None
    wells.append(w)

# ---- суточная история (Сводная таблица) ----
OBJ = [('Общий', 1), ('12 ввг', 6), ('12А ввг', 10), ('33 птв', 14), ('34 птв', 18), ('16 птв', 22), ('27 птв', 26),
       ('30 птв', 30), ('30А птв', 34), ('31 птв', 38), ('32 птв', 42)]
DEV0 = 68   # колонки 69..78 (0-based 68..77): откл. Qн от техрежима по 10 объектам
wsv = wb['Сводная таблица']
hdr = list(wsv.iter_rows(min_row=1, max_row=1, max_col=127, values_only=True))[0]
rows = []
for r in wsv.iter_rows(min_row=5, max_col=127, values_only=True):
    d0 = r[0]
    if isinstance(d0, dt.datetime) and r[1] is not None:
        rows.append(r)
dates = [r[0].date() for r in rows]
if dates:
    print('Сводная таблица: %s .. %s, строк %d' % (dates[0], dates[-1], len(dates)))


def cells(r, c):
    return [rnd(r[c], 1), rnd(r[c + 1], 1), rnd(r[c + 2], 1), rnd(r[c + 3], 0)]   # Qж, Qн, %, Ндин


# дневные (последние 400 дней)
LAST = 400
sel = list(zip(dates, rows))[-LAST:]
daily = {'d': [d.isoformat() for d, _ in sel],
         'fact': [rnd(r[5], 1) for _, r in sel],             # «Добыча факт» (т/сут по сводке)
         'obj': {nm: [cells(r, c) for _, r in sel] for nm, c in OBJ},
         'dev': {OBJ[i + 1][0]: [rnd(r[DEV0 + i], 1) for _, r in sel] for i in range(10)}}
# помесячные средние за всю историю (Общий)
mon = collections.OrderedDict()
for d, r in zip(dates, rows):
    k = d.strftime('%Y-%m')
    m = mon.setdefault(k, [[], [], [], [], []])
    for i, c in enumerate((1, 2, 3, 4)):
        v = num(r[c])
        if v is not None:
            m[i].append(v)
    v = num(r[5])
    if v:
        m[4].append(v)
mean = lambda a: round(sum(a) / len(a), 1) if a else None
monthly = {'m': list(mon.keys()), 'qj': [mean(v[0]) for v in mon.values()], 'qn': [mean(v[1]) for v in mon.values()],
           'wc': [mean(v[2]) for v in mon.values()], 'nd': [mean(v[3]) for v in mon.values()], 'fact': [mean(v[4]) for v in mon.values()]}

D = {'meta': {'date': REP.isoformat(), 'src': SRC.replace('\\', '/').split('/')[-1], 'n': len(wells)},
     'wells': wells, 'daily': daily, 'monthly': monthly}
json.dump(D, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))

# ---- самопроверка: итоги должны совпасть со строкой «текущий» Сводной таблицы ----
tot_q = sum(w['q'] or 0 for w in wells)
tot_n = sum(w['n'] or 0 for w in wells)
cur = list(wsv.iter_rows(min_row=1, max_row=1, max_col=5, values_only=True))[0]
print('Дата среза:', REP, '| скважин:', len(wells))
print('Σ Qж по скважинам %.1f  (Сводная «текущий»: %.1f)' % (tot_q, cur[1] or 0))
print('Σ Qн по скважинам %.1f  (Сводная «текущий»: %.1f)' % (tot_n, cur[2] or 0))
print('по категориям:', dict(collections.Counter(w['c'] for w in wells)))
ok = abs(tot_q - (cur[1] or 0)) < 5 and abs(tot_n - (cur[2] or 0)) < 5
print('СВЕРКА:', 'OK' if ok else 'РАСХОЖДЕНИЕ!')
sys.exit(0 if ok else 3)
