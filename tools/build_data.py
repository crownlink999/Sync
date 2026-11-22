#!/usr/bin/env python3
"""Сборка data.json для приложения «Основные сводки».

Запуск:
    python3 tools/build_data.py <8-сводка.xlsx> <9-сводка.xlsx> <план-год.xlsx> [out=data.json]

8-сводка  — «Ежесуточные показатели» (лист '1-Мәліметтер' и далее)
9-сводка  — суточный рапорт на дату (Добыча, КРС, бурение, ...)
план-год  — помесячный план добычи/сдачи (лист 'Добыча Сдача ')
"""
import sys, os, json, re, datetime as dt
import openpyxl

MONTHS_KK = ['қаңтар', 'ақпан', 'наурыз', 'сәуір', 'мамыр', 'маусым', 'шілде', 'тамыз', 'қыркүйек', 'қазан', 'қараша', 'желтоқсан']
MONTHS_RU = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']

# ---- плановые данные (месячные планы пара/воды/ОЗПВ) лежат в data/plan.json ----

# ---- словарь казахских причин (КРС) ----
KK_REASON = {
    'сорап ақауы': 'дефект насоса', 'тұтылған': 'заклинивание/прихват', 'скқ ақауы': 'дефект НКТ',
    'саға ақауы': 'дефект устья', 'шыбық сынған': 'обрыв штанг',
}


PLAN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'plan.json')


def load_plan(f3):
    """План: data/plan.json (год, нефть, сдача, пар, Волжская вода, ОЗПВ). Если дан файл плана - обновляет нефть/сдачу."""
    P = json.load(open(PLAN_FILE, encoding='utf-8'))
    if f3 and f3 != '-' and os.path.exists(f3):
        wb3 = openpyxl.load_workbook(f3, data_only=True)
        p = wb3['Добыча Сдача ']
        P['oil'] = [p.cell(8 + m, 3).value for m in range(1, 13)]
        P['sd'] = [p.cell(8 + m, 5).value for m in range(1, 13)]
        json.dump(P, open(PLAN_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    return P


def clean(v):
    if isinstance(v, str):
        if v.startswith('#'):
            return None
        return v
    if isinstance(v, float):
        return round(v, 2)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, dt.time):
        return v.strftime('%H:%M')
    return v


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def lab(*vals):
    s = ' | '.join(str(v).strip().replace('\n', ' ') for v in vals if v not in (None, ''))
    # убираем китайские иероглифы
    s = re.sub(r'[一-鿿（）、，：；·]+', '', s)
    return re.sub(r'\s+', ' ', s).strip(' |')


def row_vals(ws, r, c0, n=31):
    out = []
    for c in range(c0, c0 + n):
        v = ws.cell(r, c).value
        out.append(round(v, 2) if isinstance(v, (int, float)) and not isinstance(v, bool) else None)
    return out


def trim(arr):
    while arr and arr[-1] is None:
        arr = arr[:-1]
    return arr


def find_month_rows(ws, col, names):
    res = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str):
            k = v.strip().lower()
            for i, n in enumerate(names):
                if k == n.lower() or k == 'параметры ' + n.lower():
                    res.setdefault(i + 1, r)
    return res


def main():
    # usage: build_data.py <8.xlsx> <9.xlsx> [<план-год.xlsx>|-] [out=data.json]
    f8, f9 = sys.argv[1:3]
    f3 = sys.argv[3] if len(sys.argv) > 3 else '-'
    out = sys.argv[4] if len(sys.argv) > 4 else 'data.json'
    wb8 = openpyxl.load_workbook(f8, data_only=True)
    wb9 = openpyxl.load_workbook(f9, data_only=True)
    PLAN = load_plan(f3)
    STEAM_PLAN_MONTH = PLAN['steam']; VOLGA_PLAN_MONTH = PLAN['volga']

    # ---------- дата ----------
    s1 = wb8['1-Мәліметтер']
    d = s1['AL1'].value
    day = int(s1['AM1'].value)
    month = d.month
    dim = (dt.date(d.year + (month == 12), month % 12 + 1, 1) - dt.timedelta(days=1)).day
    D = {'meta': {'date': d.strftime('%Y-%m-%d'), 'day': day, 'month': month, 'year': d.year, 'dim': dim,
                  'months': MONTHS_RU}}

    # ---------- план по месяцам (файл 3) ----------
    oil_plan = PLAN['oil']; sd_plan = PLAN['sd']; gas_plan = PLAN.get('gas')
    if PLAN.get('year') and PLAN['year'] != d.year:
        print('ВНИМАНИЕ: план в data/plan.json на %s год, а сводка за %s — обновите план' % (PLAN['year'], d.year))

    # ---------- основные ряды из листа 1-Мәліметтер ----------
    def series_first(row0, plan_month, name, unit, with_plan_cols=True):
        s = {'name': name, 'unit': unit, 'daily': {}, 'dayPlan': {}, 'monthPlan': {}}
        for m in range(1, 13):
            r = row0 + m - 1
            vals = row_vals(s1, r, 4, 31)
            vals = trim(vals)
            if m <= month:
                s['daily'][m] = vals
            if with_plan_cols:
                s['dayPlan'][m] = num(s1.cell(r, 35).value)
            s['monthPlan'][m] = plan_month[m - 1] if plan_month else None
        return s

    series = {}
    series['oil'] = series_first(17, oil_plan, 'Добыча нефти', 'т')
    series['sd'] = series_first(41, sd_plan, 'Сдача нефти', 'т')
    # пар
    steam = {'name': 'Закачка пара', 'unit': 'т', 'daily': {}, 'dayPlan': {}, 'monthPlan': {}}
    for m in range(1, 13):
        if m <= month:
            steam['daily'][m] = trim(row_vals(s1, 160 + m, 4, 31))
        dm = (dt.date(d.year + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).day
        steam['monthPlan'][m] = STEAM_PLAN_MONTH[m - 1]
        steam['dayPlan'][m] = round(STEAM_PLAN_MONTH[m - 1] / dm, 2)
    series['steam'] = steam
    # ОЗПВ (прием опресненной воды) – лист ОЗПВ
    sw = wb8['ОЗПВ']
    oz_rows = find_month_rows(sw, 1, MONTHS_KK)
    ozpv = {'name': 'Опреснённая вода ОЗПВ', 'unit': 'м³', 'daily': {}, 'dayPlan': {}, 'monthPlan': {}}
    for m in range(1, month + 1):
        ozpv['daily'][m] = trim(row_vals(sw, oz_rows[m] + 1, 2, 31))
    for mm, pv in PLAN['ozpv_month'].items():
        ozpv['monthPlan'][int(mm)] = pv
    if ozpv['monthPlan'].get(month):
        ozpv['dayPlan'][month] = round(ozpv['monthPlan'][month] / dim, 2)
    series['ozpv'] = ozpv
    # Волжская вода – лист Волжская вода
    sv = wb8['Волжская вода']
    vv_rows = find_month_rows(sv, 1, MONTHS_RU + [])
    volga = {'name': 'Волжская вода', 'unit': 'м³', 'daily': {}, 'dayPlan': {}, 'monthPlan': {}}
    for m in range(1, month + 1):
        st = 3 + 36 * (m - 1)
        assert 'Прием Волжской' in str(sv.cell(st + 3, 1).value), (m, sv.cell(st + 3, 1).value)
        volga['daily'][m] = trim(row_vals(sv, st + 3, 5, 31))
    for m in range(1, 13):
        dm = (dt.date(d.year + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).day
        volga['monthPlan'][m] = VOLGA_PLAN_MONTH[m - 1]
        volga['dayPlan'][m] = round(VOLGA_PLAN_MONTH[m - 1] / dm, 2)
    series['volga'] = volga
    # Сдача на битум
    bit = {'name': 'Сдача нефти на битум', 'unit': 'т', 'monthFact': {}, 'monthPlan': {}}
    for m in range(1, month + 1):
        pr = 185 + 2 * (m - 1)
        bit['monthPlan'][m] = num(s1.cell(pr, 37).value)
        bit['monthFact'][m] = num(s1.cell(pr + 1, 37).value)
    series['bitum'] = bit
    # закачка воды в пласт (внутриконтурная по Siemens) - план из 9 файла
    D['series'] = series

    # ---------- сводные цифры (как в еженедельном отчёте ПУ) ----------
    def summ(key):
        s = series[key]
        dm = dim
        mf = sum(x or 0 for x in s['daily'][month][:day])
        mp = (s['dayPlan'][month] or 0) * day
        ytd_f = sum(sum(x or 0 for x in s['daily'][m][:(day if m == month else 31)]) for m in range(1, month + 1))
        ytd_p = sum((s['monthPlan'].get(m) or 0) for m in range(1, month)) + mp
        yp = sum((s['monthPlan'].get(m) or 0) for m in range(1, 13))
        return {'monthPlan': s['monthPlan'][month], 'mtdPlan': round(mp), 'mtdFact': round(mf), 'mtdDev': round(mf - mp),
                'yearPlan': round(yp), 'ytdPlan': round(ytd_p), 'ytdFact': round(ytd_f), 'ytdDev': round(ytd_f - ytd_p),
                'dayPlan': s['dayPlan'][month], 'dayFact': s['daily'][month][day - 1]}

    summary = {k: summ(k) for k in ('oil', 'sd', 'steam', 'volga')}
    oz = summ('ozpv')
    oz_pm = {int(k): v for k, v in PLAN['ozpv_month'].items()}
    oz_known = PLAN['ozpv_before'] + sum(v for m_, v in oz_pm.items() if m_ < month)   # янв-окт одной суммой + полные месяцы
    oz['yearPlan'] = PLAN['ozpv_before'] + sum(oz_pm.values())
    oz['ytdPlan'] = round(oz_known + oz['mtdPlan']) if month >= min(oz_pm) else None
    oz['ytdDev'] = oz['ytdFact'] - oz['ytdPlan'] if oz['ytdPlan'] is not None else None
    summary['ozpv'] = oz
    bf = sum(bit['monthFact'].values() or [0])
    summary['bitum'] = {'monthPlan': None, 'mtdPlan': bit['monthPlan'].get(month) or 0, 'mtdFact': bit['monthFact'].get(month),
                        'mtdDev': bit['monthFact'].get(month), 'yearPlan': None, 'ytdPlan': None, 'ytdFact': bf, 'ytdDev': bf}
    D['summary'] = summary

    # ---------- фонды ----------
    sf = wb8['Доб.фонд']
    rows = find_month_rows(sf, 1, MONTHS_KK)
    fund_p = {}
    offs = {'exploit': 2, 'active': 5, 'bd': 8, 'develop': 9, 'working': 10, 'idle': 14, 'idleLiquid': 15, 'idleOil': 16}
    for m in range(1, month + 1):
        st = rows[m]
        fund_p[m] = {k: trim(row_vals(sf, st + o, 4, 31)) for k, o in offs.items()}
        fund_p[m]['_labels'] = {k: lab(sf.cell(st + o, 1).value, sf.cell(st + o, 2).value, sf.cell(st + o, 3).value) for k, o in offs.items()} if m == month else None
    D['fundProd'] = fund_p
    sfi = wb8['Нагнет.фонд']
    rowsi = find_month_rows(sfi, 1, MONTHS_KK)
    fund_i = {}
    inj_blocks = {}
    for m in range(1, month + 1):
        st = rowsi[m]
        rs = []
        for r in range(st + 2, st + 40):
            l = lab(sfi.cell(r, 1).value)
            if not l and not any(sfi.cell(r, c).value is not None for c in range(3, 8)):
                break
            if str(sfi.cell(r, 1).value or '').strip() in MONTHS_KK or str(sfi.cell(r, 1).value or '').startswith('Параметры'):
                break
            rs.append({'l': l or lab(sfi.cell(r, 2).value), 'v': trim(row_vals(sfi, r, 3, 31))})
        inj_blocks[m] = rs
    D['fundInj'] = inj_blocks

    # ---------- простои добывающих ----------
    sp = wb8['Простой доб. скв.']
    idle_p = {}
    for m in range(1, month + 1):
        st = 4 + 18 * (m - 1)
        causes = []
        for r in range(st + 3, st + 17):
            l = lab(sp.cell(r, 2).value)
            if not l:
                continue
            wells = [clean(sp.cell(r, 3 + 2 * i).value) for i in range(31)]
            loss = [clean(sp.cell(r, 4 + 2 * i).value) for i in range(31)]
            causes.append({'l': l, 'w': trim([num(x) for x in wells]), 'q': trim([num(x) for x in loss])})
        idle_p[m] = causes
    D['idleProd'] = idle_p
    sn = wb8['Простой нагн. скв.']
    idle_i = {}
    for m in range(1, month + 1):
        st = 4 + 17 * (m - 1)
        causes = []
        for r in range(st + 3, st + 17):
            l = lab(sn.cell(r, 2).value)
            if not l:
                continue
            causes.append({'l': l, 'w': trim(row_vals(sn, r, 3, 31))})
        idle_i[m] = causes
    D['idleInj'] = idle_i

    # ---------- закачка воды (ППД) ----------
    sz = wb8['Закачка воды(сброс)']
    zrows = find_month_rows(sz, 1, MONTHS_KK)
    inj = {}
    for m in range(1, month + 1):
        st = zrows[m]
        rs = []
        for r in range(st + 1, st + 66):
            a = sz.cell(r, 1).value
            if isinstance(a, str) and a.startswith('Параметры'):
                break
            l = lab(a, sz.cell(r, 2).value, sz.cell(r, 3).value)
            vals = trim(row_vals(sz, r, 4, 31))
            if l or vals:
                rs.append({'l': l, 'v': vals})
        inj[m] = rs
    D['inj'] = inj

    # ---------- пар: ППГ / СПГУ / МПГУ ----------
    gens = {}
    for name in ('ППГ', 'СПГУ', 'МПГУ'):
        ws = wb8[name]
        gm = {}
        for m in range(1, month + 1):
            rs = []
            first = None
            for r in range(1, 200):
                v = ws.cell(r, 4).value
                if isinstance(v, str) and v.strip().lower() == MONTHS_KK[m - 1] and first is None and r > 4:
                    first = r
                if first is not None and r >= first:
                    if r - first > 6:
                        break
                    l = lab(ws.cell(r, 1).value, ws.cell(r, 3).value)
                    rs.append({'l': l, 'v': trim(row_vals(ws, r, 5, 31))})
            gm[m] = rs
        gens[name] = gm
    # простои скважин ППГ
    ws = wb8['ППГ']
    stops = {}
    for m in range(1, month + 1):
        r0 = 101 + 4 * (m - 1)
        rs = []
        for r in range(r0, r0 + 4):
            rs.append({'l': lab(ws.cell(r, 1).value), 'v': trim(row_vals(ws, r, 5, 31))})
        stops[m] = rs
    gens['ППГ_stop'] = stops
    D['gens'] = gens

    # ---------- водные балансы ----------
    vb = {}
    for m in range(1, month + 1):
        st = 3 + 36 * (m - 1)
        rs = []
        for r in range(st + 2, st + 34):
            l = lab(sv.cell(r, 1).value, sv.cell(r, 3).value)
            vals = trim(row_vals(sv, r, 5, 31))
            if l and vals:
                rs.append({'l': l, 'v': vals})
        vb[m] = rs
    D['volgaBal'] = vb
    oz_bal = {}
    for m in range(1, month + 1):
        st = oz_rows[m]
        rs = []
        for r in range(st + 1, st + 12):
            l = lab(sw.cell(r, 1).value)
            vals = trim(row_vals(sw, r, 2, 31))
            if l and vals:
                rs.append({'l': l, 'v': vals})
        oz_bal[m] = rs
    D['ozpvBal'] = oz_bal

    # ---------- химреагенты ----------
    sc = wb8['хим.реагент']
    crow = find_month_rows(sc, 1, MONTHS_KK)
    chem = {}
    for m in range(1, month + 1):
        st = crow[m]
        rs = []
        for r in range(st + 2, st + 21):
            l = lab(sc.cell(r, 2).value)
            vals = trim(row_vals(sc, r, 6, 31))
            if l and vals:
                rs.append({'l': l, 'v': vals})
        chem[m] = rs
    D['chem'] = chem

    # ================= 9-я сводка (суточный рапорт) =================
    r9 = wb9['Добыча']
    today = {}
    today['gas'] = {'plan': num(r9['C10'].value), 'fact': num(r9['D10'].value), 'mtdPlan': num(r9['F10'].value), 'mtdFact': num(r9['G10'].value)}
    today['water'] = {
        'podtov': {'dayPlan': num(r9['C14'].value), 'dayFact': num(r9['D14'].value), 'mtdPlan': num(r9['F14'].value), 'mtdFact': num(r9['G14'].value), 'monthPlan': num(r9['B14'].value)},
        'steam': {'dayPlan': num(r9['C15'].value), 'dayFact': num(r9['D15'].value), 'mtdPlan': num(r9['F15'].value), 'mtdFact': num(r9['G15'].value)},
    }
    today['oilDay'] = {'plan': num(r9['C5'].value), 'fact': num(r9['D5'].value), 'sdPlan': num(r9['J5'].value), 'sdFact': num(r9['K5'].value)}
    today['fundProd'] = {'exploit': r9['B20'].value, 'active': r9['C20'].value, 'idle': r9['D20'].value, 'bd': r9['E20'].value, 'develop': r9['F20'].value}
    today['fundInj'] = {'exploit': r9['G20'].value, 'active': r9['H20'].value, 'idle': r9['I20'].value, 'bd': r9['J20'].value, 'develop': r9['K20'].value}
    today['drilling'] = {'planMonth': num(r9['B25'].value), 'dayPlan': num(r9['C25'].value), 'dayFact': num(r9['D25'].value), 'mtdPlan': num(r9['F25'].value), 'mtdFact': num(r9['G25'].value),
                         'mtdDev': num(r9['H25'].value), 'inPlanWells': r9['I25'].value, 'wellsPlanMtd': r9['K25'].value, 'wellsFactMtd': r9['L25'].value}
    idle = []
    for r in range(9, 21):
        idle.append({'l': lab(r9.cell(r, 14).value), 'n': num(r9.cell(r, 18).value), 'q': num(r9.cell(r, 19).value)})
    today['idleCauses'] = idle
    pu = []
    for r in range(22, 30):
        pu.append({'l': lab(r9.cell(r, 14).value), 'n': r9.cell(r, 18).value, 'w': clean(r9.cell(r, 19).value)})
    today['pu'] = pu
    today['reasons'] = [lab(r9.cell(r, 1).value) + ' ' + lab(r9.cell(r, 6).value) for r in range(30, 34) if r9.cell(r, 6).value]
    today['acts'] = [{'cause': lab(r9.cell(r, 2).value), 'wells': r9.cell(r, 5).value, 'q': r9.cell(r, 6).value, 'qn': r9.cell(r, 7).value, 'why': lab(r9.cell(r, 8).value)} for r in range(40, 45) if r9.cell(r, 2).value]
    today['park'] = {'left': num(r9['C57'].value), 'rvs': lab(r9['A48'].value), 'note': lab(r9['F48'].value)}
    # потери (внутр.)
    # КРС
    kr = wb9['КРС']
    crews = []
    for r in range(4, 40):
        if kr.cell(r, 2).value is None or not isinstance(kr.cell(r, 1).value, (int, float)):
            continue
        reason = lab(kr.cell(r, 6).value)
        crews.append({'n': kr.cell(r, 1).value, 'org': lab(kr.cell(r, 2).value), 'crew': lab(kr.cell(r, 3).value), 'rig': lab(kr.cell(r, 4).value),
                      'well': str(kr.cell(r, 5).value), 'reason': KK_REASON.get(reason.strip().lower(), reason),
                      'plan': lab(kr.cell(r, 7).value), 'done': lab(kr.cell(r, 8).value), 'st': lab(kr.cell(r, 9).value)})
    today['krs'] = crews
    # бурение
    br = wb9['бурение']
    rigs = []
    for r in range(6, 12):
        if br.cell(r, 2).value and isinstance(br.cell(r, 1).value, (int, float)):
            rigs.append({'contractor': lab(br.cell(r, 2).value), 'rig': lab(br.cell(r, 3).value), 'gzu': lab(br.cell(r, 4).value), 'well': str(br.cell(r, 5).value),
                         'depth': num(br.cell(r, 6).value), 'start': lab(br.cell(r, 7).value), 'end': lab(br.cell(r, 8).value), 'perDay': num(br.cell(r, 9).value),
                         'bottom': num(br.cell(r, 10).value), 'mtd': num(br.cell(r, 11).value), 'work': lab(br.cell(r, 13).value)})
    today['drillRigs'] = rigs
    # свабирование
    sw9 = wb9[' свабирование']
    sv_rows = []
    for r in range(5, 14):
        if sw9.cell(r, 2).value and sw9.cell(r, 4).value:
            sv_rows.append({'n': lab(sw9.cell(r, 2).value), 'well': sw9.cell(r, 4).value, 'gzu': lab(sw9.cell(r, 5).value), 'start': clean(sw9.cell(r, 6).value),
                            'liq': num(sw9.cell(r, 8).value), 'oil': num(sw9.cell(r, 9).value), 'wc': num(sw9.cell(r, 10).value), 'sand': num(sw9.cell(r, 11).value),
                            'cum': num(sw9.cell(r, 13).value)})
    today['swab'] = {'rows': sv_rows, 'totLiq': num(sw9['H15'].value), 'totOil': num(sw9['I15'].value)}
    # техника
    tx = wb9['Выход техники ГПУ']
    today['tech'] = {'rows': [{'l': lab(tx.cell(r, 2).value), 'plan': num(tx.cell(r, 3).value), 'fact': num(tx.cell(r, 4).value), 'dev': num(tx.cell(r, 5).value)} for r in range(4, 9) if tx.cell(r, 2).value],
                     'garage': [{'l': lab(tx.cell(r, 7).value) or 'Итого', 'plan': num(tx.cell(r, 8).value), 'fact': num(tx.cell(r, 9).value), 'dev': num(tx.cell(r, 10).value)} for r in range(4, 8)],
                     'reasons': [{'l': lab(tx.cell(r, 12).value), 'v': num(tx.cell(r, 13).value)} for r in range(1, 11) if tx.cell(r, 12).value]}
    # парогенераторы: срез на сутки
    ms = wb9['МПГУ;СПГУ']
    units = []
    cur = None
    nominal = {'МПГУ-11тн': 264, 'МПГУ-23т': 552, 'СПГУ-18тн': 432}
    for r in range(10, 168):
        a = ms.cell(r, 1).value
        if isinstance(a, str) and re.match(r'(МПГУ|СПГУ)', a.strip()):
            name = re.sub(r'\s+', ' ', a.strip())
            kind = 'МПГУ-11тн' if '11тн' in name else 'МПГУ-23т' if '23т' in name else 'СПГУ-18тн' if '18тн' in name else None
            if kind is None:
                continue
            steam_d = num(ms.cell(r, 7).value)
            units.append({'name': name, 'kind': kind, 'nominal': nominal[kind], 'steam': steam_d, 'dry': num(ms.cell(r, 5).value), 'hours': num(ms.cell(r, 8).value),
                          'where': lab(ms.cell(r, 2).value, ms.cell(r + 1, 2).value, ms.cell(r + 1, 3).value, ms.cell(r + 3, 3).value) if kind != 'МПГУ-11тн' else lab(ms.cell(r, 2).value)})
    today['units'] = units
    today['ppg3'] = {'steam': num(ms['B4'].value), 'mtd': num(ms['C4'].value), 'volga': num(ms['B5'].value), 'volgaMtd': num(ms['C5'].value), 'pressure': num(ms['B6'].value)}
    today['ppgSpguTotals'] = {'mpgu': num(ms['G30'].value), 'mpgu23': num(ms['G63'].value), 'spgu': num(ms['G168'].value), 'collPPG': num(ms['G169'].value)}
    D['today'] = today
    try:
        e1 = r9['E1'].value
        D['meta']['d9'] = e1.strftime('%Y-%m-%d') if hasattr(e1, 'strftime') else None
    except Exception:
        D['meta']['d9'] = None

    # текущий месяц: обрезаем будущие дни (в источнике там нули/формулы)
    def cut(o, in_cur=False):
        if isinstance(o, dict):
            for k, v in o.items():
                if str(k) == str(month) and isinstance(v, (dict, list)):
                    cut(v, True)
                elif k in ('v', 'w', 'q') and in_cur and isinstance(v, list):
                    del v[day:]
                elif k == 'daily' or k == '_labels':
                    cut(v, in_cur)
                elif isinstance(v, (dict, list)):
                    cut(v, in_cur)
                elif isinstance(v, list):
                    pass
        elif isinstance(o, list):
            for x in o:
                if isinstance(x, (dict, list)):
                    cut(x, in_cur)

    for key in ('fundProd', 'fundInj', 'idleProd', 'idleInj', 'inj', 'volgaBal', 'ozpvBal', 'chem'):
        cut(D[key])
    for key in ('exploit', 'active', 'bd', 'develop', 'working', 'idle', 'idleLiquid', 'idleOil'):
        del D['fundProd'][month][key][day:]
    for g in D['gens'].values():
        cut(g)
    for s in D['series'].values():
        if 'daily' in s and month in s['daily']:
            del s['daily'][month][day:]

    json.dump(D, open(out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))

    # ---------- самопроверка относительно еженедельного отчёта ----------
    print('Дата:', D['meta']['date'])
    for k, v in summary.items():
        print(k, v)


if __name__ == '__main__':
    main()
