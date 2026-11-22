#!/usr/bin/env python3
"""Оркестратор сборки: файлы из inbox (присланы с рабочего ПК) -> data.json + data/wellstatus.json + version.json

    python3 tools/build_all.py <папка inbox/input> [<корень репозитория>=.]

Тип файла определяется ПО СОДЕРЖИМОМУ (имена на ПК бывают любые):
    сводка №8  - лист '1-Мәліметтер'
    сводка №9  - листы 'Добыча' и 'КРС'
    Well Status - лист 'TOTAL WELLS'
    план года   - лист 'Добыча Сдача '  (необязателен, план лежит в data/plan.json)
Если одного типа несколько файлов - берётся с самой свежей датой внутри файла.
Большие файлы приходят кусками name.xlsx.part001of002 - склеиваются здесь.
Если какого-то файла нет - соответствующая часть данных остаётся прежней (из репозитория), в version.json пишется предупреждение.
"""
import sys, os, re, json, glob, shutil, subprocess, tempfile, datetime as dt, collections
import openpyxl

INBOX = sys.argv[1]
ROOT = os.path.abspath(sys.argv[2] if len(sys.argv) > 2 else '.')
TOOLS = os.path.join(ROOT, 'tools')
TZ = dt.timezone(dt.timedelta(hours=5))
warns = []


def log(*a):
    print(*a, flush=True)


# ---------- 1. собрать файлы (склеить части) ----------
tmp = tempfile.mkdtemp(prefix='svodki_')
files = []   # (путь, исходное имя)
if os.path.isdir(INBOX):
    allf = [p for p in glob.glob(os.path.join(INBOX, '**', '*'), recursive=True) if os.path.isfile(p)]
    parts = collections.defaultdict(dict)
    for p in allf:
        name = os.path.basename(p)
        m = re.match(r'^(.*)\.part(\d+)of(\d+)$', name)
        if m:
            parts[(os.path.dirname(p), m.group(1), int(m.group(3)))][int(m.group(2))] = p
        elif re.search(r'\.(xlsx|xlsm)$', name, re.I) and not name.startswith('~$'):
            files.append((p, name))
    for (d, base, total), got in parts.items():
        if sorted(got) != list(range(1, total + 1)):
            warns.append('файл %s пришёл не полностью (%d из %d частей) - пропущен' % (base, len(got), total))
            continue
        dst = os.path.join(tmp, '%d_%s' % (len(files), base.replace('/', '_')))
        with open(dst, 'wb') as out:
            for i in range(1, total + 1):
                with open(got[i], 'rb') as f:
                    shutil.copyfileobj(f, out)
        files.append((dst, base))
log('файлов во входящей папке:', len(files))

# ---------- 2. определить тип и дату ----------
found = collections.defaultdict(list)    # type -> [(date, path, name)]
for path, name in files:
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sn = set(wb.sheetnames)
        d = None
        if '1-Мәліметтер' in sn:
            typ = 'f8'
            ws = wb['1-Мәліметтер']
            d = ws['AL1'].value
        elif 'TOTAL WELLS' in sn:
            typ = 'ws'
            d = wb['TOTAL WELLS']['A1'].value
        elif 'Добыча' in sn and 'КРС' in sn:
            typ = 'f9'
            d = wb['Добыча']['E1'].value
        elif 'Добыча Сдача ' in sn:
            typ = 'plan'
        else:
            log('  ? не опознан:', name, sorted(sn)[:6])
            continue
        wb.close()
        d = d.date() if isinstance(d, dt.datetime) else None
        log('  %-5s %s  дата %s' % (typ, name, d))
        found[typ].append((d or dt.date.min, os.path.getmtime(path), path, name))
    except Exception as e:
        warns.append('не удалось прочитать %s: %s' % (name, e))

best = {}
for typ, lst in found.items():
    lst.sort(key=lambda x: (x[0], x[1]))
    best[typ] = lst[-1]
    if len(lst) > 1:
        log('  несколько файлов типа %s - взят %s (%s)' % (typ, lst[-1][3], lst[-1][0]))

# ---------- 3. сборка ----------
def run(cmd):
    log('$', ' '.join(os.path.basename(c) if os.path.sep in c else c for c in cmd))
    r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True)
    out = (r.stdout or '') + (r.stderr or '')
    log('\n'.join('   ' + l for l in out.strip().splitlines()[-14:]))
    return r.returncode == 0


ver_old = {}
vp = os.path.join(ROOT, 'version.json')
if os.path.exists(vp):
    try:
        ver_old = json.load(open(vp, encoding='utf-8'))
    except Exception:
        ver_old = {}
src = dict(ver_old.get('src', {}))

if 'f8' in best and 'f9' in best:
    plan = best['plan'][2] if 'plan' in best else '-'
    ok = run([sys.executable, '-I', os.path.join(TOOLS, 'build_data.py'), best['f8'][2], best['f9'][2], plan, os.path.join(ROOT, 'data.json.new')])
    if ok:
        os.replace(os.path.join(ROOT, 'data.json.new'), os.path.join(ROOT, 'data.json'))
        src['f8'] = {'name': best['f8'][3], 'date': str(best['f8'][0])}
        src['f9'] = {'name': best['f9'][3], 'date': str(best['f9'][0])}
    else:
        warns.append('сводки №8/№9: ошибка сборки (data.json оставлен прежним)')
else:
    miss = [n for n, t in (('№8', 'f8'), ('№9', 'f9')) if t not in best]
    if miss:
        warns.append('нет файла сводки ' + ', '.join(miss) + ' - показаны прежние данные')

if 'ws' in best:
    ok = run([sys.executable, '-I', os.path.join(TOOLS, 'build_wellstatus.py'), best['ws'][2], os.path.join(ROOT, 'data', 'wellstatus.json.new')])
    if ok:
        os.replace(os.path.join(ROOT, 'data', 'wellstatus.json.new'), os.path.join(ROOT, 'data', 'wellstatus.json'))
        src['ws'] = {'name': best['ws'][3], 'date': str(best['ws'][0])}
    else:
        warns.append('Well Status: ошибка сборки или сверка не сошлась (прежние данные оставлены)')
        try:
            os.remove(os.path.join(ROOT, 'data', 'wellstatus.json.new'))
        except OSError:
            pass
else:
    warns.append('нет файла Well Status - вкладка «Скважины» показывает прежние данные')

now = dt.datetime.now(TZ)
json.dump({'built': now.strftime('%Y-%m-%d %H:%M'), 'src': src, 'warn': warns}, open(vp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
shutil.rmtree(tmp, ignore_errors=True)
log('предупреждения:', warns or 'нет')
