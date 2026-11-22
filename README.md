# Основные сводки
Интерактивные дашборды по суточным сводкам месторождения (PWA): добыча и сдача, план/факт/отклонение (сутки, с начала месяца, с начала года), простои скважин и потери нефти, ППД, пар (ППГ/СПГУ/МПГУ), Волжская и опреснённая вода, КРС, бурение, свабирование, химреагенты.

- `index.html` — приложение, `data.json` — данные, `sw.js` + `manifest.json` — PWA.
- Обновление данных: `python3 tools/build_data.py <8-файл.xlsx> <9-файл.xlsx> <план-год.xlsx> data.json`, затем commit/push.
- Установка: Android — Chrome → «Добавить на главный экран»; iPhone — Safari → «Поделиться» → «На экран Домой».

## Автообновление с рабочего ПК

ПК (PowerShell, задача «SVODKI upload») → Google Apps Script (`pc/Code.gs`) → GitHub, ветка `inbox` → Actions (`.github/workflows/build.yml`, `tools/build_all.py`) → `data.json`, `data/wellstatus.json`, `version.json` в `main` → GitHub Pages.

Папка `pc/` копируется на ПК целиком. Источники — `pc/sources.txt` (сводки №8 и №9 из Dispatcher, Well Status из архива; без подпапок).
`pc/send.txt` в репозитории — шаблон: на ПК в 1-ю строку адрес веб-приложения Apps Script, во 2-ю KEY. Ключ и токен в репозиторий не кладём.

Apps Script → Project Settings → Script properties: `GH_TOKEN` (доступ к Sync, Contents: read/write), `KEY`, `REPO=crownlink999/Sync`, `BRANCH=inbox`. Deploy → Web app → Execute as: Me, Access: Anyone.
Тип файла определяется по листам внутри, не по имени. Нет какого-то файла — остаются прежние данные, предупреждение в `version.json`.
