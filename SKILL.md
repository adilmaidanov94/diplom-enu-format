---
name: "diplom-enu-format"
description: "Оформление прикреплённой дипломной работы/проекта (.docx) по Положению ЕНУ ПОЛ ЕНУ 2.014-KV (ред. 31.08.2026): поля, шрифт, заголовки, рисунки, таблицы, формулы, нумерация, без пустот."
---

# Оформление дипломной работы по Положению ЕНУ (ПОЛ ЕНУ 2.014-KV)

Пользователь прикрепляет дипломную работу (проект) в .docx. Нужно вернуть **полностью оформленный по Положению** документ, в котором:
- нет пустот на страницах (пустых абзацев, ручных разрывов посреди главы, «дыр» перед рисунками и таблицами);
- **ни одна страница основной части не начинается с рисунка, таблицы, формулы, подписи или «где»**: сверху минимум 2 строки описательного текста, а перед каждой вставкой стоит текст со ссылкой на её номер («…представлена на рисунке 3», «…в таблице 2», «…по формуле (4)»). Исключение — «Продолжение таблицы N».

Содержание работы не переписывать. Допустимые правки текста: вставка описательных предложений со ссылкой на вставку, приведение подписей к форме «Рисунок N. Название», регистр заголовков, двоеточие после «где».

## Требования Положения (что делает скрипт)

| Параметр | Требование | Пункт |
|---|---|---|
| Бумага, шрифт | А4, Times New Roman 14, интервал одинарный, по ширине | 10.5 |
| Поля | левое 30, правое 10, верхнее 20, нижнее 25 мм | 10.6 |
| Абзацный отступ | 1,25 см | 10.6 |
| Заголовки | с абзацного отступа, с прописной, без точки в конце, без подчёркивания, не в кавычках; «1 Название», «1.1 Название» (номер без точки) | 10.9, 10.14 |
| Разделы | каждый раздел (глава) и структурный элемент — с новой страницы; от заголовка до текста и между подразделами — интервал (по умолчанию одна пустая строка через интервал абзаца, без пустых абзацев) | 10.15 |
| Нумерация страниц | арабскими, внизу по центру, без точки; обложка, титул, декларация, отчёт антиплагиата, задание, аннотация не нумеруются, но считаются; номер появляется с «Введения» | 10.10–10.12 |
| Рисунки | сразу после текста с первым упоминанием или на следующей странице; ссылка в тексте обязательна; под рисунком: пояснения → «Примечание - …»/«Источник - …» → «Рисунок N. Название» | 10.16 |
| Таблицы | после текста с первым упоминанием; ссылка обязательна; «Таблица N. Название» над таблицей; примечание/источник после таблицы; при переносе на новой странице над частью — «Продолжение таблицы N» (в работе на казахском — «N-кестенің жалғасы»), нижняя черта первой части не проводится; шрифт в таблице можно меньше | 10.17 |
| Формулы | отдельной строкой, сверху и снизу ≥ 1 свободная строка; номер (N) у правого края; пояснения начинаются со «где» без двоеточия, каждый символ с новой строки | 10.19 |
| Ссылки | [3], [3, с. 45], [5; 8; 12]; список — в порядке первого упоминания; ≥ 30 источников; у каждого URL/DOI в виде активной гиперссылки | 9.15, 10.20–10.22 |
| Приложения | каждое с новой страницы; «Приложение N» вверху справа, название — следующей строкой по центру; нумерация арабскими | 10.23 |
| Объём | работа ≥ 50 с., проект ≥ 40 с. (введение+основная часть+заключение); введение и заключение 1–3 с.; аннотация 150–200 слов | 9.6–9.14 |

Противоречия в самом Положении и принятые решения: п. 10.17 говорит «название под таблицей», но образец и правило переноса («слева над первой частью») ставят название **над** таблицей. Скрипт ставит над, с абзацного отступа (`--tbl-caption-indent 1`). Подпись рисунка в образце стоит у левого края без отступа (`--fig-caption-indent 0`). Если кафедра требует иначе, поменять флаги.

Язык работы скрипт определяет сам (по доле казахских букв ә, і, ң, ғ, ү, ұ, қ, ө, һ) и пишет отчёт `lang`. В работе на казахском: подписи «1-сурет. Атауы», «1-кесте. Атауы» (порядок слов автора сохраняется, приводится только пунктуация), продолжение таблицы — «1-кестенің жалғасы», ссылки в тексте вида «1-суретте», «2 және 3-кестеде» засчитываются; структурные заголовки — Мазмұны, Кіріспе, Қорытынды, Пайдаланылған әдебиеттер тізімі, Қосымша. Неправильная подпись продолжения (например, русская в казахской работе) заменяется на правильную.

Титульные страницы (всё до «Содержания»/«Введения») скрипт не трогает: у них свои формы (Приложение 2 к Положению).

## Порядок работы

Вести список задач (виджет справа) по шагам ниже.

### 1. Подготовка окружения

```bash
python3 -c "import docx" 2>/dev/null || pip install --break-system-packages -q python-docx
python3 -c "import pymupdf" 2>/dev/null || pip install --break-system-packages -q pymupdf
# без модуля Math LibreOffice не рисует формулы, и вёрстка считается неверно
ls /usr/lib/libreoffice/program/libsmlo.so >/dev/null 2>&1 || (apt-get update -q >/dev/null 2>&1; apt-get install -y -q libreoffice-math >/dev/null 2>&1)
fc-list | grep -qi "Liberation Serif" || echo "ВНИМАНИЕ: нет Liberation Serif (метрический аналог Times New Roman) — вёрстка будет неточной"
```

Записать оба скрипта из раздела «Скрипты» ниже в рабочую папку `vkr/` **дословно** (инструментом Write, без изменений): `vkr/vkr_format.py`, `vkr/vkr_layout.py`. Они должны лежать в одной папке.

### 2. Входной файл

- Работать с копией. `.doc` сначала конвертировать: `soffice --headless --convert-to docx file.doc`. PDF оформить нельзя: попросить .docx.
- Определить вид: «ДИПЛОМНЫЙ ПРОЕКТ» на титуле → `--kind project`, иначе `work`.
- Если в файле нет титульных страниц и он начинается сразу с «Содержания»/«Введения», спросить (или оценить по составу п. 9.2), с какого номера должна начинаться страница «Введение», и передать `--start-page N`. Если титульные страницы в файле есть, номер считается автоматически.

### 3. Первый проход форматирования

```bash
python3 vkr/vkr_format.py input.docx step1.docx --report fmt.json
```

Отчёт `fmt.json`: счётчики исправлений и `issues`. Разобрать каждую проблему:

| type | Что делать |
|---|---|
| `no_text_before` | вставка стоит сразу после заголовка или другой вставки. **Написать описательный текст** (1–2 предложения, не короче ~160 знаков, чтобы вышло ≥ 2 строки) с номером вставки, по смыслу подписи и окружающего текста, академическим стилем, на языке работы. Режим `insert`. |
| `no_reference` | текст перед вставкой есть, но номера вставки в нём нет. Если предыдущий абзац о том же — дописать к нему предложение (`append`), иначе вставить новый абзац (`insert`). Для формулы обычно хватает `append` с текстом «(N):», если абзац кончается словами «по формуле». |
| `no_caption` | рисунок/таблица без подписи: составить подпись по смыслу («Рисунок N. …»/«Таблица N. …») с правильным номером и вставить вручную через python-docx (абзац после рисунка / перед таблицей), затем повторить шаг 3. |
| `formula_no_number` | сообщить пользователю (нумерация формул и ссылки на них — решение автора); автоматически не перенумеровывать. |
| `numbering` | нумерация не сквозная: перечислить пользователю, не перенумеровывать молча. |
| `heading_caps` | заголовок прописными: перевести в «С прописной», сохранив аббревиатуры (ИИ, РК, IoT…), через `F.set_text(p, text)`. |
| `appendix_letter`, `refs_*` | сообщить пользователю в итоговом отчёте (номера приложений, порядок/число источников, нет URL/DOI). |

Файл `refs.json`, ключи берутся из отчёта (`key`):

```json
[
 {"key": "Рисунок 2", "text": "Модуль локализации объединяет данные лидара и одометрии и формирует оценку положения робота. Структура модуля представлена на рисунке 2."},
 {"key": "Таблица 3", "mode": "append", "text": "Сравнение методов приведено в таблице 3."},
 {"key": "Формула 4", "mode": "append", "text": "(4):"},
 {"key": "Формула #2", "text": "…"}
]
```

Ключи всегда пишутся по-русски («Рисунок 2», «Таблица 3») и в казахской работе тоже — это внутренние метки; сам текст писать на языке работы (каз.: «… құрылымы 2-суретте көрсетілген.»). `Формула #k` — k-я формула без номера. Не выдумывать факты: описательный текст пересказывает то, что уже сказано в подписи и окружающем тексте.

```bash
python3 vkr/vkr_format.py step1.docx step2.docx --refs refs.json --report fmt2.json
```

Повторять, пока в `issues` не останутся только пункты для пользователя.

### 4. Постраничная вёрстка

```bash
python3 vkr/vkr_layout.py step2.docx final.docx --kind work --pdf preview.pdf --report layout.json [--fmt "--start-page 9"]
```

Скрипт в цикле рендерит PDF и исправляет первую проблемную страницу сверху вниз:
- пустота > 3 строк перед рисунком → рисунок переносится на абзац ниже (до 3 раз, не через заголовок), иначе уменьшается (не меньше 60 %);
- таблица перешла на новую страницу → разрыв с подписью «Продолжение таблицы N» (каз. «N-кестенің жалғасы») над перенесённой частью, повтором шапки и без нижней черты первой части; устаревший разрыв (после правок выше) склеивается и делается заново;
- небольшой таблице (≤ 12 строк), которая не помещается и оставляет пустоту, разрешается разрыв;
- формулы и таблицы **никогда** не переносятся (формула — часть предложения; таблица идёт сразу после ссылки).

Основная защита работает в самом Word: перед каждой вставкой абзац с «не отрывать от следующего» + «запрет висячих строк», поэтому в Word страница не начнётся со вставки и после правок студента. LibreOffice переносит такой абзац целиком, а Word только его последние строки, поэтому для рендера скрипт делает техническую копию, имитирующую Word. `preview.pdf` — это рендер этой копии, только для проверки (в нём бывают «рваные» строки), пользователю его не отдавать как итоговый PDF.

`layout.json`: `actions` — что сделано; `remaining` — нерешённое; `minor_voids` — пустоты ≤ 6 строк от неразрывного блока «2 строки + подпись + 2 строки таблицы» (допустимо); `volume_pages` и `volume_issues` — объём по главам.

### 5. Проверка глазами (обязательно)

```bash
pdftoppm -r 40 -png preview.pdf pg
# собрать контактный лист по 8 страниц в ряд и посмотреть (Read)
```

Проверить: начало каждой страницы, места вставок, «Продолжение таблицы», номера страниц, отсутствие пустот. Нерешённое из `remaining` исправить вручную (python-docx по тем же правилам) и повторить шаг 4.

### 6. Сдача

- Имя: `<исходное имя>_оформлено_v1.docx`; каждая следующая правка — `_v2`, `_v3`…
- Если подключена рабочая папка проекта на компьютере пользователя, сохранить файл туда; в любом случае отправить через SendUserFile.
- Кратко, по-русски: что исправлено (счётчики из отчётов), какие описательные тексты вставлены (списком: куда и что), что осталось студенту (нумерация, источники, объём, формулы без номеров, `remaining`). Предупредить: при открытии Word спросит об обновлении полей — ответить «Да», чтобы обновились номера страниц в содержании. Финальную проверку пагинации делать в Word (LibreOffice даёт близкую, но не идентичную вёрстку).

## Скрипты

### vkr/vkr_format.py

```python
#!/usr/bin/env python3
"""Оформление дипломной работы (проекта) .docx по ПОЛ ЕНУ 2.014-KV (ред. 31.08.2026).

python3 vkr_format.py in.docx out.docx [--refs refs.json] [--start-page N]
        [--fig-caption-indent 0|1] [--tbl-caption-indent 0|1] [--heading-gap 1]
        [--report report.json]

Идемпотентен: можно прогонять повторно на своём же результате.
"""
import argparse, copy, json, re, sys
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Mm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W_NS, 'm': M_NS,
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'v': 'urn:schemas-microsoft-com:vml', 'o': 'urn:schemas-microsoft-com:office:office',
      'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'}

# ---- параметры Положения -------------------------------------------------
FONT = 'Times New Roman'
SIZE = 14
INDENT = Mm(12.5)                       # абзацный отступ 1,25 см
MARGINS = dict(left=Mm(30), right=Mm(10), top=Mm(20), bottom=Mm(25))
TEXT_W_MM = 170                         # 210 - 30 - 10
TEXT_H_MM = 252                         # 297 - 20 - 25
LINE_PT = 16                            # высота строки 14 pt при одинарном интервале
CHARS_PER_LINE = 78                     # оценка знаков в строке
MAX_IMG_H_MM = 195                      # рисунок + 2 строки текста + подпись помещаются на страницу
SMALL_TABLE_ROWS = 12                   # такие таблицы не разрываются

STRUCT_RE = re.compile(
    r'^(введение|заключение|содержание|оглавление|аннотация|abstract|аңдатпа|андатпа|'
    r'список\s+(использованн\w+\s+)?(литературы|источников)|список\s+литературы|'
    r'библиографический\s+список|нормативные\s+ссылки|'
    r'определения(\s*,?\s*обозначения\s+и\s+сокращения)?|обозначения\s+и\s+сокращения|'
    r'перечень\s+(сокращений|условных\s+обозначений)|кіріспе|қорытынды|мазмұны|пайдаланылған\s+әдебиеттер\s+тізімі|'
    r'introduction|conclusion|references|contents)$', re.I)
APPX_RE = re.compile(r'^(приложение|қосымша|appendix)(\s+[а-яёa-z0-9]{1,3})?$', re.I)
INTRO_RE = re.compile(r'^(введение|кіріспе|introduction)$', re.I)
TOC_RE = re.compile(r'^(содержание|оглавление|мазмұны|contents)$', re.I)
REFS_RE = re.compile(r'^(список\s+(использованн\w+\s+)?(литературы|источников)|список\s+литературы|'
                     r'библиографический\s+список|пайдаланылған\s+әдебиеттер\s+тізімі|references)$', re.I)
NUMH_RE = re.compile(r'^(\d{1,2})((?:\.\d{1,2}){0,3})\.?\s+(\S.*)$')
class _M:
    """Результат разбора подписи: group(1) — слово, group(2) — номер, group(3) — название."""
    def __init__(self, word, num, title, kz):
        self._g = (None, word, num, title); self.kz = kz
    def group(self, i):
        return self._g[i]

class CapRe:
    """Подпись в двух формах: «Рисунок 1. Название» и казахская «1-сурет. Название»."""
    def __init__(self, words, kz_words):
        self.ru = re.compile(r'^\s*(' + words + r')\s*(\d+(?:\.\d+)?)\s*[.\-–—:]?\s*(.*)$', re.I)
        self.kz = re.compile(r'^\s*(\d+(?:\.\d+)?)\s*[-–]?\s*(' + kz_words + r')(?![а-яәіңғүұқөһё])\s*[.\-–—:]?\s*(.*)$', re.I)
    def match(self, text):
        m = self.ru.match(text)
        if m:
            return _M(m.group(1), m.group(2), m.group(3), False)
        m = self.kz.match(text)
        if m:
            return _M(m.group(2), m.group(1), m.group(3), True)
        return None

class ContRe:
    """«Продолжение таблицы 1», «Окончание таблицы 1», «1-кестенің жалғасы», «Кестенің жалғасы 1», «Continuation of Table 1»."""
    PATS = [re.compile(r'^\s*(Продолжение|Окончание)\s+(?:таблицы|табл\.)\s*(\d+(?:\.\d+)?)', re.I),
            re.compile(r'^\s*(\d+(?:\.\d+)?)\s*[-–]?\s*кестенің\s+(жалғасы|соңы)', re.I),
            re.compile(r'^\s*(Кестенің\s+(?:жалғасы|соңы))\s*(\d+(?:\.\d+)?)', re.I),
            re.compile(r'^\s*(Continuation\s+of\s+Table|Table)\s*(\d+(?:\.\d+)?)\s*\(?\s*continued', re.I)]
    def match(self, text):
        for k, p in enumerate(self.PATS):
            m = p.match(text)
            if m:
                num = m.group(1) if k == 1 else m.group(2)
                return _M(m.group(0), num, '', k in (1, 2))
        return None

FIGCAP_RE = CapRe(r'Рисунок|Рис\.?|Сурет|Figure|Fig\.', r'сурет')
TBLCAP_RE = CapRe(r'Таблица|Табл\.|Кесте|Table', r'кесте')
CONT_RE = ContRe()

KZ_LETTERS = set('әіңғүұқөһӘІҢҒҮҰҚӨҺ')

def doc_lang(doc):
    """Язык работы: kk — казахский, en — английский, иначе ru."""
    txt = ''.join(t.text or '' for t in X(doc.element.body, './/w:t'))[:300000]
    cyr = sum(1 for ch in txt if 'а' <= ch.lower() <= 'я' or ch in 'ёЁ' or ch in KZ_LETTERS)
    lat = sum(1 for ch in txt if 'a' <= ch.lower() <= 'z')
    kz = sum(1 for ch in txt if ch in KZ_LETTERS)
    if lat > 2 * cyr:
        return 'en'
    return 'kk' if cyr and kz / cyr > 0.01 else 'ru'

def cont_caption(num, lang):
    """Подпись над перенесённой частью таблицы на языке работы."""
    return {'kk': f'{num}-кестенің жалғасы', 'en': f'Continuation of Table {num}'}.get(lang, f'Продолжение таблицы {num}')

NOTE_RE = re.compile(r'^\s*(Примечани[ея]|Источник|Ескерту|Дереккөз|Note|Source)\b', re.I)
WHERE_RE = re.compile(r'^\s*(где|мұндағы|where)\b\s*:?', re.I)
FNUM_RE = re.compile(r'^[\s,.;:]*(?:\(\s*(\d+(?:\.\d+)?)\s*\))?[\s,.;]*$')
WORD_FIX = {'рисунок': 'Рисунок', 'рис': 'Рисунок', 'рис.': 'Рисунок', 'сурет': 'Сурет', 'figure': 'Figure',
            'fig.': 'Figure', 'таблица': 'Таблица', 'табл.': 'Таблица', 'кесте': 'Кесте', 'table': 'Table'}

# ---- низкоуровневые помощники ---------------------------------------------
from lxml import etree
_XP = {}
def X(el, path):
    xp = _XP.get(path)
    if xp is None:
        xp = _XP[path] = etree.XPath(path, namespaces=NS)
    return xp(el)

def ptext(p):
    return ''.join(t.text or '' for t in X(p, './/w:t[not(ancestor::w:txbxContent)]'))

def tag(el):
    return el.tag.split('}')[1]

def has_drawing(p):
    return bool(X(p, './/w:drawing|.//w:pict')) or any(not is_eq_object(o) for o in X(p, './/w:object'))

def is_eq_object(o):
    return any('equation' in (e.get('ProgID') or '').lower() for e in X(o, './/o:OLEObject'))

def has_math(p):
    return bool(X(p, './/m:oMath')) or any(is_eq_object(o) for o in X(p, './/w:object'))

def get_pPr(p):
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr'); p.insert(0, pPr)
    return pPr

# порядок дочерних элементов w:pPr по схеме
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
             'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
             'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
             'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
             'sectPr', 'pPrChange']

def set_ppr(p, name, attrs=None, remove=False):
    pPr = get_pPr(p)
    old = pPr.find(qn('w:' + name))
    if old is not None:
        pPr.remove(old)
    if remove:
        return None
    el = OxmlElement('w:' + name)
    for k, v in (attrs or {}).items():
        el.set(qn('w:' + k), str(v))
    idx = PPR_ORDER.index(name)
    pos = 0
    for i, ch in enumerate(pPr):
        n = tag(ch)
        if n in PPR_ORDER and PPR_ORDER.index(n) < idx:
            pos = i + 1
    pPr.insert(pos, el)
    return el

def flag(p, name, on):
    if on:
        set_ppr(p, name)
    else:
        set_ppr(p, name, remove=True)
        # явно выключить, если стиль включает
        el = set_ppr(p, name, {'val': '0'})

def fmt_par(p, align='both', first=INDENT, left=0, before=0, after=0, kwn=False, keep=False, pbb=False):
    set_ppr(p, 'spacing', {'before': int(before * 20), 'after': int(after * 20), 'line': 240, 'lineRule': 'auto'})
    set_ppr(p, 'ind', {'left': int(left), 'firstLine': int(Emu(first).twips) if first else 0, 'right': 0})
    set_ppr(p, 'jc', {'val': align})
    set_ppr(p, 'widowControl')
    flag(p, 'keepNext', kwn)
    flag(p, 'keepLines', keep)
    flag(p, 'pageBreakBefore', pbb)
    set_ppr(p, 'contextualSpacing', remove=True)

def fmt_runs(p, size=SIZE, bold=None, italic=None, color_auto=True, keep_size_le=None):
    for r in X(p, './/w:r'):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
        fonts = rPr.find(qn('w:rFonts'))
        if fonts is None:
            fonts = OxmlElement('w:rFonts'); rPr.insert(0 if rPr.find(qn('w:rStyle')) is None else 1, fonts)
        for a in list(fonts.attrib):
            if a.endswith('Theme'):
                del fonts.attrib[a]
        for a in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
            fonts.set(qn('w:' + a), FONT)
        cur = rPr.find(qn('w:sz'))
        sz = size
        if keep_size_le and cur is not None:
            v = int(cur.get(qn('w:val'))) / 2
            sz = v if 10 <= v <= keep_size_le else keep_size_le
        for n in ('sz', 'szCs'):
            e = rPr.find(qn('w:' + n))
            if e is None:
                e = OxmlElement('w:' + n); _rpr_insert(rPr, e)
            e.set(qn('w:val'), str(int(sz * 2)))
        for n in ('spacing', 'w', 'position', 'highlight', 'shd', 'kern', 'u' if bold is not None else '__'):
            e = rPr.find(qn('w:' + n))
            if e is not None:
                rPr.remove(e)
        in_link = r.getparent().tag == qn('w:hyperlink')
        if color_auto and not in_link:
            e = rPr.find(qn('w:color'))
            if e is not None:
                rPr.remove(e)
        for n, val in (('b', bold), ('i', italic)):
            if val is None:
                continue
            for nn in (n, n + 'Cs'):
                e = rPr.find(qn('w:' + nn))
                if e is not None:
                    rPr.remove(e)
            if val:
                e = OxmlElement('w:' + n); _rpr_insert(rPr, e)

RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
             'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
             'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']

def _rpr_insert(rPr, el):
    idx = RPR_ORDER.index(tag(el))
    pos = 0
    for i, ch in enumerate(rPr):
        n = tag(ch)
        if n in RPR_ORDER and RPR_ORDER.index(n) < idx:
            pos = i + 1
    rPr.insert(pos, el)

def set_text(p, text):
    """Заменить текст абзаца, сохранив формат первого текстового прогона."""
    runs = [r for r in X(p, './w:r|./w:hyperlink/w:r') if X(r, './w:t')]
    if not runs:
        r = OxmlElement('w:r'); t = OxmlElement('w:t'); r.append(t); p.append(r); runs = [r]
    first = runs[0]
    ts = X(first, './w:t')
    ts[0].text = text
    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for t in ts[1:]:
        first.remove(t)
    for r in runs[1:]:
        for t in X(r, './w:t'):
            r.remove(t)
        if not X(r, './/w:drawing|.//w:object|.//w:pict|.//w:fldChar|.//w:instrText|.//w:tab'):
            r.getparent().remove(r)

def new_par(text='', like=None):
    p = OxmlElement('w:p')
    if text:
        r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); r.append(t); p.append(r)
    return p

def est_lines(p):
    return max(1, -(-len(ptext(p)) // CHARS_PER_LINE))

def norm(s):
    return re.sub(r'\s+', ' ', s.replace('\xa0', ' ')).strip()

# ---- разбор документа -------------------------------------------------------
class Item:
    def __init__(self, el, kind, **kw):
        self.el, self.kind, self.zone, self.text, self.level = el, kind, 'front', '', None
        self.__dict__.update(kw)

def style_name(doc, p):
    ps = X(p, './w:pPr/w:pStyle')
    if not ps:
        return ''
    sid = ps[0].get(qn('w:val'))
    try:
        return doc.styles.get_by_id(sid, 1).name.lower()
    except Exception:
        return sid.lower()

def outline_level(doc, p):
    o = X(p, './w:pPr/w:outlineLvl')
    if o:
        v = int(o[0].get(qn('w:val')))
        return v + 1 if v < 9 else None
    sn = style_name(doc, p)
    m = re.match(r'(heading|заголовок)\s*(\d)', sn)
    if m:
        return int(m.group(2))
    return None

def all_bold(p):
    runs = [r for r in X(p, './/w:r') if (''.join(X(r, './w:t/text()'))).strip()]
    if not runs:
        return False
    for r in runs:
        b = X(r, './w:rPr/w:b')
        if not b or b[0].get(qn('w:val')) in ('0', 'false'):
            return False
    return True

def heading_level(doc, p):
    t = norm(ptext(p))
    if not t or len(t) > 220 or has_drawing(p) or has_math(p):
        return None
    tl = t.rstrip('.').strip()
    if STRUCT_RE.match(tl) or APPX_RE.match(tl):
        return 1
    lvl = outline_level(doc, p)
    m = NUMH_RE.match(t)
    if m and not t.endswith((':', ';', ',')):
        n_lvl = 1 + m.group(2).count('.')
        if lvl or all_bold(p):
            return n_lvl
    if lvl and lvl <= 3 and len(t) < 200:
        return lvl
    return None

def analyze(doc):
    body = doc.element.body
    els = [e for e in body if tag(e) in ('p', 'tbl')]
    # границы
    start = intro = refs = None
    for i, e in enumerate(els):
        if tag(e) != 'p':
            continue
        t = norm(ptext(e)).rstrip('.').strip()
        if start is None and TOC_RE.match(t):
            start = i
        if intro is None and INTRO_RE.match(t) and heading_level(doc, e):
            intro = i
            if start is None:
                start = i
    if start is None:
        start = 0
    items = []
    for i, e in enumerate(els):
        if i < start:
            items.append(Item(e, 'front')); continue
        if tag(e) == 'tbl':
            items.append(Item(e, 'tbl')); continue
        t = norm(ptext(e))
        lvl = heading_level(doc, e)
        if lvl:
            tl = t.rstrip('.').strip()
            kind = 'appx' if APPX_RE.match(tl) else 'heading'
            items.append(Item(e, kind, level=lvl, text=tl)); continue
        if has_drawing(e):
            items.append(Item(e, 'fig')); continue
        if has_math(e):
            m = FNUM_RE.match(t)
            if m:
                items.append(Item(e, 'formula', num=m.group(1))); continue
        if not t and not X(e, './/w:fldChar|.//w:instrText|./w:pPr/w:sectPr'):
            items.append(Item(e, 'empty')); continue
        items.append(Item(e, 'text', text=t))
    # зона содержания и списка литературы
    zone = 'front'
    for it in items:
        if it.kind == 'front':
            continue
        if it.kind in ('heading', 'appx') and it.level == 1:
            if TOC_RE.match(it.text):
                zone = 'toc'
            elif REFS_RE.match(it.text):
                zone = 'refs'
            elif it.kind == 'appx':
                zone = 'appx'
            else:
                zone = 'body'
        it.zone = zone
    return items, intro

def body_items(doc):
    items, intro = analyze(doc)
    return items, intro

def build_blocks(items):
    """Находит вставки: рисунки, таблицы, формулы, их подписи и примечания."""
    blocks = []
    n = len(items)
    used = set()
    i = 0
    while i < n:
        it = items[i]
        if it.kind in ('front', 'empty') or it.zone == 'toc':
            i += 1; continue
        # --- рисунок (или таблица-раскладка с подписью «Рисунок») ---
        is_fig_tbl = False
        if it.kind == 'tbl':
            for j in range(i + 1, min(n, i + 4)):
                if items[j].kind == 'text' and FIGCAP_RE.match(items[j].text):
                    prev = items[i - 1] if i else None
                    if not (prev and prev.kind == 'text' and TBLCAP_RE.match(prev.text)):
                        is_fig_tbl = True
                    break
                if items[j].kind != 'text' or len(items[j].text) > 300:
                    break
        if it.kind == 'fig' or is_fig_tbl:
            j = i
            while j + 1 < n and (items[j + 1].kind == 'fig' or (items[j + 1].kind == 'text' and re.match(r'^[а-яa-z]\)', items[j + 1].text) and len(items[j + 1].text) < 200)):
                j += 1
            members = list(range(i, j + 1))
            cap = None
            k = j + 1
            mids = []
            while k < n and k <= j + 5:
                if items[k].kind == 'text' and FIGCAP_RE.match(items[k].text):
                    cap = k; break
                if items[k].kind == 'text' and len(items[k].text) < 300 and not TBLCAP_RE.match(items[k].text):
                    mids.append(k); k += 1; continue
                if items[k].kind == 'empty':
                    k += 1; continue
                break
            above = None
            if cap is None:
                mids = []
                if i > 0 and items[i - 1].kind == 'text' and FIGCAP_RE.match(items[i - 1].text):
                    above = i - 1
            blocks.append(dict(type='fig', fig=members, mids=mids if cap else [], cap=cap, cap_above=above))
            for x in members + (mids if cap else []) + ([cap] if cap else []) + ([above] if above is not None else []):
                used.add(x)
            i = (cap if cap else j) + 1
            continue
        if it.kind == 'tbl':
            cap = i - 1 if i > 0 and items[i - 1].kind == 'text' and (TBLCAP_RE.match(items[i - 1].text) or CONT_RE.match(items[i - 1].text)) else None
            below = None
            if cap is None and i + 1 < n and items[i + 1].kind == 'text' and TBLCAP_RE.match(items[i + 1].text):
                below = i + 1
            note = None
            k = i + 1 if below is None else i + 2
            if k < n and items[k].kind == 'text' and NOTE_RE.match(items[k].text):
                note = k
            cont = cap is not None and CONT_RE.match(items[cap].text)
            blocks.append(dict(type='tbl', tbl=i, cap=cap, cap_below=below, note=note, cont=bool(cont)))
            i = (note or below or i) + 1
            continue
        if it.kind == 'formula':
            where = i + 1 if i + 1 < n and items[i + 1].kind == 'text' and WHERE_RE.match(items[i + 1].text) else None
            blocks.append(dict(type='formula', f=i, where=where, num=it.num))
            i += 1
            continue
        i += 1
    return blocks

def block_start(b):
    if b['type'] == 'fig':
        return b['fig'][0]
    if b['type'] == 'tbl':
        return b['cap'] if b['cap'] is not None else b['tbl']
    return b['f']

def block_end(b):
    if b['type'] == 'fig':
        return b['cap'] if b['cap'] is not None else b['fig'][-1]
    if b['type'] == 'tbl':
        return b['note'] if b['note'] is not None else b['tbl']
    return b['where'] if b['where'] is not None else b['f']

def block_key(b, items):
    if b['type'] == 'fig':
        c = b['cap'] if b['cap'] is not None else b['cap_above']
        m = FIGCAP_RE.match(items[c].text) if c is not None else None
        return 'Рисунок ' + m.group(2) if m else None
    if b['type'] == 'tbl':
        c = b['cap'] if b['cap'] is not None else b['cap_below']
        if c is None:
            return None
        m = TBLCAP_RE.match(items[c].text) or CONT_RE.match(items[c].text)
        return ('Продолжение ' if b.get('cont') else '') + 'Таблица ' + m.group(2)
    return 'Формула ' + b['num'] if b['num'] else None

# ---- ссылки в тексте -----------------------------------------------------
def _nums(s):
    out = set()
    for part in re.split(r'\s*(?:,|\bи\b|\band\b|;)\s*', s):
        m = re.match(r'^(\d+(?:\.\d+)?)\s*[–\-—]\s*(\d+(?:\.\d+)?)$', part)
        if m:
            a, b = m.groups()
            if '.' not in a and '.' not in b:
                out.update(str(x) for x in range(int(a), int(b) + 1))
            else:
                pa, pb = a.split('.'), b.split('.')
                if len(pa) == 2 and len(pb) == 2 and pa[0] == pb[0]:
                    out.update(f'{pa[0]}.{x}' for x in range(int(pa[1]), int(pb[1]) + 1))
                out.update([a, b])
        elif re.match(r'^\d+(?:\.\d+)?$', part):
            out.add(part)
    return out

REF_PAT = {
    'fig': re.compile(r'(?:рис(?:унк\w*|унок|\.)?|сурет\w*|figures?|fig\.)\s*№?\s*((?:\d+(?:\.\d+)?)(?:\s*(?:,|–|-|—|\bи\b|\band\b)\s*\d+(?:\.\d+)?)*)', re.I),
    'tbl': re.compile(r'(?:табл(?:иц\w*|\.)?|кесте\w*|tables?)\s*№?\s*((?:\d+(?:\.\d+)?)(?:\s*(?:,|–|-|—|\bи\b|\band\b)\s*\d+(?:\.\d+)?)*)', re.I),
    'formula': re.compile(r'\((\d+(?:\.\d+)?)\)'),
}

REF_PAT_KZ = {
    'fig': re.compile(r'((?:\d+(?:\.\d+)?)(?:\s*(?:,|–|-|—|\bжәне\b)\s*\d+(?:\.\d+)?)*)\s*[-–]\s*сурет', re.I),
    'tbl': re.compile(r'((?:\d+(?:\.\d+)?)(?:\s*(?:,|–|-|—|\bжәне\b)\s*\d+(?:\.\d+)?)*)\s*[-–]\s*кесте', re.I),
}

def mentions(text, typ):
    out = set()
    if typ in REF_PAT_KZ:
        for m in REF_PAT_KZ[typ].finditer(text):
            out |= _nums(m.group(1).replace('және', ','))
    for m in REF_PAT[typ].finditer(text):
        if typ == 'formula':
            out.add(m.group(1))
        else:
            out |= _nums(m.group(1))
    return out

# ---- основная обработка --------------------------------------------------
def remove_empty(items, report):
    cnt = 0
    for idx, it in enumerate(items):
        if it.kind != 'empty' or it.zone in ('front',):
            continue
        p = it.el
        bms = X(p, './w:bookmarkStart|./w:bookmarkEnd')
        nxt = p.getnext()
        while nxt is not None and tag(nxt) != 'p':
            nxt = nxt.getnext()
        if bms and nxt is not None:
            pPr = nxt.find(qn('w:pPr'))
            for b in bms:
                if pPr is not None:
                    pPr.addnext(b)
                else:
                    nxt.insert(0, b)
        p.getparent().remove(p)
        cnt += 1
    report['empty_removed'] = cnt

def remove_page_breaks(doc, items, report):
    cnt = 0
    for it in items:
        if it.kind == 'front' or it.kind == 'tbl':
            continue
        for br in X(it.el, './/w:br[@w:type="page"]'):
            r = br.getparent()
            r.remove(br)
            if not X(r, './*[not(self::w:rPr)]'):
                r.getparent().remove(r)
            cnt += 1
    report['page_breaks_removed'] = cnt

def isolate_drawings(doc, items, report):
    """Рисунок, вставленный внутрь текстового абзаца, выносится в свой абзац."""
    cnt = 0
    for it in items:
        if it.kind != 'fig' or it.zone in ('front', 'toc'):
            continue
        p = it.el
        if not norm(ptext(p)):
            continue
        druns = [r for r in X(p, './w:r') if X(r, './/w:drawing|.//w:pict|.//w:object')]
        np_ = new_par()
        for r in druns:
            np_.append(r)
        p.addnext(np_)
        cnt += 1
    report['drawings_isolated'] = cnt

def anchors_to_inline(doc, report):
    cnt = 0
    for anc in X(doc.element.body, './/wp:anchor'):
        inline = OxmlElement('wp:inline')
        for a in ('distT', 'distB', 'distL', 'distR'):
            inline.set(a, '0')
        for name in ('extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr'):
            e = anc.find('{%s}%s' % (NS['wp'], name))
            if e is not None:
                inline.append(copy.deepcopy(e))
        if inline.find('{%s}effectExtent' % NS['wp']) is None:
            ee = OxmlElement('wp:effectExtent')
            for a in ('l', 't', 'r', 'b'):
                ee.set(a, '0')
            inline.insert(1, ee)
        g = anc.find('{%s}graphic' % NS['a'])
        inline.append(copy.deepcopy(g))
        anc.getparent().replace(anc, inline)
        cnt += 1
    report['floating_to_inline'] = cnt

def cap_images(doc, report, max_h_mm=MAX_IMG_H_MM):
    cnt = 0
    maxw = TEXT_W_MM * 36000
    maxh = max_h_mm * 36000
    for inl in X(doc.element.body, './/wp:inline'):
        ext = inl.find('{%s}extent' % NS['wp'])
        if ext is None:
            continue
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        k = min(1.0, maxw / cx if cx else 1, maxh / cy if cy else 1)
        if k < 0.999:
            scale_inline(inl, k); cnt += 1
    report['images_resized'] = cnt

def scale_inline(inl, k):
    ext = inl.find('{%s}extent' % NS['wp'])
    cx, cy = int(ext.get('cx')), int(ext.get('cy'))
    ext.set('cx', str(int(cx * k))); ext.set('cy', str(int(cy * k)))
    for e in X(inl, './/pic:spPr/a:xfrm/a:ext'):
        e.set('cx', str(int(int(e.get('cx')) * k))); e.set('cy', str(int(int(e.get('cy')) * k)))

def fix_caption_text(p, regex, report, key):
    t = norm(ptext(p))
    m = regex.match(t)
    title = m.group(3).strip()
    title = re.sub(r'\.$', '', title) if not title.endswith('..') else title
    if m.kz:
        head = f'{m.group(2)}-{m.group(1).lower()}'
    else:
        head = f'{WORD_FIX.get(m.group(1).lower(), m.group(1))} {m.group(2)}'
    new = f'{head}. {title}' if title else head
    if new != t:
        set_text(p, new)
        report.setdefault('captions_normalized', 0)
        report['captions_normalized'] += 1

def set_sect_type(sectPr, val):
    t = sectPr.find(qn('w:type'))
    if t is None:
        t = OxmlElement('w:type'); sectPr.insert(0, t)
    t.set(qn('w:val'), val)

def page_setup(doc, intro_el, start_page, report):
    body = doc.element.body
    # убрать лишние разрывы разделов «со следующей страницы» внутри основной части
    sects = X(body, './w:p/w:pPr/w:sectPr')
    final = body.find(qn('w:sectPr'))
    def orient(s):
        pg = s.find(qn('w:pgSz'))
        return pg.get(qn('w:orient')) if pg is not None else None
    passed_intro = False
    for s in sects:
        p = s.getparent().getparent()
        # раздел после введения?
        if intro_el is not None and _before(intro_el, p):
            nxt_s = _next_sect(s, final)
            if orient(s) == orient(nxt_s) and orient(s) != 'landscape':
                set_sect_type(nxt_s, 'continuous')
    # разрыв раздела перед «Введением» — для нумерации с введения
    if intro_el is not None:
        prev = intro_el.getprevious()
        while prev is not None and tag(prev) not in ('p', 'tbl'):
            prev = prev.getprevious()
        if prev is not None and not (tag(prev) == 'p' and X(prev, './w:pPr/w:sectPr')):
            if tag(prev) == 'tbl':
                holder = new_par(); prev.addnext(holder); prev = holder
            s = copy.deepcopy(final)
            for ref in X(s, './w:headerReference|./w:footerReference|./w:titlePg|./w:pgNumType'):
                s.remove(ref)
            t = s.find(qn('w:type'))
            if t is not None:
                s.remove(t)
            get_pPr(prev).append(s)
            report['section_break_before_intro'] = True
        nxt = intro_el
        # раздел, начинающийся с введения, должен начинаться с новой страницы
        s_of_intro = _sect_of(intro_el, final)
        t = s_of_intro.find(qn('w:type'))
        if t is not None and t.get(qn('w:val')) != 'nextPage':
            s_of_intro.remove(t)
    for sec in doc.sections:
        sec.left_margin, sec.right_margin = MARGINS['left'], MARGINS['right']
        sec.top_margin, sec.bottom_margin = MARGINS['top'], MARGINS['bottom']
        sec.header_distance = Mm(12.5); sec.footer_distance = Mm(12.5)
        if sec.orientation == 0:
            sec.page_width, sec.page_height = Mm(210), Mm(297)
    # нумерация
    secs = list(doc.sections)
    intro_idx = None
    if intro_el is not None:
        allp = list(body)
        pos = allp.index(intro_el)
        for k, sec in enumerate(secs):
            sp = sec._sectPr
            holder = sp.getparent().getparent() if sp.getparent().tag == qn('w:pPr') else None
            if holder is None or allp.index(holder) >= pos:
                intro_idx = k; break
    if intro_idx is None:
        intro_idx = 0
    for k, sec in enumerate(secs):
        tp = sec._sectPr.find(qn('w:titlePg'))
        if tp is not None and k >= intro_idx:
            sec._sectPr.remove(tp)
        if k < intro_idx:
            sec.footer.is_linked_to_previous = False
            for p in sec.footer.paragraphs:
                for r in list(p._p):
                    if tag(r) != 'pPr':
                        p._p.remove(r)
        elif k == intro_idx:
            sec.footer.is_linked_to_previous = False
            ftr = sec.footer
            ps = ftr.paragraphs
            for extra in ps[1:]:
                extra._p.getparent().remove(extra._p)
            p = ftr.paragraphs[0]._p
            for r in list(p):
                if tag(r) != 'pPr':
                    p.remove(r)
            set_ppr(p, 'jc', {'val': 'center'}); set_ppr(p, 'ind', {'firstLine': 0, 'left': 0})
            set_ppr(p, 'spacing', {'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto'})
            for kind, txt in (('begin', None), (None, ' PAGE '), ('separate', None), (None, '1'), ('end', None)):
                r = OxmlElement('w:r')
                if kind:
                    fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), kind); r.append(fc)
                else:
                    t = OxmlElement('w:instrText' if txt.strip() == 'PAGE' else 'w:t'); t.text = txt
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); r.append(t)
                p.append(r)
            fmt_runs(p)
            pn = sec._sectPr.find(qn('w:pgNumType'))
            if pn is None:
                pn = OxmlElement('w:pgNumType'); _sect_insert(sec._sectPr, pn)
            for a in list(pn.attrib):
                del pn.attrib[a]
            if start_page:
                pn.set(qn('w:start'), str(start_page))
        else:
            sec.footer.is_linked_to_previous = True
            pn = sec._sectPr.find(qn('w:pgNumType'))
            if pn is not None and pn.get(qn('w:start')):
                del pn.attrib[qn('w:start')]
    st = doc.settings.element
    for e in st.findall(qn('w:evenAndOddHeaders')):
        st.remove(e)
    if st.find(qn('w:updateFields')) is None:
        u = OxmlElement('w:updateFields'); u.set(qn('w:val'), 'true')
        after = ['hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr',
                 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats',
                 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType',
                 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator']
        nxt = next((ch for ch in st if tag(ch) in after), None)
        if nxt is not None:
            nxt.addprevious(u)
        else:
            st.append(u)

SECT_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar',
              'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote',
              'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange']

def _sect_insert(sp, el):
    idx = SECT_ORDER.index(tag(el)); pos = 0
    for i, ch in enumerate(sp):
        n = tag(ch)
        if n in SECT_ORDER and SECT_ORDER.index(n) < idx:
            pos = i + 1
    sp.insert(pos, el)

def _before(a, b):
    """a раньше b в документе"""
    allp = list(a.getparent())
    return allp.index(a) < allp.index(b)

def _sect_of(el, final):
    n = el
    while n is not None:
        if tag(n) == 'p' and X(n, './w:pPr/w:sectPr'):
            return X(n, './w:pPr/w:sectPr')[0]
        n = n.getnext()
    return final

def _next_sect(s, final):
    p = s.getparent().getparent().getnext()
    return _sect_of(p, final) if p is not None else final

def normalize_styles(doc):
    st = doc.styles['Normal']
    st.font.name = FONT; st.font.size = Pt(SIZE)
    rpr = st.element.get_or_add_rPr()
    f = rpr.find(qn('w:rFonts'))
    for a in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        f.set(qn('w:' + a), FONT)
    for a in list(f.attrib):
        if a.endswith('Theme'):
            del f.attrib[a]
    pf = st.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = 1.0
    # шрифт по умолчанию документа
    for f in X(doc.styles.element, './w:docDefaults//w:rFonts'):
        for a in list(f.attrib):
            del f.attrib[a]
        for a in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
            f.set(qn('w:' + a), FONT)

def add_hyperlinks(doc, items, report):
    url_re = re.compile(r'(https?://[^\s<>«»"]+[^\s<>«».,;:)"\]]|doi\.org/[^\s<>«»"]+[^\s.,;:)"\]]|\bDOI:?\s*10\.\d{4,9}/[^\s<>«»"]+[^\s.,;:)"\]])', re.I)
    cnt = 0
    for it in items:
        if it.zone != 'refs' or it.kind != 'text':
            continue
        p = it.el
        if X(p, './w:hyperlink'):
            continue
        text = ptext(p)
        matches = list(url_re.finditer(text))
        for m in reversed(matches):
            s, e = m.span()
            target = m.group(0)
            if target.lower().startswith('doi'):
                target = 'https://doi.org/' + re.sub(r'^(doi\.org/|doi:?\s*)', '', target, flags=re.I)
            runs = _split_runs_at(p, [s, e])
            pos = 0; inside = []
            for r in runs:
                rt = ''.join(X(r, './w:t/text()'))
                if pos >= s and pos + len(rt) <= e and rt:
                    inside.append(r)
                pos += len(rt)
            if not inside:
                continue
            rid = doc.part.relate_to(target, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
            h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), rid)
            inside[0].addprevious(h)
            for r in inside:
                h.append(r)
                rPr = r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
                for n in ('color', 'u'):
                    old = rPr.find(qn('w:' + n))
                    if old is not None:
                        rPr.remove(old)
                c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); _rpr_insert(rPr, c)
                u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); _rpr_insert(rPr, u)
            cnt += 1
    report['hyperlinks_added'] = cnt

def _split_runs_at(p, offsets):
    for off in sorted(set(offsets)):
        pos = 0
        for r in list(X(p, './w:r')):
            ts = X(r, './w:t')
            rt = ''.join(t.text or '' for t in ts)
            if pos < off < pos + len(rt):
                cut = off - pos
                r2 = copy.deepcopy(r)
                for t in X(r, './w:t')[1:]:
                    r.remove(t)
                for t in X(r2, './w:t')[1:]:
                    r2.remove(t)
                X(r, './w:t')[0].text = rt[:cut]
                X(r2, './w:t')[0].text = rt[cut:]
                for t in (X(r, './w:t')[0], X(r2, './w:t')[0]):
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                r.addnext(r2)
                break
            pos += len(rt)
    return X(p, './w:r')

def fix_tables(items, blocks, split_ok, report):
    for b in blocks:
        if b['type'] != 'tbl':
            continue
        tbl = items[b['tbl']].el
        key = b.get('key')
        tblPr = tbl.find(qn('w:tblPr'))
        # ширина таблицы — не шире текста
        if tblPr is not None:
            tw = tblPr.find(qn('w:tblW'))
            if tw is not None and tw.get(qn('w:type')) == 'dxa' and int(tw.get(qn('w:w'), '0')) > TEXT_W_MM * 56.7:
                tw.set(qn('w:type'), 'pct'); tw.set(qn('w:w'), '5000')
            ind = tblPr.find(qn('w:tblInd'))
            if ind is not None:
                tblPr.remove(ind)
        rows = X(tbl, './w:tr')
        small = len(rows) <= SMALL_TABLE_ROWS and key not in split_ok and not b.get('cont')
        for ri, tr in enumerate(rows):
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr'); tr.insert(0 if tr.find(qn('w:tblPrEx')) is None else 1, trPr)
            if trPr.find(qn('w:cantSplit')) is None:
                trPr.insert(0, OxmlElement('w:cantSplit'))
            for th in trPr.findall(qn('w:tblHeader')):
                trPr.remove(th)   # шапку при переносе повторяет vkr_layout.py вместе с «Продолжение таблицы»
            last = ri == len(rows) - 1
            kwn = (not last) if small else (ri < 1)
            if last and b['note'] is not None:
                kwn = True
            for p in X(tr, './/w:p'):
                al = X(p, './w:pPr/w:jc')
                al = al[0].get(qn('w:val')) if al else 'left'
                if al in ('both', 'distribute'):
                    al = 'left'
                fmt_par(p, align=al, first=0, kwn=kwn)
                fmt_runs(p, keep_size_le=SIZE)

def apply_keeps(items, blocks):
    """«Не отрывать от предыдущего текста»: перед каждой вставкой ≥2 строк текста на той же странице."""
    pos = {id(it.el): k for k, it in enumerate(items)}
    for b in blocks:
        s = block_start(b)
        k = s - 1
        lines = 0
        while k >= 0 and lines < 2:
            it = items[k]
            if it.kind == 'text' and not FIGCAP_RE.match(it.text) and not NOTE_RE.match(it.text):
                flag(it.el, 'keepNext', True)
                lines += est_lines(it.el)
                k -= 1
            else:
                break

def merge_continuations(doc, only_num=None):
    """Склеить таблицы, ранее разорванные с «Продолжение таблицы N» (перед новой вёрсткой)."""
    n = 0
    for p in list(doc.element.body.iterchildren(qn('w:p'))):
        m = CONT_RE.match(norm(ptext(p)))
        if not m or (only_num and m.group(2) != only_num):
            continue
        prev, nxt = p.getprevious(), p.getnext()
        if prev is None or nxt is None or tag(prev) != 'tbl' or tag(nxt) != 'tbl':
            continue
        # первая часть таблицы — ищем её шапку
        head = X(prev, './w:tr')[0]
        rows = X(nxt, './w:tr')
        def rt(tr):
            return [norm(ptext(tc)) for tc in X(tr, './w:tc')]
        # шапка самой первой части (prev может быть тоже продолжением — шапка у него та же)
        if rows and rt(rows[0]) == rt(head):
            nxt.remove(rows[0]); rows = rows[1:]
        for b in X(X(prev, './w:tr')[-1], './w:tc/w:tcPr/w:tcBorders/w:bottom[@w:val="nil"]'):
            b.getparent().remove(b)
        for tr in rows:
            prev.append(tr)
        p.getparent().remove(p); nxt.getparent().remove(nxt)
        n += 1
    return n

def process(doc, refs=None, start_page=None, fig_cap_indent=False, tbl_cap_indent=True, gap_lines=1,
            split_ok=(), move_after=None):
    report = {}
    lang = doc_lang(doc)
    report['lang'] = lang
    normalize_styles(doc)
    anchors_to_inline(doc, report)
    cap_images(doc, report)
    items, intro = analyze(doc)
    isolate_drawings(doc, items, report)
    items, intro = analyze(doc)
    remove_page_breaks(doc, items, report)
    remove_empty(items, report)
    items, intro = analyze(doc)
    blocks = build_blocks(items)
    # подписи: таблица — над таблицей, рисунок — под рисунком
    moved = 0
    for b in blocks:
        if b['type'] == 'tbl' and b['cap_below'] is not None:
            items[b['tbl']].el.addprevious(items[b['cap_below']].el); moved += 1
        if b['type'] == 'fig' and b['cap_above'] is not None:
            items[b['fig'][-1]].el.addnext(items[b['cap_above']].el); moved += 1
    report['captions_moved'] = moved
    items, intro = analyze(doc)
    blocks = build_blocks(items)
    # вставка описательного текста (refs.json)
    if refs:
        keyed = {}
        for b in blocks:
            kk = block_key(b, items)
            if kk:
                keyed[kk.lower()] = b
        unnamed_f = [b for b in blocks if b['type'] == 'formula' and not b['num']]
        ins = 0
        for r in refs:
            key = r['key'].strip().lower()
            b = keyed.get(key)
            if b is None and key.startswith('формула #'):
                n = int(key.split('#')[1]) - 1
                b = unnamed_f[n] if 0 <= n < len(unnamed_f) else None
            if b is None:
                report.setdefault('refs_not_found', []).append(r['key']); continue
            mode = r.get('mode', 'insert')
            target = items[block_start(b)].el
            if mode == 'append':
                prev = target.getprevious()
                if prev is not None and tag(prev) == 'p' and norm(ptext(prev)):
                    rr = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = ' ' + r['text'].strip()
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); rr.append(t); prev.append(rr)
                    ins += 1; continue
            target.addprevious(new_par(r['text'].strip()))
            ins += 1
        report['ref_texts_inserted'] = ins
        items, intro = analyze(doc)
        blocks = build_blocks(items)
    # перенос вставки на абзац ниже (по указанию vkr_layout.py)
    if move_after:
        mv = 0
        for key in move_after:
            for b in blocks:
                if (block_key(b, items) or '').lower() == key.lower():
                    e = block_end(b)
                    nxt = e + 1
                    if nxt < len(items) and items[nxt].kind == 'text' and not CONT_RE.match(items[nxt].text):
                        anchor = items[nxt].el
                        for idx in range(block_start(b), e + 1)[::-1]:
                            anchor.addnext(items[idx].el)
                        mv += 1
                    break
            items, intro = analyze(doc)
            blocks = build_blocks(items)
        report['blocks_moved'] = mv
    for b in blocks:
        b['key'] = block_key(b, items)
    # ---- форматирование абзацев -------------------------------------------
    gap = LINE_PT * gap_lines
    for k, it in enumerate(items):
        if it.kind == 'front' or it.kind == 'tbl':
            continue
        p = it.el
        if it.zone == 'toc' and it.kind != 'heading':
            fmt_runs(p, color_auto=False); continue
        if it.kind in ('heading', 'appx'):
            prev = items[k - 1] if k else None
            after_heading = prev is not None and prev.kind in ('heading', 'appx')
            t = it.text
            if it.level == 1 and t.isupper() and (STRUCT_RE.match(t) or APPX_RE.match(t)):
                ws = t.split()
                t = ' '.join([ws[0].capitalize()] + [w if (APPX_RE.match(t) and len(w) <= 3) else w.lower() for w in ws[1:]])
                set_text(p, t)
            if norm(ptext(p)).endswith('.') and not norm(ptext(p)).endswith('..'):
                set_text(p, norm(ptext(p))[:-1])
            if it.kind == 'appx':
                fmt_par(p, align='right', first=0, after=0, kwn=True, keep=True, pbb=True)
            else:
                pbb = it.level == 1 and not INTRO_RE.match(t)
                fmt_par(p, align='left', first=INDENT,
                        before=0 if (it.level == 1 or after_heading) else gap,
                        after=gap, kwn=True, keep=True, pbb=pbb)
            set_ppr(p, 'outlineLvl', {'val': min(it.level, 9) - 1}) if it.kind == 'heading' else None
            fmt_runs(p, bold=None)
            for u in X(p, './/w:rPr/w:u'):
                u.getparent().remove(u)
            # «Введение»: начало раздела уже даёт новую страницу
            continue
        if it.kind == 'fig':
            fmt_par(p, align='center', first=0, before=6, after=0, kwn=True, keep=True)
            continue
        if it.kind == 'formula':
            fmt_par(p, align='center', first=0, before=gap, after=gap, keep=True)
            fmt_formula(p, it.num)
            continue
        if it.kind == 'text':
            prev = items[k - 1] if k else None
            if prev is not None and prev.kind == 'appx':
                fmt_par(p, align='center', first=0, after=gap, kwn=True, keep=True)
                fmt_runs(p); continue
            fmt_par(p, align='both', first=INDENT)
            if X(p, './w:pPr/w:numPr'):
                set_ppr(p, 'ind', {'left': 0, 'firstLine': int(Emu(INDENT).twips), 'right': 0})
            fmt_runs(p)
    # ---- вставки ---------------------------------------------------------
    for b in blocks:
        if b['type'] == 'fig':
            for x in b['fig']:
                fmt_par(items[x].el, align='center', first=0, before=6, kwn=True, keep=True)
            for x in b['mids']:
                fmt_par(items[x].el, align='left', first=0, kwn=True)
                fmt_runs(items[x].el)
            if b['cap'] is not None:
                c = items[b['cap']].el
                fix_caption_text(c, FIGCAP_RE, report, b['key'])
                fmt_par(c, align='left', first=INDENT if fig_cap_indent else 0, after=gap, keep=True)
                fmt_runs(c, bold=False, italic=False)
            else:
                fmt_par(items[b['fig'][-1]].el, align='center', first=0, before=6, after=gap, keep=True)
        elif b['type'] == 'tbl':
            if b['cap'] is not None:
                c = items[b['cap']].el
                cm = CONT_RE.match(norm(ptext(c)))
                if cm:
                    want = cont_caption(cm.group(2), lang)
                    if norm(ptext(c)) != want:
                        set_text(c, want)
                    fmt_par(c, align='left', first=INDENT if tbl_cap_indent else 0, kwn=True, keep=True, pbb=True)
                else:
                    fix_caption_text(c, TBLCAP_RE, report, b['key'])
                    fmt_par(c, align='left', first=INDENT if tbl_cap_indent else 0, before=6, kwn=True, keep=True)
                fmt_runs(c, bold=False, italic=False)
            if b['note'] is not None:
                n_ = items[b['note']].el
                fmt_par(n_, align='left', first=INDENT if tbl_cap_indent else 0, after=gap)
                fmt_runs(n_)
        elif b['type'] == 'formula' and b['where'] is not None:
            flag(items[b['f']].el, 'keepNext', True)
            w = items[b['where']].el
            fmt_par(w, align='left', first=0)
            t = ptext(w)
            m = re.match(r'^(\s*)(где|мұндағы)\s*[:,]', t, re.I)
            if m:
                set_text(w, m.group(2) + t[m.end():] if t[m.end():].startswith(' ') else m.group(2) + ' ' + t[m.end():].lstrip())
            fmt_runs(w)
    fix_tables(items, blocks, set(split_ok), report)
    apply_keeps(items, blocks)
    page_setup(doc, items[intro].el if intro is not None else None, start_page, report)
    add_hyperlinks(doc, items, report)
    items, intro = analyze(doc)
    blocks = build_blocks(items)
    for b in blocks:
        b['key'] = block_key(b, items)
    check(items, blocks, intro, report)
    return report

def fmt_formula(p, num):
    """Формула по центру, номер (N) — у правого края через табуляцию."""
    for mp in X(p, './m:oMathPara'):
        for om in X(mp, './m:oMath'):
            mp.addprevious(om)
        mp.getparent().remove(mp)
    if not num:
        return
    set_ppr(p, 'jc', {'val': 'left'})
    tabs = set_ppr(p, 'tabs')
    for val, pos in (('center', int(TEXT_W_MM / 2 * 56.7)), ('right', int(TEXT_W_MM * 56.7))):
        t = OxmlElement('w:tab'); t.set(qn('w:val'), val); t.set(qn('w:pos'), str(pos)); tabs.append(t)
    for r in X(p, './w:r[w:tab]'):
        r.getparent().remove(r)
    first_math = X(p, './m:oMath|./w:r[w:object]')
    num_runs = [r for r in X(p, './w:r') if re.search(r'\(\s*' + re.escape(num) + r'\s*\)', ''.join(X(r, './w:t/text()')))]
    def tabrun():
        r = OxmlElement('w:r'); r.append(OxmlElement('w:tab')); return r
    if first_math:
        first_math[0].addprevious(tabrun())
    if num_runs:
        nr = num_runs[0]
        txt = ''.join(X(nr, './w:t/text()'))
        X(nr, './w:t')[0].text = txt.strip()
        nr.addprevious(tabrun())

def check(items, blocks, intro, report):
    issues = []
    if intro is None:
        issues.append({'type': 'no_intro', 'msg': 'Не найден заголовок «Введение»: нумерация страниц и границы основной части не определены'})
    seen = {'fig': set(), 'tbl': set(), 'formula': set()}
    order = {'fig': [], 'tbl': [], 'formula': []}
    bstarts = {block_start(b): b for b in blocks}
    textacc = []
    k = 0
    pos_text = []
    for k, it in enumerate(items):
        if it.kind == 'text' and it.zone not in ('toc', 'front', 'refs'):
            if not FIGCAP_RE.match(it.text) and not TBLCAP_RE.match(it.text) and not NOTE_RE.match(it.text):
                for typ in seen:
                    seen[typ] |= mentions(it.text, typ)
        if k in bstarts:
            b = bstarts[k]
            if b['type'] == 'tbl' and b.get('cont'):
                continue
            typ = b['type']
            num = None
            key = b['key']
            if key:
                num = key.split()[-1]
                order[typ].append(num)
            prev = items[k - 1] if k else None
            ctx = prev.text[-200:] if prev is not None and prev.kind == 'text' else ''
            cap_text = ''
            if typ == 'fig' and b['cap'] is not None:
                cap_text = items[b['cap']].text
            if typ == 'tbl' and b['cap'] is not None:
                cap_text = items[b['cap']].text
            label = key or ({'fig': 'Рисунок без подписи', 'tbl': 'Таблица без подписи', 'formula': 'Формула без номера'}[typ])
            if prev is None or prev.kind != 'text' or FIGCAP_RE.match(prev.text) or NOTE_RE.match(prev.text) or TBLCAP_RE.match(prev.text):
                what = {'heading': 'сразу после заголовка', 'appx': 'сразу после заголовка приложения'}.get(prev.kind if prev else '', 'сразу после другой вставки')
                issues.append({'type': 'no_text_before', 'key': label, 'caption': cap_text, 'msg': f'{label}: стоит {what} — нужен описательный текст с номером вставки', 'context': ctx})
            elif num and num not in seen[typ] and it.zone != 'appx':
                issues.append({'type': 'no_reference', 'key': label, 'caption': cap_text, 'msg': f'{label}: в тексте до вставки нет ссылки с её номером', 'context': ctx})
            if typ == 'fig' and b['cap'] is None:
                issues.append({'type': 'no_caption', 'key': label, 'msg': 'Рисунок без подписи «Рисунок N. Название»', 'context': ctx})
            if typ == 'tbl' and b['cap'] is None:
                issues.append({'type': 'no_caption', 'key': label, 'msg': 'Таблица без подписи «Таблица N. Название»', 'context': ctx})
            if typ == 'formula' and not b['num']:
                n_unnum = sum(1 for bb in blocks if bb['type'] == 'formula' and not bb['num'] and block_start(bb) <= k)
                issues.append({'type': 'formula_no_number', 'key': f'Формула #{n_unnum}', 'msg': 'Формула без номера (N) у правого края', 'context': ctx})
    for typ, name in (('fig', 'рисунков'), ('tbl', 'таблиц'), ('formula', 'формул')):
        nums = order[typ]
        flat = [n for n in nums if '.' not in n]
        if flat and flat != [str(i) for i in range(1, len(flat) + 1)]:
            issues.append({'type': 'numbering', 'msg': f'Нумерация {name} не сквозная/не по порядку: {", ".join(nums)}'})
    # приложения
    for it in items:
        if it.kind == 'appx':
            m = APPX_RE.match(it.text)
            if m and m.group(2) and not m.group(2).strip().isdigit():
                issues.append({'type': 'appendix_letter', 'msg': f'«{it.text}»: по п.10.23 приложения нумеруются арабскими цифрами'})
        if it.kind == 'heading' and it.text.isupper() and len(it.text) > 3 and not STRUCT_RE.match(it.text):
            issues.append({'type': 'heading_caps', 'msg': f'Заголовок прописными: «{it.text[:80]}» — нужен с прописной буквы'})
    # литература и ссылки
    refs_items = [it for it in items if it.zone == 'refs' and it.kind == 'text']
    n_refs = len(refs_items)
    cited = []
    for it in items:
        if it.kind == 'text' and it.zone in ('body',):
            for m in re.finditer(r'\[([\d\s,;–\-—сСpP\.]+)\]', it.text):
                for part in re.split(r';', m.group(1)):
                    first = re.match(r'\s*(\d+)\s*(?:[–\-—]\s*(\d+))?', part)
                    if first:
                        a = int(first.group(1)); b_ = int(first.group(2)) if first.group(2) and ',' not in part else a
                        for x in range(a, b_ + 1):
                            if x not in cited:
                                cited.append(x)
    if n_refs:
        if n_refs < 30:
            issues.append({'type': 'refs_count', 'msg': f'В списке литературы {n_refs} источников — нужно не менее 30 (п.9.15)'})
        missing = [i for i in range(1, n_refs + 1) if i not in cited]
        extra = [c for c in cited if c > n_refs]
        if missing:
            issues.append({'type': 'refs_uncited', 'msg': f'Источники без ссылок в тексте: {missing[:40]}'})
        if extra:
            issues.append({'type': 'refs_extra', 'msg': f'Ссылки на номера, которых нет в списке: {extra}'})
        if cited and cited != sorted(cited):
            issues.append({'type': 'refs_order', 'msg': 'Источники в списке не в порядке первого упоминания (п.10.22): порядок ссылок ' + ', '.join(map(str, cited[:30]))})
        no_url = [i + 1 for i, it in enumerate(refs_items) if not re.search(r'https?://|doi', it.text, re.I)]
        if no_url:
            issues.append({'type': 'refs_no_url', 'msg': f'Источники без URL/DOI (п.10.21): № {no_url[:40]}'})
    # объём по словам
    def words(zone_filter):
        return sum(len(it.text.split()) for it in items if it.kind == 'text' and zone_filter(it))
    report['words_body'] = words(lambda it: it.zone == 'body')
    report['issues'] = issues
    report['blocks'] = [b['key'] or b['type'] for b in blocks]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('inp'); ap.add_argument('out')
    ap.add_argument('--refs'); ap.add_argument('--start-page', type=int)
    ap.add_argument('--fig-caption-indent', type=int, default=0)
    ap.add_argument('--tbl-caption-indent', type=int, default=1)
    ap.add_argument('--heading-gap', type=float, default=1)
    ap.add_argument('--split-ok', default='', help='ключи таблиц, которым разрешён разрыв, через ;')
    ap.add_argument('--move-after', default='', help='ключи вставок, перенести на абзац ниже, через ;')
    ap.add_argument('--report')
    a = ap.parse_args()
    doc = Document(a.inp)
    refs = json.load(open(a.refs, encoding='utf-8')) if a.refs else None
    rep = process(doc, refs=refs, start_page=a.start_page, fig_cap_indent=bool(a.fig_caption_indent),
                  tbl_cap_indent=bool(a.tbl_caption_indent), gap_lines=a.heading_gap,
                  split_ok=[s for s in a.split_ok.split(';') if s],
                  move_after=[s for s in a.move_after.split(';') if s])
    doc.save(a.out)
    js = json.dumps(rep, ensure_ascii=False, indent=1)
    if a.report:
        open(a.report, 'w', encoding='utf-8').write(js)
    print(js)

if __name__ == '__main__':
    main()
```

### vkr/vkr_layout.py

```python
#!/usr/bin/env python3
"""Постраничная вёрстка дипломной работы после vkr_format.py.

python3 vkr_layout.py formatted.docx final.docx [--kind work|project] [--max-iter 40]
        [--report layout.json] [--pdf final.pdf] [+ любые флаги vkr_format.py через --fmt "..."]

Цикл: рендер в PDF (LibreOffice) → поиск первой проблемной страницы → исправление → снова рендер.
Проблемы: страница начинается с рисунка / таблицы / формулы / «где» / подписи;
пустота внизу страницы > 3 строк, когда следующая страница не начинается с нового раздела;
таблица перешла на новую страницу без «Продолжение таблицы N» (каз.: «N-кестенің жалғасы»).
"""
import argparse, copy, json, os, re, shlex, shutil, subprocess, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pymupdf
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import vkr_format as F

MM = 72 / 25.4
TOP = 20 * MM
BOTTOM = 841.89 - 25 * MM
LINE = 16.1
VOID_LINES = 3

def word_like_copy(docx, outdir):
    """LibreOffice переносит абзац с «не отрывать от следующего» ЦЕЛИКОМ, а Word — только последние строки
    (с запретом висячих строк — 2). Для честной имитации Word в технической копии последние ~2 строки
    такого абзаца выделяются в отдельный абзац. Сам документ не меняется."""
    doc = Document(docx)
    for p in list(doc.element.body.iterchildren(qn('w:p'))):
        kn = p.find(qn('w:pPr') + '/' + qn('w:keepNext'))
        if kn is None or kn.get(qn('w:val')) in ('0', 'false'):
            continue
        txt = F.ptext(p)
        if len(txt) < 4 * F.CHARS_PER_LINE or F.X(p, './/w:drawing|.//m:oMath|.//w:fldChar'):
            continue
        cut = txt.rfind(' ', 0, len(txt) - int(1.6 * F.CHARS_PER_LINE))
        if cut <= 0:
            continue
        runs = F._split_runs_at(p, [cut + 1])
        tail = F.new_par()
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            tail.append(copy.deepcopy(pPr))
            F.set_ppr(tail, 'ind', {'left': 0, 'firstLine': 0, 'right': 0})
        pos = 0
        for r in runs:
            rt = ''.join(F.X(r, './w:t/text()'))
            if pos >= cut + 1:
                tail.append(r)
            pos += len(rt)
        F.flag(p, 'keepNext', False)
        F.set_ppr(p, 'jc', {'val': 'both'})
        p.addnext(tail)
    sim = os.path.join(outdir, 'sim.docx')
    doc.save(sim)
    return sim

def render(docx, outdir):
    docx = word_like_copy(docx, outdir)
    prof = tempfile.mkdtemp(prefix='lo_')
    subprocess.run(['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'pdf',
                    '--outdir', outdir, docx], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    shutil.rmtree(prof, ignore_errors=True)
    return os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + '.pdf')

MATH_FONTS = ('opensymbol', 'cambria math', 'stixmath', 'symbol')

def page_model(page):
    lines = []
    d = page.get_text('dict')
    for b in d['blocks']:
        if b['type'] != 0:
            continue
        for l in b['lines']:
            txt = ''.join(s['text'] for s in l['spans']).strip()
            if not txt:
                continue
            x0, y0, x1, y1 = l['bbox']
            if y0 > BOTTOM + 4 or y1 < TOP - 8:
                continue        # колонтитулы, номер страницы
            fonts = {s['font'].lower() for s in l['spans']}
            lines.append(dict(y0=y0, y1=y1, x0=x0, x1=x1, text=txt, math=any(any(m in f for m in MATH_FONTS) for f in fonts)))
    # объединить фрагменты одной строки
    lines.sort(key=lambda l: (round(l['y0']), l['x0']))
    merged = []
    for l in lines:
        if merged and abs(merged[-1]['y0'] - l['y0']) < 3:
            m = merged[-1]; m['text'] += ' ' + l['text']; m['x1'] = max(m['x1'], l['x1']); m['y1'] = max(m['y1'], l['y1'])
            m['math'] = m['math'] or l['math']
        else:
            merged.append(dict(l))
    graphics = []
    for im in page.get_image_info():
        x0, y0, x1, y1 = im['bbox']
        if y1 - y0 > 8 and y0 < BOTTOM + 4:
            graphics.append(dict(y0=y0, y1=y1, x0=x0, x1=x1, kind='image'))
    for dr in page.get_drawings():
        r = dr['rect']
        if r.y0 > BOTTOM + 4 or r.y1 < TOP - 8:
            continue
        if r.width < 3 and r.height < 3:
            continue
        graphics.append(dict(y0=r.y0, y1=r.y1, x0=r.x0, x1=r.x1, kind='draw', w=r.width, h=r.height))
    return merged, graphics

def is_grid(graphics, y_from):
    """Есть ли вверху табличная сетка (много тонких горизонтальных/вертикальных линий)."""
    g = [x for x in graphics if x['kind'] == 'draw' and x['y0'] < y_from + 40]
    thin = [x for x in g if x['h'] < 2.5 or x['w'] < 2.5]
    return len(thin) >= 3

def classify(pages):
    out = []
    for pi, (lines, graphics) in enumerate(pages):
        info = dict(page=pi + 1, first='', start=None, bottom=TOP, n_lines=len(lines))
        tops = []
        if lines:
            tops.append(('text', lines[0]['y0']))
        if graphics:
            gtop = min(g['y0'] for g in graphics)
            tops.append(('graphic', gtop))
        if not tops:
            info['start'] = 'blank'; out.append(info); continue
        bottom = max([l['y1'] for l in lines] + [g['y1'] for g in graphics])
        info['bottom'] = bottom
        info['first'] = lines[0]['text'] if lines else ''
        first_kind = min(tops, key=lambda t: t[1])
        if first_kind[0] == 'graphic' and (not lines or first_kind[1] < lines[0]['y0'] - 1):
            gtop = first_kind[1]
            if is_grid(graphics, gtop) and lines and any(g['kind'] == 'draw' for g in graphics if g['y0'] < lines[0]['y0']):
                info['start'] = 'table'
            else:
                info['start'] = 'figure'
        else:
            t = lines[0]['text']
            if F.CONT_RE.match(t):
                info['start'] = 'cont'
            elif F.TBLCAP_RE.match(t) and len(t) < 250:
                info['start'] = 'tblcap'
            elif F.FIGCAP_RE.match(t) and len(t) < 250 and (len(lines) < 2 or not re.match(r'^[а-яё]', lines[1]['text'])):
                info['start'] = 'figcap'
            elif F.WHERE_RE.match(t) and not re.match(r'^где-', t, re.I):
                info['start'] = 'where'
            elif lines[0]['math'] or (re.search(r'\(\d+(\.\d+)?\)\s*$', t) and len(t) < 90 and (lines[0]['x0'] > 150)):
                info['start'] = 'formula'
            else:
                info['start'] = 'text'
                # сколько строк текста до первой графики
                gs = [g['y0'] for g in graphics if g['y0'] > lines[0]['y0'] and (g['kind'] == 'image' or g.get('h', 0) > 8)]
                if gs:
                    gy = min(gs)
                    info['lines_before_graphic'] = sum(1 for l in lines if l['y1'] <= gy + 1)
        info['lines'] = [l['text'] for l in lines[:6]]
        info['all_lines'] = lines
        out.append(info)
    return out

HEAD1 = re.compile(r'^(\d{1,2}\s+\S|' + F.STRUCT_RE.pattern[1:-1] + r'|приложение\b)', re.I)

def find_problem(cls, heading_pages, intro_page, skip=()):
    """Первая проблема сверху вниз (страницы до «Введения» не трогаем)."""
    for i, c in enumerate(cls):
        if c['page'] < intro_page:
            continue
        s = c['start']
        sig = lambda t: (t, c['first'][:40], c['n_lines'] and c['all_lines'][-1]['text'][:40])
        if s in ('figure', 'table', 'tblcap', 'figcap', 'where', 'formula') and c['page'] not in heading_pages \
                and sig('starts_with_' + s) not in skip:
            return dict(page=c['page'], type='starts_with_' + s, sig=sig('starts_with_' + s))
        if c.get('lines_before_graphic') == 1 and sig('one_line') not in skip:
            return dict(page=c['page'], type='one_line_before_graphic', sig=sig('one_line'))
        if i + 1 < len(cls) and cls[i + 1]['page'] not in heading_pages and cls[i + 1]['start'] != 'blank':
            void = BOTTOM - c['bottom']
            if void > VOID_LINES * LINE and sig('void') not in skip:
                return dict(page=c['page'], type='void', void=void, sig=sig('void'))
    return None

def heading_pages_of(cls, doc_headings):
    """Страницы, начинающиеся с заголовка 1-го уровня (там пустота на предыдущей странице допустима)."""
    pages = set()
    for c in cls:
        if not c.get('lines'):
            continue
        f = F.norm(c['lines'][0]).lower()
        for h in doc_headings:
            if f.startswith(h.lower()[:25]):
                pages.add(c['page']); break
    return pages

def doc_structure(path):
    doc = Document(path)
    items, intro = F.analyze(doc)
    blocks = F.build_blocks(items)
    for b in blocks:
        b['key'] = F.block_key(b, items)
    h1 = [it.text for it in items if it.kind in ('heading', 'appx') and it.level == 1]
    return doc, items, blocks, h1

def block_on_page(c, blocks, items):
    """Какая вставка на странице c (первая), по подписи."""
    keys = []
    for l in c.get('all_lines', []):
        t = l['text']
        m = F.FIGCAP_RE.match(t)
        if m and len(t) < 250:
            keys.append(('Рисунок ' + m.group(2), l['y0']))
        m = F.TBLCAP_RE.match(t)
        if m and len(t) < 250:
            keys.append(('Таблица ' + m.group(2), l['y0']))
        m = re.search(r'\((\d+(?:\.\d+)?)\)\s*$', t)
        if m and (l['math'] or l['x0'] > 150) and len(t) < 90:
            keys.append(('Формула ' + m.group(1), l['y0']))
    known = {b['key'] for b in blocks if b['key']}
    keys = [k for k in keys if k[0] in known]
    return keys

def scale_block_image(doc, items, blocks, key, factor):
    for b in blocks:
        if b['key'] == key and b['type'] == 'fig':
            done = False
            for x in b['fig']:
                for inl in F.X(items[x].el, './/wp:inline'):
                    F.scale_inline(inl, factor); done = True
            return done
    return False

TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
              'textDirection', 'tcFitText', 'vAlign', 'hideMark']

def _tcpr_child(tcPr, name):
    el = tcPr.find(qn('w:' + name))
    if el is None:
        el = OxmlElement('w:' + name)
        idx = TCPR_ORDER.index(name); pos = 0
        for i, ch in enumerate(tcPr):
            n = F.tag(ch)
            if n in TCPR_ORDER and TCPR_ORDER.index(n) < idx:
                pos = i + 1
        tcPr.insert(pos, el)
    return el

def split_table(doc, items, blocks, num, first_texts):
    """Разрыв таблицы num перед строкой, с которой началась новая страница."""
    parts = [b for b in blocks if b['type'] == 'tbl' and b['key'] in ('Таблица ' + num, 'Продолжение Таблица ' + num)]
    cand = [F.norm(t).lower() for t in first_texts if F.norm(t)]
    def row_matches(cells, line):
        toks = line.split()
        if not toks or toks[0] != cells[0].split()[0]:
            return False
        pos = 0
        for c in cells[:5]:
            k = line.find(c[:min(8, len(c))], pos)
            if k < 0:
                return False
            pos = k + 1
        return True
    for b in reversed(parts):
        tbl = items[b['tbl']].el
        rows = F.X(tbl, './w:tr')
        for ri, tr in enumerate(rows):
            if ri == 0:
                continue
            cells = [F.norm(F.ptext(tc)).lower() for tc in F.X(tr, './w:tc')]
            cells = [c for c in cells if c]
            if not cells:
                continue
            if any(row_matches(cells, t) for t in cand):
                new = copy.deepcopy(tbl)
                for r in F.X(new, './w:tr')[:ri]:
                    new.remove(r)
                # повторить шапку
                if ri > 1 and not b.get('cont'):
                    hdr = copy.deepcopy(rows[0])
                    F.X(new, './w:tr')[0].addprevious(hdr)
                elif b.get('cont'):
                    hdr = copy.deepcopy(rows[0])
                    F.X(new, './w:tr')[0].addprevious(hdr)
                for r in rows[ri:]:
                    tbl.remove(r)
                # вертикальные объединения в первой перенесённой строке
                body_rows = F.X(new, './w:tr')
                first_data = body_rows[1] if len(body_rows) > 1 and (ri > 1 or b.get('cont')) else body_rows[0]
                for vm in F.X(first_data, './w:tc/w:tcPr/w:vMerge'):
                    vm.set(qn('w:val'), 'restart')
                # нижнюю черту первой части не проводят
                for tc in F.X(F.X(tbl, './w:tr')[-1], './w:tc'):
                    tcPr = tc.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = OxmlElement('w:tcPr'); tc.insert(0, tcPr)
                    brd = _tcpr_child(tcPr, 'tcBorders')
                    bt = brd.find(qn('w:bottom'))
                    if bt is None:
                        bt = OxmlElement('w:bottom'); brd.append(bt)
                    bt.set(qn('w:val'), 'nil')
                cap = F.new_par(F.cont_caption(num, F.doc_lang(doc)))
                tbl.addnext(cap)
                cap.addnext(new)
                return True
    return False

def run(args):
    work = tempfile.mkdtemp(prefix='vkr_')
    cur = os.path.join(work, 'cur.docx')
    shutil.copy(args.inp, cur)
    d0 = Document(cur)
    if F.merge_continuations(d0):
        d0.save(cur)
    fmt_extra = shlex.split(args.fmt or '')
    split_ok, moved, scaled, log = set(), {}, {}, []
    seen = {}
    skip = set()
    final_cls = None
    for it in range(args.max_iter):
        pdf = render(cur, work)
        pdoc = pymupdf.open(pdf)
        pages = [page_model(p) for p in pdoc]
        cls = classify(pages)
        doc, items, blocks, h1 = doc_structure(cur)
        hp = heading_pages_of(cls, h1)
        intro_page = min([c['page'] for c in cls if c.get('lines') and F.INTRO_RE.match(F.norm(c['lines'][0]))] or [1])
        prob = find_problem(cls, hp, intro_page, skip)
        final_cls = (cls, hp, intro_page, h1)
        if prob is None:
            break
        sig = (prob['page'], prob['type'])
        seen[sig] = seen.get(sig, 0) + 1
        if seen[sig] > 4:
            log.append(dict(prob, action='не удалось исправить автоматически'))
            skip.add(prob['sig']); continue
        c = cls[prob['page'] - 1]
        action = None
        if prob['type'] == 'void' and prob['page'] < len(cls) and cls[prob['page']]['start'] == 'cont':
            num = F.CONT_RE.match(cls[prob['page']]['first']).group(2)
            if F.merge_continuations(doc, num):
                action = f'таблица {num}: разрыв устарел — таблица склеена для нового разрыва'
                log.append(dict(prob, action=action)); doc.save(cur)
                subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), 'vkr_format.py'), cur, cur,
                                '--split-ok', ';'.join(split_ok)] + fmt_extra, check=True, stdout=subprocess.DEVNULL)
                continue
        if prob['type'] == 'starts_with_table':
            # таблица перешла на новую страницу без «Продолжение таблицы»
            num = None
            for pc in reversed(cls[:prob['page'] - 1]):
                for l in reversed(pc.get('all_lines', [])):
                    m = F.TBLCAP_RE.match(l['text']) or F.CONT_RE.match(l['text'])
                    if m:
                        num = m.group(2); break
                if num:
                    break
            if num:
                lines = c['all_lines']
                top = [l['text'] for l in lines if l['y0'] < lines[0]['y0'] + 3] + [l['text'] for l in lines[:3]]
                if split_table(doc, items, blocks, num, top):
                    action = f'таблица {num}: разрыв с подписью «{F.cont_caption(num, F.doc_lang(doc))}»'
                    split_ok.add('Таблица ' + num)
        elif prob['type'] in ('void', 'starts_with_figure', 'starts_with_figcap', 'starts_with_tblcap',
                              'starts_with_formula', 'starts_with_where', 'one_line_before_graphic'):
            tgt_page = prob['page'] + 1 if prob['type'] == 'void' else prob['page']
            keys = block_on_page(cls[tgt_page - 1], blocks, items) if tgt_page <= len(cls) else []
            if prob['type'] != 'void' and prob['page'] > 1 and not keys:
                keys = block_on_page(cls[prob['page'] - 2], blocks, items)[-1:]
            if keys:
                key = keys[0][0]
                b = next(b for b in blocks if b['key'] == key)
                nxt = F.block_end(b) + 1
                can_move = b['type'] == 'fig' and nxt < len(items) and items[nxt].kind == 'text' and not F.CONT_RE.match(items[nxt].text) \
                    and moved.get(key, 0) < 3 and items[nxt].zone == items[F.block_start(b)].zone
                if prob['type'] == 'void' and b['type'] == 'tbl' and key not in split_ok and 3 < len(F.X(items[b['tbl']].el, './w:tr')) <= F.SMALL_TABLE_ROWS:
                    split_ok.add(key); action = f'{key}: разрешён разрыв таблицы (пустота {prob["void"]/LINE:.0f} строк)'
                elif prob['type'] == 'void' and can_move:
                    moved[key] = moved.get(key, 0) + 1
                    fmt_extra_once = ['--move-after', key]
                    action = f'{key}: перенесена на абзац ниже (пустота {prob["void"]/LINE:.0f} строк на стр. {prob["page"]})'
                    doc.save(cur)
                    subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), 'vkr_format.py'), cur, cur,
                                    '--split-ok', ';'.join(split_ok)] + fmt_extra + fmt_extra_once,
                                   check=True, stdout=subprocess.DEVNULL)
                    log.append(dict(prob, action=action))
                    continue
                elif b['type'] == 'fig' and scaled.get(key, 1.0) > 0.62:
                    k = 0.85
                    if prob['type'] == 'void':
                        imgs = [g for g in pages[tgt_page - 1][1] if g['kind'] == 'image']
                        if imgs:
                            h = imgs[0]['y1'] - imgs[0]['y0']
                            need = (imgs[0]['y1'] - TOP) + 3 * LINE - prob['void']
                            k = max(0.6 / scaled.get(key, 1.0), min(0.95, (h - need) / h)) if h > need else 0.85
                    if scale_block_image(doc, items, blocks, key, k):
                        scaled[key] = scaled.get(key, 1.0) * k
                        action = f'{key}: уменьшен до {scaled[key]*100:.0f}%'
        if action is None:
            log.append(dict(prob, action='не удалось исправить автоматически'))
            skip.add(prob['sig']); continue
        log.append(dict(prob, action=action))
        doc.save(cur)
        subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), 'vkr_format.py'), cur, cur,
                        '--split-ok', ';'.join(split_ok)] + fmt_extra, check=True, stdout=subprocess.DEVNULL)
    # финальный отчёт
    pdf = render(cur, work)
    pdoc = pymupdf.open(pdf)
    pages = [page_model(p) for p in pdoc]
    cls = classify(pages)
    doc, items, blocks, h1 = doc_structure(cur)
    hp = heading_pages_of(cls, h1)
    intro_page = min([c['page'] for c in cls if c.get('lines') and F.INTRO_RE.match(F.norm(c['lines'][0]))] or [1])
    remaining, minor = [], []
    for i, c in enumerate(cls):
        if c['page'] < intro_page:
            continue
        if c['start'] in ('figure', 'table', 'tblcap', 'figcap', 'where', 'formula') and c['page'] not in hp:
            remaining.append(f'стр. {c["page"]}: начинается с «{c["start"]}» — {c["first"][:60]}')
        if c.get('lines_before_graphic') == 1:
            remaining.append(f'стр. {c["page"]}: перед вставкой только 1 строка текста')
        if i + 1 < len(cls) and cls[i + 1]['page'] not in hp and cls[i + 1]['start'] != 'blank':
            void = BOTTOM - c['bottom']
            if void > VOID_LINES * LINE:
                nxt_start = cls[i + 1]['start']
                if void <= 6.5 * LINE:
                    minor.append(f'стр. {c["page"]}: пустота ≈{void/LINE:.0f} строк — неразрывный блок «текст + вставка» не помещается')
                else:
                    remaining.append(f'стр. {c["page"]}: пустота внизу ≈{void/LINE:.0f} строк')
    # объём по страницам
    starts = []
    for h in h1:
        for c in cls:
            if c.get('lines') and F.norm(c['lines'][0]).lower().startswith(h.lower()[:25]):
                starts.append((h, c['page'])); break
    vol = {}
    for k, (h, p) in enumerate(starts):
        end = starts[k + 1][1] - 1 if k + 1 < len(starts) else len(cls)
        vol[h] = end - p + 1
    intro_p = next((p for h, p in starts if F.INTRO_RE.match(h)), None)
    refs_p = next((p for h, p in starts if F.REFS_RE.match(h)), None)
    total = ((refs_p or len(cls) + 1) - intro_p) if intro_p else None
    vol_issues = []
    minimum = 50 if args.kind == 'work' else 40
    if total is not None and total < minimum:
        vol_issues.append(f'Объём от введения до заключения включительно: {total} стр. — нужно не менее {minimum} (п.9.14)')
    for h, n in vol.items():
        if F.INTRO_RE.match(h) and not 1 <= n <= 3:
            vol_issues.append(f'Введение: {n} стр. — рекомендуется 1–3 (п.9.9)')
        if re.match(r'^(заключение|қорытынды|conclusion)$', h, re.I) and not 1 <= n <= 3:
            vol_issues.append(f'Заключение: {n} стр. — рекомендуется 1–3')
    shutil.copy(cur, args.out)
    if args.pdf:
        shutil.copy(pdf, args.pdf)
    for a_ in log: a_.pop('sig', None)
    rep = dict(pages=len(cls), iterations=len(log), actions=log, remaining=remaining, minor_voids=minor, volume_pages=vol,
               main_pages=total, volume_issues=vol_issues)
    js = json.dumps(rep, ensure_ascii=False, indent=1)
    if args.report:
        open(args.report, 'w', encoding='utf-8').write(js)
    print(js)

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('inp'); ap.add_argument('out')
    ap.add_argument('--kind', default='work', choices=['work', 'project'])
    ap.add_argument('--max-iter', type=int, default=40)
    ap.add_argument('--report'); ap.add_argument('--pdf')
    ap.add_argument('--fmt', default='', help='доп. флаги для vkr_format.py, например "--start-page 9"')
    run(ap.parse_args())
```