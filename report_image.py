"""Генератор PNG-таблиц с рейтингом магазинов (худший → лучший).
Покраска по-ячеечно: каждая метрика красится по своему значению vs норма.
"""
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams['font.family'] = 'DejaVu Sans'

HDR_BG = '#2C3E50'
HDR_FG = 'white'
RED    = '#FFAAAA'   # значение плохое
YELLOW = '#FFF5AA'   # близко к норме / погранично
GREEN  = '#AAFFAA'   # норма выполнена
WHITE  = '#FFFFFF'   # нейтральная ячейка (нет нормы)


# ── Функции оценки цвета ──────────────────────────────────────────────────────

def _c_plan(v):
    """ПЛАН%: <50 → красный, 50-70 → жёлтый, 70-85 → светло-жёлтый, ≥85 → зелёный."""
    if v < 50:  return RED
    if v < 70:  return '#FFCC88'   # оранжевый
    if v < 85:  return YELLOW
    return GREEN

def _c_norm(v, norm, bad=0.8):
    """Сравнение с нормой: <80% нормы → красный, 80-100% → жёлтый, ≥100% → зелёный."""
    if norm <= 0:  return WHITE
    if v < norm * bad:  return RED
    if v < norm:        return YELLOW
    return GREEN

def _c_delta(v, bad=-15):
    """Динамика (%, может быть отрицательной): <bad → красный, отриц. → жёлтый, ≥0 → зелёный."""
    if v <= bad:  return RED
    if v < 0:     return YELLOW
    return GREEN


# ── Рендер таблицы ────────────────────────────────────────────────────────────

def _draw(title, col_labels, rows, cell_colors, date_str, time_str):
    """
    cell_colors — 2D список [строка][колонка] цветов.
    Первый столбец (Магазин) всегда WHITE.
    """
    n_rows = len(rows)
    n_cols = len(col_labels)
    fig_w  = max(9, n_cols * 1.75)
    fig_h  = max(5, n_rows * 0.42 + 1.8)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis('off')

    tbl = ax.table(
        cellText=rows,
        colLabels=col_labels,
        cellColours=cell_colors,
        loc='center',
        cellLoc='center',
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8.5)
    tbl.scale(1, 1.5)

    for j in range(n_cols):
        cell = tbl[(0, j)]
        cell.set_facecolor(HDR_BG)
        cell.set_text_props(color=HDR_FG, fontweight='bold')

    ax.set_title(
        f'{title}\n{date_str}  |  {time_str}',
        fontsize=10, fontweight='bold', pad=14, color='#1A252F'
    )

    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150, facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ── Форматирование значений ───────────────────────────────────────────────────

def _p(v):   return f'{v:.1f}%'
def _rp(v):  return f'{v:+.1f}%'
def _r(v):   return f'{int(v):,}'.replace(',', '\u202f')
def _pr(v):  return f'{v:.2f}'


# ── Генерация таблиц ─────────────────────────────────────────────────────────

def generate_images(stores, norms, date_str, time_str):
    """
    Возвращает список (label, png_bytes).
    stores  — список словарей из bot_analyzer.load()
    norms   — словарь норм из bot_analyzer.load()
    """
    images = []

    # ── Таблица 1: Основные — ПЛАН, КОП, ТО к нед./вчера ─────────────────────
    s1 = sorted(stores, key=lambda x: x['plan'])
    rows1, colors1 = [], []
    for st in s1:
        rows1.append([st['name'], _p(st['plan']), _p(st['kop']),
                      _rp(st['to_ned']), _rp(st['to_vch']), _pr(st['pvch'])])
        colors1.append([
            WHITE,
            _c_plan(st['plan']),
            _c_norm(st['kop'],    norms['kop']),
            _c_delta(st['to_ned'], bad=-20),
            _c_delta(st['to_vch'], bad=-15),
            WHITE,
        ])
    images.append(('main', _draw(
        'Основные — ПЛАН, КОП, ТО к нед./вчера',
        ['Магазин', 'ПЛАН%', 'КОП%', 'ТО/нед', 'ТО/вчера', 'ПвЧ'],
        rows1, colors1, date_str, time_str
    )))

    # ── Таблица 2: Допродажи ──────────────────────────────────────────────────
    def upsell(st):
        return (st['kosm'] + st['steli'] + st['yui']) / 3

    s2 = sorted(stores, key=upsell)
    rows2, colors2 = [], []
    for st in s2:
        rows2.append([st['name'], _p(st['kosm']), _p(st['steli']),
                      _p(st['yui']), _p(st['sreb']), _p(st['zol'])])
        colors2.append([
            WHITE,
            _c_norm(st['kosm'],  norms['kosm']),
            _c_norm(st['steli'], norms['steli']),
            _c_norm(st['yui'],   norms['yui']),
            _c_norm(st['sreb'],  norms['sreb']),
            _c_norm(st['zol'],   norms['zol']),
        ])
    images.append(('upsell', _draw(
        'Допродажи — Косм., Стельки, ЮИ, Серебро, Золото',
        ['Магазин', 'Косм%', 'Стельки%', 'ЮИ%', 'Серебро%', 'Золото%'],
        rows2, colors2, date_str, time_str
    )))

    # ── Таблица 3: Услуги и чек ───────────────────────────────────────────────
    s3 = sorted(stores, key=lambda x: (x['rass'] + x['sbp']) / 2)
    rows3, colors3 = [], []
    for st in s3:
        rows3.append([st['name'], _p(st['rass']), _p(st['sbp']),
                      _p(st['kozha']), _r(st['sch_ob']), _r(st['sch'])])
        colors3.append([
            WHITE,
            _c_norm(st['rass'],   norms['rass']),
            _c_norm(st['sbp'],    norms['sbp']),
            WHITE,   # Кожа — нет нормы, информационно
            _c_norm(st['sch_ob'], norms['sch_ob']),
            WHITE,
        ])
    images.append(('services', _draw(
        'Услуги и чек — Рассрочка, СБП, Кожа, СЧ обуви',
        ['Магазин', 'Рассрочка%', 'СБП%', 'Кожа%', 'СЧ обувь', 'СЧ'],
        rows3, colors3, date_str, time_str
    )))

    return images
