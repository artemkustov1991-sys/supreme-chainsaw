"""Генератор PNG-таблиц с рейтингом магазинов (худший → лучший)."""
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams['font.family'] = 'DejaVu Sans'

HDR_BG  = '#2C3E50'
HDR_FG  = 'white'
RED     = '#FFAAAA'
GREEN   = '#AAFFAA'
STRIPE1 = '#F2F2F2'
STRIPE2 = '#FFFFFF'


def _row_colors(n, bad=3, good=3):
    out = []
    for i in range(n):
        if i < bad:
            out.append(RED)
        elif i >= n - good:
            out.append(GREEN)
        else:
            out.append(STRIPE1 if i % 2 == 0 else STRIPE2)
    return out


def _draw(title, col_labels, rows, row_colors, date_str, time_str):
    n_rows = len(rows)
    n_cols = len(col_labels)
    fig_w  = max(9, n_cols * 1.7)
    fig_h  = max(5, n_rows * 0.42 + 1.8)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis('off')

    cell_colors = [c for c in ([c] * n_cols for c in row_colors)]

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


def _p(v):   return f'{v:.1f}%'
def _rp(v):  return f'{v:+.1f}%'
def _r(v):   return f'{int(v):,}'.replace(',', '\u202f')
def _pr(v):  return f'{v:.2f}'


def generate_images(stores, norms, date_str, time_str):
    """
    Возвращает список (label, png_bytes).
    stores  — список словарей из bot_analyzer.load()
    norms   — словарь норм из bot_analyzer.load()
    """
    images = []

    # ── Таблица 1: Основные — ПЛАН, КОП, ТО, ПвЧ ────────────────────────────
    s1 = sorted(stores, key=lambda x: x['plan'])
    rows1 = [[st['name'], _p(st['plan']), _p(st['kop']),
              _rp(st['to_ned']), _rp(st['to_vch']), _pr(st['pvch'])]
             for st in s1]
    img1 = _draw(
        'Основные — ПЛАН, КОП, ТО к нед./вчера',
        ['Магазин', 'ПЛАН%', 'КОП%', 'ТО/нед', 'ТО/вчера', 'ПвЧ'],
        rows1, _row_colors(len(s1)), date_str, time_str
    )
    images.append(('main', img1))

    # ── Таблица 2: Допродажи ─────────────────────────────────────────────────
    def upsell(st):
        return (st['kosm'] + st['steli'] + st['yui']) / 3

    s2 = sorted(stores, key=upsell)
    rows2 = [[st['name'], _p(st['kosm']), _p(st['steli']),
              _p(st['yui']), _p(st['sreb']), _p(st['zol'])]
             for st in s2]
    img2 = _draw(
        'Допродажи — Косм., Стельки, ЮИ, Серебро, Золото',
        ['Магазин', 'Косм%', 'Стельки%', 'ЮИ%', 'Серебро%', 'Золото%'],
        rows2, _row_colors(len(s2)), date_str, time_str
    )
    images.append(('upsell', img2))

    # ── Таблица 3: Услуги и чек ──────────────────────────────────────────────
    s3 = sorted(stores, key=lambda x: (x['rass'] + x['sbp']) / 2)
    rows3 = [[st['name'], _p(st['rass']), _p(st['sbp']),
              _p(st['kozha']), _r(st['sch_ob']), _r(st['sch'])]
             for st in s3]
    img3 = _draw(
        'Услуги и чек — Рассрочка, СБП, Кожа, СЧ обуви',
        ['Магазин', 'Рассрочка%', 'СБП%', 'Кожа%', 'СЧ обувь', 'СЧ'],
        rows3, _row_colors(len(s3)), date_str, time_str
    )
    images.append(('services', img3))

    return images
