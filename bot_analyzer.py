import pandas as pd

DIRECTORS = {
    10045: "Ярмолюк Марина Вячеславовна",
    10262: "Ломакина Анна Владимировна",
    10347: "Кумановская Анна Сергеевна",
    10775: "Рожкова Виктория Викторовна",
    10942: "Харченко Оксана Николаевна",
    10969: "Королева Виктория Викторовна",
    11193: "Касьянова Елена Ивановна",
    11208: "Коропченко Галина Александровна",
    11267: "Долгополова Наталья Николаевна",
    11338: "Мелехина Юлия Олеговна",
    11497: "Моисеева Татьяна Алексеевна",
    11666: "Грибова Елизавета Николаевна",
    11667: "Троценко Жанна Васильевна",
    11682: "Харченко Дарья Сергеевна",
    11694: "Коваленко Светлана Викторовна",
    11840: "Гайворонская Елена Сергеевна",
    11895: "Гаркушина Анна Вадимовна",
    13034: "Шавкунова Татьяна Владимировна",
    13125: "Улюмджиева Валентина Александровна",
    13159: "Кравченко Анастасия Николаевна",
}

GROUPS = {
    10045: "10045 Шахты Ростов",
    10262: "10262 Кореновск Ростов",
    10347: "10347 Батайск Ростов",
    10775: "10775 Тихорецк Ростов",
    10942: "10942 Новочеркасск Ростов",
    10969: "10969 Шахты Ростов",
    11193: "11193 Белая Калитва Ростов",
    11208: "11208 Азов (Ростов)",
    11267: "11267 Донецк Ростов",
    11338: "11338 Динская Ростов",
    11497: "11497 Каменск Ростов",
    11666: "11666 Тимашевск Ростов",
    11667: "11667 Батайск Черноморское",
    11682: "11682 Новочеркасск Ростов",
    11694: "11694 Гуково Ростов",
    11840: "11840 Миллерово Ростов",
    11895: "11895_Новошахтинск Ростов",
    13034: "13034 Волгодонск Ростов",
    13125: "13125_Элиста_Ростов",
    13159: "13159_Сальск_Ростов",
}

def s(v, mult=100):
    try: return round(float(v) * mult, 1)
    except: return 0.0

def sv(v):
    try: return round(float(v))
    except: return 0

def f(v, norm):
    if v < norm * 0.8: return '🔴'
    if v < norm: return '🟡'
    return '✅'

def fd(v, bad=-15):
    if v <= bad: return '🔴'
    if v < 0: return '🟡'
    return '✅'

def first_name(full):
    parts = full.split()
    return parts[1] if len(parts) > 1 else full

def load(filepath):
    df = pd.read_excel(filepath, sheet_name='Подразделение', header=None)
    date_str = str(df.iloc[1][0]).replace('Дата формирования: ', '').strip()
    time_str = str(df.iloc[2][0]).replace('Время: ', '').strip()
    bench = df.iloc[35]

    norms = {
        'kop':    s(bench[7]),
        'sch_ob': sv(bench[13]),
        'kosm':   s(bench[29]),
        'steli':  s(bench[30]),
        'rass':   s(bench[26]),
        'sbp':    s(bench[28]),
        'yui':    s(bench[34]),
        'sreb':   s(bench[35]),
        'zol':    s(bench[36]),
    }

    total = {
        'to':     sv(bench[2]),
        'kop':    s(bench[7]),
        'pvch':   round(s(bench[10], mult=1), 2),
        'sch':    sv(bench[15]),
        'traf':   sv(bench[17]),
        'to_ned': s(bench[3]),
        'to_vch': s(bench[4]),
    }

    stores = []
    for i in range(5, 25):
        row = df.iloc[i]
        sid = int(s(row[0], mult=1))
        stores.append({
            'id':       sid,
            'name':     str(row[1]),
            'group':    GROUPS.get(sid, f'{sid}'),
            'director': DIRECTORS.get(sid, 'Директор'),
            'to':       sv(row[2]),
            'to_ned':   s(row[3]),
            'to_vch':   s(row[4]),
            'plan':     s(row[5]),
            'kop':      s(row[7]),
            'pvch':     round(s(row[10], mult=1), 2),
            'sch_ob':   sv(row[13]),
            'sch':      sv(row[15]),
            'traf_ned': s(row[18]),
            'traf_vch': s(row[19]),
            'kozha':    s(row[20]),
            'rass':     s(row[26]),
            'sch_rass': sv(row[27]),
            'sbp':      s(row[28]),
            'kosm':     s(row[29]),
            'steli':    s(row[30]),
            'yui':      s(row[34]),
            'sreb':     s(row[35]),
            'zol':      s(row[36]),
        })

    return stores, date_str, time_str, norms, total


def make_general(stores, date_str, time_str, total):
    def anti(key, label, fmt='pct'):
        ranked = sorted(stores, key=lambda x: x[key])[:5]
        nums = ['1️⃣','2️⃣','3️⃣','4️⃣','5️⃣']
        lines = [f'📉 <b>Антирейтинг — {label}:</b>']
        for i, st in enumerate(ranked):
            v = st[key]
            if fmt == 'pct':   val = f'{v:+.1f}%'
            elif fmt == 'rub': val = f'{v:,}₽'
            elif fmt == 'par': val = f'{v:.2f} пар'
            lines.append(f'{nums[i]}  {st["name"]}: {val}')
        return '\n'.join(lines)

    return f'''🔖 <b>ОТЧЁТ ПО ЧАСУ ПРОДАЖ</b>
{date_str} | 🕐 {time_str}
━━━━━━━━━━━━━━━━━━━━━

📈 <b>ИТОГО по подразделению:</b>
💰 ТО: {total["to"]:,}₽
🧾 КОП (конв.): {total["kop"]}%
👟 ПвЧ: {total["pvch"]} пар
💳 СЧ: {total["sch"]:,}₽
👥 Трафик: {total["traf"]}
📊 ТО к неделе: {total["to_ned"]:+.1f}%
📊 ТО к вчера: {total["to_vch"]:+.1f}%
━━━━━━━━━━━━━━━━━━━━━

{anti("to_ned", "ТО к неделе")}
━━━━━━━━━━━━━━━━━━━━━

{anti("kop", "КОП")}
━━━━━━━━━━━━━━━━━━━━━

{anti("pvch", "ПвЧ", fmt="par")}
━━━━━━━━━━━━━━━━━━━━━

{anti("sch", "СЧ", fmt="rub")}'''


def make_praise(stores, date_str):
    def top1(key, label, fmt='pct'):
        best = max(stores, key=lambda x: x[key])
        v = best[key]
        if fmt == 'pct':   val = f'{v:.1f}%'
        elif fmt == 'rub': val = f'{v:,}₽'
        elif fmt == 'par': val = f'{v:.2f} пар'
        return f'🏆 <b>{label}</b>\n   {best["name"]} — {val}'

    lines = [
        f'🌟 <b>ИТОГИ ДНЯ — {date_str}</b>',
        '━━━━━━━━━━━━━━━━━━━━━',
        'Молодцы! Выделяем лучших за день:',
        '',
        top1('to',   'Лучший товарооборот',  fmt='rub'),
        top1('pvch', 'Лучшие пары в чеке',   fmt='par'),
        top1('sch',  'Лучший средний чек',   fmt='rub'),
        top1('yui',  'Лучшие ЮИ',            fmt='pct'),
        top1('zol',  'Лучшее золото',        fmt='pct'),
        top1('sreb', 'Лучшее серебро',       fmt='pct'),
        '',
        'Так держать! Завтра — новый день и новые возможности. 💪',
    ]
    return '\n'.join(lines)


def make_store_msg(st, time_str, norms):
    """Полный разбор — для всех магазинов"""

    # Заголовок с цветом по плану
    if st['plan'] >= 85:
        plan_icon = '✅'
    elif st['plan'] >= 70:
        plan_icon = '🟡'
    else:
        plan_icon = '🔴'

    plan_line = f'🎯 ПЛАН ДНЯ: {st["plan"]}% {plan_icon}'
    if st['plan'] < 100:
        plan_line += f'  (-{round(100-st["plan"],1)}% от цели)'

    # Задачи — только для тех у кого есть реальные проблемы
    tasks = []
    n = 1

    if st['plan'] < 50:
        tasks.append(f'{n}️⃣ ПЛАН {st["plan"]}% — КРИТИЧНО. Лично в зал, контролировать каждого продавца')
        tasks.append(f'   📋 Объяснительная: причина + действия + план на завтра')
        n += 1
    elif st['plan'] < 70:
        tasks.append(f'{n}️⃣ ПЛАН {st["plan"]}% — конкретные действия прямо сейчас')
        n += 1

    if st['kop'] < norms['kop']:
        tasks.append(f'{n}️⃣ КОП {st["kop"]}% — каждый покупатель должен уйти с покупкой')
        n += 1
    if st['kosm'] < norms['kosm'] * 0.8:
        tasks.append(f'{n}️⃣ Косметика {st["kosm"]}% — уход к каждой паре, без исключений')
        n += 1
    if st['steli'] < norms['steli'] * 0.8:
        tasks.append(f'{n}️⃣ Стельки {st["steli"]}% — при примерке: «Возьмём стельки — сразу почувствуете разницу»')
        n += 1
    if st['yui'] < norms['yui'] * 0.8:
        tasks.append(f'{n}️⃣ ЮИ {st["yui"]}% — при каждой продаже предлагать украшения')
        n += 1
    if st['rass'] < norms['rass'] * 0.5 and st['sch_ob'] > 2500:
        tasks.append(f'{n}️⃣ Рассрочка {st["rass"]}% — при сумме от 3 000 ₽: «Удобнее разбить на части?»')
        n += 1

    tasks_block = ''
    if tasks:
        tasks_block = f'''
━━━━━━━━━━━━━━━━━━━━━
⚡️ ЗАДАЧИ ДО ЗАКРЫТИЯ:

{chr(10).join(tasks)}'''

    return f'''📍 М.{st["id"]} | {st["name"]}
👤 {st["director"]}
━━━━━━━━━━━━━━━━━━━━━

{plan_line}

💰 ВЫРУЧКА
├ ТО: {st["to"]:,} руб
├ к вчера:  {st["to_vch"]:+.1f}% {fd(st["to_vch"])}
└ к неделе: {st["to_ned"]:+.1f}% {fd(st["to_ned"],-20)}

👣 ТРАФИК
├ к вчера:  {st["traf_vch"]:+.1f}% {fd(st["traf_vch"])}
└ к неделе: {st["traf_ned"]:+.1f}% {fd(st["traf_ned"],-20)}

🛒 КАЧЕСТВО ПРОДАЖ
├ КОП: {st["kop"]}% {f(st["kop"],norms["kop"])}  (норма {norms["kop"]}%)
├ СЧ обуви: {st["sch_ob"]:,} руб {f(st["sch_ob"],norms["sch_ob"])}  (норма {norms["sch_ob"]:,})
└ Кожа: {st["kozha"]}%

➕ ДОПРОДАЖИ
├ Косметика: {st["kosm"]}% {f(st["kosm"],norms["kosm"])}  (норма {norms["kosm"]}%)
├ Стельки:   {st["steli"]}% {f(st["steli"],norms["steli"])}  (норма {norms["steli"]}%)
├ ЮИ:        {st["yui"]}% {f(st["yui"],norms["yui"])}  (норма {norms["yui"]}%)
├ Серебро:   {st["sreb"]}% {f(st["sreb"],norms["sreb"])}  (норма {norms["sreb"]}%)
├ Золото:    {st["zol"]}% {f(st["zol"],norms["zol"])}  (норма {norms["zol"]}%)
├ Рассрочка: {st["rass"]}% {f(st["rass"],norms["rass"])}  (норма {norms["rass"]}%)
└ СБП:       {st["sbp"]}% {f(st["sbp"],norms["sbp"])}  (норма {norms["sbp"]}%){tasks_block}'''


def run(filepath):
    stores, date_str, time_str, norms, total = load(filepath)

    return {
        "general": {
            "target": "ОБЩИЙ ЧАТ ПОДРАЗДЕЛЕНИЯ",
            "message": make_general(stores, date_str, time_str, total),
            "parse_mode": "HTML"
        },
        "stores": [
            {
                "store_id":   st['id'],
                "group_name": st['group'],
                "plan_pct":   st['plan'],
                "message":    make_store_msg(st, time_str, norms),
                "parse_mode": "HTML"
            }
            for st in sorted(stores, key=lambda x: x['plan'])
        ]
    }
