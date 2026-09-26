#!/usr/bin/env python3
"""WhatsApp-Auswertung (2 Seiten, Handy-Hochformat) fuer einen Doppelkopf-Abend.

Nutzung:  python3 doko_whatsapp_pdf.py <Doppelkopf_Datenbank.xlsx> [Ausgabe.pdf] [TT.MM.JJJJ]
Ohne Datum wird der letzte Abend in der DB genommen. Rendering per Node-Playwright/Chromium
(optional Env PW_MODULE = Pfad zum playwright-Modul, PW_CHROMIUM = Chromium-Binary).
Legt neben dem PDF je Seite ein PNG zur Sichtkontrolle ab.
"""
import sys, collections, html, subprocess, os, tempfile, json, itertools
import openpyxl

PLAYERS = ['Sebastian', 'Andy', 'Frede', 'Fabian', 'Chris']
COL = {'Sebastian': 4, 'Andy': 6, 'Frede': 8, 'Fabian': 10, 'Chris': 12}  # Punkte-Spalte (1-basiert); Letter = rechts daneben
COLORS = {'Sebastian': '#e8874a', 'Andy': '#4a9ee8', 'Frede': '#4caf80', 'Fabian': '#c47ee8', 'Chris': '#e84a4a'}
HOST = {'Basti': 'Sebastian'}
SOLO = {'B': 'Buben', 'D': 'Damen', 'T': 'Trumpf', 'S': 'Stilles', 'F': 'Solo', 'A': 'Solo'}
PAGE_W, PAGE_H = 454, 960


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    sessions = collections.OrderedDict()
    for r in range(5, ws.max_row + 1):
        v = [c.value for c in ws[r]]
        if v[0]:
            sessions.setdefault(v[0].strftime('%d.%m.%Y'), []).append(v)
    return sessions


def pts(x, p):
    v = x[COL[p] - 1]
    if v is None or (v == 0 and str(x[COL[p]] or '') == 'M'):
        return None
    return v


def letter(x, p):
    return str(x[COL[p]] or '').strip()


def solo_player(x):
    t = x[16]
    if not t:
        return None
    c = [p for p in PLAYERS if letter(x, p) == t]
    if len(c) == 1:
        return c[0]
    vals = {p: pts(x, p) for p in PLAYERS if pts(x, p) not in (None, 0)}
    c = [p for p, v in vals.items() if all(abs(v) == 3 * abs(u) for q, u in vals.items() if q != p)]
    return c[0] if len(c) == 1 else None


def f(v):
    return f'+{v}' if v > 0 else str(v)


def dot(p):
    return f'<i class="dot" style="background:{COLORS[p]}"></i>'


def name(p, bold=True):
    return f'<span class="pn">{dot(p)}{"<b>" if bold else ""}{p}{"</b>" if bold else ""}</span>'


def sgn(v, cls=True):
    c = 'pos' if v > 0 else ('neg' if v < 0 else 'zero')
    return f'<span class="{c}">{f(v)}</span>'


def chart(rows, present, rank, solos):
    W, H = 418, 212
    L, R, T, B = 30, 114, 10, 22
    cum = {p: 0 for p in present}
    hist = {p: [0] for p in present}
    for x in rows:
        for p in present:
            cum[p] += pts(x, p) or 0
            hist[p].append(cum[p])
    lo = min(min(v) for v in hist.values()); hi = max(max(v) for v in hist.values())
    pad = (hi - lo) * 0.05 or 5
    lo, hi = lo - pad, hi + pad
    n = len(rows)
    X = lambda i: L + i * (W - L - R) / n
    Y = lambda v: T + (hi - v) * (H - T - B) / (hi - lo)
    o = [f'<svg viewBox="0 0 {W} {H}" width="100%" font-family="DejaVu Sans, Arial">']
    # Bock-Phasen hinterlegen
    i = 0
    while i < n:
        if rows[i][17]:
            j = i
            while j < n and rows[j][17]:
                j += 1
            o.append(f'<rect x="{X(i):.1f}" y="{T}" width="{X(j)-X(i):.1f}" height="{H-T-B}" fill="#f3ead2"/>')
            i = j
        else:
            i += 1
    # Gitter
    step = 50 if hi - lo > 150 else 25
    g = int(lo // step) * step
    while g <= hi:
        if g >= lo:
            o.append(f'<line x1="{L}" x2="{W-R}" y1="{Y(g):.1f}" y2="{Y(g):.1f}" stroke="{"#9c9480" if g == 0 else "#e6dfcc"}" stroke-width="{1 if g == 0 else 0.7}"/>'
                     f'<text x="{L-4}" y="{Y(g)+3:.1f}" font-size="9" fill="#8a8471" text-anchor="end">{f(g) if g else 0}</text>')
        g += step
    for r in range(0, n + 1, 5):
        if r:
            o.append(f'<text x="{X(r):.1f}" y="{H-8}" font-size="9" fill="#8a8471" text-anchor="middle">{r}</text>')
    o.append(f'<text x="{W-R}" y="{H-8}" font-size="9" fill="#8a8471" text-anchor="end" dx="14">Rd.</text>')
    # Linien (Sieger zuletzt/obenauf)
    for p in reversed(rank):
        pl = ' '.join(f'{X(k):.1f},{Y(v):.1f}' for k, v in enumerate(hist[p]))
        o.append(f'<polyline fill="none" stroke="#fff" stroke-width="4.5" stroke-linejoin="round" points="{pl}"/>')
        o.append(f'<polyline fill="none" stroke="{COLORS[p]}" stroke-width="2.4" stroke-linejoin="round" points="{pl}"/>')
    # Solo-Marker
    for r, p, t, v in solos:
        o.append(f'<circle cx="{X(r):.1f}" cy="{Y(hist[p][r]):.1f}" r="3.6" fill="{"#fff" if v < 0 else COLORS[p]}" stroke="{COLORS[p]}" stroke-width="1.8"/>')
    # Direkte Endlabels mit Kollisionsvermeidung
    ends = sorted(((Y(hist[p][-1]), p) for p in present))
    ys = []
    for y, p in ends:
        y = max(y, ys[-1] + 15) if ys else y
        ys.append(y)
    over = ys[-1] - (H - B) if ys and ys[-1] > H - B else 0
    ys = [y - over for y in ys]
    for (y0, p), y in zip(ends, ys):
        xe = X(n)
        o.append(f'<circle cx="{xe:.1f}" cy="{y0:.1f}" r="3.5" fill="{COLORS[p]}" stroke="#fff" stroke-width="1.5"/>')
        o.append(f'<line x1="{xe+4:.1f}" y1="{y0:.1f}" x2="{xe+10:.1f}" y2="{y:.1f}" stroke="{COLORS[p]}" stroke-width="1"/>')
        o.append(f'<circle cx="{xe+15:.1f}" cy="{y:.1f}" r="4" fill="{COLORS[p]}"/>')
        o.append(f'<text x="{xe+22:.1f}" y="{y+3.5:.1f}" font-size="9.5" font-weight="700" fill="#23201a">{p} {f(hist[p][-1])}</text>')
    o.append('</svg>')
    legend = ('<div class="leg">' + ''.join(f'<span>{dot(p)}{p}</span>' for p in rank) +
              '<span><i class="sw bock"></i>Bock</span><span><i class="ring"></i>Solo (voll = gewonnen)</span></div>')
    return ''.join(o) + legend


def build(sessions, date):
    rows = sessions[date]
    n = len(rows)
    host = HOST.get(rows[0][1], rows[0][1])
    present = [p for p in PLAYERS if any(pts(x, p) is not None for x in rows)]
    tot = {p: sum(pts(x, p) or 0 for x in rows) for p in present}
    rank = sorted(present, key=lambda p: -tot[p])
    played = {p: sum(pts(x, p) is not None for x in rows) for p in present}
    won = {p: sum((pts(x, p) or 0) > 0 for x in rows) for p in present}
    lostr = {p: sum((pts(x, p) or 0) < 0 for x in rows) for p in present}
    best_p = {p: max(pts(x, p) for x in rows if pts(x, p) is not None) for p in present}
    worst_p = {p: min(pts(x, p) for x in rows if pts(x, p) is not None) for p in present}

    cells = [(pts(x, p), p, i + 1) for i, x in enumerate(rows) for p in present if pts(x, p) is not None]
    best, worst = max(cells), min(cells)
    streak, lstreak = {}, {}
    for p in present:
        b = c = bl = cl = 0
        for x in rows:
            v = pts(x, p)
            if v is None:
                continue
            c = c + 1 if v > 0 else 0
            cl = cl + 1 if v < 0 else 0
            b, bl = max(b, c), max(bl, cl)
        streak[p], lstreak[p] = b, bl
    sp = max(present, key=lambda p: streak[p])
    lp = max(present, key=lambda p: lstreak[p])

    solos = []
    for i, x in enumerate(rows):
        p = solo_player(x)
        if p:
            solos.append((i + 1, p, x[16], pts(x, p)))
    swon = sum(1 for s in solos if s[3] > 0)
    hz = [(i + 1, p, pts(x, p)) for i, x in enumerate(rows) for p in present if letter(x, p) == 'H']
    ar = [(i + 1, p, pts(x, p)) for i, x in enumerate(rows) for p in present if letter(x, p) == 'A']
    bockrows = [x for x in rows if x[17]]
    schw = sum(1 for x in rows if x[14])

    # Bock vs. normal
    bock_pts = {p: sum(pts(x, p) or 0 for x in rows if x[17]) for p in present}
    norm_pts = {p: tot[p] - bock_pts[p] for p in present}

    # Phasen (Drittel)
    thirds = [(0, round(n / 3)), (round(n / 3), round(2 * n / 3)), (round(2 * n / 3), n)]
    phase = {p: [sum(pts(x, p) or 0 for x in rows[a:b]) for a, b in thirds] for p in present}

    # Duos: gleiche Seite (gleiches Vorzeichen) in Runden ohne Solo
    duo = collections.defaultdict(lambda: [0, 0])
    for x in rows:
        if solo_player(x):
            continue
        act = [p for p in present if pts(x, p) not in (None, 0)]
        for a, b in itertools.combinations(act, 2):
            if (pts(x, a) > 0) == (pts(x, b) > 0):
                duo[(a, b)][0 if pts(x, a) > 0 else 1] += 1
    duos = sorted(duo.items(), key=lambda kv: (-(kv[1][0] - kv[1][1]), -kv[1][0]))

    # Historie
    dates = list(sessions)
    idx = dates.index(date)
    def sess_tot(d):
        return {p: sum(pts(x, p) or 0 for x in sessions[d]) for p in PLAYERS}
    all_tot = {d: sess_tot(d) for d in dates[:idx + 1]}
    def pres(d):
        return [p for p in PLAYERS if any(pts(x, p) is not None for x in sessions[d])]
    before = {p: sum(all_tot[d][p] for d in dates[:idx]) for p in PLAYERS}
    after = {p: before[p] + all_tot[date][p] for p in PLAYERS}
    rb = {p: i for i, p in enumerate(sorted(PLAYERS, key=lambda p: -before[p]))}
    ra = sorted(PLAYERS, key=lambda p: -after[p])
    wins = collections.Counter(max(pres(d), key=lambda p: all_tot[d][p]) for d in dates[:idx + 1])
    evening_scores = sorted((all_tot[d][p] for d in dates[:idx + 1] for p in pres(d)), reverse=True)
    w0 = rank[0]
    rec_rank = evening_scores.index(tot[w0]) + 1
    low_rank = sorted(evening_scores).index(tot[rank[-1]]) + 1
    maxbar = max(abs(v) for v in after.values()) or 1

    # ---------- Seite 1 ----------
    mx = max(abs(v) for v in tot.values()) or 1
    lines = []
    for i, p in enumerate(rank, 1):
        w = abs(tot[p]) / mx * 48
        side = f'left:50%;width:{w}%' if tot[p] >= 0 else f'left:{50-w}%;width:{w}%'
        lines.append(f'<div class="rk"><span class="pos0">{i}.</span><span class="nm">{name(p)}</span>'
                     f'<span class="track"><span class="mid"></span><span class="bar" style="{side};background:{COLORS[p]}"></span></span>'
                     f'<span class="pt">{sgn(tot[p])}</span></div>')

    by_player = collections.defaultdict(list)
    for r, p, t, v in solos:
        by_player[p].append((r, t, v))
    solo_rows = ''.join(
        f'<div class="sl"><span class="nm2">{name(p)}</span>'
        + ''.join(f'<span class="chip {"w" if v > 0 else "l"}">R{r} {SOLO.get(t, t)} {f(v)}</span>' for r, t, v in by_player[p]) + '</div>'
        for p in rank if by_player[p])

    page1 = f"""
<div class="pg">
<div class="hd"><div class="k">DOPPELKOPF · SPIELABEND</div><h1>{date}</h1><div class="s">bei {html.escape(host)} · {n} Runden · {len(present)} Spieler · {len(bockrows)} Bockrunden</div></div>
<div class="win"><div class="cr">👑</div><div><div class="k2">Sieger des Abends</div><div class="n">{dot(w0)}{w0}</div><div class="sub">{won[w0]} von {played[w0]} Runden gewonnen</div></div><div class="p">{f(tot[w0])}</div></div>
<div class="card"><h2>Abendwertung</h2>{''.join(lines)}</div>
<div class="card"><h2>Punkteverlauf</h2>{chart(rows, present, rank, solos)}</div>
<div class="grid">
 <div class="t"><div class="lb">🚀 Beste Runde</div><div class="v pos">{f(best[0])}</div><div class="d">{name(best[1], False)} · R{best[2]}</div></div>
 <div class="t"><div class="lb">💥 Bitterste Runde</div><div class="v neg">{worst[0]}</div><div class="d">{name(worst[1], False)} · R{worst[2]}</div></div>
 <div class="t"><div class="lb">🔥 Längste Siegesserie</div><div class="v">{streak[sp]} Runden</div><div class="d">{name(sp, False)}</div></div>
 <div class="t"><div class="lb">🧊 Längste Pechsträhne</div><div class="v">{lstreak[lp]} Runden</div><div class="d">{name(lp, False)}</div></div>
</div>
PHASES_CARD
<div class="ft">Seite 1/2 · Details auf Seite 2</div>
</div>"""

    # ---------- Seite 2 ----------
    det = ''.join(
        f'<tr><td>{name(p)}</td><td class="r">{played[p]}</td><td class="r">{won[p]}–{lostr[p]}</td>'
        f'<td class="r">{round(100*won[p]/played[p])}%</td><td class="r">{sgn(round(tot[p]/played[p],1))}</td>'
        f'<td class="r">{sgn(best_p[p])}</td><td class="r">{sgn(worst_p[p])}</td></tr>' for p in rank)

    bmx = max(max(abs(bock_pts[p]), abs(norm_pts[p])) for p in present) or 1
    def mini(v):
        w = abs(v) / bmx * 48
        side = f'left:50%;width:{w}%' if v >= 0 else f'left:{50-w}%;width:{w}%'
        col = '#3f8f5f' if v >= 0 else '#c0392b'
        return f'<span class="track sm"><span class="mid"></span><span class="bar" style="{side};background:{col}"></span></span>'
    bock_tbl = ''.join(f'<tr><td>{name(p)}</td><td>{mini(norm_pts[p])}</td><td class="r">{sgn(norm_pts[p])}</td>'
                       f'<td>{mini(bock_pts[p])}</td><td class="r">{sgn(bock_pts[p])}</td></tr>' for p in rank)

    def heat(v):
        a = min(abs(v) / 60, 1) * 0.55 + 0.08
        c = f'rgba(46,125,79,{a:.2f})' if v > 0 else (f'rgba(192,57,43,{a:.2f})' if v < 0 else '#f3efe4')
        return f'<td class="hm" style="background:{c}">{f(v)}</td>'
    ph_head = ''.join(f'<th>R{a+1}–{b}</th>' for a, b in thirds)
    phase_tbl = ''.join(f'<tr><td>{name(p)}</td>{"".join(heat(v) for v in phase[p])}</tr>' for p in rank)

    def duo_row(kv):
        (a, b), (w, l) = kv
        return f'<div class="duo">{dot(a)}{a}&nbsp;&amp;&nbsp;{dot(b)}{b}<span class="dr">{w}–{l}</span></div>'
    duo_best = ''.join(duo_row(kv) for kv in duos[:2])
    duo_worst = ''.join(duo_row(kv) for kv in duos[-2:][::-1])

    extra = ' · '.join(filter(None, [
        '💍 ' + ', '.join(f'R{r} {p} {f(v)}' for r, p, v in hz) if hz else '',
        '🤲 ' + ', '.join(f'R{r} {p} {f(v)}' for r, p, v in ar) if ar else '',
        f'🐷 {schw}× Schwein' if schw else '']))

    def arrow(p):
        d = rb[p] - ra.index(p)
        return f'<span class="up">▲{d}</span>' if d > 0 else (f'<span class="dn">▼{-d}</span>' if d < 0 else '<span class="eq">–</span>')
    table = ''
    for i, p in enumerate(ra):
        w = abs(after[p]) / maxbar * 48
        side = f'left:50%;width:{w}%' if after[p] >= 0 else f'left:{50-w}%;width:{w}%'
        table += (f'<tr><td class="c0">{i+1}.</td><td>{name(p)}</td><td class="trk"><span class="track sm"><span class="mid"></span>'
                  f'<span class="bar" style="{side};background:{COLORS[p]}"></span></span></td><td class="r"><b>{f(after[p])}</b></td>'
                  f'<td class="r small">{f(all_tot[date][p]) if p in present else "–"}</td><td class="r small">{wins[p]}🏆</td><td class="c">{arrow(p)}</td></tr>')

    page1 = page1.replace('PHASES_CARD', '<div class="card"><h2>Abend in drei Phasen</h2><table class="ph"><tr><th></th>{ph_head}</tr>{phase_tbl}</table></div>'.format(ph_head=ph_head, phase_tbl=phase_tbl))
    umsatz = sum(v for x in rows for p in present for v in [pts(x, p)] if v and v > 0)
    solo_net = collections.Counter()
    for r, p, t, v in solos:
        solo_net[p] += v
    sk = max(solo_net, key=lambda p: solo_net[p]) if solos else None
    bp = max(present, key=lambda p: bock_pts[p])
    facts = ''.join(f'<div class="fact"><span>{a}</span><b>{b}</b></div>' for a, b in [
        ('💰 Punkte-Umsatz (Summe aller Gewinne)', f'{umsatz}'),
        ('🎭 Solo-König (Netto aus Solos)', f'{sk} {f(solo_net[sk])}' if sk else '–'),
        ('🔔 Bock-Profiteur', f'{bp} {f(bock_pts[bp])}'),
    ])
    rec = f'{w0}s {f(tot[w0])} ist das {rec_rank}.-beste Abendergebnis aller {len(evening_scores)} Einzelwertungen.' if rec_rank <= 5 else ''
    rec2 = f'{rank[-1]}s {tot[rank[-1]]} ist das {low_rank}.-schlechteste.' if low_rank <= 5 else ''

    page2 = f"""
<div class="pg">
<div class="hd sm2"><div class="k">DETAILS · {date} · bei {html.escape(host)}</div></div>
<div class="card"><h2>Spieler im Detail</h2><table class="det"><tr><th></th><th class="r">Rd.</th><th class="r">S–N</th><th class="r">Quote</th><th class="r">Ø/Rd.</th><th class="r">Top</th><th class="r">Flop</th></tr>{det}</table></div>
<div class="card"><h2>Solos · {len(solos)} gespielt · {swon} gewonnen</h2>{solo_rows or '<div class="sl">keine</div>'}
<div class="ev">{extra}</div></div>
<div class="grid2">
<div class="card m0"><h2>🤝 Traumduo</h2>{duo_best}<div class="note">gemeinsam gewonnen–verloren</div></div>
<div class="card m0"><h2>🙈 Pechduo</h2>{duo_worst}<div class="note">ohne Solorunden</div></div>
</div>
<div class="card"><h2>Normal vs. Bock ({n-len(bockrows)} / {len(bockrows)} Runden)</h2><table class="bk"><tr><th></th><th colspan="2">normal</th><th colspan="2">🔔 Bock</th></tr>{bock_tbl}</table></div>
<div class="card"><h2>Kurz & knapp</h2>{facts}</div>
<div class="card"><h2>Ewige Tabelle · {idx+1} Abende</h2><table class="et">{table}</table>
<div class="note">Gesamtpunkte · Abend · Abendsiege · Platzveränderung{('<br>🏅 ' + rec) if rec else ''}{(' ' + rec2) if rec2 else ''}</div></div>
<div class="ft">Seite 2/2 · Alle Abende & Runden: sebaredo.github.io/DokoRunde</div>
</div>"""

    css = f"""
    @page{{margin:0;size:{PAGE_W}px {PAGE_H}px}}
    *{{box-sizing:border-box}}body{{margin:0;font-family:'DejaVu Sans',Arial,sans-serif;background:#f6f2e8;color:#23201a}}
    .pg{{width:{PAGE_W}px;height:{PAGE_H}px;padding:16px 16px 10px;overflow:hidden;page-break-after:always;position:relative;background:#f6f2e8}}
    .pg:last-child{{page-break-after:auto}}
    .hd{{background:#1f3a2c;color:#f6f2e8;border-radius:14px;padding:12px 16px}}
    .hd .k{{font-size:10px;letter-spacing:1.6px;opacity:.7}}.hd h1{{margin:2px 0 0;font-size:24px}}.hd .s{{opacity:.85;font-size:12px;margin-top:3px}}
    .hd.sm2{{padding:9px 14px}}
    .win{{margin-top:10px;background:#fff;border-radius:14px;padding:10px 14px;border:2px solid #c9a84c;display:flex;align-items:center;gap:12px}}
    .win .cr{{font-size:32px}}.k2{{font-size:9.5px;letter-spacing:1px;text-transform:uppercase;color:#8a8471}}
    .win .n{{font-size:21px;font-weight:800}}.win .p{{margin-left:auto;font-size:30px;font-weight:800;color:#2e7d4f}}
    .win .sub{{font-size:11px;color:#6d6758}}
    .card{{margin-top:9px;background:#fff;border-radius:14px;padding:9px 13px}}
    .card h2{{margin:0 0 6px;font-size:11.5px;text-transform:uppercase;letter-spacing:.8px;color:#8a8471}}
    .dot{{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px;vertical-align:0}}
    .rk{{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:13.5px}}.pos0{{width:16px;color:#8a8471}}.nm{{width:92px}}
    .track{{position:relative;flex:1;height:11px;display:block}}.track.sm{{height:8px;min-width:60px}}
    .mid{{position:absolute;left:50%;top:-2px;bottom:-2px;border-left:1px solid #cfc8b6}}
    .bar{{position:absolute;top:0;bottom:0;border-radius:4px}}.pt{{width:46px;text-align:right;font-weight:800}}
    .pos{{color:#2e7d4f}}.neg{{color:#c0392b}}.zero{{color:#8a8471}}
    .leg{{display:flex;flex-wrap:wrap;gap:4px 10px;font-size:10.5px;color:#4b463b;margin-top:2px}}
    .sw{{display:inline-block;width:12px;height:9px;margin-right:4px;vertical-align:-1px;border-radius:2px}}.bock{{background:#f3ead2;border:1px solid #e3d6b0}}
    .ring{{display:inline-block;width:8px;height:8px;border-radius:50%;border:1.8px solid #6d6758;margin-right:4px;vertical-align:-1px}}
    .grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:9px}}.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}
    .m0{{margin-top:9px}}
    .t{{background:#fff;border-radius:12px;padding:8px 10px}}.t .lb{{font-size:10px;color:#8a8471;text-transform:uppercase;letter-spacing:.5px}}
    .t .v{{font-size:19px;font-weight:800;margin-top:1px}}.t .d{{font-size:11.5px;color:#4b463b}}
    .sl{{font-size:12px;margin:3px 0;display:flex;flex-wrap:wrap;align-items:center;gap:3px}}.nm2{{width:88px}}
    .chip{{display:inline-block;border-radius:7px;padding:1px 5px;font-size:10.5px;font-weight:700}}
    .w{{background:#dff1e5;color:#1f6b40}}.l{{background:#f8dfdb;color:#a0301f}}
    .ev{{font-size:11px;margin-top:5px;color:#4b463b}}
    table{{width:100%;border-collapse:collapse;font-size:12px}}td,th{{padding:3px 2px;border-bottom:1px solid #eee7d6}}
    th{{font-size:10px;color:#8a8471;font-weight:600;text-align:left}}.r{{text-align:right}}.c{{text-align:center;width:28px}}.c0{{width:18px;color:#8a8471}}
    .small{{font-size:10.5px;color:#6d6758}}.trk{{width:34%}}
    .hm{{text-align:center;font-weight:700;border:2px solid #fff;border-radius:4px}}
    .ph th{{text-align:center}}
    .up{{color:#2e7d4f;font-weight:700}}.dn{{color:#c0392b;font-weight:700}}.eq{{color:#b9b3a3}}
    .duo{{font-size:11.5px;margin:3px 0;display:flex;align-items:center}}.dr{{margin-left:auto;font-weight:800}}
        .fact{{display:flex;justify-content:space-between;font-size:12px;padding:2px 0;border-bottom:1px solid #f1ebdc}}
    .note{{font-size:10px;color:#8a8471;margin-top:4px}}
    .ft{{position:absolute;bottom:8px;left:0;right:0;text-align:center;font-size:10px;color:#8a8471}}
    """
    return f'<!doctype html><meta charset="utf-8"><style>{css}</style>{page1}{page2}'


def main():
    db = sys.argv[1]
    sessions = load(db)
    date = sys.argv[3] if len(sys.argv) > 3 else list(sessions)[-1]
    out = os.path.abspath(sys.argv[2] if len(sys.argv) > 2 else f'Doko-Abend {date}.pdf')
    tmp = tempfile.mkdtemp()
    hp = os.path.join(tmp, 'wa.html')
    open(hp, 'w').write(build(sessions, date))
    js = f"""const pw=require(process.env.PW_MODULE||'playwright');(async()=>{{
const b=await pw.chromium.launch(process.env.PW_CHROMIUM?{{executablePath:process.env.PW_CHROMIUM}}:{{}});
const p=await b.newPage({{viewport:{{width:{PAGE_W},height:{PAGE_H}}}}});await p.goto('file://'+{json.dumps(hp)});
const ov=await p.evaluate(()=>[...document.querySelectorAll('.pg')].map(pg=>{{const r=pg.getBoundingClientRect();const f=pg.querySelector('.ft').getBoundingClientRect();
 const kids=[...pg.children].filter(c=>!c.classList.contains('ft'));const last=Math.max(...kids.map(c=>c.getBoundingClientRect().bottom));return Math.round(last-(f.top-4));}}));
console.log('Ueberlauf je Seite (px, >0 = zu voll):',JSON.stringify(ov));
await p.pdf({{path:{json.dumps(out)},width:'{PAGE_W}px',height:'{PAGE_H}px',printBackground:true}});
const pgs=await p.$$('.pg');for(let i=0;i<pgs.length;i++)await pgs[i].screenshot({{path:{json.dumps(out[:-4])}+'-S'+(i+1)+'.png'}});
await b.close();}})();"""
    jp = os.path.join(tmp, 'r.js')
    open(jp, 'w').write(js)
    subprocess.run(['node', jp], check=True)
    print(out)


if __name__ == '__main__':
    main()
