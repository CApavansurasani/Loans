#!/usr/bin/env python3
"""
publish_loans.py — Loans.xlsx → loans-data.json (amortization engine, v2)

Lender-verified conventions (validated to the penny against the Loan 8242
lender amortization schedule):
• OWNER CONVENTION: the sheet's 'Loan Start Date' is the FIRST PAYMENT date;
  origination is exactly one month earlier (engine derives it for every loan).
• OWNER CONVENTION: the sheet's 'Monthly P&I' column holds ESCROW payments
  (tax/insurance), kept separate; the note payment is computed from terms.
• Payment = annuity from note terms (P·r/(1−(1+r)^−n)), rounded half-up to cents.
• Interest is IN ARREARS: the payment on the 1st of month M carries interest
  accrued over month M−1, day-weighted 30/360. The balance path over a month
  starts at (previous close − that month's scheduled principal) and steps down
  at each event on its day (segment boundary = day − 1, 30-day month).
  First payment's interest = note × rate/12.
• SKY REI release rule (owner-specified; NOT the 120% rule): a released
  property retires at its PREVIOUS-month-end slice; the excess of the actual
  amount paid reduces surviving slices by ORIGINAL proportions renormalized
  among survivors; same-month releases share that base and are removed
  together; later months run on the revised proportions.
• Scheduled principal in a release month is borne by survivors only (the
  released slice is settled in full by its payoff); release-month interest is
  attributed by month-start (pre-release) proportions.
• Loans whose released property sits at ~0% in Allocations are ALREADY-REVISED:
  slices seed from the given proportions at the release-month close (earlier
  months are note-level only, pending original data from owner).
• 'Extra Principal' events split across current slices by current proportions
  (or hit one property if named). 'Payoff' events settle the whole note.
• Hard money / interest-only: flat balance, payment = interest, until a
  Payoff/Release row arrives.

Highlights anomalies, never modifies source data. Exits non-zero WITHOUT
writing if any blocking gate fails.
Usage: python3 publish_loans.py [--loans "My files/Loans.xlsx"] [--gl pl-data.json] [--out loans-data.json]
"""
import sys, json, argparse, datetime, calendar, re
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd

CENT = Decimal('0.01'); DZ = Decimal('0'); EPS = Decimal('1e-9')

def D(x): return Decimal(str(x))
def q2(x): return x.quantize(CENT, rounding=ROUND_HALF_UP)
def f(x): return float(x)

def fail(msgs):
    print('\n✗ PUBLISH BLOCKED — fix these and re-run:')
    for m in msgs: print('   •', m)
    sys.exit(1)

def month_str(y, m): return f'{y:04d}-{m:02d}'
def month_of(ts): return month_str(ts.year, ts.month)
def add_month(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    return month_str(y + (m == 12), m % 12 + 1)

def annuity(P0, r_m, n):
    if r_m <= 0: return q2(P0 / n)
    return q2(P0 * r_m / (1 - (1 + r_m) ** -n))

def spread(total, weights, keys):
    """Split `total` over `keys` by `weights` (Decimals), cents-quantized;
    residual cent(s) go to the largest weight so the parts sum exactly."""
    wsum = sum(weights[k] for k in keys)
    if not keys or wsum <= 0: return {}
    parts = {k: q2(total * weights[k] / wsum) for k in keys}
    resid = q2(total) - sum(parts.values())
    if resid != 0:
        parts[max(keys, key=lambda k: weights[k])] += resid
    return parts

def month_date(ym, day):
    y, mo = int(ym[:4]), int(ym[5:7])
    return datetime.date(y, mo, min(day, calendar.monthrange(y, mo)[1]))

def simulate(lid, P0, r_m, pay, io, start_m, props, orig_prop, events,
             horizon_end, gl_start, warnings, errors, start_day=1,
             day_count='30/360', term=0):
    """Run one note month-by-month. Returns (sched, door_monthly, releases_out,
    slices_at_end, close, pay_changes). All money as Decimal.
    day_count '30/360': lender-verified 8242 convention — payment on the 1st,
    interest in arrears on the prior month's day-weighted (30-day) balance path.
    day_count 'Actual/360': lender-verified Pinnacle convention — payment on the
    note's anniversary day; interest = balance × rate/360 × actual calendar days
    between payment dates, segmented at each event's exact date."""
    ev_by_m = {}
    for e in events: ev_by_m.setdefault(e['m'], []).append(e)
    end_m = max([horizon_end] + list(ev_by_m))
    pre_revised = any(e['type'] == 'Sale Release' and e['prop']
                      and orig_prop.get(e['prop'], DZ) < EPS for e in events)
    actual = (day_count == 'Actual/360')
    r_yr = r_m * 12

    slices = None
    if props and not pre_revised:
        slices = spread(P0, orig_prop, list(orig_prop))
    init_slices = dict(slices) if slices else {}
    sched, releases_out, door_monthly, pay_changes, prepays_out = [], [], {}, [], []
    close_prev, prev_avg = P0, P0
    payments_made = 0
    # Actual/360 state: exact-date accrual anchored at origination
    anchor = month_date(start_m, start_day)
    acc = DZ            # Σ balance × days since last payment
    bal_run = P0
    m = start_m
    while m <= end_m:
        mev = sorted(ev_by_m.get(m, []), key=lambda e: (e['day'], e['type']))
        interest = prin = DZ
        if actual:
            # ---- exact-date engine ----
            def advance(to):
                nonlocal acc, anchor
                if to > anchor: acc += bal_run * (to - anchor).days; anchor = to
            paying = m > start_m and close_prev > 0
            pay_date = month_date(m, start_day) if paying else None
            for e in mev:
                if paying and e['dt'] >= pay_date: continue
                advance(e['dt']); bal_run -= e['amt']
            if paying:
                advance(pay_date)
                interest = q2(acc * r_yr / 360); acc = DZ
                prin = pay - interest
                if prin < 0:
                    warnings.append(f'loan {lid}: {m} interest ${interest:,.2f} exceeds payment ${pay:,.2f} — negative amortization clamped')
                    prin = DZ
                if prin > bal_run: prin = bal_run                  # final payment
                bal_run -= prin; payments_made += 1
            for e in mev:
                if not (paying and e['dt'] >= pay_date): continue
                advance(e['dt']); bal_run -= e['amt']
            close = bal_run
        else:
            # ---- 30/360 engine (unchanged, lender-verified vs 8242) ----
            if io:
                prin = DZ; interest = None                         # set after the path
            elif m == start_m:
                interest = prin = DZ
            else:
                interest = q2(prev_avg * r_m)
                prin = pay - interest
                if prin < 0:
                    warnings.append(f'loan {lid}: {m} interest ${interest:,.2f} exceeds payment ${pay:,.2f} — negative amortization clamped')
                    prin = DZ
                if prin > close_prev: prin = close_prev            # final payment
                payments_made += 1
            bal = close_prev - prin
            avg_num, cur_b = DZ, 0
            for e in mev:
                b = min(max(e['day'] - 1, 0), 30)
                if b > cur_b: avg_num += (b - cur_b) * bal; cur_b = b
                bal -= e['amt']
            avg_num += (30 - cur_b) * bal
            close = bal
            avg = avg_num / 30
            if io:
                if m == start_m and start_day > 1:                 # prorate from funding day
                    avg -= close_prev * min(start_day - 1, 30) / 30
                interest = q2(avg * r_m)
            prev_avg = avg
        if close < D('-0.5'):
            warnings.append(f'loan {lid}: {m} balance would go ${close:,.2f} negative — clamped to 0; check payoff/release amounts')
            close = DZ
            if actual: bal_run = DZ

        # ---- slice accounting ----
        rel_this = [e for e in mev if e['type'] == 'Sale Release']
        xp_this  = [e for e in mev if e['type'] == 'Extra Principal']
        po_this  = [e for e in mev if e['type'] == 'Payoff']
        if slices is None and pre_revised and rel_this:
            live = [p for p in props if orig_prop[p] > EPS]
            slices = spread(close, orig_prop, live)
            init_slices = dict(slices)
            for e in rel_this:
                releases_out.append({'m': m, 'prop': e['prop'], 'paid': f(e['amt']),
                                     'base': None, 'excess': None, 'preRevised': True})
                door_monthly.setdefault(e['prop'], {})
            warnings.append(f'loan {lid}: already-revised allocations — slices seeded at {m} close ${close:,.2f}; earlier per-door history pending original proportions')
        elif slices is not None:
            prev_sl = dict(slices)
            holders = [p for p in props if prev_sl.get(p, DZ) > DZ]
            released = [e['prop'] for e in rel_this]
            survivors = [p for p in holders if p not in released]
            # 1) releases: slice retires at prev-month-end; excess → survivors by ORIGINAL proportions
            excess_cut = {p: DZ for p in survivors}
            osum = sum(orig_prop[p] for p in survivors)
            for e in rel_this:
                base = prev_sl.get(e['prop'], DZ)
                excess = e['amt'] - base
                rec_out = {'m': m, 'prop': e['prop'], 'paid': f(e['amt']),
                           'base': f(base), 'excess': f(excess)}
                slices[e['prop']] = DZ
                if survivors and osum > 0:
                    warnings.append(f'loan {lid}: {m} released "{e["prop"]}" — paid ${e["amt"]:,.2f} vs prev-month-end slice ${base:,.2f} → excess ${excess:,.2f} to {len(survivors)} survivor(s) by original proportions')
                    e_cuts = spread(excess, orig_prop, survivors)
                    rec_out['cuts'] = {p: f(v) for p, v in e_cuts.items()}
                    for p in survivors: excess_cut[p] += excess * orig_prop[p] / osum
                elif abs(close) <= 1:
                    warnings.append(f'loan {lid}: {m} released "{e["prop"]}" and the note settled to $0 — paid ${e["amt"]:,.2f} vs prev-month-end slice ${base:,.2f}; the month\'s scheduled payment covered the ${-excess:,.2f} difference')
                elif abs(excess) > 1:
                    warnings.append(f'loan {lid}: {m} released "{e["prop"]}" — paid ${e["amt"]:,.2f} vs prev-month-end slice ${base:,.2f}, excess ${excess:,.2f} has no surviving slices and ${close:,.2f} remains on the note — check the amount')
                releases_out.append(rec_out)
            # 2) scheduled principal borne by survivors (current proportions); blank-property
            #    extra principal likewise; a NAMED property's extra principal hits that slice only
            blank_xp = sum((e['amt'] for e in xp_this if not e['prop']), DZ)
            named_xp = {}
            for e in xp_this:
                if not e['prop']: continue
                if e['prop'] in slices:
                    named_xp[e['prop']] = named_xp.get(e['prop'], DZ) + e['amt']
                    if named_xp[e['prop']] > prev_sl.get(e['prop'], DZ) + D('0.01'):
                        errors.append(f'loan {lid}: {m} extra principal ${e["amt"]:,.2f} on "{e["prop"]}" exceeds its ${prev_sl.get(e["prop"], DZ):,.2f} slice — check amount/property')
                else:
                    errors.append(f'loan {lid}: {m} extra principal on unknown property "{e["prop"]}"')
            ssum = sum(prev_sl[p] for p in survivors)
            w_surv = {p: prev_sl[p] for p in survivors}
            prin_cut = spread(prin, w_surv, survivors) if ssum > 0 else {}
            xp_cut = spread(blank_xp, w_surv, survivors) if (ssum > 0 and blank_xp > 0) else {}
            for p in survivors:
                slices[p] = q2(prev_sl[p] - excess_cut[p] - named_xp.get(p, DZ)) - prin_cut.get(p, DZ) - xp_cut.get(p, DZ)
            if blank_xp > 0 or named_xp:
                pc = {}
                for p, v in xp_cut.items(): pc[p] = pc.get(p, DZ) + v
                for p, v in named_xp.items(): pc[p] = pc.get(p, DZ) + v
                prepays_out.append({'m': m, 'amt': f(blank_xp + sum(named_xp.values(), DZ)),
                                    'cuts': {p: f(v) for p, v in pc.items() if v > 0}})
            if po_this:
                for p in list(slices): slices[p] = DZ
                if abs(close) > 1:
                    warnings.append(f'loan {lid}: {m} Payoff leaves ${close:,.2f} on the note — check the amount')
                close = min(close, DZ) if close < 0 else close
            # exact tie-out: slices must sum to the note close
            resid = close - sum(slices.values())
            if survivors:
                if abs(resid) > D('0.25'):
                    warnings.append(f'loan {lid}: {m} slice residual ${resid:,.2f} folded into largest survivor')
                slices[max(survivors, key=lambda p: slices[p])] += resid
            elif abs(resid) > D('0.25'):
                errors.append(f'loan {lid}: {m} slices ended ${resid:,.2f} away from note balance with no survivors')
            # 3) attribution: interest by month-start proportions (released doors included)
            int_cut = spread(interest, {p: prev_sl[p] for p in holders}, holders)
            if True:  # full history — month filters and trends need every month
                for p in holders:
                    dprin = prev_sl[p] - slices.get(p, DZ)
                    door_monthly.setdefault(p, {})[m] = {
                        'interest': f(int_cut.get(p, DZ)),
                        'principal': f(q2(dprin)),
                        'close': f(slices.get(p, DZ))}
            if abs(sum(slices.values()) - close) > D('0.005'):
                errors.append(f'loan {lid}: {m} INVARIANT BROKEN — slices ${sum(slices.values()):,.2f} ≠ note ${close:,.2f}')

        # ---- EMI recast: a Notes-flagged extra principal re-levels the payment ----
        recast = [e for e in mev if e.get('recast')]
        if recast and not io:
            remaining = max(term - payments_made, 1) if term else 0
            explicit = next((e['newPay'] for e in recast if e.get('newPay')), None)
            if explicit:
                newpay = q2(explicit)
                if term and abs(newpay - annuity(close, r_m, remaining)) > 5:
                    warnings.append(f'loan {lid}: {m} Notes give new EMI ${newpay:,.2f} but re-amortizing ${close:,.2f} over the remaining {remaining} months gives ${annuity(close, r_m, remaining):,.2f} — using the Notes figure')
            elif term:
                newpay = annuity(close, r_m, remaining)
            else:
                newpay = pay
                errors.append(f'loan {lid}: {m} EMI-reduce flagged but no Term and no new EMI in Notes — cannot recompute the payment')
            if newpay != pay:
                pay_changes.append({'m': m, 'from': f(pay), 'to': f(newpay)})
                warnings.append(f'loan {lid}: {m} extra principal recasts the EMI ${pay:,.2f} → ${newpay:,.2f} (tenure preserved)')
                pay = newpay

        sched.append({'m': m, 'open': f(close_prev), 'interest': f(interest),
                      'principal': f(prin), 'events': f(sum(e['amt'] for e in mev)),
                      'close': f(close)})
        close_prev = close
        if close <= D('0.005') and not any(k > m for k in ev_by_m): break
        m = add_month(m)
    return sched, door_monthly, releases_out, slices, close_prev, pay_changes, init_slices, prepays_out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--loans', default='My files/Loans.xlsx')
    ap.add_argument('--gl', default='pl-data.json')
    ap.add_argument('--out', default='loans-data.json')
    a = ap.parse_args()
    errors, warnings = [], []

    gl = json.load(open(a.gl))
    GLP = gl['periods']; HORIZON_END = GLP[-1]

    loans = pd.read_excel(a.loans, sheet_name='Loans')
    # 'None'/'NA' are pandas default NA markers — re-read the penalty column verbatim so owner text survives
    try:
        _pp = pd.read_excel(a.loans, sheet_name='Loans', usecols=['Loan ID', 'Prepay Penalty'], keep_default_na=False)
        if 'Prepay Penalty' in loans.columns:
            loans['Prepay Penalty'] = _pp['Prepay Penalty'].astype(str).str.strip().replace({'': None})
    except ValueError:
        pass  # column absent — stays optional
    allo = pd.read_excel(a.loans, sheet_name='Allocations')
    rel = pd.read_excel(a.loans, sheet_name='Releases')
    for df in (loans, allo, rel):
        first = df.columns[0]
        df.drop(df[df[first].astype(str).str.startswith('EXAMPLE')].index, inplace=True)
        df.dropna(how='all', inplace=True)
    def canon_id(v):
        if pd.isna(v): return ''
        if isinstance(v, float) and v == int(v): return str(int(v))
        return str(v).strip()
    for df in (loans, allo, rel):
        df['Loan ID'] = df['Loan ID'].map(canon_id)
    allo['Property (exact index name)'] = allo['Property (exact index name)'].astype(str).str.strip()

    # ---- duplicate or blank loan IDs: split allocations against each note ----
    counts = loans['Loan ID'].value_counts()
    def shared(lid): return lid == '' or counts.get(lid, 0) > 1
    loan_rows = [(f'{lid}#{i}' if shared(lid) else lid, lid, L)
                 for i, (_, L) in enumerate(loans.iterrows())
                 for lid in [L['Loan ID']]]
    n_blank = sum(1 for _, lid, _ in loan_rows if lid == '')
    if n_blank:
        warnings.append(f'{n_blank} loan(s) have blank Loan ID (matched to allocations by amount) — fill in loan numbers when available')
    alloc_by_key = {}
    for key, lid, L in loan_rows:
        rows = allo[allo['Loan ID'] == lid]
        if not shared(lid):
            alloc_by_key[key] = rows; continue
        target = float(L['Original Loan Amount'])
        # 1) a single allocation row equal to the note amount (hard money 1:1)
        single = [j for j, r in rows.iterrows() if abs(float(r['Allocated Amount']) - target) <= 1]
        if len(single) >= 1:
            taken = single[:1]
        else:
            # 2) order-preserving contiguous accumulation (multi-property notes)
            taken, s = [], 0.0
            for j, r in rows.iterrows():
                taken.append(j); s += float(r['Allocated Amount'])
                if abs(s - target) <= 1: break
            if abs(s - target) > 1:
                errors.append(f'loan ID "{lid}" (shared/blank): cannot match allocations to the ${target:,.2f} note — rename the notes distinctly (e.g. {lid}-P / {lid}-R)')
                alloc_by_key[key] = rows.iloc[0:0]; continue
        alloc_by_key[key] = rows.loc[taken]
        allo = allo.drop(index=taken)
        if lid and counts.get(lid, 0) > 1:
            w = f'duplicate Loan ID "{lid}" auto-split by allocation sums — please rename the notes distinctly in Loans.xlsx'
            if w not in warnings: warnings.append(w)

    out_loans, doors_out = [], {}
    for key, lid, L in loan_rows:
        lender = str(L.get('Lender', '') or '')
        ltype = str(L['Loan Type']).strip()
        try:
            first_pay = pd.Timestamp(L['Loan Start Date'])
        except Exception:
            errors.append(f'loan {lid}: bad start date {L["Loan Start Date"]!r}'); continue
        # owner convention: entered date = FIRST PAYMENT; origination is one month earlier
        start = first_pay - pd.DateOffset(months=1)
        rate = float(L['Interest Rate %'])
        rate = rate / 100.0 if rate >= 1 else rate
        r_m = D(rate) / 12
        P0 = q2(D(float(L['Original Loan Amount'])))
        if P0 <= 0:
            errors.append(f'loan {lid}: non-positive Original Loan Amount'); continue
        term = int(L['Term (months)']) if pd.notna(L['Term (months)']) and float(L['Term (months)']) > 0 else 0
        io = str(L.get('Interest Only (Y/N)', 'N')).strip().upper() == 'Y' or ltype == 'Hard Money'
        escrow = D(float(L['Monthly P&I'])) if pd.notna(L['Monthly P&I']) else DZ   # owner: this column = escrow
        # optional per-note columns (blank = defaults)
        ppv = L.get('Prepay Penalty')
        pp_raw = '' if (ppv is None or (isinstance(ppv, float) and pd.isna(ppv))) else str(ppv).strip()
        prepay = {'raw': pp_raw, 'known': False, 'steps': [], 'freeMonths': 0}
        if pp_raw:
            txt = pp_raw.lower().replace('%', '').strip()
            free_m = 0
            mmod = re.search(r'\(\s*till\s+(\d+)\s+months?\s+before\s+maturity\s*\)', txt)
            if mmod:
                free_m = int(mmod.group(1))
                txt = (txt[:mmod.start()] + txt[mmod.end():]).strip()
            if txt in ('0', 'none', 'no', 'nil', 'na', 'n/a'):
                prepay = {'raw': pp_raw, 'known': True, 'steps': [], 'freeMonths': 0}
            elif txt in ('pending', 'tbd', '?'):
                prepay = {'raw': pp_raw, 'known': False, 'steps': [], 'freeMonths': 0}
            else:
                try:
                    steps = [float(x) for x in re.split(r'[,/]', txt) if x.strip() != '']
                    if not steps or any(x < 0 or x > 25 for x in steps): raise ValueError
                    prepay = {'raw': pp_raw, 'known': True, 'steps': steps, 'freeMonths': free_m}
                except ValueError:
                    errors.append(f'loan {lid}: Prepay Penalty {pp_raw!r} not understood — use step-down "5/4/3/2/1" (or commas), flat "3", "None", "NA", "Pending"; optional suffix "(till N months before Maturity)"')
        pov = L.get('Payment Override')
        pay_override = q2(D(float(pov))) if pd.notna(pov) and float(pov) > 0 else None
        dcv = L.get('Day Count')
        dc_raw = '' if pd.isna(dcv) else str(dcv).strip().lower().replace(' ', '')
        day_count = 'Actual/360' if dc_raw in ('actual/360', 'act/360', 'actual360') else '30/360'
        if dc_raw and day_count == '30/360' and dc_raw not in ('30/360', '30e/360', '360/360'):
            errors.append(f'loan {lid}: unknown Day Count {L.get("Day Count")!r} (use 30/360 or Actual/360)'); continue
        if io and day_count != '30/360':
            errors.append(f'loan {lid}: Actual/360 with interest-only is not supported — talk to Claude before publishing'); continue
        pay = DZ if io else (pay_override or (annuity(P0, r_m, term) if term else DZ))
        if not io and pay <= 0:
            errors.append(f'loan {lid}: no Term (months) and no Payment Override — cannot compute the note payment (P&I column is escrow per owner rule)'); continue
        if not io and pay_override and term:
            calc = annuity(P0, r_m, term)
            if abs(pay_override - calc) > 25:
                warnings.append(f'loan {lid}: Payment Override ${pay_override:,.2f} is far from the ${calc:,.2f} annuity — double-check the terms')

        arows = alloc_by_key[key]
        props = arows['Property (exact index name)'].tolist()
        fr = [D(x) for x in arows['Proportion %'].astype(float).tolist()]
        tot = sum(fr)
        if tot > 2: fr = [x / 100 for x in fr]; tot = sum(fr)
        if props and abs(tot - 1) > D('0.005'):
            errors.append(f'loan {lid}: proportions sum to {float(tot)*100:.2f}% (must be 100%)'); continue
        orig_prop = {p: x for p, x in zip(props, fr)}          # renormalized exactly at use time
        amts = [float(x) for x in arows['Allocated Amount'].astype(float).tolist()]
        if props and abs(sum(amts) - float(P0)) > 1:
            warnings.append(f'loan {lid}: allocated amounts sum ${sum(amts):,.2f} ≠ note ${P0:,.2f} — proportions govern (owner rule: trust original proportions)')

        # releases & extra principal for this note
        ev = rel[rel['Loan ID'] == lid].copy()
        if shared(lid) and len(ev):
            ev = ev[ev['Property (exact index name)'].astype(str).str.strip().isin(props) |
                    ev['Property (exact index name)'].isna()]
        events = []
        for _, R in ev.iterrows():
            try:
                d = pd.Timestamp(R['Date'])
            except Exception:
                errors.append(f'loan {lid}: bad release date {R["Date"]!r}'); continue
            pr = str(R['Property (exact index name)']).strip() if pd.notna(R['Property (exact index name)']) else None
            typ = str(R['Type']).strip()
            amt = D(float(R['Amount Paid']))
            if typ not in ('Sale Release', 'Extra Principal', 'Payoff'):
                errors.append(f'loan {lid}: unknown release Type {typ!r} (use Sale Release / Extra Principal / Payoff)'); continue
            if typ == 'Sale Release' and not pr:
                errors.append(f'loan {lid}: Sale Release on {d.date()} has no property'); continue
            if pr and typ == 'Sale Release' and pr not in props:
                errors.append(f'loan {lid}: release property "{pr}" not in its allocations'); continue
            if amt <= 0:
                errors.append(f'loan {lid}: non-positive Amount Paid on {d.date()}'); continue
            if month_of(d) < month_str(start.year, start.month):
                errors.append(f'loan {lid}: event on {d.date()} predates loan start'); continue
            note = str(R.get('Notes') or '')
            recast = bool(re.search(r'(?:reduce|revise|recast|lower|new)[\s\w]{0,24}(?:emi|payment)|(?:emi|payment)[\s\w]{0,24}(?:reduce|revise|recast|lower)', note, re.I))
            if recast and typ != 'Extra Principal':
                errors.append(f'loan {lid}: EMI-reduce note on a {typ!r} row ({d.date()}) — recast applies to Extra Principal rows only'); continue
            new_pay = None
            if recast:
                mnum = re.search(r'(?:emi|payment)[^\d]{0,20}([\d,]+(?:\.\d{1,2})?)', note, re.I)
                if mnum:
                    try: new_pay = D(mnum.group(1).replace(',', ''))
                    except Exception: new_pay = None
            events.append({'m': month_of(d), 'day': int(d.day), 'dt': datetime.date(d.year, d.month, d.day),
                           'prop': pr, 'type': typ, 'amt': amt, 'recast': recast, 'newPay': new_pay})
        events.sort(key=lambda e: (e['m'], e['day']))

        start_m = month_str(start.year, start.month)
        # OWNER (2026-07): the month axis runs to the LAST NOTE DATE — project every note to
        # its own maturity (term months past origination) instead of stopping at the GL close.
        note_horizon = HORIZON_END
        if term:
            mat_m = start_m
            for _ in range(term): mat_m = add_month(mat_m)
            note_horizon = max(note_horizon, mat_m)
        sched, dm, rel_out, slices, _out_final, pay_changes, init_slices, prepays = simulate(
            lid, P0, r_m, pay, io, start_m, props, orig_prop,
            events, note_horizon, GLP[0], warnings, errors, start_day=int(start.day),
            day_count=day_count, term=term)
        # Reporting anchor stays at the GL close: 'outstanding' = balance at HORIZON_END (glEnd),
        # NOT the projected maturity balance — preserves the BS cross-page tie and the golden
        # portfolio total to the cent even though the schedule now runs years past glEnd.
        outstanding = DZ
        for r in sched:
            if r['m'] <= HORIZON_END: outstanding = D(str(r['close']))
            else: break
        if not sched:      # fully past horizon with no rows at all (defensive)
            outstanding = DZ
            slices, init_slices = {}, {}
        elif start_m > HORIZON_END:
            warnings.append(f'loan {lid or (props[0] if props else "?")}: originates {start_m}, after the GL close {HORIZON_END} — $0 outstanding now, projected forward on the axis')

        if io and term:
            mat = month_str(start.year, start.month)
            for _ in range(term): mat = add_month(mat)
            if outstanding > 0 and HORIZON_END >= mat:
                warnings.append(f'loan {lid} ({ltype}): past its {mat} maturity with ${outstanding:,.2f} outstanding — add a Payoff/extension row when known')

        # fold per-door months into global map (a door can carry several notes)
        for p in set(list(dm.keys()) + list(init_slices.keys())):
            rec = doors_out.setdefault(p, {'loans': [], 'latest': 0.0, 'initial': 0.0, 'monthly': {}})
            if key not in rec['loans']: rec['loans'].append(key)
            for mth, v in dm.get(p, {}).items():
                t = rec['monthly'].setdefault(mth, {'interest': 0.0, 'principal': 0.0, 'close': 0.0})
                for k2 in ('interest', 'principal', 'close'):
                    t[k2] = round(t[k2] + v[k2], 2)
            rec['initial'] = round(rec['initial'] + f(init_slices.get(p, DZ)), 2)

        display = lid if lid else (props[0] if props else f'HM-{key}')
        out_loans.append({'id': key, 'displayId': display, 'lender': lender, 'type': ltype,
                          'rate': rate, 'start': month_str(start.year, start.month),
                          'firstPay': month_str(first_pay.year, first_pay.month),
                          'escrow': f(escrow), 'term': term,
                          'dayCount': day_count, 'payOverride': bool(pay_override),
                          'prepay': prepay,
                          'paymentChanges': pay_changes,
                          'orig': f(P0),
                          'payment': (pay_changes[-1]['to'] if pay_changes else f(pay)),
                          'io': io,
                          'preRevised': any(r.get('preRevised') for r in rel_out),
                          'releases': rel_out, 'prepays': prepays, 'schedule': sched,
                          'outstanding': f(outstanding),
                          'origProps': {p: float(orig_prop[p]) for p in props}})

    # door 'latest' = balance at the GL close (glEnd), read from the folded monthly path so it
    # stays the reported current position even though schedules now project to maturity.
    for _p, _rec in doors_out.items():
        _rec['latest'] = round(_rec['monthly'].get(HORIZON_END, {}).get('close', 0.0), 2)

    # allocation properties must exist in the Property Index (via GL meta.doorWindows)
    idx_doors = set(gl.get('meta', {}).get('doorWindows', {}))
    if idx_doors:
        miss = sorted({p for p in doors_out if p not in idx_doors})
        if miss:
            warnings.append(f'{len(miss)} allocation propert(ies) not found in the Property Index — align names or add them: ' + ', '.join(miss))

    # doors on multiple live notes at once → highlight (possible missing payoff row)
    for p, rec in doors_out.items():
        if len(rec['loans']) > 1:
            per_loan = [l for l in out_loans if l['id'] in rec['loans'] and l['outstanding'] > 0.005]
            if len(per_loan) > 1:
                warnings.append(f'door "{p}" carries balances on {len(per_loan)} notes at once ({", ".join(x["displayId"] for x in per_loan)}) — if one was refinanced/paid off, add its Payoff row')

    if errors: fail(sorted(set(errors)))

    # ---- GL interest reconciliation (highlight only, never a gate) ----
    gl_int = {}
    for e, d, acct, vals in gl['rows']:
        if acct == 'Interest Cost':
            for i, v in enumerate(vals):
                if abs(v) > 0.005:
                    gl_int.setdefault(d, {})[GLP[i]] = round(gl_int.get(d, {}).get(GLP[i], 0) + v, 2)
    tot_diff, n_comp, worst = 0.0, 0, []
    for d, rec in doors_out.items():
        for mth, mm in rec['monthly'].items():
            gv = gl_int.get(d, {}).get(mth)
            if gv is None: continue
            diff = round(mm['interest'] - gv, 2)
            tot_diff += abs(diff); n_comp += 1
            if abs(diff) > 25: worst.append((abs(diff), d, mth, mm['interest'], gv))
    worst.sort(reverse=True)

    tot_out = round(sum(l['outstanding'] for l in out_loans), 2)
    by_type = {}
    for l in out_loans:
        by_type[l['type']] = round(by_type.get(l['type'], 0) + l['outstanding'], 2)
    # debt service / escrow measured at the GL close (glEnd), consistent with 'outstanding'
    def _row_at_glend(l):
        r = None
        for x in l['schedule']:
            if x['m'] <= HORIZON_END: r = x
            else: break
        return r
    active = [l for l in out_loans if l['outstanding'] > 0.005 and l['schedule']]
    esc_total = round(sum(l['escrow'] for l in active), 2)
    def _svc(l):
        if not l['io']: return l['payment']
        r = _row_at_glend(l); return r['interest'] if r else 0.0
    svc_total = round(sum(_svc(l) for l in active), 2)

    # projection axis: schedules now run to each note's maturity. asOf = present month (clamped
    # to the axis) is the page's default reporting month per the owner's 2026-07 spec.
    sched_end = HORIZON_END
    for l in out_loans:
        if l['schedule']: sched_end = max(sched_end, l['schedule'][-1]['m'])
    now_m = datetime.datetime.now().strftime('%Y-%m')
    as_of = min(max(now_m, GLP[0]), sched_end)

    out = {'schema': 'loans.v2',
           'meta': {'published': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'convention': 'sheet start date = first payment (origination one month earlier); interest-in-arrears day-weighted 30/360; release at prev-month-end slice, excess to survivors by original proportions; P&I column = escrow',
                    'horizonEnd': HORIZON_END, 'glEnd': HORIZON_END, 'asOf': as_of,
                    'scheduleEnd': sched_end, 'glWindow': [GLP[0], GLP[-1]],
                    'projection': f'actuals + GL reconciliation through {HORIZON_END}; months after {HORIZON_END} (up to {sched_end}) are scheduled amortization projections — no new releases, prepays or GL',
                    'loans': len(out_loans), 'totalOutstanding': tot_out, 'byType': by_type,
                    'monthlyDebtService': svc_total, 'monthlyEscrow': esc_total,
                    'reconciliation': {'comparisons': n_comp, 'totalAbsDiff': round(tot_diff, 2),
                                       'flagged': len(worst)}},
           'loans': out_loans,
           'doors': doors_out}
    json.dump(out, open(a.out, 'w'), separators=(',', ':'))

    print(f'loans: {len(out_loans)} | outstanding at {HORIZON_END}: ${tot_out:,.2f} | by type: ' +
          ', '.join(f'{k} ${v:,.0f}' for k, v in sorted(by_type.items())))
    print(f'axis: schedules project to {sched_end} | page default asOf = {as_of} (present month)')
    print(f'GL interest reconciliation: {n_comp} door-months · total |diff| ${tot_diff:,.2f} · {len(worst)} beyond $25')
    for w in worst[:12]:
        print(f'   ⚠ {w[1]} {w[2]}: engine ${w[3]:,.2f} vs GL ${w[4]:,.2f} (Δ ${w[0]:,.2f})')
    for w in warnings: print('  ⚠', w)
    print(f'✓ wrote {a.out}')

if __name__ == '__main__':
    main()
