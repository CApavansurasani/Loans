#!/usr/bin/env python3
"""qbo_export.py — SKY REI loan payments -> QuickBooks (DataDear) journal import.

Builds one balanced journal per note per month for the selected period, filling a copy of
the DataDear "Template for Export" sheet. Reads the persistent qbo_map.json so property names
map to the exact QBO customers every time (no re-mapping).

Per note / month:
  Interest   -> debit  709 9. Interest Paid           (one line per property, Customer = mapped)
  Principal  -> debit  320 1. Real Estate Mortgages   (SCHEDULED principal only; events excluded)
  Escrow     -> debit  130.04 ... Tax & Insurance Escrow   (one line = note escrow x SKY share)
  Credit     -> 100.01 Bank of America                (= total payment incl. escrow)
Journal Number = "MMM-YY loan#" on the first line; Entry Date = 1st of the month.

Gates (never emit a misplaced or unbalanced line): every in-scope door must have a mapping,
every Customer must exist in the QBO customer list, each journal must balance to the cent,
and each note's interest/principal must tie to its schedule.
Usage: python3 qbo_export.py --months 2026-05[,2026-06] [--out FILE] [--template T] [--map qbo_map.json]
"""
import sys, json, argparse, datetime, shutil, openpyxl
from decimal import Decimal, ROUND_HALF_UP

CENT = Decimal('0.01')
def q2(x): return float(Decimal(str(x)).quantize(CENT, rounding=ROUND_HALF_UP))
MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--loans', default='loans-data.json')
    ap.add_argument('--gl', default='pl-data.json')
    ap.add_argument('--map', default='qbo_map.json')
    ap.add_argument('--template', required=True)
    ap.add_argument('--months', required=True, help='comma list, e.g. 2026-05,2026-06')
    ap.add_argument('--out', default='qbo_journal.xlsx')
    a = ap.parse_args()

    d = json.load(open(a.loans)); gl = json.load(open(a.gl)); cfg = json.load(open(a.map))
    DW = gl.get('meta', {}).get('doorWindows', {})
    DOORS = d['doors']
    def entOf(x): return (DW.get(x, [None, None, '—'])[2]) if x in DW else '—'
    cust_by_door = cfg['customerByDoor']
    scope = set(cfg['entityScope'])
    locmap = cfg['locationByEntity']; ACC = cfg['accounts']; CLS = cfg['className']
    months = [m.strip() for m in a.months.split(',') if m.strip()]

    # QBO customer master (from the template's DataDear sheet) — validate every emitted name
    tw = openpyxl.load_workbook(a.template, data_only=True)
    dd = tw['DataDear']
    qbo_customers = set()
    ci = openpyxl.utils.column_index_from_string('Q')
    for r in range(1, dd.max_row + 1):
        v = dd.cell(r, ci).value
        if v not in (None, ''): qbo_customers.add(str(v).strip())

    def note_row(l, m): return next((r for r in l['schedule'] if r['m'] == m), None)

    errors, warns = [], []
    journals = []            # each = dict(jn, date, lines=[dict(acct,cust,loc,dr,cr)])

    for m in months:
        y, mo = int(m[:4]), int(m[5:7])
        jdate = datetime.datetime(y, mo, 1)
        jtag = MON[mo-1] + '-' + m[2:4]
        for l in d['loans']:
            props_all = list(l.get('origProps', {}).keys())
            props_sky = [p for p in props_all if entOf(p) in scope]
            if not props_sky:
                continue
            r = note_row(l, m)
            if not r:
                continue
            # active SKY doors this month (released doors have no monthly row -> excluded)
            active = [p for p in props_sky if m in DOORS.get(p, {}).get('monthly', {})]
            note_int = r['interest']; note_prin = r['principal']          # scheduled, events excluded
            if q2(note_int) == 0 and q2(note_prin) == 0 and q2(l['escrow']) == 0:
                continue
            event = abs(r['events']) > 0.005
            multinote = any(len(DOORS[p]['loans']) > 1 for p in active)
            clean = (not event) and (not multinote)
            if not active:
                # note has schedule activity but all SKY doors released/absent this month
                if q2(note_int) != 0 or q2(note_prin) != 0:
                    warns.append(f'{l["displayId"]} {m}: has P&I but no active SKY door this month — skipped')
                continue

            lines = []
            di = {}; dp = {}
            if clean:
                for p in active:
                    di[p] = DOORS[p]['monthly'][m]['interest']
                    dp[p] = DOORS[p]['monthly'][m]['principal']     # event-free -> scheduled
            else:
                # allocate straight from the note schedule by ORIGINAL proportion (absolute share)
                for p in active:
                    w = l['origProps'][p]
                    di[p] = q2(note_int * w); dp[p] = q2(note_prin * w)
                warns.append(f'{l["displayId"]} {m}: {"mixed-entity" if not event else "event month"}'
                             f'{"/multi-note door" if multinote else ""} — split by original proportion')

            sky_share = sum(l['origProps'][p] for p in active)
            escrow_sky = q2(l['escrow'] * sky_share) if l['escrow'] > 0.005 else 0.0

            tot = 0.0
            for p in active:                                    # interest lines
                cust = cust_by_door.get(p)
                if cust is None:
                    errors.append(f'{m}: door "{p}" (note {l["displayId"]}) has NO QBO customer mapping'); continue
                if cust not in qbo_customers:
                    errors.append(f'{m}: mapped customer "{cust}" for door "{p}" is not in the QBO customer list')
                if q2(di[p]) != 0:
                    lines.append(dict(acct=ACC['interest'], cust=cust, loc=locmap.get(entOf(p), entOf(p)), dr=q2(di[p]), cr=None))
                    tot += q2(di[p])
            for p in active:                                    # principal lines
                cust = cust_by_door.get(p)
                if cust is None: continue
                if q2(dp[p]) != 0:
                    lines.append(dict(acct=ACC['principal'], cust=cust, loc=locmap.get(entOf(p), entOf(p)), dr=q2(dp[p]), cr=None))
                    tot += q2(dp[p])
            if escrow_sky > 0:
                lines.append(dict(acct=ACC['escrow'], cust='', loc=locmap.get('SKY', 'SKY'), dr=escrow_sky, cr=None))
                tot += escrow_sky
            if not lines:
                continue
            lines.append(dict(acct=ACC['credit'], cust='', loc=locmap.get('SKY', 'SKY'), dr=None, cr=q2(tot)))
            # balance gate
            drs = sum(x['dr'] for x in lines if x['dr'] is not None)
            crs = sum(x['cr'] for x in lines if x['cr'] is not None)
            if abs(drs - crs) > 0.005:
                errors.append(f'{l["displayId"]} {m}: journal does not balance ({drs:.2f} vs {crs:.2f})')
            jn = cfg['journalNumberFormat'].format(mon=MON[mo-1], yy=m[2:4], loan=l['displayId'])
            journals.append(dict(jn=jn, date=jdate, lines=lines))

    # ---- independent accuracy gate: emitted totals vs the engine's own door figures ----
    exp_int = sum(x['dr'] for j in journals for x in j['lines'] if x['acct'] == ACC['interest'])
    exp_prin = sum(x['dr'] for j in journals for x in j['lines'] if x['acct'] == ACC['principal'])
    ref_int = ref_prin_raw = evt = 0.0
    sky_all = [p for p in DOORS if entOf(p) in scope]
    for m in months:
        for p in sky_all:
            mm = DOORS[p]['monthly'].get(m)
            if mm: ref_int += mm['interest']; ref_prin_raw += mm['principal']
        for l in d['loans']:
            if any(entOf(p) in scope for p in l.get('origProps', {})):
                r = note_row(l, m)
                if r and abs(r['events']) > 0.005:
                    sh = sum(l['origProps'][p] for p in l['origProps'] if entOf(p) in scope and m in DOORS[p]['monthly'])
                    evt += r['events'] * sh
    ref_int = q2(ref_int); ref_prin_sched = q2(ref_prin_raw - evt)
    if abs(exp_int - ref_int) > 0.05 or abs(exp_prin - ref_prin_sched) > 0.10:
        errors.append(f'ACCURACY GATE: interest export {exp_int:.2f} vs engine {ref_int:.2f}; '
                      f'principal export {exp_prin:.2f} vs engine-scheduled {ref_prin_sched:.2f}')

    if errors:
        print('✗ EXPORT BLOCKED:')
        for e in sorted(set(errors))[:40]: print('   •', e)
        sys.exit(1)

    # ---- write into a copy of the template ----
    shutil.copy(a.template, a.out)
    wb = openpyxl.load_workbook(a.out)
    ws = wb['Template for Export']
    # columns: A Post? B JournalNumber C EntryDate D Desc E Account F Customer G Vendor H Employee
    #          I Location J Class K TaxRate L TaxApp M Currency N Debit O Credit
    START = 11    # data begins row 11
    # examples aren't needed on export: blank rows 7-10 (keeps their formatting, e.g. the grey row 8)
    # and clear the data region below. Header row 6 is left untouched.
    for r in range(7, ws.max_row + 1):
        for c in range(1, 18): ws.cell(r, c).value = None
    row = START
    tot_dr = tot_cr = 0.0
    for j in journals:
        for ln in j['lines']:
            ws.cell(row, 2, j['jn'])                                            # Journal Number on EVERY row
            dt = ws.cell(row, 3, j['date']); dt.number_format = 'MM/DD/YYYY'    # Entry Date on every row
            ws.cell(row, 5, ln['acct'])
            if ln['cust']: ws.cell(row, 6, ln['cust'])
            if ln['loc']:  ws.cell(row, 9, ln['loc'])
            ws.cell(row, 10, CLS)
            if ln['dr'] is not None:
                c = ws.cell(row, 14, ln['dr']); c.number_format = '#,##0.00'; tot_dr += ln['dr']
            if ln['cr'] is not None:
                c = ws.cell(row, 15, ln['cr']); c.number_format = '#,##0.00'; tot_cr += ln['cr']
            row += 1
    wb.save(a.out)

    print('✓ wrote %s' % a.out)
    print('  period: %s | journals: %d | data rows: %d' % (','.join(months), len(journals), row - START))
    print('  total debit  {:,.2f}'.format(tot_dr))
    print('  total credit {:,.2f}'.format(tot_cr))
    print('  balanced: %s' % (abs(tot_dr - tot_cr) < 0.005))
    print('  interest  export {:,.2f} = engine {:,.2f}  TIE'.format(exp_int, ref_int))
    print('  principal export {:,.2f} = engine-scheduled {:,.2f}  TIE  (events excluded {:,.2f})'.format(exp_prin, ref_prin_sched, evt))
    for w in warns[:20]: print('  ⚠', w)

if __name__ == '__main__':
    main()
