#!/usr/bin/env python3
"""build_qbo_map.py — turn the owner's completed Property-Mapping workbook into the
persistent qbo_map.json the exporter reads every time (so no re-mapping is ever needed).
Usage: python3 build_qbo_map.py <completed_mapping.xlsx> [qbo_map.json]"""
import sys, json, openpyxl, datetime

src = sys.argv[1]
out = sys.argv[2] if len(sys.argv) > 2 else 'qbo_map.json'
wb = openpyxl.load_workbook(src, data_only=True)
pm = wb['Property Mapping']
# cols: 1 #, 2 our property, 3 entity, 4 match, 5 suggested, 6 your QBO customer
cust_by_door = {}
for r in range(3, pm.max_row + 1):
    ours = pm.cell(r, 2).value
    filled = pm.cell(r, 6).value
    if ours and filled not in (None, ''):
        cust_by_door[str(ours).strip()] = str(filled).strip()

cfg = {
    'version': 1,
    'note': 'Persistent QBO journal-export config. Edit customerByDoor to add/adjust a property; '
            'the exporter applies it automatically. Only NEW properties ever need a line here.',
    'entityScope': ['SKY'],                       # 7 Star / Ritwik excluded (separate books)
    'locationByEntity': {'SKY': 'SKY'},
    'accounts': {
        'interest':  '709 9. Interest Paid',
        'principal': '320 1. Real Estate Mortgages',
        'escrow':    '130.04 Lender Reserves:Tax & Insurance Escrow',
        'credit':    '100.01 Bank of America',
    },
    'className': 'Rental',
    'journalNumberFormat': '{mon}-{yy} {loan}',   # e.g. May-26 10050323
    'customerByDoor': cust_by_door,
}
json.dump(cfg, open(out, 'w'), indent=1)
print('wrote %s | %d door->customer mappings' % (out, len(cust_by_door)))
