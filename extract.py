# -*- coding: utf-8 -*-
import json, os, time
from datetime import datetime, timedelta
import win32com.client as win32

SRC = 'MIS FINCART FY- 2026-2027_29 Jun-26 -.xlsb'

FY_SERIALS = [46113,46143,46174,46204,46235,46266,46296,46327,46357,46388,46419,46447]
def ser2month(n):
    return (datetime(1899,12,30)+timedelta(days=float(n))).strftime('%b-%Y')
FY_MONTHS = [ser2month(s) for s in FY_SERIALS]

def num(v):
    if v is None: return 0.0
    if isinstance(v,(int,float)): return float(v)
    sv = str(v).strip().replace(',','')
    if sv == '': return 0.0
    try: return float(sv)
    except: return 0.0

def s(v):
    if v is None: return ''
    if hasattr(v,'hour'): return ''          # datetime/pywintypes time cell
    if isinstance(v,float) and v==int(v): return str(int(v))
    return str(v).strip()

def hms_to_hours(v):
    if v is None: return 0.0
    if hasattr(v,'hour'):                    # datetime or pywintypes.datetime
        return v.hour + v.minute/60.0 + v.second/3600.0
    if isinstance(v,(int,float)): return float(v)*24.0
    t = str(v).strip()
    if t=='' or t.lower()=='none': return 0.0
    parts = t.split(':')
    try:
        parts=[float(p) for p in parts]
        while len(parts)<3: parts.append(0)
        return parts[0]+parts[1]/60.0+parts[2]/3600.0
    except: return 0.0

def sheet_data(wb, name, ncols, last_col=1):
    """Read only actual data rows × ncols columns — avoids UsedRange memory blow-up."""
    ws = wb.Sheets(name)
    # xlUp = -4162: find last non-empty row from the bottom in `last_col`
    last_row = ws.Cells(ws.Rows.Count, last_col).End(-4162).Row
    if last_row < 1: return []
    data = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, ncols)).Value
    if data is None: return []
    if not isinstance(data, tuple): return [(data,)]     # single cell
    if data and not isinstance(data[0], tuple): return (data,)  # single row
    return data

t0 = time.time()
print(f'Opening {SRC} ...', flush=True)

xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False

wb = xl.Workbooks.Open(os.path.abspath(SRC), UpdateLinks=False, ReadOnly=True)
print(f'  opened in {time.time()-t0:.1f}s', flush=True)

try:
    # -------- INS --------
    print('Reading INS...', end=' ', flush=True)
    ins=[]
    for i,row in enumerate(sheet_data(wb,'INS', ncols=65)):
        if i==0: continue
        def g(ix): return row[ix] if ix<len(row) else None
        if g(0) in (None,''): continue
        if s(g(1))!='B2C': continue
        mo=s(g(48))
        if mo and mo not in FY_MONTHS: continue      # exclude pre-FY rows
        ins.append({
            'dtype':s(g(1)),'rm':s(g(2)),'client':s(g(3)),
            'platform':s(g(8)) or 'Unknown',
            'category':s(g(10)) or 'Unknown',
            'insType':s(g(11)) or 'Unknown',
            'partner':s(g(14)) or 'Unknown',
            'premium':num(g(27)),'month':s(g(48)),'confirmed':s(g(50)),
            'type':s(g(51)) or 'Unknown','revenue':num(g(52)),
            'pct':num(g(54)),'teamLeader':s(g(56)),
            'status':s(g(63)) or 'Unknown',
        })
    print(f'{len(ins)} rows  ({time.time()-t0:.1f}s)', flush=True)

    # -------- FEES --------
    print('Reading FEES...', end=' ', flush=True)
    fees=[]
    for i,row in enumerate(sheet_data(wb,'FEES', ncols=35)):
        if i==0: continue
        def g(ix): return row[ix] if ix<len(row) else None
        if g(0) in (None,''): continue
        if s(g(1))!='B2C': continue
        fees.append({
            'dtype':s(g(1)),'rm':s(g(2)),'client':s(g(3)),
            'platform':s(g(8)) or 'Unknown',
            'category':s(g(10)) or 'Unknown',
            'asset':s(g(11)) or 'Unknown',
            'prodType':s(g(12)) or 'Unknown',
            'amount':num(g(14)),
            'month':(s(g(29)).title() if s(g(29)) else ''),
            'transType':s(g(31)) or 'Unknown',
            'confirmed':s(g(32)),'revenue':num(g(33)),
        })
    print(f'{len(fees)} rows  ({time.time()-t0:.1f}s)', flush=True)

    # -------- PMS --------
    print('Reading PMS...', end=' ', flush=True)
    pms_rows=[]
    for i,row in enumerate(sheet_data(wb,'PMS', ncols=35)):
        if i==0: continue
        def g(ix): return row[ix] if ix<len(row) else None
        if g(0) in (None,''): continue
        if s(g(1))!='B2C': continue
        raw_mo=g(28)
        mo=s(raw_mo).title() if raw_mo else ''
        pms_rows.append({
            'dtype':s(g(1)),'rm':s(g(2)),'client':s(g(3)),
            'platform':s(g(8)) or 'Unknown',
            'category':s(g(10)) or 'Unknown',
            'asset':s(g(11)) or 'Unknown',
            'scheme':s(g(12)) or 'Unknown',
            'punchAmount':num(g(16)),'month':mo,
            'confirmed':s(g(32)) or 'Unknown','revenue':num(g(33)),
        })
    print(f'{len(pms_rows)} rows  ({time.time()-t0:.1f}s)', flush=True)

    # -------- Talk --------
    GROUPS={'talktime':6,'netSale':21,'sip':35,'pms':50,'clients':64}
    print('Reading Talk...', end=' ', flush=True)
    talk_rows=[]
    for i,row in enumerate(sheet_data(wb,'Talk', ncols=77, last_col=3)):
        if i<2: continue
        def g(ix): return row[ix] if ix<len(row) else None
        if (g(0) in (None,'')) and (g(2) in (None,'')): continue
        rec={'empCode':s(g(0)),'team':s(g(1)) or 'Unassigned','name':s(g(2)),'level':s(g(4)),'role':s(g(5))}
        for key,start in GROUPS.items():
            arr=[]
            for m in range(12):
                cell=g(start+m)
                arr.append(round(hms_to_hours(cell),4) if key=='talktime' else num(cell))
            rec[key]=arr
        talk_rows.append(rec)
    print(f'{len(talk_rows)} rows  ({time.time()-t0:.1f}s)', flush=True)

    # -------- REVENUE SUMMARY V 2.0 --------
    print('Reading REVENUE SUMMARY...', end=' ', flush=True)
    revsum_rows=[]
    revsum_mtd_month=''
    for i,row in enumerate(sheet_data(wb,'REVENUE SUMMARY V 2.0', ncols=37)):
        def g(ix): return row[ix] if ix<len(row) else None
        if i==0: revsum_mtd_month=s(g(3)); continue
        if i<3: continue
        if not isinstance(g(0),(int,float)): continue
        nm=s(g(3))
        if nm=='' or nm.lower().startswith('pool'): continue
        revsum_rows.append({
            'sn':num(g(0)),'empCode':s(g(1)),'team':s(g(2)) or 'Unassigned',
            'name':nm,'level':s(g(5)),'kra':s(g(6)),
            'mtd':{
                'revTarget':num(g(7)),'revAch':num(g(8)),'pct':num(g(9)),
                'cliTarget':num(g(10)),'totalClients':num(g(11)),'cliAcq':num(g(12)),
                'aumTarget':num(g(13)),'totalAum':num(g(14)),'mfNetSale':num(g(15)),'pmsAif':num(g(16)),
                'sipTarget':num(g(17)),'sipAdd':num(g(18)),'gtTarget':num(g(19)),'gtTime':round(hms_to_hours(g(20)),3)
            },
            'ytd':{
                'revTarget':num(g(22)),'revAch':num(g(23)),'revDeff':num(g(24)),
                'cliTarget':num(g(25)),'totalClients':num(g(26)),'cliAcq':num(g(27)),
                'aumTarget':num(g(28)),'totalAum':num(g(29)),'mfNetSale':num(g(30)),'pmsAif':num(g(31)),
                'sipTarget':num(g(32)),'sipAdd':num(g(33)),'gtTarget':num(g(34)),'gtTime':round(hms_to_hours(g(35)),3)
            }
        })
    print(f'{len(revsum_rows)} rows  ({time.time()-t0:.1f}s)', flush=True)

finally:
    wb.Close(False)
    xl.Quit()

data={
    'meta':{
        'source':SRC,'fyMonths':FY_MONTHS,
        'generated':datetime.now().strftime('%Y-%m-%d %H:%M'),
        'insRows':len(ins),'feesRows':len(fees),'pmsRows':len(pms_rows),
        'talkRows':len(talk_rows),'revsumRows':len(revsum_rows),
        'revsumMtdMonth':revsum_mtd_month,
    },
    'ins':ins,'fees':fees,'pms':pms_rows,
    'talk':{'months':FY_MONTHS,'rows':talk_rows},
    'revsum':{'mtdMonth':revsum_mtd_month,'rows':revsum_rows},
}

print('Writing data.js...', end=' ', flush=True)
with open('data.js','w',encoding='utf-8') as f:
    f.write('window.REVENUE_DATA = ')
    json.dump(data,f,ensure_ascii=False)
    f.write(';')
print('done', flush=True)

print()
print('INS',len(ins),'FEES',len(fees),'PMS',len(pms_rows),'TALK',len(talk_rows),'REVSUM',len(revsum_rows))
print('INS revenue total', round(sum(r['revenue'] for r in ins)))
print('FEES revenue total', round(sum(r['revenue'] for r in fees)))
print('PMS revenue total', round(sum(r['revenue'] for r in pms_rows)))
print(f'Total time: {time.time()-t0:.1f}s')
