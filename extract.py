# -*- coding: utf-8 -*-
import json, os, time, glob, subprocess
from datetime import datetime, timedelta
import win32com.client as win32

# Auto-detect the most recently modified .xlsb in this folder
_xlsb_files = sorted([f for f in glob.glob('*.xlsb') if not os.path.basename(f).startswith('~$')], key=os.path.getmtime, reverse=True)
if not _xlsb_files:
    raise FileNotFoundError('No .xlsb file found in this folder. Copy the MIS file here first.')
SRC = _xlsb_files[0]
print(f'Source file: {SRC}', flush=True)

def ser2month(n):
    return (datetime(1899,12,30)+timedelta(days=float(n))).strftime('%b-%Y')
_today=datetime.now()
_fy_start=_today.year if _today.month>=4 else _today.year-1
FY_MONTHS=[datetime(_fy_start+(1 if (3+i)%12<3 else 0),(3+i)%12+1,1).strftime('%b-%Y') for i in range(12)]

import re
def fy_months_from_label(mtd):
    """Derive the 12 FY months (Apr..Mar) from an MTD label like 'JUN-2026'. None on failure."""
    mn=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    MN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    t=(mtd or '').strip().lower(); mo=-1
    for j in range(12):
        if t.startswith(mn[j]): mo=j; break
    m=re.search(r'(\d{4})', t); yr=int(m.group(1)) if m else None
    if mo<0 or yr is None: return None
    fy_start = yr if mo>=3 else yr-1              # Apr+ → same year; Jan-Mar → previous year
    return [MN[(3+i)%12]+'-'+str(fy_start+(1 if (3+i)%12<3 else 0)) for i in range(12)]

def num(v):
    if v is None: return 0.0
    if isinstance(v,(int,float)): return float(v)
    sv = str(v).strip().replace(',','')
    if sv == '': return 0.0
    try: return float(sv)
    except: return 0.0

def num_excel_sum(v):
    """Numeric-only extraction that mirrors Excel's SUM() semantics —
    ignores text-formatted numbers (e.g. '1242775.00' stored as text).
    Users validate against Excel's SUM in the source sheet, so text-numbers
    have to be excluded to match those reference totals; otherwise premium
    totals over-count by picking up entries Excel itself treats as blank."""
    return float(v) if isinstance(v,(int,float)) else 0.0

def s(v):
    if v is None: return ''
    if hasattr(v,'hour'): return ''          # datetime/pywintypes time cell
    if isinstance(v,float) and v==int(v): return str(int(v))
    return str(v).strip()

def iso_date(v):
    """Return 'YYYY-MM-DD' from a datetime cell or Excel serial; '' if empty."""
    if v is None: return ''
    if hasattr(v,'year'):                    # datetime / pywintypes.datetime
        return '%04d-%02d-%02d' % (v.year, v.month, v.day)
    if isinstance(v,(int,float)):
        try:
            d = datetime(1899,12,30)+timedelta(days=float(v))
            return d.strftime('%Y-%m-%d')
        except: return ''
    sv = str(v).strip()
    # trim trailing time from an ISO 'YYYY-MM-DD hh:mm:ss' string
    if len(sv) >= 10 and sv[4]=='-' and sv[7]=='-':
        return sv[:10]
    return sv

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
    """Read only actual data rows × ncols columns — avoids UsedRange memory blow-up.
    Checks several candidate anchor columns (not just `last_col`) and takes the max
    last-row found, because some sheets have a serial-number/id column that goes
    blank near the bottom even though real data (RM name, revenue, etc.) continues
    further down — trusting a single column silently truncates real rows."""
    ws = wb.Sheets(name)
    # xlUp = -4162: find last non-empty row from the bottom, per column
    candidate_cols = sorted(set(c for c in (last_col, 2, 3, ncols) if 1 <= c <= ncols))
    last_row = 0
    for c in candidate_cols:
        r = ws.Cells(ws.Rows.Count, c).End(-4162).Row
        if r > last_row: last_row = r
    if last_row < 1: return []
    data = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, ncols)).Value
    if data is None: return []
    if not isinstance(data, tuple): return [(data,)]     # single cell
    if data and not isinstance(data[0], tuple): return (data,)  # single row
    return data

t0 = time.time()

# Unblock file (removes "downloaded from internet" flag that causes Excel's Protected View to freeze)
try:
    subprocess.run(
        ['powershell', '-Command', f'Unblock-File -Path "{os.path.abspath(SRC)}"'],
        capture_output=True, timeout=10
    )
    print('File unblocked (Protected View disabled)', flush=True)
except Exception as e:
    print(f'Unblock skipped: {e}', flush=True)

print(f'Opening {SRC} ...', flush=True)

xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.AutomationSecurity = 3   # disable macros without showing dialog
xl.EnableEvents = False

wb = xl.Workbooks.Open(
    os.path.abspath(SRC),
    UpdateLinks=0,
    ReadOnly=True,
    IgnoreReadOnlyRecommended=True,
    Notify=False,
    AddToMru=False
)
print(f'  opened in {time.time()-t0:.1f}s', flush=True)

# Derive the fiscal year dynamically from the REVENUE SUMMARY header month,
# so the INS pre-FY filter auto-adjusts when the file rolls into a new FY.
try:
    _hdr = s(wb.Sheets('REVENUE SUMMARY V 2.0').Cells(1,4).Value)
    _dyn = fy_months_from_label(_hdr)
    if _dyn:
        FY_MONTHS = _dyn
        print(f'FY detected from REVSUM header "{_hdr}": {FY_MONTHS[0]} .. {FY_MONTHS[-1]}', flush=True)
    else:
        print(f'FY auto-detect failed (header="{_hdr}"), using default {FY_MONTHS[0]}..{FY_MONTHS[-1]}', flush=True)
except Exception as _e:
    print(f'FY auto-detect error ({_e}), using default', flush=True)

try:
    # -------- INS --------
    print('Reading INS...', end=' ', flush=True)
    ins=[]
    for i,row in enumerate(sheet_data(wb,'INS', ncols=65)):
        if i==0: continue
        def g(ix): return row[ix] if ix<len(row) else None
        if g(2) in (None,''): continue   # RmName blank = truly empty row (s.no can be blank on real rows)
        bt=s(g(1)).upper()
        if bt=='B2B': continue
        mo=s(g(48))
        if mo and mo not in FY_MONTHS: continue      # exclude pre-FY rows
        ins.append({
            'dtype':s(g(1)) or 'B2C','rm':s(g(2)),'client':s(g(3)),
            'platform':s(g(8)) or 'Unknown',
            'category':s(g(10)) or 'Unknown',
            'insType':s(g(11)) or 'Unknown',
            'partner':s(g(14)) or 'Unknown',
            'premium':num_excel_sum(g(27)),'createDate':iso_date(g(47)),
            'month':s(g(48)),'confirmed':s(g(50)),
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
        if g(2) in (None,''): continue   # RmName blank = truly empty row
        bt=s(g(1)).upper()
        if bt=='B2B': continue
        fees.append({
            'dtype':s(g(1)) or 'B2C','rm':s(g(2)),'client':s(g(3)),
            'platform':s(g(8)) or 'Unknown',
            'category':s(g(10)) or 'Unknown',
            'asset':s(g(11)) or 'Unknown',
            'prodType':s(g(12)) or 'Unknown',
            'amount':num(g(14)),'createDate':iso_date(g(28)),
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
        if g(2) in (None,''): continue   # RM name blank = truly empty row
        bt=s(g(1)).upper()
        if bt=='B2B': continue
        raw_mo=g(28)
        mo=s(raw_mo).title() if raw_mo else ''
        pms_rows.append({
            'dtype':s(g(1)) or 'B2C','rm':s(g(2)),'client':s(g(3)),
            'platform':s(g(8)) or 'Unknown',
            'category':s(g(10)) or 'Unknown',
            'asset':s(g(11)) or 'Unknown',
            'scheme':s(g(12)) or 'Unknown',
            'punchAmount':num(g(16)),'createDate':iso_date(g(27)),'month':mo,
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

    def revsum_metrics(g, base_mtd, base_ytd):
        """Read the 12 target/achievement metrics starting at 0-indexed
        column `base_mtd` (MTD block) and `base_ytd` (YTD block, which is
        always the MTD block's columns +15)."""
        return {
            'mtd':{
                'revTarget':num(g(base_mtd)),'revAch':num(g(base_mtd+1)),'pct':num(g(base_mtd+2)),
                'cliTarget':num(g(base_mtd+3)),'totalClients':num(g(base_mtd+4)),'cliAcq':num(g(base_mtd+5)),
                'aumTarget':num(g(base_mtd+6)),'totalAum':num(g(base_mtd+7)),'mfNetSale':num(g(base_mtd+8)),'pmsAif':num(g(base_mtd+9)),
                'sipTarget':num(g(base_mtd+10)),'sipAdd':num(g(base_mtd+11)),'gtTarget':round(hms_to_hours(g(base_mtd+12)),3),'gtTime':round(hms_to_hours(g(base_mtd+13)),3)
            },
            'ytd':{
                'revTarget':num(g(base_ytd)),'revAch':num(g(base_ytd+1)),'revDeff':num(g(base_ytd+2)),
                'cliTarget':num(g(base_ytd+3)),'totalClients':num(g(base_ytd+4)),'cliAcq':num(g(base_ytd+5)),
                'aumTarget':num(g(base_ytd+6)),'totalAum':num(g(base_ytd+7)),'mfNetSale':num(g(base_ytd+8)),'pmsAif':num(g(base_ytd+9)),
                'sipTarget':num(g(base_ytd+10)),'sipAdd':num(g(base_ytd+11)),'gtTarget':round(hms_to_hours(g(base_ytd+12)),3),'gtTime':round(hms_to_hours(g(base_ytd+13)),3)
            }
        }

    # REVSUM is laid out as repeating team blocks: named RM rows, then a
    # "Pool<Team>" row (unassigned/house business for that team — a real
    # contributor, NOT a subtotal), then a blank-name team-subtotal row,
    # ending in a final "Overall" grand-total row. Pool rows carry no Team
    # of their own (column C is blank), so their team is carried forward
    # from the last named RM row above them in the same block.
    def read_revsum_snapshot():
        """Read the REVSUM sheet's current live state and return (month_label, rows, overall)."""
        rows_out=[]
        overall_out=revsum_metrics(lambda ix: None, 7, 22)
        month_lbl=''
        last_team='Unassigned'
        for i,row in enumerate(sheet_data(wb,'REVENUE SUMMARY V 2.0', ncols=37, last_col=4)):
            def g(ix): return row[ix] if ix<len(row) else None
            if i==0: month_lbl=s(g(3)); continue
            if i<3: continue
            nm=s(g(3))
            if nm=='': continue   # blank-name row = team subtotal, not a contributor
            if nm.lower().startswith('overall'):
                overall_out=revsum_metrics(g, 7, 22)
                continue
            team=s(g(2))
            if team: last_team=team
            else: team=last_team   # Pool row — inherit the preceding RM's team
            row_metrics=revsum_metrics(g, 7, 22)
            row_metrics.update({'sn':num(g(0)),'empCode':s(g(1)),'team':team,'name':nm,'level':s(g(5)),'kra':s(g(6))})
            rows_out.append(row_metrics)
        return month_lbl, rows_out, overall_out

    def revsum_sane(rows, overall):
        """Reject a snapshot if any figure looks like a stale/garbage
        recalculation (previously seen as values like -1.5e11) rather than
        a real result. YTD AUM totals aggregate across ~70+ RMs and
        legitimately reach into the hundreds of crores (~1e9-1e10), so the
        bound has to sit comfortably above that, not just above revenue
        scale figures."""
        BOUND=1e10
        vals=list(overall['mtd'].values())+list(overall['ytd'].values())
        for r in rows:
            vals.extend(r['mtd'].values()); vals.extend(r['ytd'].values())
        return all(isinstance(v,(int,float)) and abs(v)<BOUND for v in vals)

    def wait_for_calc(timeout=20):
        waited=0.0
        while xl.CalculationState!=0 and waited<timeout:
            time.sleep(0.1); waited+=0.1

    revsum_mtd_month, revsum_rows, revsum_overall = read_revsum_snapshot()
    print(f'{len(revsum_rows)} rows  ({time.time()-t0:.1f}s)', flush=True)
    print(f'  Overall target: MTD={revsum_overall["mtd"]["revTarget"]}, YTD={revsum_overall["ytd"]["revTarget"]}', flush=True)

    # -------- Full 12-month REVSUM history via DASHBOARD!E5 --------
    # REVENUE SUMMARY V 2.0's 'MTD' columns are live formulas keyed off
    # DASHBOARD!E5 (a month-label cell — see D1's formula ='DASHBOARD!$E$5).
    # 'REVENUE REPORT MONTHLY' (read further below) is driven by the same
    # cell. We step E5 through each FY month, force a full recalculation,
    # and capture a snapshot for that month — then restore E5's original
    # value before reading anything else, so later reads see the true
    # current month again. (A previous attempt at this left E5 mutated
    # for the REVENUE REPORT read afterward, which is what actually
    # produced garbage values — not an inherent COM/Excel limitation.)
    key=revsum_mtd_month.upper() if revsum_mtd_month else 'CURRENT'
    revsum_by_month={key:{'rows':revsum_rows,'overallTarget':revsum_overall}}
    try:
        dash_ws=wb.Sheets('DASHBOARD')
        e5=dash_ws.Range('E5')
        orig_e5=e5.Value
        orig_fmt=e5.NumberFormat
        # E5 holds the month as literal TEXT (e.g. "APR-2026"). Writing a plain
        # string via COM without forcing Text format first lets Excel's input
        # parser silently reinterpret "AUG-2026" as an actual date — which
        # breaks every text-keyed SUMIFS/lookup downstream and is what produced
        # garbage achievement figures in earlier attempts at this. Forcing '@'
        # (Text) number format before each write keeps it a literal string.
        e5.NumberFormat='@'
        print('Reading full-year REVSUM history via DASHBOARD!E5...', flush=True)
        for m in FY_MONTHS:
            mkey=m.upper()
            if mkey==key: continue   # already have the current month, from the proven main read above
            e5.Value=mkey
            xl.Calculate()
            wait_for_calc()
            m_lbl, m_rows, m_overall = read_revsum_snapshot()
            if revsum_sane(m_rows, m_overall):
                revsum_by_month[mkey]={'rows':m_rows,'overallTarget':m_overall}
                print(f'  {mkey}: ok ({len(m_rows)} rows)', flush=True)
            else:
                print(f'  {mkey}: SKIPPED (implausible recalculation result)', flush=True)
        e5.Value=orig_e5
        e5.NumberFormat=orig_fmt
        xl.Calculate()
        wait_for_calc()
    except Exception as ex:
        print(f'  Multi-month REVSUM skipped: {ex}', flush=True)
    print(f'  REVSUM months captured: {sorted(revsum_by_month.keys())}  ({time.time()-t0:.1f}s)', flush=True)

    # -------- REVENUE REPORT MONTHLY / YTD (reconciled product-wise) --------
    # Column layouts differ between the two sheets:
    #   MONTHLY: fees=FP New+Renewal · insurance=Total Insurance col · pms=Unlisted+PMS+Stallion · trail=MF Trail
    #   YTD:     insurance has NO pre-summed col, so sum the 7 product cols; pms/trail shift left by 1
    def read_report(sheet_name, mapping, ncols):
        print(f'Reading {sheet_name}...', end=' ', flush=True)
        rows_out=[]; hdr_month=''
        raw=sheet_data(wb, sheet_name, ncols=ncols)
        if raw and len(raw)>0 and len(raw[0])>3:
            hdr_month=s(raw[0][3])
        for i,row in enumerate(raw):
            if i<2: continue                      # skip 2 header rows
            def g(ix): return row[ix] if ix<len(row) else None
            if g(0) in (None,''): continue        # need Emp Code
            rec={'empCode':s(g(0)),'team':s(g(1)) or 'Unassigned','name':s(g(3)) or s(g(2)),
                 'pool':s(g(2)).lower().startswith('pool')}
            for k,cols in mapping.items():
                rec[k]=round(sum(num(g(c)) for c in cols),2)
            rows_out.append(rec)
        print(f'{len(rows_out)} rows  ({time.time()-t0:.1f}s)', flush=True)
        return hdr_month, rows_out

    MONTHLY_MAP={'fees':[4,5],'insurance':[13],'pms':[15,16,17],'trail':[18],'total':[19]}
    YTD_MAP    ={'fees':[4,5],'insurance':[6,7,8,9,10,11,12],'pms':[13,14,15],'trail':[16],'total':[17]}
    rr_month, rr_monthly = read_report('REVENUE REPORT MONTHLY', MONTHLY_MAP, 20)
    _,        rr_ytd     = read_report('REVENUE REPORT YTD',     YTD_MAP,     20)

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
    'revsum':{'mtdMonth':revsum_mtd_month,'rows':revsum_rows,'overallTarget':revsum_overall,'months':revsum_by_month},
    'revreport':{'mtdMonth':rr_month,'monthly':rr_monthly,'ytd':rr_ytd},
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
def rsum(rows,k): return round(sum(r[k] for r in rows))
print(f'REVENUE REPORT MONTHLY ({rr_month}): ins={rsum(rr_monthly,"insurance")} fees={rsum(rr_monthly,"fees")} pms={rsum(rr_monthly,"pms")} trail={rsum(rr_monthly,"trail")} total={rsum(rr_monthly,"total")}')
print(f'REVENUE REPORT YTD           : ins={rsum(rr_ytd,"insurance")} fees={rsum(rr_ytd,"fees")} pms={rsum(rr_ytd,"pms")} trail={rsum(rr_ytd,"trail")} total={rsum(rr_ytd,"total")}')
print(f'Total time: {time.time()-t0:.1f}s')
