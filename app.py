from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime, date
from pathlib import Path
from threading import RLock
import json, uuid, os, tempfile, shutil

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = Path(os.environ.get('METRO_CONNECT_XLSX', BASE_DIR / 'data' / 'MetroConnectData.xlsx'))
DOMAIN_FILE = BASE_DIR / 'data' / 'domains.json'
LOCK = RLock()
app = Flask(__name__)

# Daily rows in the original monthly tabs are imported as-is. Original sheets remain in the workbook.
SERVICE_ROWS = {
    'daily': {'DLF Cybercity': 'Daily', 'DLF': 'Daily', 'Keppel': 'Daily', 'ITPC': 'Daily', 'Kosmo One': 'Daily', 'SC-GBS': 'Daily'},
    'MTC': {'Ridership': 'Ridership', 'Revenue': 'Revenue', 'GAP': 'GAP'},
    'parking': {'Pallavaram parking': 'Vehicle Count', 'Thirumangalam Parking': 'Vehicle Count'}
}
HEADERS = ['Record ID','Date','Month','Year','Service or Domain','Performance Metric','Daily Value','Cumulative Value','Remarks','Last Updated','Source Sheet']


def as_date(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    if isinstance(v, str):
        for fmt in ('%Y-%m-%d','%d-%m-%Y','%d/%m/%Y'):
            try: return datetime.strptime(v, fmt).date()
            except ValueError: pass
    return None


def month_name(d): return d.strftime('%B') if d else ''


def infer_metric_rows(ws):
    """Return (row number, domain, metric) for source daily-value rows only."""
    rows=[]
    recognized = {'mtc':'MTC','pallavaram parking':'Pallavaram parking','thirumangalam parking':'Thirumangalam Parking'}
    daily_names = {x.lower():x for x in SERVICE_ROWS['daily']}
    for r in range(2, ws.max_row+1):
        label=ws.cell(r,1).value
        metric_label=ws.cell(r,2).value
        label_s=str(label).strip() if label is not None else ''
        metric_s=str(metric_label).strip() if metric_label is not None else ''
        if metric_s.lower() == 'cumulative' or not metric_s:
            continue
        metric_low=metric_s.lower()
        if label_s.lower() in daily_names and metric_low == 'daily':
            rows.append((r,daily_names[label_s.lower()], 'Daily'))
            continue
        if metric_low not in ('ridership','revenue','gap','vehicle count','daily'):
            continue
        # For blank service labels, use the nearest preceding recognized domain heading.
        domain = recognized.get(label_s.lower())
        if not domain and not label_s:
            for rr in range(r-1, 1, -1):
                prior=ws.cell(rr,1).value
                if prior and str(prior).strip().lower() in recognized:
                    domain=recognized[str(prior).strip().lower()]
                    break
        if not domain:
            continue
        if domain == 'MTC' and metric_low in ('ridership','revenue','gap'):
            metric = 'GAP' if metric_low == 'gap' else metric_s.title()
            rows.append((r,domain,metric))
        elif domain in ('Pallavaram parking','Thirumangalam Parking') and metric_low in ('vehicle count','revenue'):
            rows.append((r,domain,'Vehicle Count' if metric_low == 'vehicle count' else 'Revenue'))
    seen=set(); out=[]
    for item in rows:
        if item not in seen:
            out.append(item); seen.add(item)
    return out


def ensure_data_sheet():
    if not DATA_FILE.exists():
        raise FileNotFoundError(f'Excel data file not found: {DATA_FILE}')
    with LOCK:
        wb=load_workbook(DATA_FILE)
        if 'Data Management' in wb.sheetnames:
            ws=wb['Data Management']
            if ws.max_row >= 1 and [ws.cell(1,c).value for c in range(1,len(HEADERS)+1)] == HEADERS:
                wb.close(); return
            wb.remove(ws)
        ws=wb.create_sheet('Data Management')
        ws.append(HEADERS)
        imported=[]
        for sn in wb.sheetnames:
            if sn in ('Data Management','Pallavaram detailed analysis ','SC-GBS performance','Sheet1'):
                continue
            source=wb[sn]
            for r,domain,metric in infer_metric_rows(source):
                for c in range(3, source.max_column+1):
                    d=as_date(source.cell(1,c).value)
                    if not d: continue
                    value=source.cell(r,c).value
                    # Do not create records for blank source cells; preserve numeric zero as a real value.
                    if value is None or value == '':
                        continue
                    if not isinstance(value,(int,float)):
                        try: value=float(str(value).replace(',',''))
                        except Exception: continue
                    imported.append({'Record ID':str(uuid.uuid4()),'Date':d,'Month':month_name(d),'Year':d.year,'Service or Domain':domain,'Performance Metric':metric,'Daily Value':value,'Cumulative Value':None,'Remarks':'','Last Updated':'Imported from source workbook','Source Sheet':sn})
        imported.sort(key=lambda x:(x['Date'],x['Service or Domain'],x['Performance Metric']))
        # Compute monthly cumulative values, separately per domain/metric; blanks do not become zeros.
        running={}
        for rec in imported:
            key=(rec['Date'].year,rec['Date'].month,rec['Service or Domain'],rec['Performance Metric'])
            val=rec['Daily Value']
            if val is None:
                rec['Cumulative Value']=None
            else:
                running[key]=running.get(key,0)+val
                rec['Cumulative Value']=running[key]
            ws.append([rec[h] for h in HEADERS])
        style_data_sheet(ws)
        # Keep source tabs and their original historical values; only append the central data sheet.
        wb.save(DATA_FILE); wb.close()


def style_data_sheet(ws):
    ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:K{max(ws.max_row,1)}'
    for cell in ws[1]:
        cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='17365D'); cell.alignment=Alignment(horizontal='center')
    widths=[38,14,15,10,25,22,16,18,32,30,24]
    for i,w in enumerate(widths,1): ws.column_dimensions[chr(64+i)].width=w
    for row in ws.iter_rows(min_row=2):
        row[1].number_format='dd-mmm-yyyy'; row[6].number_format='#,##0.##'; row[7].number_format='#,##0.##'


def read_records():
    ensure_data_sheet()
    with LOCK:
        wb=load_workbook(DATA_FILE, data_only=True, read_only=True)
        ws=wb['Data Management']; out=[]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            d=as_date(row[1])
            out.append({'id':str(row[0]),'date':d.isoformat() if d else '', 'month':row[2] or (month_name(d) if d else ''),'year':row[3] or (d.year if d else None),'service':row[4] or '', 'metric':row[5] or '', 'value':row[6], 'cumulative':row[7], 'remarks':row[8] or '', 'updated':row[9] or '', 'source':row[10] or ''})
        wb.close()
        return out


def recalc_and_write(records):
    records.sort(key=lambda x:(x.get('date',''),x.get('service',''),x.get('metric','')))
    running={}
    for r in records:
        d=as_date(r.get('date'))
        if not d: r['cumulative']=None; continue
        r['month']=month_name(d); r['year']=d.year
        key=(d.year,d.month,r.get('service',''),r.get('metric',''))
        v=r.get('value')
        if v is None or v=='': r['value']=None; r['cumulative']=None
        else:
            v=float(v); r['value']=v; running[key]=running.get(key,0)+v; r['cumulative']=running[key]
    with LOCK:
        wb=load_workbook(DATA_FILE)
        ws=wb['Data Management'] if 'Data Management' in wb.sheetnames else wb.create_sheet('Data Management')
        if ws.max_row: ws.delete_rows(1,ws.max_row)
        ws.append(HEADERS)
        for r in records:
            d=as_date(r.get('date'))
            ws.append([r.get('id') or str(uuid.uuid4()),d,r.get('month'),r.get('year'),r.get('service'),r.get('metric'),r.get('value'),r.get('cumulative'),r.get('remarks',''),r.get('updated',''),r.get('source','Manual entry')])
        style_data_sheet(ws)
        # Save atomically to reduce the chance of a partially written workbook.
        tmp=DATA_FILE.with_suffix('.tmp.xlsx')
        wb.save(tmp); wb.close(); os.replace(tmp,DATA_FILE)


def options_from_records(records):
    domains = load_domain_config()
    return {'services':sorted(set(domains) | {r['service'] for r in records if r['service']}), 'metrics':sorted({r['metric'] for r in records if r['metric']} | {m for d in domains.values() for m in d.get('metrics', [])}), 'years':sorted({int(r['year']) for r in records if r.get('year')})}


def load_domain_config():
    """Persistent list of domains and their entry-enabled status."""
    records = read_records() if DATA_FILE.exists() else []
    known = sorted({r['service'] for r in records if r.get('service')})
    try:
        cfg = json.loads(DOMAIN_FILE.read_text(encoding='utf-8')) if DOMAIN_FILE.exists() else {}
    except Exception:
        cfg = {}
    domains = cfg.get('domains', {}) if isinstance(cfg, dict) else {}
    changed = False
    for name in known:
        if name not in domains:
            domains[name] = {'enabled': True, 'metrics': sorted({r['metric'] for r in records if r['service'] == name}) or ['Daily']}
            changed = True
    if changed or not DOMAIN_FILE.exists():
        DOMAIN_FILE.parent.mkdir(parents=True, exist_ok=True)
        DOMAIN_FILE.write_text(json.dumps({'domains': domains}, indent=2), encoding='utf-8')
    return domains

def save_domain_config(domains):
    DOMAIN_FILE.parent.mkdir(parents=True, exist_ok=True)
    DOMAIN_FILE.write_text(json.dumps({'domains': domains}, indent=2), encoding='utf-8')

@app.route('/')
def index(): return render_template('index.html')

@app.get('/api/domains')
def api_domains():
    domains = load_domain_config()
    return jsonify({'domains': [{'name': name, 'enabled': bool(cfg.get('enabled', True)), 'metrics': cfg.get('metrics', ['Daily'])} for name, cfg in sorted(domains.items())]})

@app.post('/api/domains')
def add_domain():
    payload = request.get_json(silent=True) or {}
    name = ' '.join(str(payload.get('name', '')).split())
    if not name or len(name) > 80:
        return jsonify({'error': 'Enter a domain name between 1 and 80 characters.'}), 400
    domains = load_domain_config()
    existing = next((n for n in domains if n.casefold() == name.casefold()), None)
    if existing:
        return jsonify({'error': f'{existing} already exists.'}), 409
    domains[name] = {'enabled': True, 'metrics': ['Daily']}
    save_domain_config(domains)
    return jsonify({'message': f'{name} added and enabled. Add records to start building its history.', 'domain': {'name': name, 'enabled': True, 'metrics': ['Daily']}})

@app.put('/api/domains/<path:name>')
def update_domain(name):
    payload = request.get_json(silent=True) or {}
    domains = load_domain_config()
    if name not in domains:
        return jsonify({'error': 'Domain not found.'}), 404
    domains[name]['enabled'] = bool(payload.get('enabled', domains[name].get('enabled', True)))
    save_domain_config(domains)
    return jsonify({'message': f'{name} ' + ('enabled' if domains[name]['enabled'] else 'disabled') + ' for new data entry.'})

@app.get('/api/records')
def api_records():
    recs=read_records()
    return jsonify({'records':recs,'options':options_from_records(recs)})

@app.post('/api/records')
def create_records():
    payload=request.get_json(silent=True) or {}
    d=as_date(payload.get('date'))
    items=payload.get('items') or []
    if not d: return jsonify({'error':'Choose a valid date before saving.'}),400
    if not items: return jsonify({'error':'Enter at least one daily value before saving.'}),400
    records=read_records()
    domains = load_domain_config()
    existing={(r['date'],r['service'].casefold(),r['metric'].casefold()):r for r in records}
    created=0; updated=0
    edit_mode=bool(payload.get('edit_mode'))
    for item in items:
        service=str(item.get('service','')).strip(); metric=str(item.get('metric','')).strip(); raw=item.get('value')
        if not service or not metric: continue
        domain_cfg = next((cfg for nm, cfg in domains.items() if nm.casefold() == service.casefold()), None)
        if domain_cfg is not None and not domain_cfg.get('enabled', True):
            return jsonify({'error': f'{service} is disabled. Enable it before entering new values.'}), 400
        if raw is None or raw=='': continue
        try: value=float(raw)
        except (ValueError,TypeError): return jsonify({'error':f'Value for {service} / {metric} must be numeric.'}),400
        if value < 0: return jsonify({'error':f'Value for {service} / {metric} cannot be negative.'}),400
        key=(d.isoformat(),service.casefold(),metric.casefold())
        if key in existing:
            if not edit_mode:
                return jsonify({'error':f'A record already exists for {d.isoformat()} — {service} / {metric}. Use Edit existing date to update it.'}),409
            r=existing[key]; r['value']=value; r['remarks']=str(item.get('remarks') or r.get('remarks','')); r['updated']=datetime.now().strftime('%Y-%m-%d %H:%M:%S'); updated+=1
        else:
            r={'id':str(uuid.uuid4()),'date':d.isoformat(),'month':month_name(d),'year':d.year,'service':service,'metric':metric,'value':value,'cumulative':None,'remarks':str(item.get('remarks') or ''),'updated':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'source':'Manual entry'}
            records.append(r); existing[key]=r; created+=1
    if created+updated==0: return jsonify({'error':'No values were entered. Fill at least one metric.'}),400
    recalc_and_write(records)
    return jsonify({'message':f'Saved successfully. {created} new record(s), {updated} updated.','created':created,'updated':updated})

@app.put('/api/records/<record_id>')
def update_record(record_id):
    payload=request.get_json(silent=True) or {}; records=read_records(); target=next((r for r in records if r['id']==record_id),None)
    if not target: return jsonify({'error':'Record not found.'}),404
    d=as_date(payload.get('date',target['date'])); service=str(payload.get('service',target['service'])).strip(); metric=str(payload.get('metric',target['metric'])).strip(); raw=payload.get('value',target['value'])
    if not d or not service or not metric: return jsonify({'error':'Date, service, and metric are required.'}),400
    try: value=None if raw is None or raw=='' else float(raw)
    except (ValueError,TypeError): return jsonify({'error':'Daily value must be numeric.'}),400
    if value is not None and value<0: return jsonify({'error':'Daily value cannot be negative.'}),400
    duplicate=next((r for r in records if r['id']!=record_id and r['date']==d.isoformat() and r['service'].casefold()==service.casefold() and r['metric'].casefold()==metric.casefold()),None)
    if duplicate: return jsonify({'error':'This date, service, and metric already exists as another record.'}),409
    target.update({'date':d.isoformat(),'month':month_name(d),'year':d.year,'service':service,'metric':metric,'value':value,'remarks':str(payload.get('remarks',target['remarks']) or ''),'updated':datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
    recalc_and_write(records)
    return jsonify({'message':'Record updated and cumulative values recalculated.'})

@app.delete('/api/records/<record_id>')
def delete_record(record_id):
    records=read_records(); new=[r for r in records if r['id']!=record_id]
    if len(new)==len(records): return jsonify({'error':'Record not found.'}),404
    recalc_and_write(new)
    return jsonify({'message':'Record deleted. Cumulative values recalculated.'})

@app.get('/api/export')
def export_file():
    ensure_data_sheet()
    return send_file(DATA_FILE,as_attachment=True,download_name='MetroConnectData.xlsx')

@app.get('/api/metrics')
def metrics():
    recs=read_records()
    metrics_map={}
    for r in recs:
        key=(r['service'],r['metric'])
        metrics_map.setdefault(key,{'service':r['service'],'metric':r['metric']})
    return jsonify({'metrics':list(metrics_map.values())})

if __name__=='__main__':
    ensure_data_sheet()
    print('\nMetro Connect Services dashboard is running at http://127.0.0.1:5000')
    print(f'Excel data file: {DATA_FILE}\n')
    app.run(host='127.0.0.1',port=5000,debug=False)
