from flask import Flask,jsonify,render_template,request
import random,math,datetime,os,io,re
from zoneinfo import ZoneInfo
import requests
from openpyxl import load_workbook
app=Flask(__name__); TZ=ZoneInfo('America/Cancun')
BASE={'seed':20260917,'admissions_day':30.5,'births_day':7.5,'cesarean_probability':.41,'admission_probability':.55,'floor_capacity':27,'initial_floor_occupancy':22,'labor_capacity':9,'initial_labor_occupancy':4,'physicians_per_shift':4,'triage_minutes':35}
PHEDS=[]
for i in range(1,25): PHEDS.append({'id':f'H{i:02d}','area':'Hospitalización GO','censable':True})
for i in range(25,28): PHEDS.append({'id':f'H{i:02d}','area':'Hospitalización GO · Aislamiento','censable':True})
for i in range(1,11): PHEDS.append({'id':f'T{i:02d}','area':'Tococirugía','censable':False})
PHEDS += [{'id':'E01','area':'Expulsión','censable':False},{'id':'M01','area':'Sala Mixta','censable':False},{'id':'UT01','area':'Urgencias Tococirugía','censable':False}]
for i in range(1,6): PHEDS.append({'id':f'RP{i:02d}','area':'Recuperación Postparto','censable':False})
def now(): return datetime.datetime.now(TZ)
def today(): return now().date()
def normdate(v):
    if isinstance(v,(datetime.date,datetime.datetime)): return v.date() if isinstance(v,datetime.datetime) else v
    if not v:return None
    s=str(v).strip()
    for f in ('%d/%m/%Y','%d/%m/%y','%Y-%m-%d'):
        try:return datetime.datetime.strptime(s,f).date()
        except:pass
    return None
def getbytes(url):
    if not url:return None
    r=requests.get(url,timeout=15); r.raise_for_status(); return r.content
def sheet_csv(sheet_id,gid):
    url=f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}'
    r=requests.get(url,timeout=15); r.raise_for_status(); return r.text
def live_births():
    try:
        sid=os.getenv('BIRTHS_SHEET_ID','1SJGGbbUGA2pTmouDi-tOFi_kVmtSVFhGFXBq70kFOTw'); gid=os.getenv('BIRTHS_GID','405341237'); import csv
        rows=list(csv.reader(io.StringIO(sheet_csv(sid,gid)))); vals=[]
        for row in rows[1:]:
            if len(row)>5:
                d=normdate(row[3]); typ=row[5].strip().lower()
                if d and d<=today() and typ in ('parto','cesarea','cesárea'): vals.append((d,typ))
        if not vals:return None
        last=max(d for d,_ in vals); day=[x for x in vals if x[0]==last]; return {'value':len(day),'vaginal':sum(t=='parto' for _,t in day),'cesareans':sum(t!='parto' for _,t in day),'last_valid_date':last.isoformat(),'status':'REAL · EN VIVO'}
    except Exception as e:return {'value':None,'last_valid_date':None,'status':'SIN CONEXIÓN','error':str(e)[:100]}
def live_xlsx(url,kind):
    try:
        data=getbytes(url)
        if not data:return None
        wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True); candidates=[]
        for ws in wb.worksheets:
            m=re.fullmatch(r'0?(\d{1,2})',ws.title.strip())
            if m:
                d=datetime.date(today().year,today().month,int(m.group(1)))
                if d<=today():candidates.append((d,ws))
        if not candidates:return None
        d,ws=max(candidates,key=lambda x:x[0]); rows=list(ws.iter_rows(values_only=True)); occ=set(); labor=0; triage=0
        for row in rows:
            txt=' | '.join(str(x or '') for x in row).upper()
            if kind=='census' and 'GINECO' in txt:
                bed=str(row[0] or '').strip() if row else ''
                if re.fullmatch(r'\d{1,2}',bed): occ.add(int(bed))
                if 'LABOR' in txt or 'TOCOCIR' in txt: labor+=1
            if kind=='triage':
                ds=[normdate(x) for x in row[:12]]
                if d in ds:triage+=1
        return {'value':len(occ) if kind=='census' else triage,'labor':labor if kind=='census' else None,'last_valid_date':d.isoformat(),'status':'REAL · EN VIVO'}
    except Exception as e:return {'value':None,'last_valid_date':None,'status':'SIN CONEXIÓN','error':str(e)[:100]}
def real_state():
    b=live_births(); c=live_xlsx(os.getenv('CENSUS_XLSX_URL'),'census'); t=live_xlsx(os.getenv('TRIAGE_XLSX_URL'),'triage')
    src=[{'name':'Nacimientos 2026','status':b['status'] if b else 'SIN CONEXIÓN','last_valid_date':b.get('last_valid_date') if b else None},{'name':'Censo Piso / Labor','status':c['status'] if c else 'CONFIGURAR URL','last_valid_date':c.get('last_valid_date') if c else None},{'name':'Urgencias / Triage','status':t['status'] if t else 'CONFIGURAR URL','last_valid_date':t.get('last_valid_date') if t else None}]
    return {'mode':'REAL','as_of':now().isoformat(),'cutoff_date':today().isoformat(),'sources':src,'births':b or {'value':None},'floor':{'value':c.get('value') if c else None,'capacity':27,'last_valid_date':c.get('last_valid_date') if c else None},'labor':{'value':c.get('labor') if c else None,'capacity':9,'last_valid_date':c.get('last_valid_date') if c else None},'triage':{'value':t.get('value') if t else None,'last_valid_date':t.get('last_valid_date') if t else None}}
def poisson(lam,rng):
    l=math.exp(-lam);k=0;q=1
    while q>l:k+=1;q*=rng.random()
    return k-1
def simulate(days=7,params=None):
    p=dict(BASE);p.update(params or {});rng=random.Random(int(p['seed']));rows=[];floor=p['initial_floor_occupancy'];labor=p['initial_labor_occupancy']
    for d in range(days):
        adm=poisson(float(p['admissions_day']),rng);births=poisson(float(p['births_day']),rng);cs=sum(rng.random()<p['cesarean_probability'] for _ in range(births)); admitted=sum(rng.random()<p['admission_probability'] for _ in range(adm));labor=max(0,min(p['labor_capacity'],labor+min(admitted,births)-births));floor=max(0,min(p['floor_capacity'],floor+admitted-poisson(max(1,floor/1.2),rng)));rows.append({'day':d+1,'emergency_arrivals':adm,'hospital_admissions':admitted,'labor_occupancy':labor,'labor_capacity':p['labor_capacity'],'births':births,'vaginal':births-cs,'cesareans':cs,'floor_occupancy':floor,'floor_pct':round(100*floor/p['floor_capacity'],1),'triage_queue':max(0,round(adm/24*p['triage_minutes']/60-p['physicians_per_shift']))})
    return {'mode':'SIMULADO','days':rows,'summary':{'emergency_arrivals':sum(x['emergency_arrivals'] for x in rows),'births':sum(x['births'] for x in rows),'peak_floor_pct':max(x['floor_pct'] for x in rows),'peak_labor_occupancy':max(x['labor_occupancy'] for x in rows)},'parameters':p}
@app.route('/')
def home():return render_template('index.html')
@app.route('/health')
def health():return jsonify({'status':'ok','server_time':now().isoformat(),'pheds_positions':len(PHEDS)})
@app.route('/api/config')
def config():return jsonify(BASE)
@app.route('/api/pheds')
def pheds():return jsonify({'total':len(PHEDS),'positions':PHEDS})
@app.route('/api/sources')
def sources():return jsonify(real_state()['sources'])
@app.route('/api/real-state')
def state():return jsonify(real_state())
@app.route('/api/simulate',methods=['POST'])
def sim():
    b=request.get_json(silent=True) or {};return jsonify(simulate(int(b.get('days',7)),b.get('parameters')))
@app.route('/api/predict')
def predict():
    r=real_state();return jsonify({'mode':'PREDICHO','plus_2h':{'floor_pct':0,'labor_pct':0,'triage_queue':0},'plus_4h':{'floor_pct':0,'labor_pct':0,'triage_queue':0},'plus_6h':{'floor_pct':0,'labor_pct':0,'triage_queue':0},'warning':'Predicción suspendida hasta disponer de las tres fuentes reales simultáneamente.'})
if __name__=='__main__':app.run(host='0.0.0.0',port=10000)
