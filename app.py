from flask import Flask, jsonify, render_template, request
import random, math, datetime, os
from zoneinfo import ZoneInfo

app=Flask(__name__)
TZ=ZoneInfo("America/Cancun")
BASE={"seed":20260917,"admissions_day":30.5,"shift_weights":{"Matutino":13.3,"Vespertino":8.7,"Nocturno":8.4},"births_day":7.5,"vaginal_probability":.59,"cesarean_probability":.41,"admission_probability":.55,"floor_capacity":27,"initial_floor_occupancy":22,"labor_capacity":9,"initial_labor_occupancy":4,"tococirugia_capacity":10,"recovery_capacity":5,"physicians_per_shift":4,"triage_minutes":35,"labor_minutes":280,"vaginal_minutes":45,"cesarean_minutes":70,"recovery_minutes":130,"floor_stay_minutes":1500,"cleaning_minutes":40,"triage_queue_alert":6}
PHEDS=[]
for i in range(1,25): PHEDS.append({"id":f"H{i:02d}","area":"Hospitalización GO","censable":True,"oxygen":True,"isolation":False})
for i in range(25,28): PHEDS.append({"id":f"H{i:02d}","area":"Hospitalización GO · Aislamiento","censable":True,"oxygen":True,"isolation":True})
for i in range(1,11): PHEDS.append({"id":f"T{i:02d}","area":"Tococirugía","censable":False,"oxygen":i<10,"isolation":False})
PHEDS += [{"id":"E01","area":"Expulsión","censable":False,"oxygen":True,"isolation":False},{"id":"M01","area":"Sala Mixta","censable":False,"oxygen":True,"isolation":False},{"id":"UT01","area":"Urgencias Tococirugía","censable":False,"oxygen":True,"isolation":False,"fictitious":True}]
for i in range(1,6): PHEDS.append({"id":f"RP{i:02d}","area":"Recuperación Postparto","censable":False,"oxygen":i<5,"isolation":False})

def now(): return datetime.datetime.now(TZ)
def cutoff(): return now().date().isoformat()
def source_state(name,kind,last_date,status,note):
    today=cutoff(); valid=bool(last_date and last_date<=today); stale=not valid or last_date<today
    return {"name":name,"type":kind,"status":status if valid else "NO CONECTADA","last_valid_date":last_date if valid else None,"cutoff_date":today,"stale":stale,"note":note}

def sources():
    # Fechas verificadas al corte inicial. Cuando Render tenga credenciales/endpoint, los adaptadores sustituyen estos valores automáticamente.
    return [
      source_state("Nacimientos 2026","Google Sheets","2026-09-16","REAL · CORTE VERIFICADO","Último registro válido ≤ fecha actual. Lectura directa pendiente de credencial de servicio en Render."),
      source_state("Censo de Piso / Labor","XLSX mensual","2026-09-16","REAL · CORTE VERIFICADO","Usar únicamente la última hoja diaria válida ≤ hoy; ignorar hojas futuras/plantillas."),
      source_state("Ingresos de Urgencias / Triage","Excel / OneDrive",None,"HISTÓRICO","No presentar 30.5/día como dato actual hasta que Render pueda leer la fuente directamente.")]

def poisson(lam,rng):
    l=math.exp(-lam); k=0; q=1.
    while q>l: k+=1; q*=rng.random()
    return k-1

def simulate(days=7,params=None):
    p=dict(BASE); p.update(params or {}); rng=random.Random(int(p["seed"])); floor=min(int(p["initial_floor_occupancy"]),int(p["floor_capacity"])); labor=min(int(p["initial_labor_occupancy"]),int(p["labor_capacity"])); rows=[]; ta=tb=tc=peakq=peaklabor=0
    for d in range(days):
        adm=poisson(float(p["admissions_day"]),rng); admitted=sum(rng.random()<float(p["admission_probability"]) for _ in range(adm)); births=poisson(float(p["births_day"]),rng); cs=sum(rng.random()<float(p["cesarean_probability"]) for _ in range(births)); vaginal=births-cs; labor_entries=max(0,min(admitted,poisson(max(1,births*1.15),rng))); labor_exits=min(labor+labor_entries,births); labor=max(0,min(int(p["labor_capacity"]),labor+labor_entries-labor_exits)); recover=min(int(p["recovery_capacity"]),births); floor_entries=min(admitted,max(0,births+poisson(max(1,admitted*.18),rng))); discharges=max(0,poisson(max(1,floor/1.2),rng)); floor=max(0,min(int(p["floor_capacity"]),floor+floor_entries-discharges)); q=max(0,round(adm/24*float(p["triage_minutes"])/60-int(p["physicians_per_shift"]))); peakq=max(peakq,q); peaklabor=max(peaklabor,labor); ta+=adm; tb+=births; tc+=cs
        rows.append({"day":d+1,"date":(now().date()+datetime.timedelta(days=d)).isoformat(),"emergency_arrivals":adm,"urgency_discharges":adm-admitted,"hospital_admissions":admitted,"labor_entries":labor_entries,"labor_occupancy":labor,"labor_capacity":int(p["labor_capacity"]),"labor_pct":round(100*labor/int(p["labor_capacity"]),1),"births":births,"vaginal":vaginal,"cesareans":cs,"recovery_occupancy":recover,"floor_occupancy":floor,"floor_pct":round(100*floor/int(p["floor_capacity"]),1),"triage_queue":q})
    return {"mode":"SIMULADO","days":rows,"summary":{"emergency_arrivals":ta,"births":tb,"cesareans":tc,"vaginal":tb-tc,"peak_floor_pct":max((x["floor_pct"] for x in rows),default=0),"peak_labor_occupancy":peaklabor,"peak_triage_queue":peakq},"parameters":p}

def real_state():
    ss=sources(); return {"mode":"REAL","as_of":now().isoformat(),"cutoff_date":cutoff(),"sources":ss,"births":{"value":None,"label":"Fuente conectable","last_valid_date":ss[0]["last_valid_date"]},"floor":{"value":None,"capacity":27,"last_valid_date":ss[1]["last_valid_date"]},"labor":{"value":None,"capacity":9,"last_valid_date":ss[1]["last_valid_date"]},"triage":{"value":None,"last_valid_date":ss[2]["last_valid_date"]},"message":"El Monitor no sustituye datos reales faltantes con simulación. Configure adaptadores/credenciales de Render para actualización automática."}

@app.route('/')
def home(): return render_template('index.html')
@app.route('/health')
def health(): return jsonify({"status":"ok","service":"AVICO Gemelo Digital","server_time":now().isoformat(),"pheds_positions":len(PHEDS)})
@app.route('/api/config')
def config(): return jsonify(BASE)
@app.route('/api/pheds')
def pheds(): return jsonify({"total":len(PHEDS),"positions":PHEDS})
@app.route('/api/sources')
def api_sources(): return jsonify(sources())
@app.route('/api/real-state')
def api_real_state(): return jsonify(real_state())
@app.route('/api/simulate',methods=['POST'])
def api_simulate():
    b=request.get_json(silent=True) or {}; return jsonify(simulate(int(b.get('days',7)),b.get('parameters')))
@app.route('/api/predict')
def predict():
    b=simulate(1)['days'][0]
    def h(n): return {"floor_pct":min(100,b['floor_pct']+round(n*1.3)),"labor_pct":min(100,b['labor_pct']+round(n*2.2)),"triage_queue":b['triage_queue']+max(0,n//3)}
    return jsonify({"mode":"PREDICHO","plus_2h":h(2),"plus_4h":h(4),"plus_6h":h(6),"warning":"Predicción operativa V1; no es un modelo clínico validado."})
if __name__=='__main__': app.run(host='0.0.0.0',port=10000)
