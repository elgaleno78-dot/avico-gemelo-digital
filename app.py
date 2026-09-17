from flask import Flask, jsonify, render_template, request
import random, math, datetime

app = Flask(__name__)

BASE = {
    "seed": 20260917,
    "admissions_day": 30.5,
    "shift_weights": {"Matutino": 13.3, "Vespertino": 8.7, "Nocturno": 8.4},
    "births_day": 7.5,
    "vaginal_probability": 0.59,
    "cesarean_probability": 0.41,
    "admission_probability": 0.55,
    "floor_capacity": 27,
    "initial_floor_occupancy": 22,
    "labor_capacity": 9,
    "initial_labor_occupancy": 4,
    "tococirugia_capacity": 10,
    "recovery_capacity": 5,
    "physicians_per_shift": 4,
    "triage_minutes": 35,
    "labor_minutes": 280,
    "vaginal_minutes": 45,
    "cesarean_minutes": 70,
    "recovery_minutes": 130,
    "floor_stay_minutes": 1500,
    "cleaning_minutes": 40,
    "triage_queue_alert": 6
}

PHEDS = []
for i in range(1, 25): PHEDS.append({"id":f"H{i:02d}","area":"Hospitalización GO","censable":True,"oxygen":True,"isolation":False})
for i in range(25, 28): PHEDS.append({"id":f"H{i:02d}","area":"Hospitalización GO · Aislamiento","censable":True,"oxygen":True,"isolation":True})
for i in range(1, 11): PHEDS.append({"id":f"T{i:02d}","area":"Tococirugía","censable":False,"oxygen":i < 10,"isolation":False})
PHEDS += [{"id":"E01","area":"Expulsión","censable":False,"oxygen":True,"isolation":False},{"id":"M01","area":"Sala Mixta","censable":False,"oxygen":True,"isolation":False},{"id":"UT01","area":"Urgencias Tococirugía","censable":False,"oxygen":True,"isolation":False,"fictitious":True}]
for i in range(1,6): PHEDS.append({"id":f"RP{i:02d}","area":"Recuperación Postparto","censable":False,"oxygen":i < 5,"isolation":False})

SOURCES=[{"name":"Nacimientos 2026","type":"Google Sheets","status":"registrada","note":"Producción y distribución temporal."},{"name":"Censo de Piso","type":"Google Sheets / XLSX mensual","status":"registrada","note":"Ocupación y dinámica de 27 camas censables."},{"name":"Ingresos de Urgencias / Triage","type":"Excel / OneDrive","status":"registrada","note":"Demanda de entrada al servicio y distribución por turno."}]

def poisson(lam,rng):
    l=math.exp(-lam); k=0; q=1.0
    while q>l: k+=1; q*=rng.random()
    return k-1

def simulate(days=7,params=None):
    p=dict(BASE)
    if params: p.update(params)
    rng=random.Random(int(p["seed"])); floor=min(int(p["initial_floor_occupancy"]),int(p["floor_capacity"])); labor=min(int(p["initial_labor_occupancy"]),int(p["labor_capacity"]))
    rows=[]; ta=tb=tc=peakq=peaklabor=0
    for d in range(days):
        adm=poisson(float(p["admissions_day"]),rng); admitted=sum(1 for _ in range(adm) if rng.random()<float(p["admission_probability"])); discharged_urg=adm-admitted
        births=poisson(float(p["births_day"]),rng); cs=sum(1 for _ in range(births) if rng.random()<float(p["cesarean_probability"])); vaginal=births-cs
        labor_entries=max(0,min(admitted,poisson(max(1,births*1.15),rng))); labor_exits=min(labor+labor_entries,births); labor=max(0,min(int(p["labor_capacity"]),labor+labor_entries-labor_exits))
        recover=min(int(p["recovery_capacity"]),births)
        floor_entries=min(admitted,max(0,births+poisson(max(1,admitted*.18),rng))); discharges=max(0,poisson(max(1,floor/1.2),rng)); floor=max(0,min(int(p["floor_capacity"]),floor+floor_entries-discharges))
        q=max(0,round(adm/24*float(p["triage_minutes"])/60-int(p["physicians_per_shift"])))
        peakq=max(peakq,q); peaklabor=max(peaklabor,labor); ta+=adm; tb+=births; tc+=cs
        rows.append({"day":d+1,"date":(datetime.date.today()+datetime.timedelta(days=d)).isoformat(),"emergency_arrivals":adm,"urgency_discharges":discharged_urg,"hospital_admissions":admitted,"labor_entries":labor_entries,"labor_occupancy":labor,"labor_capacity":int(p["labor_capacity"]),"labor_pct":round(100*labor/int(p["labor_capacity"]),1),"births":births,"vaginal":vaginal,"cesareans":cs,"recovery_occupancy":recover,"floor_occupancy":floor,"floor_pct":round(100*floor/int(p["floor_capacity"]),1),"triage_queue":q})
    return {"days":rows,"summary":{"emergency_arrivals":ta,"births":tb,"cesareans":tc,"vaginal":tb-tc,"peak_floor_pct":max((x["floor_pct"] for x in rows),default=0),"peak_labor_occupancy":peaklabor,"peak_triage_queue":peakq},"parameters":p}

@app.route("/")
def home(): return render_template("index.html")
@app.route("/health")
def health(): return jsonify({"status":"ok","service":"AVICO Gemelo Digital","pheds_positions":len(PHEDS),"labor_positions":BASE["labor_capacity"]})
@app.route("/api/config")
def config(): return jsonify(BASE)
@app.route("/api/pheds")
def pheds(): return jsonify({"total":len(PHEDS),"positions":PHEDS})
@app.route("/api/sources")
def sources(): return jsonify(SOURCES)
@app.route("/api/simulate",methods=["POST"])
def api_simulate():
    b=request.get_json(silent=True) or {}; return jsonify(simulate(int(b.get("days",7)),b.get("parameters")))
@app.route("/api/predict")
def predict():
    b=simulate(1)["days"][0]
    def h(n): return {"floor_pct":min(100,b["floor_pct"]+round(n*1.3)),"labor_pct":min(100,b["labor_pct"]+round(n*2.2)),"triage_queue":b["triage_queue"]+max(0,n//3)}
    return jsonify({"plus_2h":h(2),"plus_4h":h(4),"plus_6h":h(6),"warning":"Predicción operativa V1 de simulación; no es un modelo clínico validado."})
if __name__=="__main__": app.run(host="0.0.0.0",port=10000)
