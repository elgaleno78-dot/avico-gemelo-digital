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
    "tococirugia_capacity": 10,
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
for i in range(1, 25):
    PHEDS.append({"id": f"H{i:02d}", "area": "Hospitalización GO", "censable": True, "oxygen": True, "isolation": False})
for i in range(25, 28):
    PHEDS.append({"id": f"H{i:02d}", "area": "Hospitalización GO · Aislamiento", "censable": True, "oxygen": True, "isolation": True})
for i in range(1, 11):
    PHEDS.append({"id": f"T{i:02d}", "area": "Tococirugía", "censable": False, "oxygen": i < 10, "isolation": False})
PHEDS += [
    {"id":"E01","area":"Expulsión","censable":False,"oxygen":True,"isolation":False},
    {"id":"M01","area":"Sala Mixta","censable":False,"oxygen":True,"isolation":False},
    {"id":"UT01","area":"Urgencias Tococirugía","censable":False,"oxygen":True,"isolation":False,"fictitious":True},
]
for i in range(1, 6):
    PHEDS.append({"id":f"RP{i:02d}","area":"Recuperación Postparto","censable":False,"oxygen":i < 5,"isolation":False})

SOURCES = [
    {"name":"Nacimientos 2026","type":"Google Sheets","status":"registrada","note":"Fuente histórica para producción y distribución temporal."},
    {"name":"Censo de Piso","type":"Google Sheets / XLSX mensual","status":"registrada","note":"Ocupación y dinámica de las 27 camas censables."},
    {"name":"Admisión / Triage","type":"Excel / OneDrive","status":"registrada","note":"Demanda de entrada al servicio."}
]

def poisson(lam, rng):
    l = math.exp(-lam); k = 0; p = 1.0
    while p > l:
        k += 1; p *= rng.random()
    return k - 1

def simulate(days=7, params=None):
    p = dict(BASE)
    if params: p.update(params)
    rng = random.Random(int(p["seed"]))
    occupancy = min(int(p["initial_floor_occupancy"]), int(p["floor_capacity"]))
    rows=[]; total_adm=total_births=total_cs=peak_q=0
    for d in range(days):
        date=(datetime.date.today()+datetime.timedelta(days=d)).isoformat()
        adm=poisson(float(p["admissions_day"]),rng)
        births=poisson(float(p["births_day"]),rng)
        cs=sum(1 for _ in range(births) if rng.random()<float(p["cesarean_probability"]))
        admitted=sum(1 for _ in range(adm) if rng.random()<float(p["admission_probability"]))
        discharges=max(0, poisson(max(1, occupancy/1.2),rng))
        occupancy=max(0,min(int(p["floor_capacity"]),occupancy+admitted-discharges))
        q=max(0, round(adm/24*float(p["triage_minutes"])/60 - int(p["physicians_per_shift"])))
        peak_q=max(peak_q,q); total_adm+=adm; total_births+=births; total_cs+=cs
        rows.append({"day":d+1,"date":date,"admissions":adm,"births":births,"vaginal":births-cs,"cesareans":cs,"floor_occupancy":occupancy,"floor_pct":round(100*occupancy/int(p["floor_capacity"]),1),"triage_queue":q})
    return {"days":rows,"summary":{"admissions":total_adm,"births":total_births,"cesareans":total_cs,"vaginal":total_births-total_cs,"peak_floor_pct":max((x["floor_pct"] for x in rows),default=0),"peak_triage_queue":peak_q},"parameters":p}

@app.route("/")
def home(): return render_template("index.html")

@app.route("/health")
def health(): return jsonify({"status":"ok","service":"AVICO Gemelo Digital","pheds_positions":len(PHEDS)})

@app.route("/api/config")
def config(): return jsonify(BASE)

@app.route("/api/pheds")
def pheds(): return jsonify({"total":len(PHEDS),"positions":PHEDS})

@app.route("/api/sources")
def sources(): return jsonify(SOURCES)

@app.route("/api/simulate", methods=["POST"])
def api_simulate():
    body=request.get_json(silent=True) or {}
    return jsonify(simulate(int(body.get("days",7)), body.get("parameters")))

@app.route("/api/predict")
def predict():
    base=simulate(1)["days"][0]
    return jsonify({"plus_2h":{"floor_pct":min(100,base["floor_pct"]+2),"triage_queue":base["triage_queue"]},"plus_4h":{"floor_pct":min(100,base["floor_pct"]+5),"triage_queue":base["triage_queue"]+1},"plus_6h":{"floor_pct":min(100,base["floor_pct"]+8),"triage_queue":base["triage_queue"]+2},"warning":"Predicción operativa V1 de simulación; no es un modelo clínico validado."})

if __name__ == "__main__": app.run(host="0.0.0.0", port=10000)
