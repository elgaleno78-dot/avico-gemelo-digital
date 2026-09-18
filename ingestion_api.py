import io, re, unicodedata
from flask import jsonify, request
import pandas as pd

MAX_BYTES = 20 * 1024 * 1024
ID_WORDS = ('expediente','curp','folio','nombre','paciente','nacimiento','fecha','hora')

def clean_name(value):
    text = unicodedata.normalize('NFD', str(value or '').strip().lower())
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]+', '_', text).strip('_') or 'sin_nombre'

def classify(series):
    s = series.dropna()
    if not len(s): return 'vacía'
    if pd.api.types.is_datetime64_any_dtype(s): return 'fecha/hora'
    if pd.api.types.is_numeric_dtype(s): return 'numérica continua' if s.nunique() > 12 else 'numérica discreta'
    parsed = pd.to_datetime(s.astype(str), errors='coerce', dayfirst=True)
    if len(s) and parsed.notna().mean() >= .80: return 'fecha/hora'
    ratio = s.astype(str).nunique() / max(len(s), 1)
    if s.astype(str).nunique() <= 20 or ratio < .12: return 'categórica'
    return 'identificador/texto' if any(k in clean_name(series.name) for k in ID_WORDS) else 'texto libre'

def read_book(raw, filename):
    ext = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''
    if ext not in ('xlsx','xls'): raise ValueError('Formato no admitido. Use .xlsx o .xls')
    return pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype_backend='numpy_nullable')

def profile(raw, filename):
    book = read_book(raw, filename); sheets=[]; all_columns=set(); total=0
    for sheet_name, df in book.items():
        df = df.dropna(how='all'); total += len(df)
        columns=[]
        for col in df.columns:
            key=clean_name(col); all_columns.add(key); s=df[col]
            columns.append({'original':str(col),'key':key,'type':classify(s),'rows':int(len(s)),'non_null':int(s.notna().sum()),'missing_pct':round(float(s.isna().mean()*100),1),'unique':int(s.nunique(dropna=True))})
        sheets.append({'name':str(sheet_name),'rows':int(len(df)),'columns':len(columns),'variables':columns})
    keys=sorted(k for k in all_columns if any(word in k for word in ID_WORDS))
    return {'file':filename,'sheets':sheets,'rows_total':total,'variables_total':sum(x['columns'] for x in sheets),'candidate_link_keys':keys,'privacy':'No se devuelven valores de pacientes'}

def common_links(items):
    if len(items)<2:return []
    sets=[]
    for item in items:
        cols={v['key'] for s in item['sheets'] for v in s['variables']};sets.append(cols)
    common=set.intersection(*sets)
    preferred=sorted(common,key=lambda x:(not any(w in x for w in ID_WORDS),x))
    return [{'key':k,'strength':'alta' if any(x in k for x in ('expediente','curp','folio')) else 'media' if any(x in k for x in ('nombre','nacimiento')) else 'exploratoria'} for k in preferred[:30]]

def register(app, getbytes):
    @app.post('/api/ingest/upload')
    def upload():
        files=request.files.getlist('files')
        if not files:return jsonify({'error':'Seleccione uno o más archivos Excel'}),400
        if len(files)>8:return jsonify({'error':'Máximo 8 archivos por lote'}),400
        out=[]
        try:
            for f in files:
                raw=f.read(MAX_BYTES+1)
                if len(raw)>MAX_BYTES:raise ValueError(f'{f.filename}: excede 20 MB')
                out.append(profile(raw,f.filename or 'archivo.xlsx'))
            return jsonify({'files':out,'common_link_keys':common_links(out),'next_step':'Revise las variables y confirme las claves antes de vincular pacientes.'})
        except Exception as e:return jsonify({'error':str(e)[:240]}),400

    @app.post('/api/ingest/url')
    def ingest_url():
        body=request.get_json(silent=True) or {};url=str(body.get('url') or '').strip()
        if not url:return jsonify({'error':'Pegue un enlace compartido de Excel o Drive'}),400
        try:
            raw=getbytes(url)
            if len(raw)>MAX_BYTES:raise ValueError('El archivo excede 20 MB')
            return jsonify(profile(raw,body.get('name') or 'fuente_drive.xlsx'))
        except Exception as e:return jsonify({'error':str(e)[:240]}),400
