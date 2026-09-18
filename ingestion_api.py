import io, re, unicodedata, datetime, statistics, hashlib, tempfile, os
from collections import Counter
from flask import jsonify, request
import pandas as pd
from openpyxl import load_workbook
import gdown

MAX_BYTES = 20 * 1024 * 1024
SAMPLE_ROWS = 5000
EMPTY_STOP = 80
ID_WORDS = ('expediente','curp','folio','nombre','paciente','nacimiento','fecha','hora')
DRIVE_CACHE={}

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

def classify_values(name, values):
    vals=[v for v in values if v is not None and str(v).strip()!='']
    if not vals:return 'vacía'
    if sum(isinstance(v,(datetime.date,datetime.datetime,datetime.time)) for v in vals)/len(vals)>=.8:return 'fecha/hora'
    numeric=sum(isinstance(v,(int,float)) and not isinstance(v,bool) for v in vals)
    unique=len({str(v)[:200] for v in vals})
    if numeric/len(vals)>=.8:return 'numérica continua' if unique>12 else 'numérica discreta'
    if unique<=20 or unique/max(len(vals),1)<.12:return 'categórica'
    return 'identificador/texto' if any(k in clean_name(name) for k in ID_WORDS) else 'texto libre'

def safe_value(value):
    if isinstance(value,(datetime.datetime,datetime.date,datetime.time)):return value.isoformat()
    if isinstance(value,float):return round(value,2)
    return str(value)[:100]

def summarize_values(name, values, kind):
    vals=[v for v in values if v is not None and str(v).strip()!='']
    if not vals:return {'summary':'Sin datos'}
    if kind.startswith('numérica'):
        nums=[float(v) for v in vals if isinstance(v,(int,float)) and not isinstance(v,bool)]
        if not nums:return {'summary':'Sin valores numéricos válidos'}
        modes=statistics.multimode(nums);mode=modes[0] if len(modes)==1 else None
        return {'mean':round(statistics.fmean(nums),2),'median':round(statistics.median(nums),2),'mode':round(mode,2) if mode is not None else 'Sin moda única','min':round(min(nums),2),'max':round(max(nums),2),'summary':f'Media {statistics.fmean(nums):.2f} · Mediana {statistics.median(nums):.2f} · Moda {mode if mode is not None else "no única"}'}
    if kind=='fecha/hora':
        ordered=sorted(safe_value(v) for v in vals);return {'min':ordered[0],'max':ordered[-1],'summary':f'{ordered[0]} → {ordered[-1]}'}
    if kind in ('categórica','numérica discreta'):
        counts=Counter(safe_value(v) for v in vals);top=[]
        for value,count in counts.most_common(8):top.append({'value':value,'count':count,'percent':round(100*count/len(vals),1)})
        return {'mode':top[0]['value'] if top else None,'frequencies':top,'summary':' · '.join(f'{x["value"]}: {x["count"]} ({x["percent"]}%)' for x in top[:3])}
    return {'summary':'Variable identificadora o texto libre; no se muestran valores'}

def read_book(raw, filename):
    ext = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''
    if ext not in ('xlsx','xls'): raise ValueError('Formato no admitido. Use .xlsx o .xls')
    return pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype_backend='numpy_nullable')

def profile_xlsx(raw, filename):
    wb=load_workbook(io.BytesIO(raw),read_only=True,data_only=True);sheets=[];all_columns=set();total=0
    for ws in wb.worksheets:
        header=None;header_row=0;last_col=0
        rows=ws.iter_rows(min_row=1,max_row=30,values_only=True)
        for rn,row in enumerate(rows,1):
            if row is None:break
            populated=[i for i,v in enumerate(row,1) if v is not None and str(v).strip()!='']
            if len(populated)>=2:
                header_row=rn;last_col=max(populated);header=[str(v).strip() if v is not None and str(v).strip() else f'Columna {i+1}' for i,v in enumerate(row[:last_col])];break
        if not header:
            sheets.append({'name':ws.title,'rows':0,'columns':0,'variables':[],'sampled':False});continue
        cols=[[] for _ in header];sampled=0;empty_run=0
        for row in ws.iter_rows(min_row=header_row+1,max_col=last_col,values_only=True):
            if not any(v is not None and str(v).strip()!='' for v in row):
                empty_run+=1
                if empty_run>=EMPTY_STOP:break
                continue
            empty_run=0
            if sampled>=SAMPLE_ROWS:break
            sampled += 1
            for i in range(len(header)):cols[i].append(row[i] if i<len(row) else None)
        actual=sampled;total+=actual;variables=[]
        for name,values in zip(header,cols):
            key=clean_name(name);all_columns.add(key);non_null=sum(v is not None and str(v).strip()!='' for v in values)
            kind=classify_values(name,values)
            variables.append({'original':name,'key':key,'type':kind,'rows':actual,'non_null':non_null,'missing_pct':round(100*(actual-non_null)/max(actual,1),1),'unique':len({str(v)[:200] for v in values if v is not None}),'sample_size':actual,'statistics':summarize_values(name,values,kind)})
        sheets.append({'name':ws.title,'rows':actual,'columns':len(variables),'variables':variables,'sampled':actual>=SAMPLE_ROWS,'sample_size':actual,'declared_dimensions':f'{ws.max_row} × {ws.max_column}','detected_dimensions':f'{actual} × {last_col}'})
    wb.close();keys=sorted(k for k in all_columns if any(word in k for word in ID_WORDS))
    return {'file':filename,'sheets':sheets,'rows_total':total,'variables_total':sum(x['columns'] for x in sheets),'candidate_link_keys':keys,'privacy':'No se devuelven identificadores individuales','method':f'Detección de rango real; máximo {SAMPLE_ROWS} registros por hoja'}

def profile(raw, filename):
    if filename.lower().endswith('.xlsx'):return profile_xlsx(raw,filename)
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

    @app.post('/api/drive/folder')
    def drive_folder():
        body=request.get_json(silent=True) or {};url=str(body.get('url') or '').strip()
        if not re.search(r'drive\.google\.com/drive/(?:u/\d+/)?folders/[A-Za-z0-9_-]+',url):return jsonify({'error':'Use un enlace válido de carpeta de Google Drive'}),400
        try:
            found=gdown.download_folder(url=url,skip_download=True,quiet=True,remaining_ok=True) or []
            items=[];counts={};mapping={}
            for f in found:
                if not str(f.path).lower().endswith(('.xlsx','.xls')):continue
                month=(str(f.path).split('/')[0] or 'SIN MES').strip().upper();counts[month]=counts.get(month,0)+1
                token=hashlib.sha256(str(f.id).encode()).hexdigest()[:20];mapping[token]=str(f.id)
                items.append({'token':token,'month':month,'alias':f'HC-{counts[month]:04d}.xlsx'})
            cache_key=hashlib.sha256(url.encode()).hexdigest()[:16];DRIVE_CACHE[cache_key]=mapping
            return jsonify({'folder_key':cache_key,'total_files':len(items),'months':[{'month':m,'files':n} for m,n in sorted(counts.items())],'files':items,'privacy':'Nombres originales ocultos; archivos pseudonimizados en pantalla'})
        except Exception as e:return jsonify({'error':'No fue posible enumerar la carpeta: '+str(e)[:180]}),400

    @app.post('/api/drive/analyze')
    def drive_analyze():
        body=request.get_json(silent=True) or {};folder_key=str(body.get('folder_key') or '');token=str(body.get('token') or '')
        file_id=DRIVE_CACHE.get(folder_key,{}).get(token)
        if not file_id:return jsonify({'error':'La sesión de Drive expiró. Vuelva a conectar la carpeta.'}),400
        path=None
        try:
            fd,path=tempfile.mkstemp(suffix='.xlsx');os.close(fd)
            result=gdown.download(id=file_id,output=path,quiet=True)
            if not result:raise ValueError('Drive no entregó el archivo')
            with open(path,'rb') as h:raw=h.read(MAX_BYTES+1)
            if len(raw)>MAX_BYTES:raise ValueError('El archivo excede 20 MB')
            return jsonify(profile(raw,'Historia clínica pseudonimizada.xlsx'))
        except Exception as e:return jsonify({'error':'No fue posible analizar el archivo de Drive: '+str(e)[:180]}),400
        finally:
            if path and os.path.exists(path):
                try:os.unlink(path)
                except:pass
