import base64,datetime,io,re
from urllib.parse import urlsplit,urlunsplit,parse_qsl,urlencode
import requests
import app as core


def _xlsx(data):
    return bool(data and data[:2]==b'PK')


def _query_download(url):
    p=urlsplit(url.strip()); q=dict(parse_qsl(p.query,keep_blank_values=True)); q['download']='1'
    return urlunsplit((p.scheme,p.netloc,p.path,urlencode(q),''))


def _graph_share(url):
    token=base64.urlsafe_b64encode(url.encode()).decode().rstrip('=')
    return 'https://api.onedrive.com/v1.0/shares/u!'+token+'/root/content'


def getbytes(url):
    if not url:return None
    url=url.strip(); candidates=[]
    if any(x in url.lower() for x in ('onedrive.live.com','1drv.ms','sharepoint.com')):
        candidates += [_query_download(url),_graph_share(url)]
    candidates.append(url)
    last=''
    for u in candidates:
        try:
            r=requests.get(u,timeout=35,allow_redirects=True,headers={'User-Agent':'Mozilla/5.0','Accept':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*'})
            r.raise_for_status()
            if _xlsx(r.content):return r.content
            last='respuesta no XLSX: '+(r.headers.get('content-type') or 'sin content-type')
        except Exception as e:last=str(e)
    raise ValueError('La fuente no entregó un XLSX descargable: '+last[:140])


def _bed_state(ws):
    expected={str(300+i):f'{i:02d}' for i in range(1,25)}
    expected.update({f'AISL{i}':f'{24+i:02d}' for i in range(1,4)})
    seen={}; labor=0
    for row in ws.iter_rows(values_only=True):
        if len(row)<3:continue
        raw=row[0]
        if isinstance(raw,(int,float)) and float(raw).is_integer():bed=str(int(raw))
        else:bed=str(raw or '').strip().upper()
        service=str(row[1] or '').strip().upper()
        if bed in expected and bed not in seen and 'GINECO' in service:
            seen[bed]=bool(str(row[2] or '').strip())
        txt=' | '.join(str(x or '') for x in row[:8]).upper()
        if ('LABOR' in txt or 'TOCOCIR' in txt) and any(str(x or '').strip() for x in row[2:5]):labor+=1
    return expected,seen,labor


def parse_census(wb):
    valid=[]
    for ws in wb.worksheets:
        m=re.fullmatch(r'0?(\d{1,2})',ws.title.strip())
        if not m:continue
        try:d=datetime.date(core.today().year,core.today().month,int(m.group(1)))
        except:continue
        if d>core.today():continue
        expected,seen,labor=_bed_state(ws)
        # A real daily census must contain the expected GO bed block; blank future/template sheets are skipped.
        if len(seen)<10:continue
        occupied=sorted((expected[k] for k,v in seen.items() if v),key=int)
        valid.append((d,occupied,labor,len(seen)))
    if not valid:return None
    d,occupied,labor,recognized=max(valid,key=lambda x:x[0])
    return {'value':len(occupied),'occupied_beds':occupied,'recognized_beds':recognized,'labor':labor,'last_valid_date':d.isoformat(),'status':'REAL · EN VIVO'}

# Patch only source transport/parsing. Routes and privacy behavior remain in app.py.
core.getbytes=getbytes
core.parse_census=parse_census
app=core.app
