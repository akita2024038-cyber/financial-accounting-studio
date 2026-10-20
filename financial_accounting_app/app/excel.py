"""Excel generation and validation. All editable data requires teacher confirmation.
The source OOXML is kept by injecting only LibreOffice-computed cached values.
"""
import io, json, math, os, random, re, shutil, subprocess, tempfile, zipfile, threading
from concurrent.futures import ThreadPoolExecutor
from collections import Counter
from copy import copy
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET
from openpyxl import load_workbook, Workbook
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.formula import Tokenizer

MAX_UPLOAD=8*1024*1024
MAX_CELLS=45000
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DANGER_PARTS=('externalLinks/','embeddings/','activeX/','connections.xml','vbaProject')

class WorkbookError(ValueError):pass

def validate_xlsx(data):
    if len(data)>MAX_UPLOAD:raise WorkbookError('Файл превышает лимит 8 МБ')
    if not zipfile.is_zipfile(io.BytesIO(data)):raise WorkbookError('Не является действительным XLSX')
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names=z.namelist()
            if len(names)>2500 or sum(x.file_size for x in z.infolist())>MAX_UPLOAD*15:
                raise WorkbookError('Подозрительный размер распакованного XLSX')
            if '[Content_Types].xml' not in names or 'xl/workbook.xml' not in names:
                raise WorkbookError('Некорректная структура книги Excel')
            if any(any(p.lower() in n.lower() for p in DANGER_PARTS) for n in names):
                raise WorkbookError('Макросы, внешние подключения и вложенные объекты запрещены')
        wb=load_workbook(io.BytesIO(data),read_only=False, data_only=False,keep_links=False)
        if sum(ws.max_row*ws.max_column for ws in wb)>MAX_CELLS:raise WorkbookError('Слишком большая книга (предел 45 000 ячеек)')
        allowed_functions={'SUM','AVERAGE','MIN','MAX','ROUND','ROUNDUP','ROUNDDOWN','IF','IFERROR','ABS','COUNT','COUNTA','COUNTIF','SUMIF','SUMIFS','AND','OR','NOT','DATE','YEAR','MONTH','DAY','PMT','FV','PV','NPV','IRR','VLOOKUP','INDEX','MATCH','POWER','SQRT','SUBTOTAL','SUMPRODUCT'}
        for ws in wb:
            for row in ws:
                for c in row:
                    if c.data_type=='f':
                        if '[' in c.value or '_xlfn.' in c.value:
                            raise WorkbookError(f'Внешняя или неподдерживаемая формула: {ws.title}!{c.coordinate}')
                        try:
                            for token in Tokenizer(c.value).items:
                                if token.type=='FUNC' and token.subtype=='OPEN' and token.value[:-1].upper() not in allowed_functions:
                                    raise WorkbookError(f'Неподдерживаемая функция: {token.value} в {ws.title}!{c.coordinate}')
                        except WorkbookError:raise
                        except Exception:raise WorkbookError(f'Некорректный синтаксис формулы {ws.title}!{c.coordinate}')
        return wb
    except WorkbookError:raise
    except Exception as e:raise WorkbookError('Ошибка чтения XLSX: '+str(e)[:180])

def load_safe(path):return validate_xlsx(Path(path).read_bytes())
def num(v):return type(v) in (int,float) and math.isfinite(v)
def same(a,b):
    return type(a)==type(b) and a==b

def inspect_pair(task_path,key_path):
    t,k=load_safe(task_path),load_safe(key_path)
    if t.sheetnames!=k.sheetnames:
        raise WorkbookError('Названия и порядок листов задания и решения должны совпадать. Сопоставьте их перед загрузкой.')
    inputs=[]; questions=[]; previews=[]; key_previews=[]
    for st,sk in zip(t,k):
        preview=[]
        for row in st.iter_rows(min_row=1,max_row=min(35,st.max_row),max_col=min(12,st.max_column)):
            preview.append([{'cell':c.coordinate,'value':str(c.value)[:170] if c.value is not None else '', 'type': 'formula' if c.data_type=='f' else ('number' if num(c.value) else 'text')} for c in row if not isinstance(c,MergedCell)])
        previews.append({'sheet':st.title,'grid':preview})
        key_preview=[]
        for row in sk.iter_rows(min_row=1,max_row=min(35,sk.max_row),max_col=min(12,sk.max_column)):
            key_preview.append([{'cell':c.coordinate,'value':str(c.value)[:170] if c.value is not None else '', 'type':'formula' if c.data_type=='f' else ('number' if num(c.value) else 'text')} for c in row if not isinstance(c,MergedCell)])
        key_previews.append({'sheet':sk.title,'grid':key_preview})
        for row in st:
            for c in row:
                if isinstance(c,MergedCell):continue
                kc=sk[c.coordinate]
                if num(c.value) and num(kc.value) and same(c.value,kc.value) and c.value not in (0,1) and c.number_format not in ('yyyy-mm-dd','mm-dd-yy'):
                    base=float(c.value)
                    low=math.floor(base*.8); high=math.ceil(base*1.2)
                    if low==high:high=low+2
                    inputs.append({'sheet':st.title,'cell':c.coordinate,'original':c.value,'min':low,'max':high,'step':1 if float(base).is_integer() else .01,'selected':False})
                if kc.value is not None and (c.value is None or (c.data_type != 'f' and kc.data_type=='f')):
                    questions.append({'id':f'Q{len(questions)+1}', 'sheet':st.title,'answer_sheet':sk.title,'cell':c.coordinate,'answer_cell':c.coordinate,'points':1.0,'formula_required':kc.data_type=='f','tolerance':.01,'description':f'{st.title}!{c.coordinate}'})
    return {'inputs':inputs,'questions':questions,'previews':previews,'key_previews':key_previews,'sheetnames':t.sheetnames,'warnings': ['Числа в текстовых строках и внешние бухгалтерские взаимосвязи не определяются автоматически. Проверьте выбранные входные данные и задайте ограничения.', 'Формулы оцениваются движком LibreOffice; неподдерживаемые формулы блокируют генерацию.']}

def _check_config(cfg,task,key):
    selected=cfg.get('inputs',[])
    if not selected:raise WorkbookError('Выберите минимум одну числовую ячейку для изменения')
    unique=set()
    for x in selected:
        sheet=x['sheet']; cell=x['cell'].upper(); ident=(sheet,cell)
        if ident in unique:raise WorkbookError('Ячейка выбрана дважды')
        unique.add(ident)
        if sheet not in task.sheetnames or not re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}',cell): raise WorkbookError('Недопустимая ячейка')
        origin=task[sheet][cell].value
        if not num(origin) or not same(origin,key[sheet][cell].value):raise WorkbookError(f'Редактирование {sheet}!{cell} небезопасно: нет одинакового числового входа в обеих книгах')
        lo,hi,step=map(float,(x['min'],x['max'],x['step']))
        if not all(map(math.isfinite,(lo,hi,step))) or lo>hi or step<=0 or (hi-lo)/step>1000000:
            raise WorkbookError('Некорректный диапазон значений')
        if int((hi-lo)/step)+1<2:raise WorkbookError(f'Недостаточно возможных значений: {sheet}!{cell}')
    questions=cfg.get('questions',[])
    if not questions:raise WorkbookError('Нужна хотя бы одна ячейка для проверки ответа')
    for q in questions:
        if q['sheet'] not in task.sheetnames or q.get('answer_sheet',q['sheet']) not in key.sheetnames:raise WorkbookError('Лист ответа отсутствует')
        if not (re.fullmatch(r'[A-Z]{1,3}[1-9]\d*',q['cell'].upper()) and re.fullmatch(r'[A-Z]{1,3}[1-9]\d*',q.get('answer_cell',q['cell']).upper())):raise WorkbookError('Неверный адрес ответа')
        answer=key[q.get('answer_sheet',q['sheet'])][q.get('answer_cell',q['cell'])].value
        if answer is None:raise WorkbookError('Ключ не содержит ответа в '+q['answer_cell'])
        if num(answer) and not q.get('constant_independent',False):
            raise WorkbookError(f'{q["sheet"]}!{q["cell"]}: постоянный числовой ответ нельзя безопасно пересчитать. Добавьте формулу в ключ или подтвердите независимость ответа (constant_independent=true).')
        if float(q.get('points',1))<=0 or float(q.get('tolerance',.01))<0:raise WorkbookError('Некорректные баллы или допуск')

def libreoffice_executable():
    for binary in ('libreoffice','soffice','soffice.exe'):
        resolved=shutil.which(binary)
        if resolved:return resolved
    if os.name=='nt':
        for key in ('ProgramFiles','ProgramFiles(x86)'):
            base=os.environ.get(key)
            if base:
                candidate=Path(base)/'LibreOffice'/'program'/'soffice.exe'
                if candidate.is_file():return str(candidate)
    return None

def recalculate_cached(source):
    """Return a new XLSX with source formatting/formulas unchanged and caches from headless LO.
    Raises if LO fails or a formula cell reports an Excel error.
    """
    libreoffice=libreoffice_executable()
    if not libreoffice:raise WorkbookError('LibreOffice Calc требуется для достоверного пересчёта формул; установите его или добавьте soffice в PATH')
    with tempfile.TemporaryDirectory(prefix='fa_recalc_') as tmp:
        original=Path(tmp)/'source.xlsx'; original.write_bytes(source)
        outdir=Path(tmp)/'calc'; outdir.mkdir(); profile=(Path(tmp)/'profile').as_uri()
        if True:
            cmd=[libreoffice,'-env:UserInstallation='+profile,'--headless','--convert-to','xlsx','--outdir',str(outdir),str(original)]
            r=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=38,env={**os.environ,'HOME':tmp})
            calc=outdir/'source.xlsx'
            if r.returncode!=0 or not calc.exists():raise WorkbookError('LibreOffice не смог пересчитать книгу: '+r.stderr.decode(errors='replace')[:200])
            wb=load_workbook(calc,data_only=True)
            source_wb=load_workbook(io.BytesIO(source),data_only=False)
            cache={}
            for src in source_wb:
                results=wb[src.title]
                for row in src:
                    for cell in row:
                        if cell.data_type=='f':
                            val=results[cell.coordinate]
                            if val.data_type=='e' or val.value is None:raise WorkbookError(f'Ошибка вычисления формулы {src.title}!{cell.coordinate}: {val.value}')
                            cache[(src.title,cell.coordinate)]=val.value
            if not cache:return source
            # Map source original package worksheet filenames through workbook rels, not sheet index.
            with zipfile.ZipFile(io.BytesIO(source)) as zin:
                sheets=source_wb.sheetnames
                out=io.BytesIO()
                with zipfile.ZipFile(out,'w') as zout:
                    for info in zin.infolist():
                        data=zin.read(info.filename)
                        m=re.fullmatch(r'xl/worksheets/sheet(\d+)\.xml',info.filename)
                        if m and int(m[1])-1<len(sheets):
                            sheet=sheets[int(m[1])-1]
                            root=ET.fromstring(data); changed=False
                            for node in root.iter(f'{{{NS}}}c'):
                                address=node.attrib.get('r'); cached=cache.get((sheet,address))
                                if cached is None:continue
                                formula=node.find(f'{{{NS}}}f')
                                if formula is None:continue
                                v=node.find(f'{{{NS}}}v')
                                if v is None:v=ET.SubElement(node,f'{{{NS}}}v')
                                if isinstance(cached,bool):node.set('t','b');v.text='1' if cached else '0'
                                elif isinstance(cached,(int,float)):node.attrib.pop('t',None);v.text=str(cached)
                                elif isinstance(cached,str):node.set('t','str');v.text=cached
                                else:raise WorkbookError('Неподдерживаемое значение кеша формулы')
                                changed=True
                            if changed:data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
                        zout.writestr(info,data)
                return out.getvalue()

def verify_constraints(workbook, constraints):
    wb=load_workbook(io.BytesIO(workbook),data_only=True)
    for c in constraints:
        if c['sheet'] not in wb.sheetnames:raise WorkbookError('Неверный лист ограничения')
        ws=wb[c['sheet']]; left=ws[c['left']].value
        right=ws[c['right']].value if c.get('right') else float(c['value'])
        if not num(left) or not num(right):return False
        tol=float(c.get('tolerance',.01)); op=c.get('operator','eq')
        if not {'eq':abs(left-right)<=tol,'le':left<=right+tol,'ge':left+tol>=right}.get(op,False):return False
    return True

def generate(task_path,key_path,cfg,count,output_dir):
    task=load_safe(task_path); key=load_safe(key_path)
    _check_config(cfg,task,key)
    if count<1 or count>150:raise WorkbookError('Количество вариантов: от 1 до 150')
    chosen=cfg['inputs']; questions=cfg['questions']; constraints=cfg.get('constraints',[])
    for c in constraints:
        if c.get('operator') not in ('eq','le','ge') or not c.get('sheet') or not c.get('left') or not (c.get('right') or c.get('value') is not None):
            raise WorkbookError('Неправильное условие проверки: ожидается sheet, left, right/value, operator')
    capacity=math.prod(min(1000000,int((float(x['max'])-float(x['min']))/float(x['step']))+1) for x in chosen)
    if capacity<count:raise WorkbookError(f'Невозможно сделать {count} разных вариантов: доступно комбинаций {capacity}')
    output_dir=Path(output_dir);output_dir.mkdir(parents=True,exist_ok=True)
    rng=random.SystemRandom(); used=set(); lock=threading.Lock()
    def create_one(i):
        for attempt in range(200):
            vals=[]
            for x in chosen:
                low,hi,step=map(Decimal,map(str,(x['min'],x['max'],x['step'])))
                max_n=int((hi-low)/step)
                val=low+step*rng.randint(0,max_n)
                vals.append((x['sheet'],x['cell'].upper(),int(val) if val==val.to_integral_value() else float(val)))
            v_key=tuple(vals)
            with lock:
                if v_key in used:continue
                used.add(v_key)
            t=load_safe(task_path); k=load_safe(key_path)
            for sheet,cell,value in vals:t[sheet][cell]=value;k[sheet][cell]=value
            t_io=io.BytesIO();k_io=io.BytesIO();t.save(t_io);k.save(k_io)
            try: solution=recalculate_cached(k_io.getvalue())
            except WorkbookError as exc:raise WorkbookError(f'Вариант {i}: {exc}')
            if not verify_constraints(solution,constraints):
                with lock:used.discard(v_key)
                continue
            # For question cells ensure all expected results exist; fail fast on invalid solution.
            cached=load_workbook(io.BytesIO(solution),data_only=True)
            for q in questions:
                ans=cached[q.get('answer_sheet',q['sheet'])][q.get('answer_cell',q['cell'])].value
                if ans is None:raise WorkbookError(f'Нет вычисленного ответа: {q["sheet"]}!{q["cell"]}')
            file_task=output_dir/f'Variant_{i:03d}_Task.xlsx'; file_key=output_dir/f'Variant_{i:03d}_KEY.xlsx'
            file_task.write_bytes(t_io.getvalue());file_key.write_bytes(solution)
            return {'number':i,'task_path':str(file_task),'key_path':str(file_key),'values':{s+'!'+c:v for s,c,v in vals}}
        raise WorkbookError(f'Не удалось найти допустимый уникальный вариант {i}; расширьте диапазоны или ослабьте ограничения')
    # LibreOffice instances have private profiles; parallel calculation avoids a long
    # serial subprocess startup for classes with 30+ students.
    with ThreadPoolExecutor(max_workers=min(3,count)) as executor:
        manifest=list(executor.map(create_one,range(1,count+1)))
    summary=Workbook();ws=summary.active;ws.title='Ответы';ws.append(['Вариант','Вопрос','Лист','Ячейка','Правильный ответ','Формула','Макс. баллов'])
    from openpyxl.styles import Font,PatternFill,Alignment
    for v in manifest:
        raw=load_workbook(v['key_path'],data_only=False);cal=load_workbook(v['key_path'],data_only=True)
        for q in questions:
            sh=q.get('answer_sheet',q['sheet']);pos=q.get('answer_cell',q['cell']);f=raw[sh][pos].value
            ws.append([v['number'],q.get('id',''),sh,pos,cal[sh][pos].value,f if isinstance(f,str) and f.startswith('=') else '',float(q.get('points',1))])
    for cell in ws[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='1E3A5F');cell.alignment=Alignment(wrap_text=True)
    for col,w in {'A':13,'B':18,'C':25,'D':16,'E':23,'F':35,'G':17}.items():ws.column_dimensions[col].width=w
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    summary.save(output_dir/'All_Answers.xlsx')
    with zipfile.ZipFile(output_dir/'All_Variants.zip','w',compression=zipfile.ZIP_DEFLATED) as z:
        for f in sorted(output_dir.glob('*.xlsx')):z.write(f,f.name)
    with zipfile.ZipFile(output_dir/'Tasks_Only.zip','w',compression=zipfile.ZIP_DEFLATED) as z:
        for f in sorted(output_dir.glob('Variant_*_Task.xlsx')):z.write(f,f.name)
    with zipfile.ZipFile(output_dir/'Keys_Only.zip','w',compression=zipfile.ZIP_DEFLATED) as z:
        for f in sorted(output_dir.glob('Variant_*_KEY.xlsx')):z.write(f,f.name)
    return manifest
