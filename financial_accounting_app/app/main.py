import hashlib, hmac, io, json, os, re, secrets, shutil, tempfile, zipfile
from datetime import datetime, timezone
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request, Depends, File, UploadFile, Form, Header, HTTPException, Response
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from sqlalchemy import select, func
from sqlalchemy.orm import Session
from .db import SessionLocal, init_db, Teacher, Assignment, Variant, Student, Submission, Audit
from .security import hash_password,verify_password,sign_teacher,read_teacher,token_hash,student_token
from .excel import validate_xlsx, inspect_pair, generate, WorkbookError, _check_config, load_safe, MAX_UPLOAD
from .grading import grade_workbooks, graded_copy, report_xlsx, report_csv

ROOT=Path(__file__).resolve().parents[1];STORAGE=Path(os.getenv('STORAGE_DIR',str(ROOT/'storage')))
STORAGE.mkdir(parents=True,exist_ok=True)

def db_session():
    with SessionLocal() as db:yield db

@asynccontextmanager
async def lifespan(app):
    init_db();yield

app=FastAPI(title='Financial Accounting Studio',lifespan=lifespan)
app.mount('/static', StaticFiles(directory=str(ROOT/'app'/'static')),name='static')

@app.exception_handler(WorkbookError)
async def excel_error(req,exc):return JSONResponse({'detail':str(exc)},status_code=422)

@app.middleware('http')
async def protect_csrf(request:Request,call_next):
    if request.method in ('POST','PUT','PATCH','DELETE') and request.url.path.startswith('/api/'):
        origin=request.headers.get('origin')
        if origin:
            from urllib.parse import urlparse
            host=urlparse(origin).netloc
            if host!=request.headers.get('host'):
                return JSONResponse({'detail':'Недопустимый Origin'},status_code=403)
    return await call_next(request)

def teacher(request:Request,db:Session=Depends(db_session)):
    cookie=request.cookies.get('fa_teacher','')
    tid=read_teacher(cookie)
    t=db.get(Teacher,tid)
    if not t:raise HTTPException(401,'Учетная запись не найдена')
    return t

def own(assignment_id,t,db):
    a=db.get(Assignment,assignment_id)
    if not a or a.teacher_id!=t.id:raise HTTPException(404,'Задание не найдено')
    return a

def logging(db,t,action,details=''):
    db.add(Audit(teacher_id=t.id if t else None,action=action,details=str(details)[:800]))

def assignment_json(a,db,details=False):
    d={'id':a.id,'title':a.title,'topic':a.topic,'instructions':a.instructions,'published':a.published,'deadline':a.deadline,
    'url':f'/s/{a.id}','variants':db.scalar(select(func.count()).select_from(Variant).where(Variant.assignment_id==a.id)) or 0,
    'students':db.scalar(select(func.count()).select_from(Student).where(Student.assignment_id==a.id)) or 0}
    if details:d.update(config=json.loads(a.config_json))
    return d

class SetupData(BaseModel):
    email:str
    password:str=Field(min_length=10)
    setup_token:str
@app.post('/api/setup')
def setup(data:SetupData,db:Session=Depends(db_session)):
    required=os.getenv('SETUP_TOKEN','')
    if not required or not secrets.compare_digest(data.setup_token,required):raise HTTPException(403,'Неверный токен первоначальной настройки')
    if db.scalar(select(func.count()).select_from(Teacher)):raise HTTPException(409,'Преподаватель уже создан')
    t=Teacher(email=data.email.lower().strip(),password_hash=hash_password(data.password));db.add(t);db.commit()
    return {'created':True}
class LoginData(BaseModel):email:str;password:str
@app.post('/api/login')
def login(d:LoginData,response:Response,db:Session=Depends(db_session)):
    user=db.scalar(select(Teacher).where(Teacher.email==d.email.lower().strip()))
    if not user or not verify_password(d.password,user.password_hash):raise HTTPException(401,'Неверная почта или пароль')
    response.set_cookie('fa_teacher',sign_teacher(user.id),httponly=True,samesite='strict',secure=os.getenv('COOKIE_SECURE')=='1',max_age=86400,path='/')
    return {'email':user.email}
@app.post('/api/logout')
def logout(response:Response):response.delete_cookie('fa_teacher',path='/');return {'ok':True}
@app.get('/api/me')
def me(t=Depends(teacher)):return {'id':t.id,'email':t.email}

async def read_upload(u:UploadFile):
    if not u.filename or not u.filename.lower().endswith('.xlsx'):raise HTTPException(422,'Поддерживаются безопасные файлы .xlsx')
    content=await u.read(MAX_UPLOAD+1)
    validate_xlsx(content)
    return content

@app.post('/api/assignments')
async def upload_assignment(title:str=Form(...),topic:str=Form(''),instructions:str=Form(''),task:UploadFile=File(...),key:UploadFile=File(...),t=Depends(teacher),db:Session=Depends(db_session)):
    tbytes=await read_upload(task);kbytes=await read_upload(key)
    slug=secrets.token_hex(12); d=STORAGE/'assignments'/slug;d.mkdir(parents=True)
    p=d/'original_task.xlsx';q=d/'original_key.xlsx';p.write_bytes(tbytes);q.write_bytes(kbytes)
    try:inspection=inspect_pair(p,q)
    except Exception:
        shutil.rmtree(d);raise
    cfg={'inputs':[], 'candidates':inspection['inputs'], 'questions':inspection['questions'], 'constraints':[]}
    a=Assignment(teacher_id=t.id,title=title[:200],topic=topic[:2500],instructions=instructions[:4000],question_path=str(p),key_path=str(q),config_json=json.dumps(cfg,ensure_ascii=False))
    db.add(a);logging(db,t,'upload_assignment',title);db.commit()
    return {**assignment_json(a,db,True),'inspection':inspection}
@app.get('/api/assignments')
def list_assignments(t=Depends(teacher),db:Session=Depends(db_session)):
    return [assignment_json(a,db) for a in db.scalars(select(Assignment).where(Assignment.teacher_id==t.id).order_by(Assignment.id.desc()))]
@app.get('/api/assignments/{aid}')
def one_assignment(aid:int,t=Depends(teacher),db:Session=Depends(db_session)):
    a=own(aid,t,db);return {**assignment_json(a,db,True),'inspection':inspect_pair(a.question_path,a.key_path), 'variants_list':[{'number':v.number,'id':v.id} for v in db.scalars(select(Variant).where(Variant.assignment_id==aid).order_by(Variant.number))]}
@app.put('/api/assignments/{aid}/config')
def save_config(aid:int,payload:dict,t=Depends(teacher),db:Session=Depends(db_session)):
    a=own(aid,t,db)
    if a.published or db.scalar(select(func.count()).select_from(Student).where(Student.assignment_id==aid)):
        raise HTTPException(409,'Нельзя менять правила после публикации или регистрации студентов')
    _check_config(payload,load_safe(a.question_path),load_safe(a.key_path))
    a.config_json=json.dumps({'inputs':payload['inputs'],'questions':payload['questions'],'constraints':payload.get('constraints',[])},ensure_ascii=False)
    logging(db,t,'configure',aid);db.commit();return {'saved':True}
class Count(BaseModel):count:int=Field(ge=1,le=150)
@app.post('/api/assignments/{aid}/generate')
def generate_variants(aid:int,request:Count,t=Depends(teacher),db:Session=Depends(db_session)):
    a=own(aid,t,db)
    if a.published or db.scalar(select(func.count()).select_from(Student).where(Student.assignment_id==aid)):
        raise HTTPException(409,'После публикации или регистрации студентов набор вариантов менять нельзя')
    cfg=json.loads(a.config_json)
    folder=Path(a.question_path).parent
    temp=folder/('new_'+secrets.token_hex(6))
    try:
        manifest=generate(a.question_path,a.key_path,cfg,request.count,temp)
        existing=list(db.scalars(select(Variant).where(Variant.assignment_id==aid)))
        for v in existing:db.delete(v)
        db.flush()
        old=folder/'generated';new=folder/'generated_prev'
        if new.exists():shutil.rmtree(new)
        if old.exists():old.rename(new)
        temp.rename(old)
        for entry in manifest:
            db.add(Variant(assignment_id=aid,number=entry['number'],task_path=str(old/Path(entry['task_path']).name),key_path=str(old/Path(entry['key_path']).name),values_json=json.dumps(entry['values'])))
        logging(db,t,'generate_variants',f'{aid}, count={request.count}');db.commit()
        if new.exists():shutil.rmtree(new)
        return {'generated':len(manifest),'files':[{'variant':v['number'],'task':f'Variant_{v["number"]:03d}_Task.xlsx','key':f'Variant_{v["number"]:03d}_KEY.xlsx'} for v in manifest], 'zip':'All_Variants.zip','combined':'All_Answers.xlsx'}
    except Exception:
        if temp.exists():shutil.rmtree(temp)
        db.rollback();raise
class Publication(BaseModel):published:bool;deadline:str=''
@app.post('/api/assignments/{aid}/publish')
def publish(aid:int,v:Publication,t=Depends(teacher),db:Session=Depends(db_session)):
    a=own(aid,t,db)
    if v.published and not db.scalar(select(func.count()).select_from(Variant).where(Variant.assignment_id==aid)):
        raise HTTPException(409,'Сначала создайте варианты')
    a.published=v.published;a.deadline=v.deadline[:35];logging(db,t,'publish',f'{aid}:{v.published}');db.commit();return {'published':a.published,'student_url':f'/s/{aid}'}
@app.get('/api/assignments/{aid}/file/{name}')
def teacher_file(aid:int,name:str,t=Depends(teacher),db:Session=Depends(db_session)):
    a=own(aid,t,db);folder=Path(a.question_path).parent/'generated'
    allowed={'All_Answers.xlsx','All_Variants.zip','Tasks_Only.zip','Keys_Only.zip'}
    for v in db.scalars(select(Variant).where(Variant.assignment_id==aid)):
        allowed.add(Path(v.task_path).name);allowed.add(Path(v.key_path).name)
    if name not in allowed:raise HTTPException(404,'Файл не найден')
    f=folder/name
    if not f.is_file():raise HTTPException(404,'Файл не найден')
    return FileResponse(f,filename=name,media_type='application/octet-stream')

class Register(BaseModel):assignment_id:int;student_id:str=Field(min_length=2,max_length=120);group:str=Field(min_length=2,max_length=120);full_name:str=''
def student_by_token(auth:str|None,db:Session):
    if not auth or len(auth)>256:raise HTTPException(401,'Нужен индивидуальный токен студента')
    target=token_hash(auth)
    s=db.scalar(select(Student).where(Student.token_hash==target))
    if not s:raise HTTPException(401,'Неверный токен студента')
    return s
@app.get('/api/public/assignments/{aid}')
def public_assignment(aid:int,db:Session=Depends(db_session)):
    a=db.get(Assignment,aid)
    if not a or not a.published:raise HTTPException(404,'Работа недоступна')
    return {'title':a.title,'topic':a.topic,'instructions':a.instructions,'deadline':a.deadline,'assignment_id':aid}
@app.post('/api/student/register')
def register(d:Register,db:Session=Depends(db_session)):
    a=db.get(Assignment,d.assignment_id)
    if not a or not a.published:raise HTTPException(404,'Работа недоступна')
    if a.deadline:
        try:
            if datetime.now(timezone.utc)>datetime.fromisoformat(a.deadline.replace('Z','+00:00')):raise HTTPException(403,'Срок сдачи истёк')
        except ValueError:raise HTTPException(422,'Срок указан неверно')
    uid=d.student_id.strip(); group=d.group.strip()
    if not re.fullmatch(r'[\w\-]{2,120}',uid,re.UNICODE) or not group:raise HTTPException(422,'Некорректный ID или группа')
    if db.scalar(select(Student).where(Student.assignment_id==d.assignment_id,Student.student_id==uid)):
        raise HTTPException(409,'Этот студенческий ID уже зарегистрирован. Обратитесь к преподавателю за восстановлением доступа.')
    vs=list(db.scalars(select(Variant).where(Variant.assignment_id==a.id).order_by(Variant.number)))
    if not vs:raise HTTPException(409,'Варианты отсутствуют')
    cnt=db.scalar(select(func.count()).select_from(Student).where(Student.assignment_id==a.id)) or 0
    v=vs[cnt%len(vs)]
    token=student_token()
    stu=Student(assignment_id=a.id,variant_id=v.id,student_id=uid,group_name=group,full_name=d.full_name.strip()[:250],token_hash=token_hash(token))
    db.add(stu);db.commit()
    return {'token':token,'student_id':uid,'group':group,'variant':v.number,'assignment':a.title}
@app.get('/api/student/me')
def student_me(x_student_token:str|None=Header(None),db:Session=Depends(db_session)):
    s=student_by_token(x_student_token,db);v=db.get(Variant,s.variant_id)
    return {'student_id':s.student_id,'group':s.group_name,'name':s.full_name,'variant':v.number,'assignment_id':s.assignment_id}
@app.get('/api/student/download')
def download_student(x_student_token:str|None=Header(None),db:Session=Depends(db_session)):
    s=student_by_token(x_student_token,db);v=db.get(Variant,s.variant_id)
    return FileResponse(v.task_path,filename=f'{s.student_id}_Variant_{v.number:03d}.xlsx')

async def submit_file(s,content,db):
    a=db.get(Assignment,s.assignment_id);v=db.get(Variant,s.variant_id)
    folder=Path(a.question_path).parent/'submissions';folder.mkdir(exist_ok=True)
    path=folder/f'{s.id}_{secrets.token_hex(8)}.xlsx';path.write_bytes(content)
    try:res=grade_workbooks(path,v.key_path,v.task_path,json.loads(a.config_json))
    except Exception:
        path.unlink(missing_ok=True);raise
    sub=Submission(student_row_id=s.id,file_path=str(path),score=res['score'],maximum=res['maximum'],result_json=json.dumps(res,ensure_ascii=False,default=str),status=res['status'])
    db.add(sub);db.commit();db.refresh(sub)
    return {'submission_id':sub.id,'score':sub.score,'maximum':sub.maximum,'percentage':res['percentage'],'status':sub.status,'rows':res['rows'],'counts':res['counts']}
@app.post('/api/student/submit')
async def student_submit(file:UploadFile=File(...),x_student_token:str|None=Header(None),db:Session=Depends(db_session)):
    s=student_by_token(x_student_token,db);a=db.get(Assignment,s.assignment_id)
    if not a.published:raise HTTPException(403,'Приём работы закрыт')
    if a.deadline and datetime.now(timezone.utc)>datetime.fromisoformat(a.deadline.replace('Z','+00:00')):raise HTTPException(403,'Срок сдачи истёк')
    if db.scalar(select(func.count()).select_from(Submission).where(Submission.student_row_id==s.id))>=int(os.getenv('MAX_SUBMISSIONS','3')):
        raise HTTPException(429,'Достигнут лимит попыток')
    return await submit_file(s,await read_upload(file),db)
@app.get('/api/student/results')
def student_results(x_student_token:str|None=Header(None),db:Session=Depends(db_session)):
    s=student_by_token(x_student_token,db)
    subs=list(db.scalars(select(Submission).where(Submission.student_row_id==s.id).order_by(Submission.id.desc())))
    return [{'id':x.id,'at':x.submitted_at.isoformat(),'result':json.loads(x.result_json)} for x in subs]
@app.get('/api/student/graded-copy/{sid}')
def student_copy(sid:int,x_student_token:str|None=Header(None),db:Session=Depends(db_session)):
    s=student_by_token(x_student_token,db);sub=db.get(Submission,sid)
    if not sub or sub.student_row_id!=s.id:raise HTTPException(404,'Отчёт недоступен')
    data=graded_copy(sub.file_path,json.loads(sub.result_json))
    return StreamingResponse(io.BytesIO(data),media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':f'attachment; filename="graded_{sid}.xlsx"'})

def roster(aid,db):
    rows=[]
    for stu in db.scalars(select(Student).where(Student.assignment_id==aid).order_by(Student.student_id)):
        variant=db.get(Variant,stu.variant_id)
        for sub in db.scalars(select(Submission).where(Submission.student_row_id==stu.id).order_by(Submission.id.desc())):
            res=json.loads(sub.result_json);c=res['counts']
            rows.append({'id':sub.id,'student_id':stu.student_id,'group':stu.group_name,'name':stu.full_name,'variant':variant.number,'submitted':sub.submitted_at.isoformat(),'score':res['score'],'maximum':res['maximum'],'percentage':res['percentage'],'status':sub.status,'correct':c['correct'],'partial':c['partial'],'incorrect':c['incorrect']+c['error'],'missing':c['missing'],'review':c['review'],'rows':res['rows']})
    return rows
@app.get('/api/assignments/{aid}/grades')
def grades(aid:int,t=Depends(teacher),db:Session=Depends(db_session)):
    own(aid,t,db);return roster(aid,db)
@app.get('/api/assignments/{aid}/grades-export')
def grades_export(aid:int,format:str='xlsx',t=Depends(teacher),db:Session=Depends(db_session)):
    own(aid,t,db)
    if format not in ('xlsx','csv'):raise HTTPException(422,'Формат xlsx или csv')
    data=report_csv(roster(aid,db)) if format=='csv' else report_xlsx(roster(aid,db))
    mime='text/csv; charset=utf-8' if format=='csv' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return StreamingResponse(io.BytesIO(data),media_type=mime,headers={'Content-Disposition':f'attachment; filename="Grades_{aid}.{format}"'})
@app.get('/api/submissions/{sid}')
def detail(sid:int,t=Depends(teacher),db:Session=Depends(db_session)):
    sub=db.get(Submission,sid)
    if not sub:raise HTTPException(404,'Нет результата')
    stu=db.get(Student,sub.student_row_id);own(stu.assignment_id,t,db)
    return {'submission_id':sid,'student_id':stu.student_id,'group':stu.group_name,'result':json.loads(sub.result_json),'submitted_at':sub.submitted_at.isoformat()}
class Override(BaseModel):question:str;awarded:float=Field(ge=0);reason:str=Field(min_length=3)
@app.post('/api/submissions/{sid}/override')
def override(sid:int,d:Override,t=Depends(teacher),db:Session=Depends(db_session)):
    sub=db.get(Submission,sid)
    if not sub:raise HTTPException(404,'Нет результата')
    stu=db.get(Student,sub.student_row_id);own(stu.assignment_id,t,db)
    res=json.loads(sub.result_json)
    row=next((r for r in res['rows'] if r['question']==d.question),None)
    if not row:raise HTTPException(404,'Вопрос не найден')
    if d.awarded>row['points']:raise HTTPException(422,'Баллы превышают максимум')
    row['awarded']=d.awarded;row['reason']='Ручная проверка: '+d.reason;row['manual_review']=False
    row['status']='correct' if d.awarded==row['points'] else 'partial' if d.awarded else 'incorrect'
    res['score']=round(sum(r['awarded'] for r in res['rows']),4);res['percentage']=round(100*res['score']/res['maximum'],2)
    res['counts']={s:sum(r['status']==s for r in res['rows']) for s in ['correct','partial','incorrect','missing','review','error']}
    res['status']='review' if any(r['manual_review'] for r in res['rows']) else 'graded'
    sub.result_json=json.dumps(res,ensure_ascii=False,default=str);sub.score=res['score'];sub.status=res['status']
    logging(db,t,'score_override',f'{sid} {d.question} ->{d.awarded}: {d.reason}');db.commit()
    return res
@app.post('/api/submissions/{sid}/regrade')
def regrade(sid:int,t=Depends(teacher),db:Session=Depends(db_session)):
    sub=db.get(Submission,sid)
    if not sub:raise HTTPException(404,'Нет результата')
    stu=db.get(Student,sub.student_row_id);a=own(stu.assignment_id,t,db);v=db.get(Variant,stu.variant_id)
    res=grade_workbooks(sub.file_path,v.key_path,v.task_path,json.loads(a.config_json))
    sub.result_json=json.dumps(res,ensure_ascii=False,default=str);sub.score=res['score'];sub.maximum=res['maximum'];sub.status=res['status']
    logging(db,t,'regrade',sid);db.commit();return res
@app.post('/api/assignments/{aid}/grade-upload')
async def grade_upload(aid:int,student_id:str=Form(...),file:UploadFile=File(...),t=Depends(teacher),db:Session=Depends(db_session)):
    own(aid,t,db);stu=db.scalar(select(Student).where(Student.assignment_id==aid,Student.student_id==student_id))
    if not stu:raise HTTPException(404,'Студент с указанным ID не зарегистрирован')
    result=await submit_file(stu,await read_upload(file),db);logging(db,t,'teacher_upload',f'{aid}:{student_id}');db.commit();return result
@app.post('/api/assignments/{aid}/grade-zip')
async def grade_zip(aid:int,file:UploadFile=File(...),t=Depends(teacher),db:Session=Depends(db_session)):
    own(aid,t,db)
    data=await file.read(40*1024*1024+1)
    if len(data)>40*1024*1024 or not zipfile.is_zipfile(io.BytesIO(data)):raise HTTPException(422,'Ожидается ZIP до 40 МБ')
    results=[]
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        members=[x for x in z.infolist() if x.filename.lower().endswith('.xlsx') and not x.is_dir()]
        if len(members)>100:raise HTTPException(422,'Не более 100 работ')
        if any(m.file_size>MAX_UPLOAD for m in members):raise HTTPException(422,'Один из XLSX превышает 8 МБ')
        if sum(m.file_size for m in members)>100*1024*1024:raise HTTPException(422,'Превышен размер распакованных файлов')
        for entry in members:
            # Identify student by exact StudentID__*.xlsx prefix; ambiguous names are reported, never guessed.
            student_id=Path(entry.filename).name.split('__')[0]
            stu=db.scalar(select(Student).where(Student.assignment_id==aid,Student.student_id==student_id))
            if not stu:
                results.append({'file':entry.filename,'error':'Студент не найден; используйте имя StudentID__filename.xlsx'});continue
            try:
                data=z.read(entry)
                validate_xlsx(data)
                result=await submit_file(stu,data,db)
                results.append({'file':entry.filename,'student_id':student_id,'score':result['score']})
            except Exception as exc:results.append({'file':entry.filename,'error':str(exc)[:180]})
    logging(db,t,'grade_zip',f'{aid}:{len(results)}');db.commit();return {'results':results}
@app.get('/',response_class=HTMLResponse)
def homepage():return FileResponse(ROOT/'app'/'static'/'index.html')
@app.get('/grading',response_class=HTMLResponse)
def grading_page():return FileResponse(ROOT/'app'/'static'/'index.html')
@app.get('/s/{aid}',response_class=HTMLResponse)
def student_page(aid:int):return FileResponse(ROOT/'app'/'static'/'index.html')
