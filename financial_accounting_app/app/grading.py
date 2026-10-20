"""Cell-aware grading; mathematical formula comparator adapted from BSD PySheetGrader.
Status and CSV aggregation pattern adapted from MIT ExcelLabGrader; its original engine skips XLSX.
"""
import csv, io, json, math, re, sys
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Font
from .excel import load_safe, recalculate_cached, WorkbookError, num
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'third_party'/'pysheetgrader-core-main'))
try:
    from pysheetgrader.formula_parser import parse_formula as pysheetgrader_parse
    from sympy import simplify
    PYSHEETGRADER_AVAILABLE=True
except Exception:
    PYSHEETGRADER_AVAILABLE=False

FORMULA_NORMALIZE=re.compile(r'\s+')
def normalized(f):return FORMULA_NORMALIZE.sub('',f or '').replace('$','').upper()
def symbolic_equivalent(a,b):
    if not PYSHEETGRADER_AVAILABLE:return False
    # PySheetGrader delegates symbolic parsing to SymPy; do not feed its parser
    # untrusted Excel tokens capable of invoking Python expressions.
    from openpyxl.formula import Tokenizer
    from openpyxl.formula.tokenizer import Token
    for formula in (a,b):
        try:
            for t in Tokenizer(formula).items:
                if t.type in ('OP_IN','OP_PRE') and t.value not in ('+','-','*','/','^'):return False
                if t.type == 'FUNC' and t.subtype == 'OPEN' and t.value[:-1].upper() not in ('SUM','ROUND','MIN','MAX','AVERAGE','ABS','SQRT'):return False
                if t.type == 'OPERAND' and t.subtype == 'RANGE' and not re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}',t.value.upper()):return False
                if t.type == 'OPERAND' and t.subtype not in ('RANGE','NUMBER'):return False
                if t.type in ('ARRAY','WSPACE'):return False
        except Exception:return False
    try:
        # Original PySheetGrader engine's symbolic parser. Unsupported functions => unknown.
        return bool(simplify(pysheetgrader_parse(a)-pysheetgrader_parse(b))==0)
    except Exception:return False

def formula_refs(formula):
    from openpyxl.formula import Tokenizer
    from openpyxl.formula.tokenizer import Token
    try:
        return sorted(t.value.upper().replace('$','') for t in Tokenizer(formula).items if t.subtype==Token.RANGE)
    except Exception:return []

def similar_values(actual,expected,tol):
    if num(actual) and num(expected):return math.isclose(float(actual),float(expected),abs_tol=float(tol),rel_tol=1e-9)
    return actual==expected

def grade_workbooks(student_path,key_path,task_path,config):
    raw=load_safe(student_path);key=load_safe(key_path);task=load_safe(task_path)
    # workbook structure and protected input checks
    sheet_issue=raw.sheetnames!=task.sheetnames
    tamper=[]
    for item in config.get('inputs',[]):
        sh,c=item['sheet'],item['cell'].upper()
        if sh not in raw.sheetnames or raw[sh][c].value!=task[sh][c].value:
            tamper.append(f'{sh}!{c}')
    try:
        calc_bytes=recalculate_cached(Path(student_path).read_bytes())
        computed=load_workbook(io.BytesIO(calc_bytes),data_only=True)
        calculation_error=''
    except WorkbookError as exc:
        # Catch LO errors and mark affected answers for manual review, not silently zero all work.
        computed=None;calculation_error=str(exc)
    expected=load_workbook(key_path,data_only=True)
    rows=[]
    for q in config['questions']:
        sheet=q['sheet'];cell=q['cell'].upper();answer_sheet=q.get('answer_sheet',sheet);answer_cell=q.get('answer_cell',cell).upper()
        points=float(q.get('points',1));tolerance=float(q.get('tolerance',.01));need_formula=bool(q.get('formula_required'))
        ev=expected[answer_sheet][answer_cell].value
        ef=key[answer_sheet][answer_cell].value
        ef=ef if isinstance(ef,str) and ef.startswith('=') else None
        studentcell=raw[sheet][cell] if sheet in raw.sheetnames else None
        sf=studentcell.value if studentcell is not None and studentcell.data_type=='f' else None
        sv=computed[sheet][cell].value if computed is not None and sheet in computed.sheetnames else (studentcell.value if studentcell is not None and sf is None else None)
        reason='';suggestion='';status='incorrect';awarded=0.0;manual=False
        if sheet_issue:
            status='review';manual=True;reason='Структура листов изменена';suggestion='Восстановите структуру исходного файла.'
        elif studentcell is None or studentcell.value is None:
            status='missing';reason='Ответ не заполнен';suggestion='Заполните указанную ячейку.'
        elif studentcell.data_type=='e' or (computed is not None and computed[sheet][cell].data_type=='e'):
            status='error';reason='Ошибка вычисления Excel';suggestion='Проверьте деление на ноль и ссылки.'
        elif computed is None and sf:
            status='review';manual=True;reason='Пересчёт не выполнен: '+calculation_error;suggestion='Проверьте книгу в Excel.'
        elif need_formula and not sf:
            if similar_values(sv,ev,tolerance):
                status='partial';awarded=round(points*.25,4);reason='Правильное число введено без обязательной формулы'
            else:
                reason='Вместо формулы введено значение';
            suggestion='Введите вычисляющую формулу, а не готовое число.'
        else:
            value_ok=similar_values(sv,ev,tolerance)
            formula_ok=not need_formula or normalized(sf)==normalized(ef) or (sf and ef and symbolic_equivalent(sf,ef))
            if formula_ok and value_ok:
                status='correct';awarded=points;reason='Формула и результат правильные' if need_formula else 'Ответ правильный'
            elif formula_ok and not value_ok:
                status='partial';awarded=round(points*.5,4);reason='Правильная формула, но неверный результат; проверьте исходные данные';suggestion='Проверьте промежуточные значения.'
            elif value_ok and sf:
                if formula_refs(sf)==formula_refs(ef) and formula_refs(sf):
                    status='review';manual=True;awarded=round(points*.75,4)
                    reason='Результат верен, но формула отличается; эквивалентность не доказана'
                else:
                    status='review';manual=True;awarded=round(points*.5,4)
                    reason='Результат верен, но ссылки в формуле отличаются'
                suggestion='Преподавателю нужно подтвердить альтернативную формулу.'
            else:
                reason='Неверная формула или результат';suggestion='Проверьте ссылки и арифметические операции.'
        if tamper:
            manual=True
            if status=='correct':status='review'
            reason+=f' | Изменены входные ячейки: {", ".join(tamper[:8])}'
        rows.append({'question':q.get('id',''), 'description':q.get('description',''), 'sheet':sheet, 'cell':cell, 'expected_formula':ef or '', 'student_formula':sf or '', 'expected_value':ev,'student_value':sv,'points':points,'awarded':awarded,'status':status,'reason':reason,'suggestion':suggestion,'manual_review':manual})
    score=round(sum(x['awarded'] for x in rows),4);maximum=round(sum(x['points'] for x in rows),4)
    counts={s:sum(x['status']==s for x in rows) for s in ['correct','partial','incorrect','missing','review','error']}
    return {'score':score,'maximum':maximum,'percentage':round(100*score/maximum,2) if maximum else 0,'status':'review' if any(r['manual_review'] for r in rows) else 'graded','rows':rows,'counts':counts,'tampered_inputs':tamper,'formula_engine':'PySheetGrader symbolic comparator + LibreOffice calculated values'}

def graded_copy(path,result):
    wb=load_safe(path)
    fills={'incorrect':'FCE4E4','error':'FCE4E4','missing':'FCE4E4','partial':'FFF0C2','review':'FFF0C2','correct':'DDF5DF'}
    for row in result['rows']:
        if row['sheet'] not in wb.sheetnames:continue
        cell=wb[row['sheet']][row['cell']]
        cell.fill=PatternFill('solid',fgColor=fills.get(row['status'],'FFF0C2'))
        cell.comment=Comment(f"{row['question']}: {row['reason']}\n{row['suggestion']}\nБаллы: {row['awarded']}/{row['points']}",'Financial Accounting Grader')
    name='Проверка'
    if name in wb.sheetnames:name='Grading_Report'
    ws=wb.create_sheet(name)
    for v in [['Проверка работы',''],['Баллы',f"{result['score']} / {result['maximum']}"],['Процент',result['percentage']],[],['Вопрос','Лист','Ячейка','Статус','Баллы','Причина','Исправление']]:ws.append(v)
    for row in result['rows']:ws.append([row['question'],row['sheet'],row['cell'],row['status'],row['awarded'],row['reason'],row['suggestion']])
    for cell in ws[5]:cell.fill=PatternFill('solid',fgColor='153A5B');cell.font=Font(color='FFFFFF',bold=True)
    for col,w in {'A':23,'B':18,'C':12,'D':18,'E':12,'F':60,'G':58}.items():ws.column_dimensions[col].width=w
    out=io.BytesIO();wb.save(out);return out.getvalue()

def csv_safe(x):
    if isinstance(x,str) and x.startswith(('=','+','-','@','\t','\r')):return "'"+x
    return x

def report_csv(results):
    dst=io.StringIO();w=csv.writer(dst)
    w.writerow(['Student ID','Group','Name','Variant','Submitted','Score','Maximum','Percent','Status','Correct','Partial','Incorrect','Missing','Review'])
    for r in results:w.writerow([csv_safe(r.get(k,'')) for k in ('student_id','group','name','variant','submitted','score','maximum','percentage','status','correct','partial','incorrect','missing','review')])
    return dst.getvalue().encode('utf-8-sig')

def report_xlsx(results):
    wb=Workbook();ws=wb.active;ws.title='Оценки'
    headers=['Student ID','Group','Name','Variant','Submitted','Score','Maximum','Percent','Status','Correct','Partial','Incorrect','Missing','Review']
    keys=['student_id','group','name','variant','submitted','score','maximum','percentage','status','correct','partial','incorrect','missing','review']
    ws.append(headers)
    for r in results:ws.append([csv_safe(r.get(k,'')) for k in keys])
    for c in ws[1]:c.fill=PatternFill('solid',fgColor='153A5B');c.font=Font(bold=True,color='FFFFFF')
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    for i,h in enumerate(headers,1):ws.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width=24 if i in (1,3,5) else 16
    out=io.BytesIO();wb.save(out);return out.getvalue()
