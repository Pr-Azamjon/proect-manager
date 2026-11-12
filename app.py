"""Proect Manager — система учёта выполненных работ по ТЗ."""
from __future__ import annotations
import os, uuid
from datetime import datetime, date, timedelta
from functools import wraps
from pathlib import Path
from io import BytesIO
from flask import Flask, render_template, redirect, url_for, request, flash, send_file, abort
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from flask_wtf.csrf import CSRFProtect
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException
from sqlalchemy import or_, func
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / '.env')
app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.getenv('SECRET_KEY', 'CHANGE-ME-before-production-2026'),
    SQLALCHEMY_DATABASE_URI=os.getenv('DATABASE_URL', f'sqlite:///{BASE_DIR / "proect_manager.db"}'),
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    MAX_CONTENT_LENGTH=20 * 1024 * 1024,
    UPLOAD_FOLDER=str(BASE_DIR / 'uploads'),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
)
db = SQLAlchemy(app)
login_manager = LoginManager(app); login_manager.login_view = 'login'; login_manager.login_message = 'Войдите в систему.'
csrf = CSRFProtect(app)
Path(app.config['UPLOAD_FOLDER']).mkdir(exist_ok=True)

ROLE_ADMIN='Администратор'; ROLE_HEAD='Руководитель департамента'; ROLE_MANAGER='Менеджер'; ROLE_EMPLOYEE='Сотрудник'
ROLES=[ROLE_ADMIN, ROLE_HEAD, ROLE_MANAGER, ROLE_EMPLOYEE]
STATUS_NEW='Новая'; STATUS_ACCEPTED='Принята'; STATUS_PROGRESS='В работе'; STATUS_REVIEW='На проверке'; STATUS_REWORK='На доработке'; STATUS_DONE='Выполнена'; STATUS_CANCELLED='Отменена'
STATUSES=[STATUS_NEW,STATUS_ACCEPTED,STATUS_PROGRESS,STATUS_REVIEW,STATUS_REWORK,STATUS_DONE,STATUS_CANCELLED]
PRIORITIES=['Низкий','Средний','Высокий','Срочный']

# --- Справочники и основные сущности (10 таблиц по предварительной схеме ТЗ) ---
class Department(db.Model):
    __tablename__='Departments'
    id=db.Column(db.Integer, primary_key=True); name=db.Column(db.String(160), unique=True, nullable=False)
    sections=db.relationship('DepartmentSection', backref='department', cascade='all, delete-orphan')
class DepartmentSection(db.Model):
    __tablename__='DepartmentsSections'
    id=db.Column(db.Integer, primary_key=True); name=db.Column(db.String(160), nullable=False)
    department_id=db.Column(db.Integer, db.ForeignKey('Departments.id'), nullable=False)
    __table_args__=(db.UniqueConstraint('name','department_id'),)
class EmployeeRole(db.Model):
    __tablename__='EmployeeRoles'
    id=db.Column(db.Integer, primary_key=True); name=db.Column(db.String(80), unique=True, nullable=False)
    description=db.Column(db.String(250))
class Employee(UserMixin, db.Model):
    __tablename__='Employees'
    id=db.Column(db.Integer, primary_key=True); fullname=db.Column(db.String(160), nullable=False)
    username=db.Column(db.String(80), unique=True, nullable=False); password_hash=db.Column(db.String(255), nullable=False)
    role_id=db.Column(db.Integer, db.ForeignKey('EmployeeRoles.id'), nullable=False)
    section_id=db.Column(db.Integer, db.ForeignKey('DepartmentsSections.id'))
    manager_id=db.Column(db.Integer, db.ForeignKey('Employees.id'))
    is_active=db.Column(db.Boolean, default=True, nullable=False)
    role=db.relationship('EmployeeRole'); section=db.relationship('DepartmentSection')
    manager=db.relationship('Employee', remote_side=[id], foreign_keys=[manager_id])
    @property
    def role_name(self): return self.role.name if self.role else ''
    def set_password(self,p): self.password_hash=generate_password_hash(p)
    def check_password(self,p): return check_password_hash(self.password_hash,p)
class WorkTaskStatus(db.Model):
    __tablename__='WorkTaskStatuses'
    id=db.Column(db.Integer, primary_key=True); name=db.Column(db.String(60), unique=True, nullable=False); is_final=db.Column(db.Boolean,default=False)
class WorkTask(db.Model):
    __tablename__='WorkTasks'
    id=db.Column(db.Integer, primary_key=True); task_number=db.Column(db.String(30), unique=True, nullable=False)
    title=db.Column(db.String(220), nullable=False); description=db.Column(db.Text, nullable=False)
    expected_result=db.Column(db.Text); result_description=db.Column(db.Text)
    created_at=db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    start_date=db.Column(db.Date, nullable=False); due_date=db.Column(db.Date, nullable=False); original_due_date=db.Column(db.Date, nullable=False)
    priority=db.Column(db.String(30), nullable=False, default='Средний')
    status_id=db.Column(db.Integer, db.ForeignKey('WorkTaskStatuses.id'), nullable=False)
    section_id=db.Column(db.Integer, db.ForeignKey('DepartmentsSections.id'), nullable=False)
    created_by_id=db.Column(db.Integer, db.ForeignKey('Employees.id'), nullable=False)
    status=db.relationship('WorkTaskStatus'); section=db.relationship('DepartmentSection'); creator=db.relationship('Employee', foreign_keys=[created_by_id])
    executors=db.relationship('WorkTaskExecutor', back_populates='task', cascade='all, delete-orphan')
    comments=db.relationship('WorkTaskComment', back_populates='task', cascade='all, delete-orphan', order_by='WorkTaskComment.created_at')
    files=db.relationship('WorkTaskFile', back_populates='task', cascade='all, delete-orphan')
    histories=db.relationship('WorkTaskHistory', back_populates='task', cascade='all, delete-orphan', order_by='WorkTaskHistory.changed_at.desc()')
    @property
    def status_name(self): return self.status.name if self.status else ''
    @property
    def executor_names(self): return ', '.join(x.employee.fullname for x in self.executors if x.employee)
    @property
    def overdue_days(self): return max(0,(date.today()-self.due_date).days) if self.status_name not in (STATUS_DONE,STATUS_CANCELLED) and self.due_date<date.today() else 0
class WorkTaskExecutor(db.Model):
    __tablename__='WorkTaskExecutors'
    id=db.Column(db.Integer,primary_key=True); task_id=db.Column(db.Integer,db.ForeignKey('WorkTasks.id'),nullable=False)
    employee_id=db.Column(db.Integer,db.ForeignKey('Employees.id'),nullable=False)
    task=db.relationship('WorkTask',back_populates='executors'); employee=db.relationship('Employee')
    __table_args__=(db.UniqueConstraint('task_id','employee_id'),)
class WorkTaskHistory(db.Model):
    __tablename__='WorkTaskHistory'
    id=db.Column(db.Integer,primary_key=True); task_id=db.Column(db.Integer,db.ForeignKey('WorkTasks.id'),nullable=False)
    changed_at=db.Column(db.DateTime,default=datetime.utcnow,nullable=False); changed_by_id=db.Column(db.Integer,db.ForeignKey('Employees.id'))
    field_name=db.Column(db.String(80),nullable=False); old_value=db.Column(db.Text); new_value=db.Column(db.Text); reason=db.Column(db.Text)
    task=db.relationship('WorkTask',back_populates='histories'); changed_by=db.relationship('Employee')
class WorkTaskComment(db.Model):
    __tablename__='WorkTaskComments'
    id=db.Column(db.Integer,primary_key=True); task_id=db.Column(db.Integer,db.ForeignKey('WorkTasks.id'),nullable=False)
    employee_id=db.Column(db.Integer,db.ForeignKey('Employees.id'),nullable=False); body=db.Column(db.Text,nullable=False)
    created_at=db.Column(db.DateTime,default=datetime.utcnow,nullable=False)
    task=db.relationship('WorkTask',back_populates='comments'); employee=db.relationship('Employee')
class WorkTaskFile(db.Model):
    __tablename__='WorkTaskFiles'
    id=db.Column(db.Integer,primary_key=True); task_id=db.Column(db.Integer,db.ForeignKey('WorkTasks.id'),nullable=False)
    employee_id=db.Column(db.Integer,db.ForeignKey('Employees.id'),nullable=False); original_name=db.Column(db.String(255),nullable=False)
    stored_name=db.Column(db.String(255),unique=True,nullable=False); uploaded_at=db.Column(db.DateTime,default=datetime.utcnow)
    task=db.relationship('WorkTask',back_populates='files'); employee=db.relationship('Employee')

@login_manager.user_loader
def load_user(uid): return db.session.get(Employee,int(uid))

def role_required(*roles):
    def deco(fn):
        @wraps(fn)
        @login_required
        def wrapped(*a,**kw):
            if current_user.role_name not in roles: abort(403)
            return fn(*a,**kw)
        return wrapped
    return deco

def audit(task, field, old, new, reason=None):
    db.session.add(WorkTaskHistory(task=task,changed_by=current_user,field_name=field,old_value=str(old or ''),new_value=str(new or ''),reason=reason))

def can_view(task):
    r=current_user.role_name
    if r==ROLE_ADMIN: return True
    if not current_user.section_id: return r==ROLE_HEAD and task.section.department_id==getattr(current_user,'department_id',None)
    if r==ROLE_HEAD: return task.section.department_id==current_user.section.department_id
    if r==ROLE_MANAGER: return task.section_id==current_user.section_id
    return any(x.employee_id==current_user.id for x in task.executors)

def scoped_tasks():
    q=WorkTask.query
    r=current_user.role_name
    if r==ROLE_ADMIN: return q
    if r==ROLE_HEAD:
        if not current_user.section: return q.filter(False)
        ids=[s.id for s in DepartmentSection.query.filter_by(department_id=current_user.section.department_id)]
        return q.filter(WorkTask.section_id.in_(ids))
    if r==ROLE_MANAGER: return q.filter_by(section_id=current_user.section_id)
    return q.join(WorkTaskExecutor).filter(WorkTaskExecutor.employee_id==current_user.id)

def visible_departments():
    if current_user.role_name == ROLE_ADMIN: return Department.query.order_by(Department.name).all()
    if current_user.section: return [current_user.section.department]
    return []

def visible_sections():
    if current_user.role_name == ROLE_ADMIN: return DepartmentSection.query.order_by(DepartmentSection.name).all()
    if current_user.role_name == ROLE_HEAD and current_user.section:
        return DepartmentSection.query.filter_by(department_id=current_user.section.department_id).all()
    if current_user.role_name == ROLE_MANAGER and current_user.section_id:
        return DepartmentSection.query.filter_by(id=current_user.section_id).all()
    return []

def visible_employees():
    if current_user.role_name == ROLE_ADMIN: return Employee.query.filter_by(is_active=True).order_by(Employee.fullname).all()
    if current_user.role_name == ROLE_HEAD and current_user.section:
        section_ids=[s.id for s in DepartmentSection.query.filter_by(department_id=current_user.section.department_id).all()]
        return Employee.query.filter(Employee.is_active.is_(True),Employee.section_id.in_(section_ids)).order_by(Employee.fullname).all()
    if current_user.role_name == ROLE_MANAGER and current_user.section_id:
        return Employee.query.filter_by(is_active=True,section_id=current_user.section_id).order_by(Employee.fullname).all()
    return [current_user] if current_user.is_active else []

def get_task(task_id):
    t=WorkTask.query.get_or_404(task_id)
    if not can_view(t): abort(403)
    return t

def validate_org_assignment(role_name, section_id, manager_id):
    if role_name != ROLE_ADMIN and not section_id:
        raise ValueError('Для этой роли необходимо назначить отдел.')
    if role_name == ROLE_EMPLOYEE:
        if not manager_id:
            raise ValueError('Для сотрудника необходимо указать непосредственного менеджера.')
        manager = db.session.get(Employee, manager_id)
        if not manager or manager.role_name != ROLE_MANAGER or manager.section_id != section_id or not manager.is_active:
            raise ValueError('Непосредственный менеджер должен быть активным менеджером того же отдела.')
    if manager_id:
        manager = db.session.get(Employee, manager_id)
        if not manager or not manager.is_active:
            raise ValueError('Выбранный руководитель не найден или отключён.')
        if role_name == ROLE_MANAGER:
            section = db.session.get(DepartmentSection, section_id)
            if manager.role_name != ROLE_HEAD or not section or not manager.section or manager.section.department_id != section.department_id:
                raise ValueError('Непосредственным руководителем менеджера может быть руководитель того же департамента.')

def change_status(task,new_status,reason=None,result=None):
    allowed={STATUS_NEW:{STATUS_ACCEPTED,STATUS_CANCELLED},STATUS_ACCEPTED:{STATUS_PROGRESS,STATUS_CANCELLED},STATUS_PROGRESS:{STATUS_REVIEW,STATUS_CANCELLED},STATUS_REVIEW:{STATUS_DONE,STATUS_REWORK},STATUS_REWORK:{STATUS_PROGRESS,STATUS_REVIEW,STATUS_CANCELLED}}
    old=task.status_name
    if new_status not in allowed.get(old,set()): raise ValueError(f'Переход «{old}» → «{new_status}» запрещён.')
    if new_status==STATUS_REWORK and not (reason or '').strip(): raise ValueError('Для возврата на доработку обязательно укажите причину.')
    if new_status==STATUS_REVIEW and not (result or '').strip(): raise ValueError('Перед отправкой на проверку опишите результат работы.')
    if result is not None: task.result_description=result
    task.status=WorkTaskStatus.query.filter_by(name=new_status).one()
    audit(task,'Статус',old,new_status,reason)
    db.session.commit()

@app.context_processor
def inject_globals(): return {'ROLES':ROLES,'STATUSES':STATUSES,'PRIORITIES':PRIORITIES,'ROLE_ADMIN':ROLE_ADMIN,'ROLE_HEAD':ROLE_HEAD,'ROLE_MANAGER':ROLE_MANAGER,'ROLE_EMPLOYEE':ROLE_EMPLOYEE,'today':date.today(),'week_later':date.today()+timedelta(days=7)}

@app.route('/login',methods=['GET','POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('dashboard'))
    if request.method=='POST':
        user=Employee.query.filter_by(username=request.form.get('username','').strip(),is_active=True).first()
        if user and user.check_password(request.form.get('password','')):
            login_user(user); return redirect(request.args.get('next') or url_for('dashboard'))
        flash('Неверный логин или пароль.','error')
    return render_template('login.html')
@app.route('/logout',methods=['POST'])
@login_required
def logout(): logout_user(); return redirect(url_for('login'))
@app.route('/')
@login_required
def dashboard():
    tasks=scoped_tasks().all(); active=[t for t in tasks if t.status_name not in (STATUS_DONE,STATUS_CANCELLED)]
    counts={s:sum(t.status_name==s for t in tasks) for s in STATUSES}
    employee_load=[]
    if current_user.role_name in (ROLE_ADMIN,ROLE_HEAD,ROLE_MANAGER):
        employees=Employee.query.filter_by(is_active=True).all()
        for e in employees:
            if current_user.role_name==ROLE_MANAGER and e.section_id!=current_user.section_id: continue
            if current_user.role_name==ROLE_HEAD and (not e.section or not current_user.section or e.section.department_id!=current_user.section.department_id): continue
            employee_load.append((e,sum(any(x.employee_id==e.id for x in t.executors) and t.status_name not in (STATUS_DONE,STATUS_CANCELLED) for t in tasks)))
    return render_template('dashboard.html',tasks=tasks[:12],total_count=len(tasks),counts=counts,overdue=[t for t in active if t.overdue_days],due_soon=[t for t in active if 0 <= (t.due_date-date.today()).days <= 3],employee_load=employee_load)
@app.route('/tasks')
@login_required
def tasks():
    q=scoped_tasks()
    status=request.args.get('status'); priority=request.args.get('priority'); section=request.args.get('section',type=int)
    if status: q=q.join(WorkTaskStatus).filter(WorkTaskStatus.name==status)
    if priority: q=q.filter(WorkTask.priority==priority)
    if section and current_user.role_name in (ROLE_ADMIN,ROLE_HEAD): q=q.filter_by(section_id=section)
    search=request.args.get('q','').strip()
    if search: q=q.filter(or_(WorkTask.title.ilike(f'%{search}%'),WorkTask.task_number.ilike(f'%{search}%')))
    return render_template('tasks.html',tasks=q.order_by(WorkTask.created_at.desc()).all(),sections=visible_sections())
@app.route('/tasks/new',methods=['GET','POST'])
@role_required(ROLE_ADMIN,ROLE_MANAGER)
def task_new():
    if request.method=='POST':
        try:
            section_id=request.form.get('section_id',type=int) or current_user.section_id
            section=DepartmentSection.query.get_or_404(section_id)
            if current_user.role_name==ROLE_MANAGER and section.id!=current_user.section_id: abort(403)
            start=date.fromisoformat(request.form['start_date']); due=date.fromisoformat(request.form['due_date'])
            if due<start: raise ValueError('Срок выполнения не может быть раньше даты начала.')
            ids=[int(x) for x in request.form.getlist('executor_ids')]
            if not ids: raise ValueError('Назначьте хотя бы одного исполнителя.')
            people=Employee.query.filter(Employee.id.in_(ids),Employee.is_active.is_(True),Employee.section_id==section.id).all()
            if len(people)!=len(set(ids)): raise ValueError('Исполнители должны быть активными сотрудниками выбранного отдела.')
            task=WorkTask(task_number='PM-'+datetime.utcnow().strftime('%Y%m%d')+'-'+uuid.uuid4().hex[:6].upper(),title=request.form['title'].strip(),description=request.form['description'].strip(),expected_result=request.form.get('expected_result','').strip(),start_date=start,due_date=due,original_due_date=due,priority=request.form['priority'],status=WorkTaskStatus.query.filter_by(name=STATUS_NEW).one(),section=section,creator=current_user)
            db.session.add(task); db.session.flush()
            for p in people: db.session.add(WorkTaskExecutor(task=task,employee=p))
            audit(task,'Создание','',task.title,'Задача создана')
            db.session.commit(); flash(f'Задача {task.task_number} создана.','success'); return redirect(url_for('task_detail',task_id=task.id))
        except HTTPException:
            db.session.rollback(); raise
        except Exception as e: db.session.rollback(); flash(str(e),'error')
    sections=DepartmentSection.query.all() if current_user.role_name==ROLE_ADMIN else DepartmentSection.query.filter_by(id=current_user.section_id).all()
    employees=Employee.query.filter_by(is_active=True).all()
    if current_user.role_name==ROLE_MANAGER: employees=[e for e in employees if e.section_id==current_user.section_id]
    return render_template('task_form.html',sections=sections,employees=employees)
@app.route('/tasks/<int:task_id>')
@login_required
def task_detail(task_id): return render_template('task_detail.html',task=get_task(task_id))
@app.route('/tasks/<int:task_id>/status',methods=['POST'])
@login_required
def task_status(task_id):
    t=get_task(task_id); new=request.form.get('status'); reason=request.form.get('reason','').strip(); result=request.form.get('result','').strip()
    # Исполнитель может принять, начать, отправить на проверку; руководитель отдела проверяет.
    if current_user.role_name==ROLE_EMPLOYEE:
        if not any(x.employee_id==current_user.id for x in t.executors): abort(403)
        if new not in (STATUS_ACCEPTED,STATUS_PROGRESS,STATUS_REVIEW): abort(403)
    elif current_user.role_name not in (ROLE_ADMIN,ROLE_MANAGER): abort(403)
    try: change_status(t,new,reason,result if new==STATUS_REVIEW else None); flash('Статус обновлён.','success')
    except Exception as e: db.session.rollback(); flash(str(e),'error')
    return redirect(url_for('task_detail',task_id=t.id))
@app.route('/tasks/<int:task_id>/deadline',methods=['POST'])
@role_required(ROLE_ADMIN,ROLE_MANAGER)
def task_deadline(task_id):
    t=get_task(task_id); reason=request.form.get('reason','').strip()
    try:
        due=date.fromisoformat(request.form['due_date'])
        if not reason: raise ValueError('Укажите причину изменения срока.')
        if due<t.start_date: raise ValueError('Срок не может быть раньше даты начала.')
        old=t.due_date; t.due_date=due; audit(t,'Срок выполнения',old,due,reason); db.session.commit(); flash('Срок изменён, изменение записано в историю.','success')
    except Exception as e: db.session.rollback(); flash(str(e),'error')
    return redirect(url_for('task_detail',task_id=t.id))
@app.route('/tasks/<int:task_id>/comment',methods=['POST'])
@login_required
def task_comment(task_id):
    t=get_task(task_id); body=request.form.get('body','').strip()
    if body:
        db.session.add(WorkTaskComment(task=t,employee=current_user,body=body)); audit(t,'Комментарий','',body[:180],'Добавлен комментарий'); db.session.commit()
    return redirect(url_for('task_detail',task_id=t.id)+'#comments')
@app.route('/tasks/<int:task_id>/upload',methods=['POST'])
@login_required
def task_upload(task_id):
    t=get_task(task_id); f=request.files.get('file')
    if not f or not f.filename: flash('Выберите файл.','error')
    else:
        original=secure_filename(f.filename)[:240]; stored=uuid.uuid4().hex+'_'+original
        f.save(Path(app.config['UPLOAD_FOLDER'])/stored); db.session.add(WorkTaskFile(task=t,employee=current_user,original_name=original,stored_name=stored)); audit(t,'Файл','',original,'Файл прикреплён'); db.session.commit(); flash('Файл прикреплён.','success')
    return redirect(url_for('task_detail',task_id=t.id))
@app.route('/files/<int:file_id>')
@login_required
def download_file(file_id):
    f=WorkTaskFile.query.get_or_404(file_id)
    if not can_view(f.task): abort(403)
    return send_file(Path(app.config['UPLOAD_FOLDER'])/f.stored_name,as_attachment=True,download_name=f.original_name)
@app.route('/calendar')
@login_required
def calendar():
    month=request.args.get('month',date.today().strftime('%Y-%m'))
    try: first=date.fromisoformat(month+'-01')
    except ValueError: first=date.today().replace(day=1)
    next_month=(first.replace(day=28)+timedelta(days=4)).replace(day=1)
    items=scoped_tasks().filter(WorkTask.due_date>=first,WorkTask.due_date<next_month).order_by(WorkTask.due_date).all()
    return render_template('calendar.html',items=items,month=first.strftime('%Y-%m'),month_label=first.strftime('%m.%Y'))
@app.route('/reports')
@role_required(ROLE_ADMIN,ROLE_HEAD,ROLE_MANAGER)
def reports():
    q=scoped_tasks(); start=request.args.get('start'); end=request.args.get('end'); employee_id=request.args.get('employee',type=int); section_id=request.args.get('section',type=int); department_id=request.args.get('department',type=int); status=request.args.get('status'); priority=request.args.get('priority')
    if start:
        try:q=q.filter(WorkTask.created_at>=datetime.combine(date.fromisoformat(start),datetime.min.time()))
        except ValueError: pass
    if end:
        try:q=q.filter(WorkTask.created_at<datetime.combine(date.fromisoformat(end)+timedelta(days=1),datetime.min.time()))
        except ValueError: pass
    if employee_id:q=q.join(WorkTaskExecutor).filter(WorkTaskExecutor.employee_id==employee_id)
    if section_id:q=q.filter(WorkTask.section_id==section_id)
    if department_id:q=q.filter(WorkTask.section_id.in_([s.id for s in DepartmentSection.query.filter_by(department_id=department_id).all()]))
    if status:q=q.join(WorkTaskStatus).filter(WorkTaskStatus.name==status)
    if priority:q=q.filter(WorkTask.priority==priority)
    result=q.distinct().all()
    staff=[]
    for e in visible_employees():
        staff.append((e,sum(any(x.employee_id==e.id for x in t.executors) for t in result),sum(any(x.employee_id==e.id for x in t.executors) and t.status_name==STATUS_DONE for t in result)))
    department_summary=[]
    for d in visible_departments():
        dep_tasks=[t for t in result if t.section.department_id==d.id]
        department_summary.append((d,len(dep_tasks),sum(t.status_name==STATUS_DONE for t in dep_tasks),sum(t.overdue_days>0 for t in dep_tasks)))
    return render_template('reports.html',tasks=result,staff=staff,employees=visible_employees(),sections=visible_sections(),departments=visible_departments(),department_summary=department_summary)
@app.route('/reports/excel')
@role_required(ROLE_ADMIN,ROLE_HEAD,ROLE_MANAGER)
def reports_excel():
    q=scoped_tasks(); start=request.args.get('start'); end=request.args.get('end'); employee_id=request.args.get('employee',type=int); section_id=request.args.get('section',type=int); department_id=request.args.get('department',type=int); status=request.args.get('status'); priority=request.args.get('priority')
    if start:
        try: q=q.filter(WorkTask.created_at>=datetime.combine(date.fromisoformat(start),datetime.min.time()))
        except ValueError: pass
    if end:
        try: q=q.filter(WorkTask.created_at<datetime.combine(date.fromisoformat(end)+timedelta(days=1),datetime.min.time()))
        except ValueError: pass
    if employee_id: q=q.join(WorkTaskExecutor).filter(WorkTaskExecutor.employee_id==employee_id)
    if section_id: q=q.filter(WorkTask.section_id==section_id)
    if department_id: q=q.filter(WorkTask.section_id.in_([s.id for s in DepartmentSection.query.filter_by(department_id=department_id).all()]))
    if status: q=q.join(WorkTaskStatus).filter(WorkTaskStatus.name==status)
    if priority: q=q.filter(WorkTask.priority==priority)
    items=q.distinct().all(); wb=Workbook(); ws=wb.active; ws.title='Задачи'
    headers=['Номер','Название','Департамент','Отдел','Исполнители','Статус','Приоритет','Дата создания','Дата начала','Первоначальный срок','Текущий срок','Дней просрочки','Результат']
    ws.append(headers)
    for t in items: ws.append([t.task_number,t.title,t.section.department.name,t.section.name,t.executor_names,t.status_name,t.priority,t.created_at.strftime('%d.%m.%Y'),t.start_date.strftime('%d.%m.%Y'),t.original_due_date.strftime('%d.%m.%Y'),t.due_date.strftime('%d.%m.%Y'),t.overdue_days,t.result_description or ''])
    staff=wb.create_sheet('Сводка по сотрудникам'); staff.append(['Сотрудник','Отдел','Всего задач','Принятые менеджером','На доработке','Просроченные'])
    for e in visible_employees():
        own=[t for t in items if any(x.employee_id==e.id for x in t.executors)]
        if own: staff.append([e.fullname,e.section.name if e.section else '',len(own),sum(t.status_name==STATUS_DONE for t in own),sum(t.status_name==STATUS_REWORK for t in own),sum(t.overdue_days>0 for t in own)])
    dep_sheet=wb.create_sheet('Сводка по отделам'); dep_sheet.append(['Департамент','Всего задач','Принятые менеджером','Просроченные'])
    for d in visible_departments():
        dep_tasks=[t for t in items if t.section.department_id==d.id]
        dep_sheet.append([d.name,len(dep_tasks),sum(t.status_name==STATUS_DONE for t in dep_tasks),sum(t.overdue_days>0 for t in dep_tasks)])
    overdue=wb.create_sheet('Просроченные'); overdue.append(headers)
    for t in items:
        if t.overdue_days: overdue.append([t.task_number,t.title,t.section.department.name,t.section.name,t.executor_names,t.status_name,t.priority,t.created_at.strftime('%d.%m.%Y'),t.start_date.strftime('%d.%m.%Y'),t.original_due_date.strftime('%d.%m.%Y'),t.due_date.strftime('%d.%m.%Y'),t.overdue_days,t.result_description or ''])
    for sheet in wb.worksheets:
        sheet.freeze_panes='A2'; sheet.auto_filter.ref=sheet.dimensions
        for c in sheet[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='244B70'); c.alignment=Alignment(wrap_text=True)
        for col in sheet.columns:
            letter=col[0].column_letter; sheet.column_dimensions[letter].width=min(max(max(len(str(c.value or '')) for c in col)+2,12),38)
    bio=BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio,as_attachment=True,download_name='Proect_Manager_Report.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/admin')
@role_required(ROLE_ADMIN)
def admin(): return render_template('admin.html',employees=Employee.query.order_by(Employee.id).all(),departments=Department.query.all(),sections=DepartmentSection.query.all())
@app.route('/admin/employee',methods=['POST'])
@role_required(ROLE_ADMIN)
def admin_employee():
    try:
        role_name=request.form['role']; section_id=request.form.get('section_id',type=int) or None; manager_id=request.form.get('manager_id',type=int) or None
        validate_org_assignment(role_name,section_id,manager_id)
        e=Employee(fullname=request.form['fullname'].strip(),username=request.form['username'].strip(),role=EmployeeRole.query.filter_by(name=role_name).one(),section_id=section_id,manager_id=manager_id,is_active=True)
        e.set_password(request.form['password']); db.session.add(e); db.session.commit(); flash('Пользователь создан.','success')
    except Exception as ex: db.session.rollback(); flash(f'Не удалось создать пользователя: {ex}','error')
    return redirect(url_for('admin'))
@app.route('/admin/employee/<int:employee_id>/update',methods=['POST'])
@role_required(ROLE_ADMIN)
def admin_employee_update(employee_id):
    e=Employee.query.get_or_404(employee_id)
    try:
        role_name=request.form['role']; section_id=request.form.get('section_id',type=int) or None; manager_id=request.form.get('manager_id',type=int) or None
        if manager_id == e.id: raise ValueError('Пользователь не может быть собственным руководителем.')
        validate_org_assignment(role_name,section_id,manager_id)
        e.fullname=request.form['fullname'].strip(); e.role=EmployeeRole.query.filter_by(name=role_name).one()
        e.section_id=section_id; e.manager_id=manager_id
        db.session.commit(); flash('Данные пользователя обновлены.','success')
    except Exception as ex: db.session.rollback(); flash(f'Не удалось обновить пользователя: {ex}','error')
    return redirect(url_for('admin'))
@app.route('/admin/employee/<int:employee_id>/toggle',methods=['POST'])
@role_required(ROLE_ADMIN)
def admin_employee_toggle(employee_id):
    e=Employee.query.get_or_404(employee_id)
    if e.id==current_user.id: flash('Нельзя отключить собственную учётную запись.','error')
    else: e.is_active=not e.is_active; db.session.commit(); flash('Статус учётной записи изменён.','success')
    return redirect(url_for('admin'))
@app.route('/admin/employee/<int:employee_id>/password',methods=['POST'])
@role_required(ROLE_ADMIN)
def admin_employee_password(employee_id):
    e=Employee.query.get_or_404(employee_id); password=request.form.get('password','')
    if len(password)<8: flash('Пароль должен содержать не менее 8 символов.','error')
    else: e.set_password(password); db.session.commit(); flash('Пароль пользователя обновлён.','success')
    return redirect(url_for('admin'))

@app.route('/admin/department',methods=['POST'])
@role_required(ROLE_ADMIN)
def admin_department():
    try: db.session.add(Department(name=request.form['name'].strip())); db.session.commit(); flash('Департамент добавлен.','success')
    except Exception as e: db.session.rollback(); flash(str(e),'error')
    return redirect(url_for('admin'))
@app.route('/admin/section',methods=['POST'])
@role_required(ROLE_ADMIN)
def admin_section():
    try: db.session.add(DepartmentSection(name=request.form['name'].strip(),department_id=int(request.form['department_id']))); db.session.commit(); flash('Отдел добавлен.','success')
    except Exception as e: db.session.rollback(); flash(str(e),'error')
    return redirect(url_for('admin'))
@app.errorhandler(403)
def forbidden(e): return render_template('error.html',code=403,message='Недостаточно прав для этого действия.'),403
@app.errorhandler(404)
def not_found(e): return render_template('error.html',code=404,message='Страница или запись не найдена.'),404

@app.cli.command('init-db')
def init_db():
    """Создать таблицы и демонстрационные учётные записи (если БД пустая)."""
    db.create_all()
    for name in ROLES:
        if not EmployeeRole.query.filter_by(name=name).first(): db.session.add(EmployeeRole(name=name))
    if not WorkTaskStatus.query.first():
        for s in STATUSES: db.session.add(WorkTaskStatus(name=s,is_final=s in (STATUS_DONE,STATUS_CANCELLED)))
    db.session.commit()
    if not Department.query.first():
        dep=Department(name='IT Департамент'); db.session.add(dep); db.session.flush()
        sec=DepartmentSection(name='Отдел разработки',department=dep); db.session.add(sec); db.session.flush()
        role_map={r.name:r for r in EmployeeRole.query.all()}
        accounts=[('Системный администратор','admin',ROLE_ADMIN),('Руководитель департамента','director',ROLE_HEAD),('Менеджер отдела','manager',ROLE_MANAGER),('Фарход','farhod',ROLE_EMPLOYEE),('Мадина','madina',ROLE_EMPLOYEE),('Нигора','nigora',ROLE_EMPLOYEE)]
        users={}
        for fullname,username,role in accounts:
            u=Employee(fullname=fullname,username=username,role=role_map[role],section=sec,is_active=True); u.set_password({'admin':'Admin123!','director':'Director123!','manager':'Manager123!','farhod':'Farhod123!','madina':'Madina123!','nigora':'Nigora123!'}[username]); db.session.add(u); users[username]=u
        db.session.flush(); users['manager'].manager_id=users['director'].id; users['farhod'].manager_id=users['manager'].id; users['madina'].manager_id=users['manager'].id; users['nigora'].manager_id=users['manager'].id
        db.session.commit()
    print('База данных инициализирована. Демо-пароли см. README.md')

with app.app_context():
    init_db()

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=os.getenv('FLASK_DEBUG', '0') == '1')

