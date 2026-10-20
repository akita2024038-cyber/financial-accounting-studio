import os
from datetime import datetime, timezone
from sqlalchemy import create_engine, String, Integer, ForeignKey, DateTime, Text, Boolean, UniqueConstraint
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, sessionmaker

DATABASE_URL = os.environ.get('DATABASE_URL', 'sqlite:///./storage/accounting.db')
engine = create_engine(DATABASE_URL, connect_args={'check_same_thread': False} if DATABASE_URL.startswith('sqlite') else {}, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
def utcnow(): return datetime.now(timezone.utc).replace(tzinfo=None)
class Base(DeclarativeBase): pass
class Teacher(Base):
    __tablename__='teachers'
    id: Mapped[int]=mapped_column(Integer,primary_key=True)
    email: Mapped[str]=mapped_column(String(255),unique=True)
    password_hash: Mapped[str]=mapped_column(String(200))
class Assignment(Base):
    __tablename__='assignments'
    id: Mapped[int]=mapped_column(Integer,primary_key=True)
    teacher_id: Mapped[int]=mapped_column(ForeignKey('teachers.id'))
    title: Mapped[str]=mapped_column(String(200))
    topic: Mapped[str]=mapped_column(Text, default='')
    instructions: Mapped[str]=mapped_column(Text, default='')
    question_path: Mapped[str]=mapped_column(Text)
    key_path: Mapped[str]=mapped_column(Text)
    config_json: Mapped[str]=mapped_column(Text, default='{}')
    published: Mapped[bool]=mapped_column(Boolean, default=False)
    deadline: Mapped[str]=mapped_column(String(40),default='')
    created_at: Mapped[datetime]=mapped_column(DateTime,default=utcnow)
    variants: Mapped[list['Variant']]=relationship(back_populates='assignment',cascade='all, delete-orphan')
class Variant(Base):
    __tablename__='variants'
    __table_args__=(UniqueConstraint('assignment_id','number'),)
    id: Mapped[int]=mapped_column(Integer,primary_key=True)
    assignment_id: Mapped[int]=mapped_column(ForeignKey('assignments.id'))
    number: Mapped[int]=mapped_column(Integer)
    task_path: Mapped[str]=mapped_column(Text)
    key_path: Mapped[str]=mapped_column(Text)
    values_json: Mapped[str]=mapped_column(Text,default='{}')
    assignment: Mapped[Assignment]=relationship(back_populates='variants')
class Student(Base):
    __tablename__='students'
    __table_args__=(UniqueConstraint('assignment_id','student_id'),)
    id: Mapped[int]=mapped_column(Integer,primary_key=True)
    assignment_id: Mapped[int]=mapped_column(ForeignKey('assignments.id'))
    variant_id: Mapped[int]=mapped_column(ForeignKey('variants.id'))
    student_id: Mapped[str]=mapped_column(String(120))
    group_name: Mapped[str]=mapped_column(String(120))
    full_name: Mapped[str]=mapped_column(String(250),default='')
    token_hash: Mapped[str]=mapped_column(String(128))
    created_at: Mapped[datetime]=mapped_column(DateTime,default=utcnow)
class Submission(Base):
    __tablename__='submissions'
    id: Mapped[int]=mapped_column(Integer,primary_key=True)
    student_row_id: Mapped[int]=mapped_column(ForeignKey('students.id'))
    file_path: Mapped[str]=mapped_column(Text)
    score: Mapped[float]=mapped_column(default=0.0)
    maximum: Mapped[float]=mapped_column(default=0.0)
    result_json: Mapped[str]=mapped_column(Text,default='{}')
    status: Mapped[str]=mapped_column(String(40),default='received')
    submitted_at: Mapped[datetime]=mapped_column(DateTime,default=utcnow)
class Audit(Base):
    __tablename__='audit'
    id: Mapped[int]=mapped_column(Integer,primary_key=True)
    teacher_id: Mapped[int]=mapped_column(Integer,nullable=True)
    action: Mapped[str]=mapped_column(String(100))
    details: Mapped[str]=mapped_column(Text,default='')
    at: Mapped[datetime]=mapped_column(DateTime,default=utcnow)
def init_db():
    Base.metadata.create_all(engine)
