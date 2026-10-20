"""Initial Financial Accounting Studio database schema.
Revision ID: 0001_initial
Revises:
"""
from alembic import op
import sqlalchemy as sa
revision='0001_initial'
down_revision=None
branch_labels=None
depends_on=None

def upgrade():
    # Databases bootstrapped by an older create_all release can be marked as baseline.
    conn=op.get_bind()
    if sa.inspect(conn).has_table('teachers'):
        return
    op.create_table('teachers',sa.Column('id',sa.Integer(),primary_key=True),sa.Column('email',sa.String(255),nullable=False,unique=True),sa.Column('password_hash',sa.String(200),nullable=False))
    op.create_table('assignments',sa.Column('id',sa.Integer(),primary_key=True),sa.Column('teacher_id',sa.Integer(),sa.ForeignKey('teachers.id'),nullable=False),sa.Column('title',sa.String(200),nullable=False),sa.Column('topic',sa.Text(),nullable=False),sa.Column('instructions',sa.Text(),nullable=False),sa.Column('question_path',sa.Text(),nullable=False),sa.Column('key_path',sa.Text(),nullable=False),sa.Column('config_json',sa.Text(),nullable=False),sa.Column('published',sa.Boolean(),nullable=False),sa.Column('deadline',sa.String(40),nullable=False),sa.Column('created_at',sa.DateTime(),nullable=False))
    op.create_table('variants',sa.Column('id',sa.Integer(),primary_key=True),sa.Column('assignment_id',sa.Integer(),sa.ForeignKey('assignments.id'),nullable=False),sa.Column('number',sa.Integer(),nullable=False),sa.Column('task_path',sa.Text(),nullable=False),sa.Column('key_path',sa.Text(),nullable=False),sa.Column('values_json',sa.Text(),nullable=False),sa.UniqueConstraint('assignment_id','number'))
    op.create_table('students',sa.Column('id',sa.Integer(),primary_key=True),sa.Column('assignment_id',sa.Integer(),sa.ForeignKey('assignments.id'),nullable=False),sa.Column('variant_id',sa.Integer(),sa.ForeignKey('variants.id'),nullable=False),sa.Column('student_id',sa.String(120),nullable=False),sa.Column('group_name',sa.String(120),nullable=False),sa.Column('full_name',sa.String(250),nullable=False),sa.Column('token_hash',sa.String(128),nullable=False),sa.Column('created_at',sa.DateTime(),nullable=False),sa.UniqueConstraint('assignment_id','student_id'))
    op.create_table('submissions',sa.Column('id',sa.Integer(),primary_key=True),sa.Column('student_row_id',sa.Integer(),sa.ForeignKey('students.id'),nullable=False),sa.Column('file_path',sa.Text(),nullable=False),sa.Column('score',sa.Float(),nullable=False),sa.Column('maximum',sa.Float(),nullable=False),sa.Column('result_json',sa.Text(),nullable=False),sa.Column('status',sa.String(40),nullable=False),sa.Column('submitted_at',sa.DateTime(),nullable=False))
    op.create_table('audit',sa.Column('id',sa.Integer(),primary_key=True),sa.Column('teacher_id',sa.Integer(),nullable=True),sa.Column('action',sa.String(100),nullable=False),sa.Column('details',sa.Text(),nullable=False),sa.Column('at',sa.DateTime(),nullable=False))

def downgrade():
    for name in ('audit','submissions','students','variants','assignments','teachers'):
        op.drop_table(name)
