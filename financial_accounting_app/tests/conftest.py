import os,tempfile
from pathlib import Path
os.environ['APP_SECRET']='a'*80
os.environ['SETUP_TOKEN']='test-setup-token'
os.environ['DATABASE_URL']='sqlite:///'+str(Path(tempfile.gettempdir())/'financial_accounting_test.db')
os.environ['STORAGE_DIR']=str(Path(tempfile.gettempdir())/'fa_test_storage')
import pytest
from fastapi.testclient import TestClient
from app.db import Base,engine
from app.main import app
@pytest.fixture
def client():
    Base.metadata.drop_all(bind=engine);Base.metadata.create_all(bind=engine)
    with TestClient(app) as c:yield c
