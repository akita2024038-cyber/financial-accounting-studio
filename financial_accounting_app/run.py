"""Convenient local launcher; reads .env if available. Never generates credentials."""
import os
from pathlib import Path
for line in (Path(__file__).parent/'.env').read_text().splitlines() if (Path(__file__).parent/'.env').exists() else []:
    if '=' in line and not line.lstrip().startswith('#'):
        k,v=line.split('=',1);os.environ.setdefault(k.strip(),v.strip())
if len(os.getenv('APP_SECRET',''))<32 or not os.getenv('SETUP_TOKEN') or os.environ['APP_SECRET'].startswith('REPLACE_') or os.environ['SETUP_TOKEN'].startswith('REPLACE_'):
    raise SystemExit('Configure APP_SECRET (32+ chars) and SETUP_TOKEN in .env before starting.')
from alembic import command
from alembic.config import Config
command.upgrade(Config(str(Path(__file__).parent/'alembic.ini')), 'head')
import uvicorn
uvicorn.run('app.main:app',host='127.0.0.1',port=int(os.getenv('PORT','8000')),reload=False)
