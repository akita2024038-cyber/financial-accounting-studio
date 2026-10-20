import hashlib, hmac, os, secrets, time
import jwt
from fastapi import HTTPException

def hash_password(p):
    salt=os.urandom(16); digest=hashlib.scrypt(p.encode(),salt=salt,n=2**14,r=8,p=1)
    return salt.hex()+':'+digest.hex()
def verify_password(p, stored):
    try:
        salt,expected=stored.split(':'); actual=hashlib.scrypt(p.encode(),salt=bytes.fromhex(salt),n=2**14,r=8,p=1)
        return hmac.compare_digest(actual,bytes.fromhex(expected))
    except Exception: return False
def jwt_secret():
    k=os.environ.get('APP_SECRET','')
    if len(k)<32 or k.startswith('REPLACE_'): raise RuntimeError('Set APP_SECRET to a random string of 32+ characters in .env')
    return k
def sign_teacher(tid): return jwt.encode({'sub':str(tid),'exp':int(time.time())+86400,'role':'teacher'},jwt_secret(),algorithm='HS256')
def read_teacher(token):
    try:
        d=jwt.decode(token,jwt_secret(),algorithms=['HS256'])
        if d.get('role')!='teacher':raise ValueError()
        return int(d['sub'])
    except Exception: raise HTTPException(401,'Авторизация преподавателя обязательна')
def student_token(): return secrets.token_urlsafe(36)
def token_hash(t): return hashlib.sha256(t.encode()).hexdigest()
