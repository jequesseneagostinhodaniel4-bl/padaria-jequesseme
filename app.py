import os, re, io
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote
from dotenv import load_dotenv
load_dotenv()
from fastapi import FastAPI, Depends, HTTPException
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.security import OAuth2PasswordBearer
from jose import jwt
from pwdlib import PasswordHash
from pydantic import BaseModel
from sqlalchemy import create_engine, String, Integer, DateTime, Text, Boolean, Float
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker, Session
from fpdf import FPDF
import openpyxl

SECRET = os.getenv("SECRET_KEY", "PADARIA_JEQUESSENE_2026")
DBURL = os.getenv("DATABASE_URL", "sqlite:///./padaria.db")
engine = create_engine(DBURL, connect_args={"check_same_thread": False} if DBURL.startswith("sqlite") else {})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
pwd = PasswordHash.recommended()
def agora(): return datetime.now(timezone.utc)
class Base(DeclarativeBase): pass

class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(primary_key=True)
    nome: Mapped[str] = mapped_column(String(120))
    telefone: Mapped[str] = mapped_column(String(30), unique=True, index=True)
    idade: Mapped[int] = mapped_column(Integer, default=18)
    email: Mapped[str] = mapped_column(String(180), default="")
    foto: Mapped[str] = mapped_column(Text, default="")
    pais: Mapped[str] = mapped_column(String(60), default="Moçambique")
    bairro: Mapped[str] = mapped_column(String(80), default="Matola")
    tipo: Mapped[str] = mapped_column(String(30), default="Cliente")
    password_hash: Mapped[str] = mapped_column(String(255))
    is_admin: Mapped[bool] = mapped_column(Boolean, default=False)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Product(Base):
    __tablename__ = "products"
    id: Mapped[int] = mapped_column(primary_key=True)
    nome: Mapped[str] = mapped_column(String(120))
    categoria: Mapped[str] = mapped_column(String(60), default="Pães")
    preco: Mapped[float] = mapped_column(Float, default=0)
    custo: Mapped[float] = mapped_column(Float, default=0)
    stock: Mapped[int] = mapped_column(Integer, default=0)
    stock_minimo: Mapped[int] = mapped_column(Integer, default=5)
    foto: Mapped[str] = mapped_column(Text, default="")
    updated_by: Mapped[str] = mapped_column(String(120), default="")
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Historico(Base):
    __tablename__ = "historico"
    id: Mapped[int] = mapped_column(primary_key=True)
    tipo: Mapped[str] = mapped_column(String(20), default="SAIDA")
    produto: Mapped[str] = mapped_column(String(120), default="")
    quantidade: Mapped[int] = mapped_column(Integer, default=0)
    valor: Mapped[float] = mapped_column(Float, default=0)
    motivo: Mapped[str] = mapped_column(String(20), default="VENDA")
    data: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Fornecedor(Base):
    __tablename__ = "fornecedores"
    id: Mapped[int] = mapped_column(primary_key=True)
    nome: Mapped[str] = mapped_column(String(120))
    telefone: Mapped[str] = mapped_column(String(30), default="")
    produto: Mapped[str] = mapped_column(String(120), default="")
    preco: Mapped[float] = mapped_column(Float, default=0)
    localizacao: Mapped[str] = mapped_column(String(120), default="")

class Venda(Base):
    __tablename__ = "vendas"
    id: Mapped[int] = mapped_column(primary_key=True)
    cliente: Mapped[str] = mapped_column(String(120), default="Cliente Balcão")
    cliente_tel: Mapped[str] = mapped_column(String(30), default="")
    vendedor: Mapped[str] = mapped_column(String(120), default="")
    produto: Mapped[str] = mapped_column(String(120))
    qtd: Mapped[int] = mapped_column(Integer, default=1)
    total: Mapped[float] = mapped_column(Float, default=0)
    custo_total: Mapped[float] = mapped_column(Float, default=0)
    lucro: Mapped[float] = mapped_column(Float, default=0)
    pagamento: Mapped[str] = mapped_column(String(40), default="M-Pesa")
    estado: Mapped[str] = mapped_column(String(30), default="Pago")
    data_pagamento: Mapped[str] = mapped_column(String(20), default="")
    whatsapp_link: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Despesa(Base):
    __tablename__ = "despesas"
    id: Mapped[int] = mapped_column(primary_key=True)
    descricao: Mapped[str] = mapped_column(String(180))
    valor: Mapped[float] = mapped_column(Float, default=0)
    categoria: Mapped[str] = mapped_column(String(60), default="Geral")
    criado_por: Mapped[str] = mapped_column(String(120), default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Compra(Base):
    __tablename__ = "compras"
    id: Mapped[int] = mapped_column(primary_key=True)
    fornecedor: Mapped[str] = mapped_column(String(120))
    produto: Mapped[str] = mapped_column(String(120))
    qtd: Mapped[int] = mapped_column(Integer, default=1)
    preco_unitario: Mapped[float] = mapped_column(Float, default=0)
    total: Mapped[float] = mapped_column(Float, default=0)
    pagamento: Mapped[str] = mapped_column(String(40), default="Dinheiro")
    comprado_por: Mapped[str] = mapped_column(String(120), default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Encomenda(Base):
    __tablename__ = "encomendas"
    id: Mapped[int] = mapped_column(primary_key=True)
    cliente: Mapped[str] = mapped_column(String(120))
    telefone: Mapped[str] = mapped_column(String(30), default="")
    tipo: Mapped[str] = mapped_column(String(120))
    data_entrega: Mapped[str] = mapped_column(String(20), default="")
    preco: Mapped[float] = mapped_column(Float, default=0)
    sinal: Mapped[str] = mapped_column(String(60), default="Sem Sinal")
    obs: Mapped[str] = mapped_column(Text, default="")
    estado: Mapped[str] = mapped_column(String(30), default="Pendente")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

class Log(Base):
    __tablename__ = "logs"
    id: Mapped[int] = mapped_column(primary_key=True)
    quem: Mapped[str] = mapped_column(String(120), default="")
    tipo: Mapped[str] = mapped_column(String(30), default="")
    acao: Mapped[str] = mapped_column(String(80))
    detalhe: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=agora)

Base.metadata.create_all(engine)

def dbdep():
    d = SessionLocal()
    try: yield d
    finally: d.close()
def log_action(d, quem, tipo, acao, detalhe=""):
    try:
        d.add(Log(quem=quem, tipo=tipo, acao=acao, detalhe=detalhe[:600])); d.commit()
    except Exception: d.rollback()
def make_token(u):
    return jwt.encode({"sub": str(u.id), "exp": datetime.now(timezone.utc) + timedelta(hours=24)}, SECRET, algorithm="HS256")
def clean_tel(tel): return re.sub(r"\D", "", tel or "")
def normalizar_tipo(tipo):
    t = (tipo or "Cliente").strip().lower()
    if "funcion" in t: return "Funcionário"
    if "admin" in t: return "Admin"
    return "Cliente"
def is_funcionario(u): return u.is_admin or normalizar_tipo(u.tipo) == "Funcionário"
def whatsapp_link(tel, msg):
    tel = clean_tel(tel)
    if not tel: return ""
    if tel.startswith("0"): tel = "258" + tel[1:]
    if not tel.startswith("258"): tel = "258" + tel
    return f"https://wa.me/{tel}?text={quote(msg)}"
def pagar_mpesa(telefone, valor): return {"ok": True, "metodo": "M-Pesa", "telefone": telefone, "valor": valor, "ref": f"MPESA-{datetime.now().strftime('%Y%m%d%H%M%S')}"}
def pagar_emola(telefone, valor): return {"ok": True, "metodo": "e-Mola", "telefone": telefone, "valor": valor, "ref": f"EMOLA-{datetime.now().strftime('%Y%m%d%H%M%S')}"}

def init_data():
    d = SessionLocal()
    try:
        if not d.query(User).filter_by(telefone="840000000").first():
            d.add(User(nome="Jequessene", telefone="840000000", idade=18, email="jequessene@padaria.co.mz", pais="Moçambique", bairro="Matola", tipo="Admin", password_hash=pwd.hash(os.getenv("ADMIN_PASSWORD", "Jequessene12345")), is_admin=True)); d.commit()
        if d.query(Product).count() == 0:
            dados = [("Pão Caseiro","Pães",80,30,40),("Pão de Forma","Pães",150,70,20),("Bolo de Chocolate","Bolos",800,400,5),("Bolo de Aniversário","Bolos",2500,1200,2),("Pastel de Nata (6)","Pastéis",300,150,20),("Croissant","Pastéis",120,60,15),("Cupcake","Doces",100,45,18),("Donut","Doces",90,40,12),("Chamuça","Salgados",60,25,30),("Sumo Natural","Bebidas",120,50,15)]
            for p in dados: d.add(Product(nome=p[0], categoria=p[1], preco=p[2], custo=p[3], stock=p[4]))
            d.commit()
        if d.query(Fornecedor).count() == 0:
            for f in [("Farinha Matola","841234567","Farinha 50kg",1800,"Matola"),("Ovos Maputo","849876543","Ovos 30 un",900,"Maputo"),("Açúcar Matola","842222333","Açúcar 25kg",1500,"Matola")]:
                d.add(Fornecedor(nome=f[0], telefone=f[1], produto=f[2], preco=f[3], localizacao=f[4]))
            d.commit()
    finally: d.close()
init_data()

oauth = OAuth2PasswordBearer(tokenUrl="/api/login")
def current_user(tok: str = Depends(oauth), d: Session = Depends(dbdep)):
    try: uid = int(jwt.decode(tok, SECRET, algorithms=["HS256"])["sub"])
    except Exception: raise HTTPException(401, "Sessão inválida.")
    u = d.get(User, uid)
    if not u or not u.is_active: raise HTTPException(401, "Conta indisponível.")
    return u
def admin_only(u: User = Depends(current_user)):
    if not u.is_admin: raise HTTPException(403, "Só Admin Jequessene.")
    return u
def funcionario_only(u: User = Depends(current_user)):
    if not is_funcionario(u): raise HTTPException(403, "Só Funcionário / Admin.")
    return u

app = FastAPI(title="PADARIA JEQUESSENE")
app.mount("/static", StaticFiles(directory="static"), name="static")
@app.get("/", response_class=HTMLResponse)
def home(): return open("static/index.html", encoding="utf-8").read()
@app.get("/api/health")
def health(): return {"status": "ok", "app": "PADARIA JEQUESSENE"}

class Register(BaseModel):
    nome: str; telefone: str; idade: int = 18; email: str = ""; foto: str = ""; pais: str = "Moçambique"; bairro: str = "Matola"; tipo: str = "Cliente"; password: str
class Login(BaseModel):
    telefone: str; password: str
class SwitchIn(BaseModel):
    target_telefone: str; admin_password: str
class ProdIn(BaseModel):
    nome: str; categoria: str = "Pães"; preco: float = 0; custo: float = 0; stock: int = 0; stock_minimo: int = 5; foto: str = ""
class FornIn(BaseModel):
    nome: str; telefone: str = ""; produto: str = ""; preco: float = 0; localizacao: str = ""
class CompraIn(BaseModel):
    fornecedor_id: int; produto: str; qtd: int = 1; preco_unitario: float = 0; pagamento: str = "Dinheiro"
class VendaIn(BaseModel):
    produto_id: int; qtd: int = 1; cliente: str = "Cliente Balcão"; cliente_tel: str = ""; pagamento: str = "M-Pesa"; data_pagamento: str = ""
class DespesaIn(BaseModel):
    descricao: str; valor: float; categoria: str = "Geral"
class EncIn(BaseModel):
    cliente: str; telefone: str = ""; tipo: str = "Bolo de Aniversário"; data_entrega: str = ""; preco: float = 0; sinal: str = "Sem Sinal"; obs: str = ""

@app.post("/api/register")
def register(x: Register, d: Session = Depends(dbdep)):
    tel = clean_tel(x.telefone)
    if len(tel) < 9: raise HTTPException(400, "Telefone inválido.")
    if d.query(User).filter_by(telefone=tel).first(): raise HTTPException(409, "Telefone já registado.")
    if not x.password or len(x.password) < 4: raise HTTPException(400, "Senha curta, usa 4+ caracteres.")
    u = User(nome=x.nome.strip(), telefone=tel, idade=x.idade, email=x.email, foto=x.foto, pais=x.pais, bairro=x.bairro, tipo=normalizar_tipo(x.tipo), password_hash=pwd.hash(x.password))
    d.add(u); d.commit(); log_action(d, u.nome, u.tipo, "registo", f"{u.nome} - {tel}")
    return {"token": make_token(u), "user": {"nome": u.nome, "tipo": u.tipo}}
@app.post("/api/login")
def login(x: Login, d: Session = Depends(dbdep)):
    tel = clean_tel(x.telefone); u = d.query(User).filter_by(telefone=tel).first()
    if not u: raise HTTPException(401, "Telefone ou senha errada.")
    try: ok = pwd.verify(x.password, u.password_hash)
    except Exception: ok = False
    if not ok: raise HTTPException(401, "Telefone ou senha errada.")
    if not u.is_active: raise HTTPException(403, "Conta bloqueada.")
    log_action(d, u.nome, u.tipo, "login", "Entrou")
    return {"token": make_token(u), "user": {"nome": u.nome, "tipo": u.tipo, "is_admin": u.is_admin}}
@app.get("/api/me")
def me(u: User = Depends(current_user)):
    return {"id": u.id, "nome": u.nome, "telefone": u.telefone, "idade": u.idade, "email": u.email, "foto": u.foto, "pais": u.pais, "bairro": u.bairro, "tipo": u.tipo, "is_admin": u.is_admin}
@app.get("/api/pessoas")
def pessoas(u: User = Depends(current_user), d: Session = Depends(dbdep)):
    lista = d.query(User).all() if u.is_admin else d.query(User).filter_by(tipo="Cliente").all()
    return [{"id": x.id, "nome": x.nome, "telefone": x.telefone, "idade": x.idade, "email": x.email, "foto": x.foto, "pais": x.pais, "bairro": x.bairro, "tipo": x.tipo, "active": x.is_active} for x in lista]
@app.post("/api/admin/switch")
def switch_account(x: SwitchIn, a: User = Depends(admin_only), d: Session = Depends(dbdep)):
    if not pwd.verify(x.admin_password, a.password_hash): raise HTTPException(401, "Senha do Admin errada!")
    t = d.query(User).filter_by(telefone=clean_tel(x.target_telefone)).first()
    if not t: raise HTTPException(404, "Pessoa não existe.")
    log_action(d, a.nome, "Admin", "troca_conta", f"Admin entrou na conta de {t.nome}")
    return {"token": make_token(t), "user": {"nome": t.nome, "tipo": t.tipo}}
@app.get("/api/products")
def get_products(u: User = Depends(current_user), d: Session = Depends(dbdep)):
    return [{"id": p.id, "nome": p.nome, "categoria": p.categoria, "preco": p.preco, "custo": p.custo, "stock": p.stock, "stock_minimo": p.stock_minimo, "foto": p.foto, "lucro": p.preco - p.custo, "alerta": p.stock <= p.stock_minimo, "updated_by": p.updated_by} for p in d.query(Product).order_by(Product.nome).all()]
@app.post("/api/products")
def add_product(x: ProdIn, u: User = Depends(funcionario_only), d: Session = Depends(dbdep)):
    p = Product(**x.model_dump()); p.updated_by = u.nome; d.add(p); d.commit(); log_action(d, u.nome, u.tipo, "produto_criar", x.nome); return {"ok": True}
@app.put("/api/products/{pid}")
def edit_product(pid: int, x: ProdIn, u: User = Depends(funcionario_only), d: Session = Depends(dbdep)):
    p = d.get(Product, pid)
    if not p: raise HTTPException(404, "Não existe")
    old = f"{p.nome} {p.preco} MT stock {p.stock}"
    for k, v in x.model_dump().items(): setattr(p, k, v)
    p.updated_by = u.nome; p.updated_at = agora(); d.commit(); log_action(d, u.nome, u.tipo, "produto_editar", f"De {old} para {p.nome} {p.preco} MT"); return {"ok": True}
@app.delete("/api/products/{pid}")
def del_product(pid: int, a: User = Depends(admin_only), d: Session = Depends(dbdep)):
    p = d.get(Product, pid)
    if not p: raise HTTPException(404, "Não existe")
    d.delete(p); d.commit(); log_action(d, a.nome, "Admin", "produto_apagar", p.nome); return {"ok": True}
@app.get("/api/fornecedores")
def get_forn(u: User = Depends(current_user), d: Session = Depends(dbdep)):
    return [{"id": f.id, "nome": f.nome, "telefone": f.telefone, "produto": f.produto, "preco": f.preco, "localizacao": f.localizacao} for f in d.query(Fornecedor).all()]
@app.post("/api/fornecedores")
def add_forn(x: FornIn, u: User = Depends(funcionario_only), d: Session = Depends(dbdep)):
    f = Fornecedor(**x.model_dump()); d.add(f); d.commit(); return {"ok": True}
@app.post("/api/compras")
def add_compra(x: CompraIn, u: User = Depends(funcionario_only), d: Session = Depends(dbdep)):
    f = d.get(Fornecedor, x.fornecedor_id)
    if not f: raise HTTPException(404, "Fornecedor não existe")
    total = x.qtd * x.preco_unitario
    c = Compra(fornecedor=f.nome, produto=x.produto, qtd=x.qtd, preco_unitario=x.preco_unitario, total=total, pagamento=x.pagamento, comprado_por=u.nome); d.add(c); d.commit(); log_action(d, u.nome, u.tipo, "compra", f"{x.produto} x{x.qtd} de {f.nome} = {total} MT")
    d.add(Historico(tipo="ENTRADA", produto=x.produto, quantidade=x.qtd, valor=total, motivo="COMPRA")); d.commit()
    return {"ok": True, "total": total, "compra_id": c.id}
@app.get("/api/financeiro")
def financeiro(u: User = Depends(current_user), d: Session = Depends(dbdep)):
    vendas = d.query(Venda).order_by(Venda.created_at.desc()).limit(300).all()
    despesas = d.query(Despesa).order_by(Despesa.created_at.desc()).limit(300).all()
    compras = d.query(Compra).order_by(Compra.created_at.desc()).limit(300).all()
    return {
      "vendas": [{"id": v.id, "cliente": v.cliente, "cliente_tel": v.cliente_tel, "vendedor": v.vendedor, "produto": v.produto, "qtd": v.qtd, "total": v.total, "lucro": v.lucro, "pagamento": v.pagamento, "estado": v.estado, "data_pagamento": v.data_pagamento, "whatsapp_link": v.whatsapp_link, "data": v.created_at.isoformat(), "hora": v.created_at.strftime("%H:%M:%S")} for v in vendas],
      "despesas": [{"id": x.id, "descricao": x.descricao, "valor": x.valor, "categoria": x.categoria, "criado_por": x.criado_por, "data": x.created_at.isoformat()} for x in despesas],
      "compras": [{"id": c.id, "fornecedor": c.fornecedor, "produto": c.produto, "qtd": c.qtd, "total": c.total, "pagamento": c.pagamento, "comprado_por": c.comprado_por, "data": c.created_at.isoformat()} for c in compras],
      "dividas": [{"id": v.id, "cliente": v.cliente, "cliente_tel": v.cliente_tel, "produto": v.produto, "total": v.total, "data": v.created_at.strftime("%d/%m/%Y %H:%M"), "data_pagamento": v.data_pagamento, "whatsapp_link": v.whatsapp_link} for v in vendas if v.estado == "Dívida"],
      "resumo": {"entradas": sum(v.total for v in vendas), "lucro_vendas": sum(v.lucro for v in vendas), "saidas": sum(x.valor for x in despesas), "compras_total": sum(c.total for c in compras), "lucro_liquido": sum(v.total for v in vendas) - sum(c.total for c in compras) - sum(x.valor for x in despesas), "dividas": sum(v.total for v in vendas if v.estado == "Dívida")}
    }
@app.post("/api/vendas")
def add_venda(x: VendaIn, u: User = Depends(current_user), d: Session = Depends(dbdep)):
    p = d.get(Product, x.produto_id)
    if not p: raise HTTPException(404, "Produto não existe")
    if p.stock < x.qtd: raise HTTPException(400, f"Stock insuficiente! Só tem {p.stock}")
    p.stock -= x.qtd; total = p.preco * x.qtd; custo = p.custo * x.qtd; lucro = total - custo
    estado = "Dívida" if "Fiado" in x.pagamento or "Dívida" in x.pagamento else "Pago"
    pag_info = ""
    if x.pagamento == "M-Pesa": pag_info = str(pagar_mpesa(x.cliente_tel, total))
    if x.pagamento == "e-Mola": pag_info = str(pagar_emola(x.cliente_tel, total))
    msg = f"Ola {x.cliente}! Obrigado pela compra na Padaria Jequessene. {p.nome} x{x.qtd} = {total} MT. Pagamento: {x.pagamento}. Volte sempre!"
    wa = whatsapp_link(x.cliente_tel, msg) if x.cliente_tel else ""
    v = Venda(cliente=x.cliente, cliente_tel=x.cliente_tel, vendedor=u.nome, produto=p.nome, qtd=x.qtd, total=total, custo_total=custo, lucro=lucro, pagamento=x.pagamento, estado=estado, data_pagamento=x.data_pagamento, whatsapp_link=wa); d.add(v); d.commit(); log_action(d, u.nome, u.tipo, "venda", f"{p.nome} x{x.qtd} = {total} MT | Lucro {lucro} MT | {x.pagamento} | {pag_info}")
    d.add(Historico(tipo="SAIDA", produto=p.nome, quantidade=x.qtd, valor=total, motivo="VENDA")); d.commit()
    return {"ok": True, "total": total, "lucro": lucro, "venda_id": v.id, "whatsapp_link": wa, "pagamento_info": pag_info}
@app.post("/api/vendas/{vid}/pagar")
def pagar(vid: int, u: User = Depends(current_user), d: Session = Depends(dbdep)):
    v = d.get(Venda, vid)
    if not v: raise HTTPException(404, "Não existe")
    v.estado = "Pago"; v.pagamento = "M-Pesa (Pago)"; d.commit(); log_action(d, u.nome, u.tipo, "divida_paga", f"{v.cliente} pagou {v.total} MT"); return {"ok": True}
@app.post("/api/despesas")
def add_despesa(x: DespesaIn, u: User = Depends(current_user), d: Session = Depends(dbdep)):
    ds = Despesa(**x.model_dump(), criado_por=u.nome); d.add(ds); d.commit(); log_action(d, u.nome, u.tipo, "despesa", f"{x.descricao} = {x.valor} MT"); return {"ok": True}
@app.get("/api/encomendas")
def get_enc(u: User = Depends(current_user), d: Session = Depends(dbdep)):
    return [{"id": e.id, "cliente": e.cliente, "telefone": e.telefone, "tipo": e.tipo, "data_entrega": e.data_entrega, "preco": e.preco, "sinal": e.sinal, "obs": e.obs, "estado": e.estado, "whatsapp_link": whatsapp_link(e.telefone, f"Ola {e.cliente}! A tua encomenda {e.tipo} esta {e.estado}. Entrega: {e.data_entrega}. Preco: {e.preco} MT - Padaria Jequessene")} for e in d.query(Encomenda).order_by(Encomenda.created_at.desc()).all()]
@app.post("/api/encomendas")
def add_enc(x: EncIn, u: User = Depends(current_user), d: Session = Depends(dbdep)):
    e = Encomenda(**x.model_dump()); d.add(e); d.commit(); log_action(d, u.nome, u.tipo, "encomenda", f"{x.tipo} para {x.cliente} = {x.preco} MT"); return {"ok": True}
@app.post("/api/encomendas/{eid}/entregar")
def entregar(eid: int, u: User = Depends(current_user), d: Session = Depends(dbdep)):
    e = d.get(Encomenda, eid)
    if not e: raise HTTPException(404, "Não existe")
    e.estado = "Entregue"; d.commit(); return {"ok": True}
@app.get("/api/dashboard")
def dashboard(u: User = Depends(current_user), d: Session = Depends(dbdep)):
    vendas = d.query(Venda).all(); produtos = d.query(Product).all(); users = d.query(User).all(); despesas = d.query(Despesa).all(); compras = d.query(Compra).all(); hoje = date.today()
    def soma(dia):
        vs = [x for x in vendas if x.created_at.date() == dia]; cs = [x for x in compras if x.created_at.date() == dia]; ds = [x for x in despesas if x.created_at.date() == dia]
        return {"data": dia.isoformat(), "label": dia.strftime("%d/%m"), "vendas": sum(x.total for x in vs), "lucro": sum(x.lucro for x in vs) - sum(x.total for x in cs) - sum(x.valor for x in ds), "n": len(vs)}
    ult7 = [soma(hoje - timedelta(days=i)) for i in range(6, -1, -1)]
    cont = {}
    for v in vendas: cont[v.produto] = cont.get(v.produto, 0) + v.qtd
    mais = sorted(cont.items(), key=lambda x: x[1], reverse=True)[:5]
    por_pag = {}
    for v in vendas: por_pag[v.pagamento] = por_pag.get(v.pagamento, 0) + v.total
    return {"hoje": ult7[-1], "semana": {"vendas": sum(x["vendas"] for x in ult7), "lucro": sum(x["lucro"] for x in ult7)}, "mes": {"vendas": sum(x.total for x in vendas if x.created_at.month == hoje.month), "lucro": sum(x.lucro for x in vendas if x.created_at.month == hoje.month) - sum(x.total for x in compras if x.created_at.month == hoje.month) - sum(x.valor for x in despesas if x.created_at.month == hoje.month)}, "grafico": ult7, "mais_vendidos": [{"produto": k, "qtd": v} for k, v in mais], "por_pagamento": por_pag, "totais": {"clientes": len([x for x in users if x.tipo == "Cliente"]), "funcionarios": len([x for x in users if normalizar_tipo(x.tipo) == "Funcionário"]), "produtos": len(produtos), "vendas_total": sum(v.total for v in vendas), "lucro_liquido": sum(v.total for v in vendas) - sum(c.total for c in compras) - sum(x.valor for x in despesas), "dividas": sum(v.total for v in vendas if v.estado == "Dívida")}}
@app.get("/api/recibo/venda/{vid}")
def recibo_venda_pdf(vid: int, d: Session = Depends(dbdep)):
    v = d.get(Venda, vid)
    if not v: raise HTTPException(404, "Não existe")
    pdf = FPDF(); pdf.add_page(); pdf.set_font("Helvetica", "B", 18); pdf.cell(0, 12, "PADARIA JEQUESSENE", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 11); pdf.cell(0, 8, f"Recibo Venda #{vid} - {v.created_at.strftime('%d/%m/%Y %H:%M')}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.cell(0, 8, f"Cliente: {v.cliente} | Vendedor: {v.vendedor}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, f"Produto: {v.produto} x{v.qtd} | Pagamento: {v.pagamento} | Estado: {v.estado}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "B", 14); pdf.cell(0, 12, f"TOTAL: {v.total} MT | Lucro: {v.lucro} MT", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 11); pdf.cell(0, 8, "Obrigado! Volte sempre - Matola, Mocambique", new_x="LMARGIN", new_y="NEXT", align="C")
    out = bytes(pdf.output()); return Response(content=out, media_type="application/pdf", headers={"Content-Disposition": f"attachment; filename=recibo_venda_{vid}.pdf"})
@app.get("/api/recibo/compra/{cid}")
def recibo_compra_pdf(cid: int, d: Session = Depends(dbdep)):
    c = d.get(Compra, cid)
    if not c: raise HTTPException(404, "Não existe")
    pdf = FPDF(); pdf.add_page(); pdf.set_font("Helvetica", "B", 18); pdf.cell(0, 12, "PADARIA JEQUESSENE", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 11); pdf.cell(0, 8, f"Recibo Compra #{cid} - {c.created_at.strftime('%d/%m/%Y %H:%M')}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.cell(0, 8, f"Fornecedor: {c.fornecedor} | Comprado por: {c.comprado_por}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, f"Produto: {c.produto} x{c.qtd} | Pagamento: {c.pagamento}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "B", 14); pdf.cell(0, 12, f"TOTAL: {c.total} MT", new_x="LMARGIN", new_y="NEXT")
    out = bytes(pdf.output()); return Response(content=out, media_type="application/pdf", headers={"Content-Disposition": f"attachment; filename=recibo_compra_{cid}.pdf"})
@app.get("/api/recibo/diario", response_class=HTMLResponse)
def recibo_diario(d: Session = Depends(dbdep), u: User = Depends(current_user)):
    hoje = date.today(); vs = [x for x in d.query(Venda).all() if x.created_at.date() == hoje]; cs = [x for x in d.query(Compra).all() if x.created_at.date() == hoje]; ds = [x for x in d.query(Despesa).all() if x.created_at.date() == hoje]
    tv = sum(x.total for x in vs); tc = sum(x.total for x in cs); td = sum(x.valor for x in ds)
    return f"<html><body style='font-family:Arial;padding:20px;max-width:480px;margin:auto'><h1>PADARIA JEQUESSENE</h1><p>Fecho {hoje.strftime('%d/%m/%Y')}</p><p>Vendas {len(vs)} - {tv} MT</p><p>Compras {tc} MT</p><p>Despesas {td} MT</p><p>M-Pesa {sum(x.total for x in vs if 'M-Pesa' in x.pagamento)} MT | e-Mola {sum(x.total for x in vs if 'e-Mola' in x.pagamento)} MT | Dinheiro {sum(x.total for x in vs if 'Dinheiro' in x.pagamento)} MT</p><h2>LUCRO LIQUIDO: {tv-tc-td} MT</h2><a href='javascript:window.print()'>Imprimir / Guardar PDF</a></body></html>"
@app.get("/api/export/csv")
def export_csv(d: Session = Depends(dbdep), u: User = Depends(current_user)):
    linhas = ["tipo,data,cliente_vendedor,produto,quantidade,total,lucro,pagamento"]
    for v in d.query(Venda).all(): linhas.append(f"VENDA,{v.created_at},{v.cliente}/{v.vendedor},{v.produto},{v.qtd},{v.total},{v.lucro},{v.pagamento}")
    for c in d.query(Compra).all(): linhas.append(f"COMPRA,{c.created_at},{c.fornecedor}/{c.comprado_por},{c.produto},{c.qtd},{c.total},0,{c.pagamento}")
    return Response(content="\n".join(linhas), media_type="text/csv", headers={"Content-Disposition": "attachment; filename=padaria_jequessene.csv"})
@app.get("/api/export/excel")
def export_excel(d: Session = Depends(dbdep), u: User = Depends(current_user)):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Vendas"; ws.append(["Data","Cliente","Vendedor","Produto","Qtd","Total MT","Lucro MT","Pagamento","Estado"])
    for v in d.query(Venda).all(): ws.append([v.created_at.strftime("%d/%m/%Y %H:%M"), v.cliente, v.vendedor, v.produto, v.qtd, v.total, v.lucro, v.pagamento, v.estado])
    ws2 = wb.create_sheet("Produtos"); ws2.append(["Produto","Categoria","Preco MT","Custo MT","Stock","Lucro MT"])
    for p in d.query(Product).all(): ws2.append([p.nome, p.categoria, p.preco, p.custo, p.stock, p.preco - p.custo])
    ws3 = wb.create_sheet("Compras"); ws3.append(["Data","Fornecedor","Produto","Qtd","Total MT","Pagamento"])
    for c in d.query(Compra).all(): ws3.append([c.created_at.strftime("%d/%m/%Y %H:%M"), c.fornecedor, c.produto, c.qtd, c.total, c.pagamento])
    ws4 = wb.create_sheet("Despesas"); ws4.append(["Data","Descricao","Valor MT","Categoria","Criado por"])
    for x in d.query(Despesa).all(): ws4.append([x.created_at.strftime("%d/%m/%Y %H:%M"), x.descricao, x.valor, x.categoria, x.criado_por])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=padaria_jequessene.xlsx"})
@app.get("/api/admin/painel")
def painel_admin(a: User = Depends(admin_only), d: Session = Depends(dbdep)):
    vendas = d.query(Venda).order_by(Venda.created_at.desc()).limit(100).all(); compras = d.query(Compra).order_by(Compra.created_at.desc()).limit(100).all(); logs = d.query(Log).order_by(Log.created_at.desc()).limit(300).all()
    total_v = sum(v.total for v in d.query(Venda).all()); total_c = sum(c.total for c in d.query(Compra).all()); total_d = sum(x.valor for x in d.query(Despesa).all())
    return {"resumo": {"total_vendas": total_v, "lucro_bruto": sum(v.lucro for v in d.query(Venda).all()), "total_compras": total_c, "total_despesas": total_d, "lucro_liquido_final": total_v - total_c - total_d, "dividas": sum(v.total for v in d.query(Venda).all() if v.estado == "Dívida"), "n_vendas": d.query(Venda).count(), "n_compras": d.query(Compra).count()},
      "ultimas_vendas": [{"id": v.id, "data": v.created_at.strftime("%d/%m/%Y"), "hora": v.created_at.strftime("%H:%M:%S"), "cliente": v.cliente, "vendedor": v.vendedor, "produto": v.produto, "qtd": v.qtd, "total": v.total, "lucro": v.lucro, "pagamento": v.pagamento, "estado": v.estado, "whatsapp_link": v.whatsapp_link} for v in vendas],
      "ultimas_compras": [{"id": c.id, "data": c.created_at.strftime("%d/%m/%Y"), "hora": c.created_at.strftime("%H:%M:%S"), "fornecedor": c.fornecedor, "produto": c.produto, "qtd": c.qtd, "total": c.total, "comprado_por": c.comprado_por} for c in compras],
      "fluxo": [{"data": l.created_at.strftime("%d/%m/%Y %H:%M:%S"), "quem": l.quem, "tipo": l.tipo, "acao": l.acao, "detalhe": l.detalhe} for l in logs]}
@app.get("/api/backup")
def backup(a: User = Depends(admin_only), d: Session = Depends(dbdep)):
    return {"users": [{"nome": u.nome, "telefone": u.telefone, "tipo": u.tipo} for u in d.query(User).all()], "products": [{"nome": p.nome, "preco": p.preco, "stock": p.stock} for p in d.query(Product).all()], "vendas": [{"cliente": v.cliente, "produto": v.produto, "total": v.total, "data": str(v.created_at)} for v in d.query(Venda).all()]}
@app.get("/api/historico")
def get_historico(d: Session = Depends(dbdep), u: User = Depends(current_user)):
    rows = d.query(Historico).order_by(Historico.id.desc()).all()
    return [{"id": h.id, "tipo": h.tipo, "produto": h.produto, "quantidade": h.quantidade, "valor": h.valor, "motivo": h.motivo, "data": str(h.data)} for h in rows]