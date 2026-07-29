import os
import math
import re
from datetime import datetime

import pandas as pd
import unicodedata
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

# =========================================================
# CONFIGURAÇÃO GERAL
# =========================================================
APP_NAME = "Safra360"

ROTAS_POR_NIVEL = {
    "admin": "/dashboard",
    "secador": "/secador",
    "pilagem": "/pilagem"
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

USUARIOS_FILE = os.path.join(DATA_DIR, "usuarios.xlsx")
DADOS_FILE = os.path.join(DATA_DIR, "dados.xlsx")

ESTOQUE_SHEET = "estoque"
PILAGEM_SHEET = "pilagem"

# =========================================================
# ROLOS DE SECAGEM
# =========================================================
ROLOS = [
    {"id": 1, "capacidade": 200, "ocupado": 0, "status": "LIVRE"},
    {"id": 2, "capacidade": 200, "ocupado": 0, "status": "LIVRE"},
    {"id": 3, "capacidade": 200, "ocupado": 0, "status": "LIVRE"},
    {"id": 4, "capacidade": 200, "ocupado": 0, "status": "LIVRE"},
    {"id": 5, "capacidade": 200, "ocupado": 0, "status": "LIVRE"},
    {"id": 6, "capacidade": 200, "ocupado": 0, "status": "LIVRE"},
    {"id": 7, "capacidade": 280, "ocupado": 0, "status": "LIVRE"},
    {"id": 8, "capacidade": 280, "ocupado": 0, "status": "LIVRE"},
    {"id": 9, "capacidade": 280, "ocupado": 0, "status": "LIVRE"},
]

for rolo in ROLOS:
    rolo.setdefault("historico_trocas", [])

FILA_ESPERA = []

# =========================================================
# FASTAPI
# =========================================================
app = FastAPI(title=APP_NAME)
templates = Jinja2Templates(directory=TEMPLATES_DIR)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# =========================================================
# UTILITÁRIOS
# =========================================================
def safe_int(value, default=0):
    try:
        v = float(value)
        if math.isnan(v) or math.isinf(v):
            return default
        return int(v)
    except:
        return default

def safe_float(value, default=0.0):
    try:
        v = float(value)
        if math.isnan(v) or math.isinf(v):
            return default
        return v
    except:
        return default

def normalizar(texto):
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    return texto.strip().lower()

# =========================================================
# USUÁRIOS (login/senha em Excel)
# =========================================================
def carregar_usuarios():
    if not os.path.exists(USUARIOS_FILE):
        pd.DataFrame(
            columns=["usuario", "senha", "nivel"]
        ).to_excel(USUARIOS_FILE, index=False)
    return pd.read_excel(USUARIOS_FILE)

def autenticar_usuario(usuario, senha):
    df = carregar_usuarios()
    user = df.loc[
        (df["usuario"].str.lower() == usuario.lower()) &
        (df["senha"].astype(str) == str(senha))
    ]
    if user.empty:
        return None
    return {
        "usuario": user.iloc[0]["usuario"],
        "nivel": user.iloc[0]["nivel"]
    }

def checar_acesso(request: Request, nivel=None):
    nivel_user = request.cookies.get("nivel")
    if not nivel_user:
        return False
    if nivel and nivel_user != nivel:
        return False
    return True

# =========================================================
# GARANTIR PLANILHA DE DADOS
# =========================================================
def garantir_dados():
    cols_estoque = [
        "cliente", "fazenda", "motorista", "rolo_origem", "rolo_final",
        "historico_troca_rolo", "peso_bruto", "sacas_inicial", "grau",
        "inicio_secagem", "fim_secagem", "secado_por", "finalizado_secagem_por",
        "inicio_pilagem", "fim_pilagem", "pilagem_iniciada_por",
        "pilagem_finalizada_por", "total_peso_pilagem", "total_sacas_pilagem",
        "data_registro"
    ]
    cols_pilagem = [
        "cliente", "fazenda", "motorista", "rolo_origem", "peso_bruto",
        "sacas", "grau", "secado_por", "finalizado_por",
        "inicio_secagem", "fim_secagem", "inicio_pilagem", "fim_pilagem",
        "pilagem_iniciada_por", "pilagem_finalizada_por", "status_pilagem",
        "pilagem_em_execucao_por", "sacas_piladas", "peso_pilado", "registros_bags"
    ]

    if not os.path.exists(DADOS_FILE):
        with pd.ExcelWriter(DADOS_FILE, engine="openpyxl") as writer:
            pd.DataFrame(columns=cols_estoque).to_excel(writer, sheet_name=ESTOQUE_SHEET, index=False)
            pd.DataFrame(columns=cols_pilagem).to_excel(writer, sheet_name=PILAGEM_SHEET, index=False)
        return

    # Garante abas
    with pd.ExcelFile(DADOS_FILE, engine="openpyxl") as xls:
        with pd.ExcelWriter(DADOS_FILE, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
            if ESTOQUE_SHEET not in xls.sheet_names:
                pd.DataFrame(columns=cols_estoque).to_excel(writer, sheet_name=ESTOQUE_SHEET, index=False)
            if PILAGEM_SHEET not in xls.sheet_names:
                pd.DataFrame(columns=cols_pilagem).to_excel(writer, sheet_name=PILAGEM_SHEET, index=False)

garantir_dados()

# =========================================================
# PILAGEM (EXCEL)
# =========================================================
def carregar_pilagem():
    garantir_dados()
    df = pd.read_excel(DADOS_FILE, sheet_name=PILAGEM_SHEET, engine="openpyxl")
    return df.to_dict(orient="records")

def salvar_pilagem(pilagem):
    df = pd.DataFrame(pilagem)
    with pd.ExcelWriter(DADOS_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=PILAGEM_SHEET, index=False)

# =========================================================
# ESTOQUE (EXCEL)
# =========================================================
def carregar_estoque():
    garantir_dados()
    return pd.read_excel(DADOS_FILE, sheet_name=ESTOQUE_SHEET, engine="openpyxl")

def salvar_estoque(df):
    if df is None or df.empty:
        return
    with pd.ExcelWriter(DADOS_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=ESTOQUE_SHEET, index=False)

def obter_saldo_cliente(cliente):
    df_estoque = carregar_estoque()
    if df_estoque.empty or "cliente" not in df_estoque.columns:
        return 0.0, 0.0
    df_estoque["cliente"] = df_estoque["cliente"].astype(str).str.lower().str.strip()
    df_cliente = df_estoque[df_estoque["cliente"] == cliente.lower().strip()]
    saldo_sacas = df_cliente["total_sacas_pilagem"].fillna(0).sum()
    saldo_kg = df_cliente["total_peso_pilagem"].fillna(0).sum()
    return float(saldo_sacas), float(saldo_kg)

# =========================================================
# FUNÇÕES DOS ROLOS
# =========================================================
def encontrar_rolo_disponivel(sacas):
    for rolo in ROLOS:
        if rolo.get("status") == "LIVRE" and rolo.get("capacidade", 0) >= sacas:
            return rolo
    return None

def rolos_disponiveis_para_troca(peso_sacas, rolo_atual_id=None):
    disponiveis = []
    for r in ROLOS:
        if r.get("status") != "LIVRE":
            continue
        if r.get("capacidade", 0) < peso_sacas:
            continue
        if rolo_atual_id and r.get("id") == rolo_atual_id:
            continue
        disponiveis.append(r)
    return disponiveis

def puxar_proxima_da_fila(rolo):
    if not FILA_ESPERA:
        return
    indice_escolhido = None
    for i, carga in enumerate(FILA_ESPERA):
        if rolo.get("capacidade", 0) >= carga.get("sacas", 0):
            indice_escolhido = i
            break
    if indice_escolhido is None:
        return
    proximo = FILA_ESPERA.pop(indice_escolhido)
    rolo["status"] = "AGUARDANDO"
    rolo["cliente"] = proximo["cliente"]
    rolo["fazenda"] = proximo["fazenda"]
    rolo["motorista"] = proximo["motorista"]
    rolo["peso_bruto"] = proximo["peso"]
    rolo["sacas"] = proximo["sacas"]

# =========================================================
# LOGIN
# =========================================================
@app.get("/", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})

@app.post("/login")
def login(request: Request, usuario: str = Form(...), senha: str = Form(...)):
    user = autenticar_usuario(usuario, senha)
    if not user:
        return templates.TemplateResponse("login.html", {"request": request, "erro": "Usuário ou senha inválidos"})
    resp = RedirectResponse(ROTAS_POR_NIVEL[user["nivel"]], status_code=303)
    resp.set_cookie("usuario", user["usuario"])
    resp.set_cookie("nivel", user["nivel"])
    return resp

@app.get("/logout")
def logout():
    resp = RedirectResponse("/", status_code=302)
    resp.delete_cookie("usuario")
    resp.delete_cookie("nivel")
    return resp

@app.get("/finalizar-turno")
def finalizar_turno():
    resp = RedirectResponse("/", status_code=303)
    resp.delete_cookie("usuario")
    resp.delete_cookie("nivel")
    return resp

# =========================================================
# DASHBOARD (ADMIN)
# =========================================================
@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    if not checar_acesso(request, "admin"):
        return RedirectResponse("/", status_code=302)

    rolos_ocupados = sum(1 for r in ROLOS if r["status"] != "LIVRE")
    rolos_livres = sum(1 for r in ROLOS if r["status"] == "LIVRE")
    fila_count = len(FILA_ESPERA)
    usuario = request.cookies.get("usuario")

    return templates.TemplateResponse("dashboard.html", {
        "request": request,
        "usuario": usuario,
        "rolos_ocupados": rolos_ocupados,
        "rolos_livres": rolos_livres,
        "fila_count": fila_count,
    })

# =========================================================
# CALCULAR DESTINO (CHAMADO EXTERNO)
# =========================================================
@app.post("/calcular-destino")
def calcular_destino(
    request: Request,
    cliente: str = Form(...),
    fazenda: str = Form(...),
    motorista: str = Form(...),
    peso: int = Form(...)
):
    if not checar_acesso(request):
        return RedirectResponse("/", status_code=302)

    sacas = round(peso / 100)
    rolo = encontrar_rolo_disponivel(sacas)

    if rolo:
        return templates.TemplateResponse("confirmar_descarga.html", {
            "request": request,
            "rolo": rolo,
            "cliente": cliente,
            "fazenda": fazenda,
            "motorista": motorista,
            "peso": peso,
            "sacas": sacas
        })

    FILA_ESPERA.append({
        "cliente": cliente,
        "fazenda": fazenda,
        "motorista": motorista,
        "peso": peso,
        "sacas": sacas,
        "data_entrada": datetime.now().strftime("%d/%m/%Y %H:%M")
    })

    return RedirectResponse(f"/dashboard?fila=1&sacas={sacas}", status_code=303)

# =========================================================
# SECADOR
# =========================================================
@app.get("/secador", response_class=HTMLResponse)
def secador(request: Request):
    nivel = request.cookies.get("nivel")
    if nivel not in ["admin", "secador"]:
        return RedirectResponse("/", status_code=302)
    return templates.TemplateResponse("secador.html", {
        "request": request,
        "rolos": ROLOS,
        "fila": FILA_ESPERA,
        "user": {"nivel": nivel}
    })

@app.post("/secador/confirmar/{rolo_id}")
def confirmar_descarga(
    request: Request,
    rolo_id: int,
    cliente: str = Form(...),
    fazenda: str = Form(...),
    motorista: str = Form(...),
    peso: int = Form(...),
    sacas: int = Form(...)
):
    nivel = request.cookies.get("nivel")
    if nivel not in ["admin", "secador"]:
        return RedirectResponse("/", status_code=302)

    rolo = next((r for r in ROLOS if r["id"] == rolo_id), None)
    if not rolo:
        return HTMLResponse("Rolo nao encontrado", status_code=404)

    rolo["status"] = "AGUARDANDO"
    rolo["cliente"] = cliente
    rolo["fazenda"] = fazenda
    rolo["motorista"] = motorista
    rolo["peso_bruto"] = peso
    rolo["sacas"] = sacas

    return RedirectResponse("/secador", status_code=303)

@app.post("/secador/iniciar")
def iniciar_secagem(request: Request, rolo_id: int = Form(...), senha: str = Form(...)):
    usuario_logado = request.cookies.get("usuario")
    nivel = request.cookies.get("nivel")

    if nivel not in ["admin", "secador"]:
        return HTMLResponse("Acesso negado", status_code=403)

    df = carregar_usuarios()
    user = df.loc[(df["usuario"] == usuario_logado) & (df["senha"].astype(str) == str(senha))]
    if user.empty:
        return HTMLResponse("Senha invalida", status_code=403)

    rolo = next((r for r in ROLOS if r["id"] == rolo_id), None)
    if not rolo:
        return HTMLResponse("Rolo nao encontrado", status_code=404)
    if rolo["status"] != "AGUARDANDO":
        return HTMLResponse("Rolo nao esta aguardando", status_code=400)

    rolo["status"] = "SECANDO"
    rolo["inicio_secagem"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    rolo["iniciado_por"] = usuario_logado

    return RedirectResponse("/secador", status_code=303)

@app.get("/secador/finalizar/{rolo_id}", response_class=HTMLResponse)
def finalizar_secagem(request: Request, rolo_id: int):
    rolo = next((r for r in ROLOS if r["id"] == rolo_id), None)
    if not rolo:
        return HTMLResponse("Rolo nao encontrado", status_code=404)

    rolo.setdefault("cliente", "Nao informado")
    rolo.setdefault("peso_bruto", 0)
    rolo.setdefault("sacas", rolo.get("peso_bruto", 0) // 100)
    rolo.setdefault("fazenda", "Nao informado")
    rolo.setdefault("motorista", "Nao informado")
    rolo.setdefault("iniciado_por", "Desconhecido")
    rolo.setdefault("inicio_secagem", "Nao informado")

    finalizado_por = request.cookies.get("usuario") or "Desconhecido"
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")

    return templates.TemplateResponse("finalizar_secagem.html", {
        "request": request,
        "r": rolo,
        "finalizado_por": finalizado_por,
        "data_hora": data_hora
    })

@app.post("/secador/registrar-grau/{rolo_id}")
def registrar_grau(request: Request, rolo_id: int, grau: str = Form(...)):
    usuario = request.cookies.get("usuario")
    rolo = next((r for r in ROLOS if r["id"] == rolo_id), None)
    if not rolo:
        return HTMLResponse("Rolo nao encontrado", status_code=404)

    pilagem = carregar_pilagem()
    dados_pilagem = {
        "cliente": rolo.get("cliente"),
        "fazenda": rolo.get("fazenda"),
        "motorista": rolo.get("motorista"),
        "sacas": rolo.get("sacas"),
        "peso_bruto": rolo.get("peso_bruto"),
        "grau": grau,
        "secado_por": rolo.get("iniciado_por"),
        "finalizado_por": usuario,
        "inicio_secagem": rolo.get("inicio_secagem"),
        "fim_secagem": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "rolo_origem": rolo_id,
        "status_pilagem": "AGUARDANDO"
    }
    pilagem.append(dados_pilagem)
    salvar_pilagem(pilagem)

    rolo["ocupado"] = 0
    rolo["status"] = "LIVRE"
    for chave in ["cliente", "fazenda", "motorista", "peso_bruto", "sacas", "inicio_secagem", "iniciado_por"]:
        rolo.pop(chave, None)

    puxar_proxima_da_fila(rolo)
    return RedirectResponse("/secador", status_code=303)

# =========================================================
# TROCAR ROLO
# =========================================================
@app.post("/secador/trocar-rolo")
def trocar_rolo(
    request: Request,
    rolo_atual_id: int = Form(...),
    novo_rolo_id: int = Form(...),
    motivo: str = Form(...),
    senha: str = Form(...)
):
    usuario = request.cookies.get("usuario")
    nivel = request.cookies.get("nivel")
    if nivel not in ["admin", "secador"]:
        return HTMLResponse("Acesso negado", status_code=403)

    df_users = carregar_usuarios()
    if df_users.loc[(df_users["usuario"] == usuario) & (df_users["senha"].astype(str) == str(senha))].empty:
        return HTMLResponse("Senha invalida", status_code=403)

    rolo_atual = next((r for r in ROLOS if r.get("id") == rolo_atual_id), None)
    novo_rolo = next((r for r in ROLOS if r.get("id") == novo_rolo_id), None)

    if not rolo_atual or not novo_rolo:
        return HTMLResponse("Rolo invalido", status_code=404)
    if rolo_atual.get("status") == "LIVRE":
        return HTMLResponse("Rolo nao esta em uso", status_code=400)
    if novo_rolo.get("status") != "LIVRE":
        return HTMLResponse("Novo rolo indisponivel", status_code=400)

    sacas = rolo_atual.get("sacas", 0)
    if novo_rolo.get("capacidade", 0) < sacas:
        return HTMLResponse("Capacidade insuficiente", status_code=400)

    for chave in ["cliente", "fazenda", "motorista", "peso_bruto", "sacas", "inicio_secagem", "iniciado_por"]:
        if chave in rolo_atual:
            novo_rolo[chave] = rolo_atual[chave]

    novo_rolo["status"] = rolo_atual["status"]
    novo_rolo.setdefault("historico_trocas", []).append({
        "de": rolo_atual_id,
        "para": novo_rolo_id,
        "motivo": motivo,
        "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "usuario": usuario
    })

    rolo_atual["status"] = "MANUTENCAO"
    rolo_atual["ocupado"] = 0
    for k in list(rolo_atual.keys()):
        if k not in ["id", "capacidade", "status", "ocupado", "historico_trocas"]:
            rolo_atual.pop(k, None)

    return RedirectResponse("/secador", status_code=303)

@app.get("/secador/trocar-rolo")
def trocar_rolo_get():
    return RedirectResponse("/secador", status_code=303)

@app.post("/secador/liberar-rolo")
def liberar_rolo(request: Request, rolo_id: int = Form(...), senha: str = Form(...)):
    usuario = request.cookies.get("usuario")
    nivel = request.cookies.get("nivel")
    if nivel not in ["admin", "secador"]:
        return HTMLResponse("Acesso negado", status_code=403)

    df_users = carregar_usuarios()
    if df_users.loc[(df_users["usuario"] == usuario) & (df_users["senha"].astype(str) == str(senha))].empty:
        return HTMLResponse("Senha invalida", status_code=403)

    rolo = next((r for r in ROLOS if r.get("id") == rolo_id), None)
    if not rolo:
        return HTMLResponse("Rolo nao encontrado", status_code=404)
    if str(rolo.get("status", "")).strip().upper() != "MANUTENCAO":
        return HTMLResponse("Rolo nao esta em manutencao", status_code=400)

    rolo["status"] = "LIVRE"
    rolo["ocupado"] = 0
    puxar_proxima_da_fila(rolo)

    return HTMLResponse("""<!DOCTYPE html><html><head><meta charset='UTF-8'>
        <script>window.location.href='/secador';</script></head>
        <body>Rolo liberado. Redirecionando...</body></html>""")

# =========================================================
# PILAGEM
# =========================================================
@app.get("/pilagem", response_class=HTMLResponse)
def pilagem(request: Request):
    pilagem_data = carregar_pilagem()
    cargas = []
    for p in pilagem_data:
        cargas.append({
            "cliente": p.get("cliente"),
            "fazenda": p.get("fazenda"),
            "motorista": p.get("motorista", "—"),
            "peso": safe_float(p.get("peso_bruto")),
            "sacas": safe_float(p.get("sacas")),
            "rolo_nome": f"Rolo {p.get('rolo_origem')}",
            "rolo_id": p.get("rolo_origem"),
            "status": p.get("status_pilagem", "AGUARDANDO"),
            "status_pilagem": p.get("status_pilagem", "AGUARDANDO"),
            "iniciado_por": p.get("pilagem_iniciada_por", "—"),
            "finalizado_por": p.get("pilagem_finalizada_por"),
            "grau": p.get("grau"),
            "registros": safe_int(p.get("registros_bags")),
        })

    total_peso = sum(c["peso"] for c in cargas)

    return templates.TemplateResponse("pilagem.html", {
        "request": request,
        "cargas": cargas,
        "total_peso": total_peso
    })

@app.post("/pilagem/iniciar", response_class=HTMLResponse)
def iniciar_pilagem(request: Request, rolo_id: int = Form(...), senha: str = Form(...)):
    usuario = request.cookies.get("usuario")
    nivel = request.cookies.get("nivel")
    if nivel not in ["admin", "pilagem"]:
        return HTMLResponse("Acesso negado", status_code=403)

    df_users = carregar_usuarios()
    if df_users.loc[(df_users["usuario"] == usuario) & (df_users["senha"].astype(str) == str(senha))].empty:
        return HTMLResponse("Senha invalida", status_code=403)

    pilagem_data = carregar_pilagem()
    carga = next((p for p in pilagem_data if p.get("rolo_origem") == rolo_id), None)
    if not carga:
        return HTMLResponse("Carga nao encontrada", status_code=404)

    carga["sacas_piladas"] = safe_float(carga.get("sacas_piladas"))
    carga["peso_pilado"] = safe_float(carga.get("peso_pilado"))
    carga["registros_bags"] = safe_int(carga.get("registros_bags"))

    if carga.get("status_pilagem") != "EM PILAGEM":
        carga["status_pilagem"] = "EM PILAGEM"
        carga["inicio_pilagem"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        carga["pilagem_iniciada_por"] = usuario

    carga["pilagem_em_execucao_por"] = usuario
    salvar_pilagem(pilagem_data)

    return templates.TemplateResponse("iniciar_pilagem.html", {
        "request": request,
        "cliente": carga.get("cliente"),
        "fazenda": carga.get("fazenda"),
        "motorista": carga.get("motorista", "—"),
        "peso": safe_float(carga.get("peso_bruto")),
        "sacas": safe_float(carga.get("sacas")),
        "rolo_nome": f"Rolo {rolo_id}",
        "rolo_id": rolo_id,
        "grau": carga.get("grau"),
        "inicio_secagem": carga.get("inicio_secagem"),
        "fim_secagem": carga.get("fim_secagem"),
        "inicio_pilagem": carga.get("inicio_pilagem"),
        "pilagem_iniciada_por": carga.get("pilagem_iniciada_por"),
        "pilagem_em_execucao_por": carga.get("pilagem_em_execucao_por"),
        "sacas_piladas": carga["sacas_piladas"],
        "peso_pilado": carga["peso_pilado"],
        "registros": carga["registros_bags"],
    })

@app.post("/pilagem/atualizar_progresso")
async def atualizar_progresso(request: Request):
    data = await request.json()
    rolo_id = data.get("rolo_id")
    pilagem_data = carregar_pilagem()
    carga = next((p for p in pilagem_data if p.get("rolo_origem") == rolo_id), None)
    if not carga:
        return JSONResponse({"error": "Carga nao encontrada"}, status_code=404)

    carga["registros_bags"] = safe_int(carga.get("registros_bags")) + 1
    carga["sacas_piladas"] = safe_float(data.get("sacas_piladas"))
    carga["peso_pilado"] = safe_float(data.get("peso_pilado"))
    salvar_pilagem(pilagem_data)

    return JSONResponse({
        "success": True,
        "sacas_piladas": carga["sacas_piladas"],
        "peso_pilado": carga["peso_pilado"],
        "registros_bags": carga["registros_bags"]
    })

@app.post("/pilagem/finalizar")
def finalizar_pilagem(
    request: Request,
    rolo_id: int = Form(...),
    total_peso: float = Form(...),
    total_sacas: float = Form(...),
    senha: str = Form(...)
):
    usuario = request.cookies.get("usuario")
    nivel = request.cookies.get("nivel")
    if nivel not in ["admin", "pilagem"]:
        return HTMLResponse("Acesso negado", status_code=403)

    df_users = carregar_usuarios()
    if df_users.loc[(df_users["usuario"] == usuario) & (df_users["senha"].astype(str) == str(senha))].empty:
        return HTMLResponse("Senha invalida", status_code=403)

    pilagem_data = carregar_pilagem()
    carga = next((p for p in pilagem_data if p.get("rolo_origem") == rolo_id), None)
    if not carga:
        return HTMLResponse("Carga nao encontrada", status_code=404)

    carga["status_pilagem"] = "FINALIZADO"
    carga["fim_pilagem"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    carga["pilagem_finalizada_por"] = usuario

    df_estoque = carregar_estoque()
    rolo_final_id = carga.get("rolo_origem")
    rolo_obj = next((r for r in ROLOS if r.get("id") == rolo_final_id), None)

    historico_str = ""
    if rolo_obj:
        historico_lista = rolo_obj.get("historico_trocas", [])
        historico_str = " | ".join([
            f"{h.get('data')} - {h.get('usuario')} trocou {h.get('de')}→{h.get('para')} ({h.get('motivo')})"
            for h in historico_lista
        ])

    nova_linha = {
        "cliente": carga.get("cliente"),
        "fazenda": carga.get("fazenda"),
        "motorista": carga.get("motorista"),
        "rolo_origem": carga.get("rolo_origem"),
        "rolo_final": rolo_final_id,
        "historico_troca_rolo": historico_str,
        "peso_bruto": carga.get("peso_bruto"),
        "sacas_inicial": carga.get("sacas"),
        "grau": carga.get("grau"),
        "inicio_secagem": carga.get("inicio_secagem"),
        "fim_secagem": carga.get("fim_secagem"),
        "secado_por": carga.get("secado_por"),
        "finalizado_secagem_por": carga.get("finalizado_por"),
        "inicio_pilagem": carga.get("inicio_pilagem"),
        "fim_pilagem": carga.get("fim_pilagem"),
        "pilagem_iniciada_por": carga.get("pilagem_iniciada_por"),
        "pilagem_finalizada_por": usuario,
        "total_peso_pilagem": total_peso,
        "total_sacas_pilagem": total_sacas,
        "data_registro": datetime.now().strftime("%d/%m/%Y %H:%M")
    }

    df_estoque = pd.concat([df_estoque, pd.DataFrame([nova_linha])], ignore_index=True)
    salvar_estoque(df_estoque)

    pilagem_data = [p for p in pilagem_data if p.get("rolo_origem") != rolo_id]
    salvar_pilagem(pilagem_data)

    return RedirectResponse("/pilagem?sucesso=1", status_code=303)

@app.get("/pilagem/continuar/{rolo_id}", response_class=HTMLResponse)
def continuar_pilagem(request: Request, rolo_id: int):
    pilagem_data = carregar_pilagem()
    carga = next((p for p in pilagem_data if p.get("rolo_origem") == rolo_id), None)
    if not carga:
        return RedirectResponse("/pilagem", status_code=303)

    usuario = request.cookies.get("usuario")
    carga["pilagem_em_execucao_por"] = usuario
    salvar_pilagem(pilagem_data)

    return templates.TemplateResponse("iniciar_pilagem.html", {
        "request": request,
        "cliente": carga.get("cliente"),
        "fazenda": carga.get("fazenda"),
        "motorista": carga.get("motorista"),
        "peso": carga.get("peso_bruto"),
        "sacas": carga.get("sacas"),
        "rolo_nome": f"Rolo {rolo_id}",
        "rolo_id": rolo_id,
        "grau": carga.get("grau"),
        "pilagem_iniciada_por": carga.get("pilagem_iniciada_por", "—"),
        "pilagem_em_execucao_por": carga.get("pilagem_em_execucao_por"),
        "inicio_pilagem": carga.get("inicio_pilagem"),
        "sacas_piladas": carga.get("sacas_piladas", 0),
        "peso_pilado": carga.get("peso_pilado", 0),
        "registros": int(carga.get("registros_bags", 0)),
    })
