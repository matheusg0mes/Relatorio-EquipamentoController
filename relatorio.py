import os
from io import BytesIO
import requests
import pandas as pd

# Configurações (com suporte a variáveis de ambiente por segurança)
API = os.getenv("API_URL", "https://projeto-equipamentocontrol.onrender.com")
USUARIO = os.getenv("API_USER", "admin@empresa.com")
SENHA = os.getenv("API_PASSWORD", "123456")

COLUNAS = [
    "equipamento.patrimonio", "equipamento.nome", "colaborador.nome",
    "colaborador.email", "dataPrevistaDevolucao", "diasDeAtraso",
]
def gerar_relatorio() -> BytesIO:
    login = requests.post(f"{API}/auth/login",
                          json={"email": USUARIO, "senha": SENHA}, timeout=90)
    login.raise_for_status()
    token = login.json().get("token")
    if not token:
        raise ValueError("Token não encontrado na resposta de login")

    resp = requests.get(f"{API}/emprestimos",
                        headers={"Authorization": f"Bearer {token}"}, timeout=90)
    resp.raise_for_status()

    tabela = pd.json_normalize(resp.json())
    for col in ("dataDevolucao", "dataPrevistaDevolucao"):
        if col not in tabela.columns:
            tabela[col] = pd.NaT

    tabela["dataPrevistaDevolucao"] = pd.to_datetime(
        tabela["dataPrevistaDevolucao"], errors="coerce")
    hoje = pd.Timestamp.today().normalize()

    abertos = tabela[tabela["dataDevolucao"].isna()]
    atrasados = abertos[abertos["dataPrevistaDevolucao"] < hoje].copy()
    atrasados["diasDeAtraso"] = (hoje - atrasados["dataPrevistaDevolucao"]).dt.days

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as arquivo:
        atrasados.reindex(columns=COLUNAS).to_excel(
            arquivo, sheet_name="Atrasados", index=False)
        if "equipamento.tipo" in abertos.columns and not abertos.empty:
            (abertos.groupby("equipamento.tipo").size().rename("quantidade")
                    .to_excel(arquivo, sheet_name="Emprestados por tipo"))

    buffer.seek(0)
    return buffer