import os
from io import BytesIO
import requests
import pandas as pd

# Configurações (com suporte a variáveis de ambiente por segurança)
API = os.getenv("API_URL", "https://projeto-equipamentocontrol.onrender.com")
USUARIO = os.getenv("API_USER", "admin@empresa.com")
SENHA = os.getenv("API_PASSWORD", "123456")

def gerar_relatorio() -> BytesIO:
    try:
         # 1. Fazer login e guardar o token com validação de status
        print("Autenticando na API...")
        response_login = requests.post(f"{API}/auth/login", json={"email": USUARIO, "senha": SENHA})
        response_login.raise_for_status()
        
        token = response_login.json().get("token")
        if not token:
           print("Erro: Token não encontrado na resposta de login.")
           return

        cabecalho = {"Authorization": f"Bearer {token}"}

        # 2. Buscar todos os empréstimos
        print("Buscando dados de empréstimos...")
        response_emp = requests.get(f"{API}/emprestimos", headers=cabecalho)
        response_emp.raise_for_status()
        
        emprestimos = response_emp.json()
        if not emprestimos:
            print("Nenhum empréstimo cadastrado. Nada pra relatar.")
            return BytesIO()

        # 3. Transformar o JSON numa tabela
        tabela = pd.json_normalize(emprestimos)
        
        # Verificar se as colunas essenciais existem antes de processar
        colunas_obrigatorias = ["dataDevolucao", "dataPrevistaDevolucao"]
        for col in colunas_obrigatorias:
            if col not in tabela.columns:
                print(f"Erro: A coluna obrigatória '{col}' não foi retornada pela API.")
                return

        tabela["dataPrevistaDevolucao"] = pd.to_datetime(tabela["dataPrevistaDevolucao"], errors="coerce")
        hoje = pd.Timestamp.today().normalize()

        # 4. Separar os que ainda não voltaram e, desses, os atrasados
        abertos = tabela[tabela["dataDevolucao"].isna()]
        atrasados = abertos[abertos["dataPrevistaDevolucao"] < hoje].copy()
        atrasados["diasDeAtraso"] = (hoje - atrasados["dataPrevistaDevolucao"]).dt.days

        # 5. Contar os emprestados por tipo (validando se a coluna existe)
        if "equipamento.tipo" in abertos.columns:
            por_tipo = abertos.groupby("equipamento.tipo").size().rename("quantidade")
        else:
            por_tipo = pd.Series(dtype=int)

        # 6. Salvar a planilha com as duas abas de forma segura
        colunas_desejadas = [
            "equipamento.patrimonio", "equipamento.nome", "colaborador.nome",
            "colaborador.email", "dataPrevistaDevolucao", "diasDeAtraso"
        ]
        
        # Filtra apenas colunas que realmente existem no DataFrame para evitar erros
        colunas_existentes = [c for c in colunas_desejadas if c in atrasados.columns]

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as arquivo:
            atrasados[colunas_existentes].to_excel(arquivo, sheet_name="Atrasados", index=False)
            if not por_tipo.empty:
                por_tipo.to_excel(arquivo, sheet_name="Emprestados por tipo")

        buffer.seek(0)
        return buffer

    except requests.exceptions.RequestException as e:
        print(f"Erro de comunicação com a API: {e}")
    except PermissionError:
        print("Erro: O arquivo 'relatorio_emprestimos.xlsx' está aberto em outro programa. Feche-o e tente novamente.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    gerar_relatorio()