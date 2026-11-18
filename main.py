import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from relatorio import gerar_relatorio

app = FastAPI()
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://fronto-equipamentocontrollz.onrender.com"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

@app.get("/relatorio")
def relatorio():
    try:
        buffer = gerar_relatorio()
    except requests.RequestException as e:
        raise HTTPException(status_code=502, detail=f"Erro ao consultar o Spring: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="relatorio_emprestimos.xlsx"'},
    )