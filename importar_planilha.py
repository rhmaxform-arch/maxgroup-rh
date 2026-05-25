"""
Importa colaboradores da planilha "1 - Check list RH_DP" para o MaxGroup RH.
Uso:
  python importar_planilha.py            -> importa para a NUVEM
  python importar_planilha.py --local    -> importa para o servidor local (porta 5000)
"""
import sys
import openpyxl
import urllib.request
import urllib.parse
import json
from datetime import datetime

CLOUD_URL = "https://maxgroup-rh-production.up.railway.app"
LOCAL_URL = "http://localhost:5000"
BASE_URL = LOCAL_URL if "--local" in sys.argv else CLOUD_URL

PLANILHA = r"C:\Users\MaxFom3\Downloads\1 - Check list RH_DP (1).xlsx"

# ── Mapeamentos ────────────────────────────────────────────────────────────────

EMPRESA_MAP = {
    "FORM":      "MaxForm",
    "MAXFORM":   "MaxForm",
    "PHARMA":    "MaxPharma",
    "MAXPHARMA": "MaxPharma",
    "FOODS RES": "MaxFoods Restaurante",
    "FOODS MKT": "MaxFoods",
    "FOODS":     "MaxFoods",
    "MAXGROUP":  "MaxGroup",
    "GROUP":     "MaxGroup",
}

FILIAL_MAP = {
    "105 S":   "105 Asa Sul",
    "105 SUL": "105 Asa Sul",
    "315 N":   "315 Asa Norte",
    "103 N":   "103 Asa Norte",
    "103 SQS": "103 Asa Sul",
    "103 SW":  "103 Sudoeste",
}

CONTRATO_MAP = {
    "CLT":   "CLT",
    "PJ":    "PJ",
    "EST":   "Estágio",
    "TERC.": "Terceirizado",
    "TERC":  "Terceirizado",
}


def fmt_date(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    return ""


def limpar_colaboradores():
    """Busca e deleta todos os colaboradores existentes para evitar duplicatas."""
    import re
    req = urllib.request.Request(
        f"{BASE_URL}/?status=todos",
        headers={"User-Agent": "MaxGroupImport/1.0"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8")
        ids = re.findall(r'/colaborador/(\d+)/excluir', html)
        if not ids:
            print("  Nenhum colaborador existente para remover.")
            return
        print(f"  Removendo {len(ids)} colaboradores existentes...")
        for cid in ids:
            del_req = urllib.request.Request(
                f"{BASE_URL}/colaborador/{cid}/excluir",
                data=b"",
                headers={"Content-Type": "application/x-www-form-urlencoded",
                         "User-Agent": "MaxGroupImport/1.0"},
                method="POST",
            )
            try:
                with urllib.request.urlopen(del_req, timeout=15):
                    pass
            except Exception:
                pass
        print(f"  Limpeza concluida.")
    except Exception as e:
        print(f"  Erro ao limpar: {e}")


def post_colaborador(data):
    encoded = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(
        f"{BASE_URL}/colaborador/novo",
        data=encoded,
        headers={"Content-Type": "application/x-www-form-urlencoded",
                 "User-Agent": "MaxGroupImport/1.0"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, ""
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode()[:200]
    except Exception as ex:
        return 0, str(ex)


def importar_aba(ws, tem_desligamento_col, label):
    ok = err = skip = 0
    for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        nome = row[3]
        if not nome or not str(nome).strip():
            continue
        nome = str(nome).strip().title()

        empresa_raw = str(row[0] or "").strip().upper()
        empresa = EMPRESA_MAP.get(empresa_raw, "")
        if not empresa:
            print(f"  [AVISO] Linha {i}: empresa desconhecida '{empresa_raw}' — {nome}")
            empresa = empresa_raw or "MaxGroup"

        filial_raw = str(row[2] or "").strip().upper()
        localizacao = FILIAL_MAP.get(filial_raw, "")
        if not localizacao:
            # tenta match parcial
            for k, v in FILIAL_MAP.items():
                if k in filial_raw:
                    localizacao = v
                    break
            if not localizacao:
                localizacao = "105 Asa Sul"
                print(f"  [AVISO] Linha {i}: filial desconhecida '{filial_raw}' -- {nome} -> padrao '105 Asa Sul'")

        contrato_raw = str(row[1] or "").strip().upper()
        tipo_contrato = CONTRATO_MAP.get(contrato_raw, contrato_raw or "")

        cargo   = str(row[4] or "").strip().title() if row[4] else ""
        horario = str(row[5] or "").strip() if row[5] else ""
        email   = str(row[9] or "").strip().lower() if row[9] else ""

        data_admissao    = fmt_date(row[6])
        data_aniversario = fmt_date(row[8])

        # coluna de desligamento (índice 26 na aba ativa, 21 nos desligados)
        data_desligamento = ""
        if tem_desligamento_col is not None:
            data_desligamento = fmt_date(row[tem_desligamento_col])

        data = {
            "nome_completo":   nome,
            "empresa":         empresa,
            "localizacao":     localizacao,
            "tipo_contrato":   tipo_contrato,
            "cargo":           cargo,
            "horario":         horario,
            "email":           email,
            "data_admissao":   data_admissao,
            "data_aniversario": data_aniversario,
            "data_desligamento": data_desligamento,
            "remuneracao":     "0",
            "premiacao":       "0",
            "vale_transporte": "0",
            "auxilio_transporte": "0",
            "vale_alimentacao": "0",
            "assiduidade":     "0",
            "comissao":        "0",
        }

        status, msg = post_colaborador(data)
        if status in (200, 302):
            print(f"  OK  [{label}] {nome} ({empresa} / {localizacao})")
            ok += 1
        else:
            print(f"  ERR [{label}] {nome} — HTTP {status}: {msg[:80]}")
            err += 1

    return ok, err


def main():
    print(f"\n  MaxGroup RH — Importação de Planilha")
    print(f"  Destino: {BASE_URL}\n")

    wb = openpyxl.load_workbook(PLANILHA, data_only=True)

    print(">> Limpando colaboradores existentes...")
    limpar_colaboradores()

    # Aba ativos — coluna 26 (índice) = Desligamento
    ws_ativos = wb["Controle RH e DP"]
    print(">> Importando colaboradores ativos...")
    ok1, err1 = importar_aba(ws_ativos, tem_desligamento_col=26, label="ATIVO")

    # Aba desligados — coluna 21 (índice) = Desligamento
    ws_deslig = wb["Desligados"]
    print("\n>> Importando colaboradores desligados...")
    ok2, err2 = importar_aba(ws_deslig, tem_desligamento_col=21, label="DESLIG")

    total_ok  = ok1 + ok2
    total_err = err1 + err2
    print(f"\n  Resultado: {total_ok} importados, {total_err} erros")
    print(f"  Acesse: {BASE_URL}\n")


if __name__ == "__main__":
    main()
