"""
Importa colaboradores da planilha "1 - Check list RH_DP" para o MaxGroup RH.
Uso:
  python importar_planilha.py            -> importa para a NUVEM
  python importar_planilha.py --local    -> importa para o servidor local (porta 5000)
"""
import sys
import re
import openpyxl
import urllib.request
import urllib.parse
import json
from datetime import datetime, date

CLOUD_URL = "https://maxgroup-rh-production.up.railway.app"
LOCAL_URL = "http://localhost:5000"
BASE_URL = LOCAL_URL if "--local" in sys.argv else CLOUD_URL

PLANILHA = r"C:\Users\MaxFom3\Downloads\1 - Check list RH_DP (1).xlsx"

# ── Mapeamentos ────────────────────────────────────────────────────────────────

# Apenas empresas do grupo MaxGroup — qualquer outra é IGNORADA
EMPRESA_MAP = {
    "FORM":           "MaxForm",
    "MAXFORM":        "MaxForm",
    "PHARMA":         "MaxPharma",
    "MAXPHARMA":      "MaxPharma",
    "FOODS RES":      "MaxFoods Restaurante",
    "FOODS MKT":      "MaxFoods",
    "FOODS":          "MaxFoods",
    "MAXFOODS RES":   "MaxFoods Restaurante",
    "MAXFOODS":       "MaxFoods",
    "MAXGROUP":       "MaxGroup",
    "GROUP":          "MaxGroup",
}

FILIAL_MAP = {
    "105 S":   "105 Asa Sul",
    "105 SUL": "105 Asa Sul",
    "315 N":   "315 Asa Norte",
    "315 ASA NORTE": "315 Asa Norte",
    "103 N":   "103 Asa Norte",
    "103 SQS": "103 Asa Sul",
    "103 SW":  "103 Sudoeste",
    "103 SUL": "103 Asa Sul",
}

CONTRATO_MAP = {
    "CLT":   "CLT",
    "PJ":    "PJ",
    "EST":   "Estagio",
    "TERC.": "Terceirizado",
    "TERC":  "Terceirizado",
}


def fmt_date(val):
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        if isinstance(val, datetime):
            return val.strftime("%Y-%m-%d")
        return val.strftime("%Y-%m-%d")
    return ""


def limpar_colaboradores():
    """Busca e deleta todos os colaboradores existentes para evitar duplicatas."""
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
        print("  Limpeza concluida.")
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


def importar_aba(ws, desligamento_col, label, forcar_desligado=False, ja_importados=None):
    """
    ws               — worksheet openpyxl
    desligamento_col — índice da coluna com data de desligamento (ou None)
    label            — "ATIVO" | "DESLIG"
    forcar_desligado — se True, garante data_desligamento != '' mesmo sem data na planilha
    ja_importados    — set de (nome, empresa) já inseridos (para deduplicar)
    """
    if ja_importados is None:
        ja_importados = set()

    ok = err = skip = 0
    hoje = date.today().strftime("%Y-%m-%d")

    for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        nome = row[3]
        if not nome or not str(nome).strip():
            continue
        nome = str(nome).strip().title()

        # ── Empresa — IGNORA se não for do grupo MaxGroup ──────────────────
        empresa_raw = str(row[0] or "").strip().upper()
        empresa = EMPRESA_MAP.get(empresa_raw, "")
        if not empresa:
            # tenta match parcial (ex: "MAXFOODS RESTAURANTE")
            for k, v in EMPRESA_MAP.items():
                if k in empresa_raw:
                    empresa = v
                    break
        if not empresa:
            # empresa desconhecida → pula linha
            skip += 1
            print(f"  [SKIP] Linha {i}: empresa '{empresa_raw}' nao e do grupo — {nome}")
            continue

        # ── Filial ─────────────────────────────────────────────────────────
        filial_raw = str(row[2] or "").strip().upper()
        localizacao = FILIAL_MAP.get(filial_raw, "")
        if not localizacao:
            for k, v in FILIAL_MAP.items():
                if k in filial_raw:
                    localizacao = v
                    break
        if not localizacao:
            localizacao = "105 Asa Sul"
            print(f"  [AVISO] Linha {i}: filial '{filial_raw}' desconhecida -- {nome} -> padrao '105 Asa Sul'")

        # ── Deduplicacao ───────────────────────────────────────────────────
        chave = (nome.lower(), empresa)
        if chave in ja_importados:
            skip += 1
            print(f"  [DUP]  Linha {i}: duplicata ignorada — {nome} ({empresa})")
            continue
        ja_importados.add(chave)

        # ── Outros campos ─────────────────────────────────────────────────
        contrato_raw = str(row[1] or "").strip().upper()
        tipo_contrato = CONTRATO_MAP.get(contrato_raw, contrato_raw or "")

        cargo   = str(row[4] or "").strip().title() if row[4] else ""
        horario = str(row[5] or "").strip()        if row[5] else ""
        email   = str(row[9] or "").strip().lower() if row[9] else ""

        data_admissao    = fmt_date(row[6])
        data_aniversario = fmt_date(row[8])

        # ── Data de desligamento ───────────────────────────────────────────
        data_desligamento = ""
        if desligamento_col is not None:
            data_desligamento = fmt_date(row[desligamento_col])

        # Para a aba Desligados: garantir que sempre haja data de desligamento
        if forcar_desligado and not data_desligamento:
            data_desligamento = hoje

        data = {
            "nome_completo":      nome,
            "empresa":            empresa,
            "localizacao":        localizacao,
            "tipo_contrato":      tipo_contrato,
            "cargo":              cargo,
            "horario":            horario,
            "email":              email,
            "data_admissao":      data_admissao,
            "data_aniversario":   data_aniversario,
            "data_desligamento":  data_desligamento,
            "remuneracao":        "0",
            "premiacao":          "0",
            "vale_transporte":    "0",
            "auxilio_transporte": "0",
            "vale_alimentacao":   "0",
            "assiduidade":        "0",
            "comissao":           "0",
        }

        status, msg = post_colaborador(data)
        if status in (200, 302):
            print(f"  OK  [{label}] {nome} ({empresa} / {localizacao})")
            ok += 1
        else:
            print(f"  ERR [{label}] {nome} -- HTTP {status}: {msg[:80]}")
            err += 1

    return ok, err, skip


def main():
    print(f"\n  MaxGroup RH -- Importacao de Planilha")
    print(f"  Destino: {BASE_URL}\n")

    wb = openpyxl.load_workbook(PLANILHA, data_only=True)

    print(">> Limpando colaboradores existentes...")
    limpar_colaboradores()

    # Conjunto compartilhado para deduplicar entre as duas abas
    ja_importados = set()

    # Aba ativos — col 26 = Desligamento (pode ter data se alguém foi desligado)
    ws_ativos = wb["Controle RH e DP"]
    print("\n>> Importando colaboradores ativos...")
    ok1, err1, skip1 = importar_aba(
        ws_ativos,
        desligamento_col=26,
        label="ATIVO",
        forcar_desligado=False,
        ja_importados=ja_importados,
    )

    # Aba desligados — col 21 = Desligamento; sempre garante data de desligamento
    ws_deslig = wb["Desligados"]
    print("\n>> Importando colaboradores desligados...")
    ok2, err2, skip2 = importar_aba(
        ws_deslig,
        desligamento_col=21,
        label="DESLIG",
        forcar_desligado=True,   # <-- garante que fiquem como desligados
        ja_importados=ja_importados,
    )

    print(f"\n  Resultado: {ok1+ok2} importados, {err1+err2} erros, {skip1+skip2} ignorados")
    print(f"    Ativos:     {ok1} ok  |  {err1} erros  |  {skip1} ignorados")
    print(f"    Desligados: {ok2} ok  |  {err2} erros  |  {skip2} ignorados")
    print(f"\n  Acesse: {BASE_URL}\n")


if __name__ == "__main__":
    main()
