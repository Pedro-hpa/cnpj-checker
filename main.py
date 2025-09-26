import json
import os
import re
import time
import requests
import pandas as pd
from datetime import datetime

# ====== CONFIGURAÇÕES ======
ARQUIVO = r"V:\Star Class-Auditoria\SIC DE VENDAS - 2025\CARTEIRA DE CLIENTES\BOLSAO SPRINTER ATIVOS\BOLSAO SPRINTER ATIVO.xlsx"
ABA = "PJ"            # nome da aba com os CNPJs
COL_IDX_CNPJ = 5      # coluna F (0-based → 5)
COL_IDX_STATUS = 6    # coluna G (0-based → 6)
LINHA_INICIO = 0      # começa na linha 2 (0-based → 1)
SALVAR_CADA = 100     # salva a planilha a cada X consultas
# ===========================

API_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"
USER_AGENT = "Mozilla/5.0 (compatible; CNPJ-Checker/2.0)"
CACHE_PATH = os.path.splitext(ARQUIVO)[0] + "_cache.json"

def apenas_digitos(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def calcula_dv_cnpj(cnpj12: str, pesos):
    soma = sum(int(a) * b for a, b in zip(cnpj12, pesos))
    resto = soma % 11
    return "0" if resto < 2 else str(11 - resto)

def valida_cnpj(cnpj: str) -> bool:
    c = apenas_digitos(cnpj).zfill(14)
    if len(c) != 14 or c == c[0] * 14:
        return False
    base = c[:12]
    dv1 = calcula_dv_cnpj(base, [5,4,3,2,9,8,7,6,5,4,3,2])
    dv2 = calcula_dv_cnpj(base + dv1, [6,5,4,3,2,9,8,7,6,5,4,3,2])
    return c[-2:] == dv1 + dv2

def normaliza_cnpj(cnpj: str) -> str:
    return apenas_digitos(str(cnpj)).zfill(14) if cnpj not in [None, "nan"] else ""

def carregar_cache(path: str):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def salvar_cache(path: str, cache: dict):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def consulta_brasilapi(cnpj: str) -> dict:
    headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
    r = requests.get(API_URL.format(cnpj), headers=headers, timeout=30)
    r.raise_for_status()
    return r.json()

def extrair_campos(payload: dict) -> dict:
    if not isinstance(payload, dict):
        return {"situacao": None, "erro": "payload inválido"}
    if "descricao_situacao_cadastral" not in payload:
        return {"situacao": None, "erro": "campo não encontrado"}
    return {"situacao": payload.get("descricao_situacao_cadastral"), "erro": None}

def salvar_planilha(df, nome_saida):
    with pd.ExcelWriter(nome_saida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=ABA, index=False)
    print(f"[AUTO-SAVE] Planilha salva em: {nome_saida}")

def formatar_tempo(segundos: float) -> str:
    if segundos < 60:
        return f"{int(segundos)}s"
    elif segundos < 3600:
        return f"{int(segundos // 60)}min"
    else:
        h = int(segundos // 3600)
        m = int((segundos % 3600) // 60)
        return f"{h}h {m}min"

def main():
    df = pd.read_excel(ARQUIVO, sheet_name=ABA, dtype=str)
    cache = carregar_cache(CACHE_PATH)

    total = len(df)
    inicio_exec = datetime.now()
    consultados = 0

    print(f"[INFO] Linhas totais na aba {ABA}: {total}")
    print(f"[INFO] Iniciando na linha Excel {LINHA_INICIO+1}")

    SAIDA = os.path.splitext(ARQUIVO)[0] + "_com_situacao.xlsx"

    try:
        for i in range(LINHA_INICIO, total):
            raw = df.iat[i, COL_IDX_CNPJ]
            if pd.isna(raw) or str(raw).strip() == "":
                df.iat[i, COL_IDX_STATUS] = ""
                continue

            cnpj = normaliza_cnpj(raw)

            if cnpj in cache:
                situacao = cache[cnpj].get("situacao")
                erro = cache[cnpj].get("erro")
                df.iat[i, COL_IDX_STATUS] = situacao if situacao else (f"ERRO: {erro}" if erro else "")
                print(f"[{i+1}/{total}] {cnpj} - CACHE → {df.iat[i, COL_IDX_STATUS]}")
                continue

            if not valida_cnpj(cnpj):
                cache[cnpj] = {"situacao": None, "erro": "CNPJ inválido"}
                df.iat[i, COL_IDX_STATUS] = "ERRO: CNPJ inválido"
                salvar_cache(CACHE_PATH, cache)
                print(f"[{i+1}/{total}] {cnpj} - CNPJ inválido")
                continue

            try:
                print(f"[{i+1}/{total}] {cnpj} - consultando...")
                payload = consulta_brasilapi(cnpj)
                dados = extrair_campos(payload)
                cache[cnpj] = dados
                situacao = dados.get("situacao") or ""
                erro = dados.get("erro")
                df.iat[i, COL_IDX_STATUS] = situacao if situacao else (f"ERRO: {erro}" if erro else "")
                salvar_cache(CACHE_PATH, cache)
                print(f"   → Situação: {df.iat[i, COL_IDX_STATUS]}")
            except Exception as e:
                cache[cnpj] = {"situacao": None, "erro": str(e)}
                df.iat[i, COL_IDX_STATUS] = f"ERRO: {e}"
                salvar_cache(CACHE_PATH, cache)
                print(f"   → ERRO: {e}")

            consultados += 1

            # Calcula ETA
            tempo_passado = (datetime.now() - inicio_exec).total_seconds()
            media = tempo_passado / consultados if consultados > 0 else 0
            faltando = total - (i + 1)
            eta = faltando * media
            print(f"   → Progresso: {consultados}/{total-LINHA_INICIO} | ETA: {formatar_tempo(eta)}")

            # Salvar parcial a cada X consultas
            if consultados % SALVAR_CADA == 0:
                salvar_planilha(df, SAIDA)

            # 👉 não tem mais delay fixo aqui
            time.sleep(0.5)  # opcional: só pra não sobrecarregar a API

    except KeyboardInterrupt:
        print("\n[INFO] Interrompido manualmente.")
        salvar_planilha(df, SAIDA)
        print(f"[OK] Cache salvo: {CACHE_PATH}")
        return

    # Se rodar até o fim → salva planilha final
    salvar_planilha(df, SAIDA)
    print(f"[OK] Cache salvo: {CACHE_PATH}")
    print(f"[OK] Processamento concluído.")

if __name__ == "__main__":
    main()
