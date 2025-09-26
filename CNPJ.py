import requests
import pandas as pd
import time
import json
import os
import re
from datetime import timedelta

# Lista de CNAEs alvo (secundários)
CNAE_PROMO = {
    "46.31-1-00","46.32-0-03","47.24-5-00","46.33-8-01","46.34-6-03","47.22-9-01",
    "46.33-8-02","46.35-4-01","46.35-4-02","10.91-1-01","46.92-3-00","47.21-1-04",
    "47.23-7-00","56.11-2-04","10.66-0-00","46.17-6-00"
}

CACHE_FILE = "cnpj_cache.json"

def carregar_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)

def consultar_cnpj(cnpj, cache):
    if cnpj in cache:
        return cache[cnpj]

    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    try:
        r = requests.get(url, timeout=30)
        if r.status_code == 200:
            dados = r.json()
            cache[cnpj] = dados
            salvar_cache(cache)
            return dados
    except Exception:
        pass
    return None

def processar_planilha(entrada, saida):
    df = pd.read_excel(entrada)
    cache = carregar_cache()

    inicio = 859  # linha 860 no Excel
    total = len(df) - inicio
    start_time = time.time()

    for idx in range(inicio, len(df)):
        valor = str(df.iloc[idx, 2]).strip()  # Coluna C (CNPJs)

        if not valor or valor.lower() == "nan":
            continue

        # Remove máscara do CNPJ
        cnpj = re.sub(r"\D", "", valor)

        if len(cnpj) != 14:
            print(f"⚠️ CNPJ inválido na linha {idx+1}: {valor}")
            continue

        dados = consultar_cnpj(cnpj, cache)
        if not dados:
            print(f"❌ Erro ao consultar {cnpj}")
            continue

        # CNAE principal → Coluna B (índice 1)
        cnae_principal = f"{dados.get('cnae_fiscal', '')} - {dados.get('cnae_fiscal_descricao', '')}"
        df.iat[idx, 1] = cnae_principal

        # CNAEs secundários → Coluna A (índice 0) → apenas o primeiro da lista de promo
        secundarios = dados.get("cnaes_secundarios", [])
        encontrado = next(
            (f"{c['codigo']} - {c['descricao']}" for c in secundarios if c["codigo"] in CNAE_PROMO),
            None
        )
        if encontrado:
            df.iat[idx, 0] = encontrado

        # Contador + previsão
        processados = idx - inicio + 1
        tempo_passado = time.time() - start_time
        tempo_medio = tempo_passado / processados
        tempo_restante = tempo_medio * (total - processados)

        print(f"✅ {processados}/{total} | "
              f"Tempo médio: {tempo_medio:.2f}s | "
              f"Restante: {str(timedelta(seconds=int(tempo_restante)))}")

        time.sleep(0.5)

        # Salva progresso parcial a cada 100
        if processados % 100 == 0:
            df.to_excel(saida, index=False)
            print(f"💾 Progresso salvo em {saida}")

    # Salva no final
    df.to_excel(saida, index=False)
    print(f"✅ Arquivo final salvo em: {saida}")

# ------------------------------
# Executar
# ------------------------------
if __name__ == "__main__":
    processar_planilha(
        r"C:\Users\phpereira\Desktop\CARTEIRA EQ VENDAS - 05-08-25 .xlsb.xlsx",
        r"C:\Users\phpereira\Desktop\CARTEIRA_EQ_VENDAS_RESULTADO.xlsx"
    )
