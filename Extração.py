import re
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# --- CONFIGURAÇÃO DAS COLUNAS ---
COL_X = "X"      # Informação 1
COL_Z = "Z"      # Informação 2
COL_V = "V"      # Informação 3
COL_XPED = "CK"   # Onde o XML joga o pedido (Tag <xPed>)
COL_DADOS = "BQ"  # Dados Adicionais (Onde o pedido às vezes é digitado)

PREFIXOS_VALIDOS = {"500", "400", "410", "590", "450"}

def extrair_pedido(txt: str) -> str:
    if not isinstance(txt, str) or not txt.strip() or txt.lower() == "nan":
        return ""
    # Encontra apenas números
    partes = re.findall(r'\d+', txt)
    # Filtra por 10 dígitos e prefixos SAP
    validos = [p for p in partes if len(p) == 10 and p[:3] in PREFIXOS_VALIDOS]
    if validos:
        return validos[0]
    if "ZVIC" in txt.upper():
        return "ZVIC"
    return ""

def col_letra_para_idx(letra: str) -> int:
    letra = letra.upper()
    idx = 0
    for c in letra:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1

def processar():
    root = tk.Tk()
    root.withdraw()
    print("📌 Selecione o arquivo Excel...")
    caminho_entrada = filedialog.askopenfilename(title="Selecione o relatório", filetypes=[("Excel files", "*.csv *.xlsx *.xls *.meso")])
    
    if not caminho_entrada: return

    try:
        # Lê a aba 2 (index 1)
        df = pd.read_excel(caminho_entrada, sheet_name=1, dtype=str)
        
        idx_x = col_letra_para_idx(COL_X)
        idx_z = col_letra_para_idx(COL_Z)
        idx_v = col_letra_para_idx(COL_V)
        idx_xped = col_letra_para_idx(COL_XPED)
        idx_dados = col_letra_para_idx(COL_DADOS)

        dados_finais = []

        for _, linha in df.iterrows():
            # Tenta pegar da coluna CK (xPed do XML) primeiro
            pedido = extrair_pedido(str(linha.iloc[idx_xped]))
            
            # Se não achou na CK, tenta na BQ (Dados Adicionais)
            if not pedido:
                pedido = extrair_pedido(str(linha.iloc[idx_dados]))

            dados_finais.append({
                "INFO_X": linha.iloc[idx_x],
                "INFO_Z": linha.iloc[idx_z],
                "INFO_V": linha.iloc[idx_v],
                "PEDIDO_FINAL": pedido
            })

        df_final = pd.DataFrame(dados_finais)
        caminho_saida = os.path.join(os.path.dirname(caminho_entrada), "resultado_extração_de_pedidos.xlsx")
        
        df_final.to_excel(caminho_saida, index=False)
        print(f"\n✅ SUCESSO!")
        print(f"📊 Colunas: {COL_X}, {COL_Z}, {COL_V} + Pedido")
        print(f"📍 Salvo em: {caminho_saida}")

    except Exception as e:
        print(f"❌ Erro: {e}")

if __name__ == "__main__":
    processar()