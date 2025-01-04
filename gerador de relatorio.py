import os
import re
from PyPDF2 import PdfReader
import pandas as pd
from decimal import Decimal

def extract_valor_servico(texto):
    padrao = r"Valor\s*doServi[çc]o.*?\n(R?\$?\s*[\d.,]+)"
    match = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
    
    if match:
        valor_str = match.group(1).replace('R$', '').strip()
        valor_str = valor_str.replace('.', '').replace(',', '.')
        try:
            return Decimal(valor_str)
        except:
            return Decimal('0')
    return Decimal('0')

def extract_nome_tomador(texto):
    padrao = r"Nome\s*/\s*Nome\s*Empresarial\s*[:\-]?\s*([A-Za-zÀ-ÿ\s]+(?:\s+[A-Za-zÀ-ÿ]+)*)(?=\s*E-mail)"
    texto = texto.replace("\n", " ").strip()
    match = re.search(padrao, texto, re.IGNORECASE)
    
    if match:
        return match.group(1).strip()
    return "Nome não encontrado"

def processar_pdf(caminho_arquivo):    
    try:
        reader = PdfReader(caminho_arquivo)
        texto_completo = ""
        for pagina in reader.pages:
            texto_completo += pagina.extract_text()
        
        valor = extract_valor_servico(texto_completo)
        nome_tomador = extract_nome_tomador(texto_completo)
        
        return {
            'arquivo': os.path.basename(caminho_arquivo),
            'caminho': caminho_arquivo,
            'valor': float(valor),
            'tomador': nome_tomador
        }
    except Exception as e:
        return None

def processar_pasta(pasta_raiz):
    resultados = []
    for pasta_atual, _, arquivos in os.walk(pasta_raiz):
        pdfs = [a for a in arquivos if a.lower().endswith('.pdf')]
        for arquivo in pdfs:
            resultado = processar_pdf(os.path.join(pasta_atual, arquivo))
            if resultado:
                resultados.append(resultado)
    return pd.DataFrame(resultados) if resultados else pd.DataFrame()

def main():
    pasta_raiz = r"Sua pasta"
    df_resultados = processar_pasta(pasta_raiz)
    
    if df_resultados.empty:
        print("Nenhum dado encontrado")
        return
    
    nome_arquivo = 'resultado_notas_fiscais.xlsx'
    df_resultados.to_excel(nome_arquivo, index=False)
    total = df_resultados['valor'].sum()
    
    print(f"\nTotal de arquivos: {len(df_resultados)}")
    print(f"Valor total: R$ {total:,.2f}")
    print(f"Resultados salvos em: {nome_arquivo}")

if __name__ == "__main__":
    main()
