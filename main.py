import os
import PyPDF2
import re
import pandas as pd
from openpyxl import load_workbook
import shutil
import pdfplumber

DIST = 'COMGÁS'

class ExtratorFaturas:
    def __init__(self):
        self.regexes = {
            'cnpj': [r'\d{2}\.\d{3}\.?\d{3}\/?\d{4}\-?\s?\d{2}'],
            'valor_total': [r'R\$\s(\d+\.?\d+\,\d{2})\s'],
            'volume_total': [r'Total\s(\d+\.?\,?\d+\.?\,?\d+?)\.?\,?\s', 
                             r'Natural\sR[$]\s\d{1,3}(?:\.\d{3})*,\d{2}\s(\d{1,3}(?:\.\d{3})*,\d{1,3})'],
            'data_emissao': [r'apresentação\s(\d{2}\.\d{2}\.\d{2,4})'],
            'data_inicio': [r'\d{2}\.\d{2}\.\d{4}\d{2}\.\d{2}\.\d{4}(\d{2}\.\d{2}\.\d{4})\d{2}\.\d{2}\.\d{2,4}'],
            'data_fim': [r'\d{2}\.\d{2}\.\d{2,4}(\d{2}\.\d{2}\.\d{2,4})\d{2}\.\d{2}\.\d{2,4}'],            
            'numero_fatura': [r'\s(\d{3}\.\d{3}\.\d{3})\s'],
            'valor_icms': [r'ICMS\s?R\$\s(\d+\.?\d+\,\d{2})\s'],     
            'correcao_pcs': [r'Total\s\d+\.?\,?\d+\.?\,?\d+?\.?\,?\s(\d+\.?\,?\d+\.?\,?\d+?)\s']
        }

    def extrair_informacoes(self, texto):
        informacoes = {}
        for chave, regex_list in self.regexes.items():
            for regex in regex_list:
                match = re.search(regex, texto)
                if match:
                    informacoes[chave] = match.group(1) if match.groups() else match.group(0)
                    break
        return informacoes

    def processar_pagina(self, texto_pagina):
        for chave, padroes in self.regexes.items():
            for padrao in padroes:
                correspondencia = re.findall(padrao, texto_pagina)
                if correspondencia:
                    print(f"{chave}: {correspondencia[0]}")
                    break
            else:
                print(f"{chave}: Não encontrado")

    def processar_pdf(self, caminho_pdf):
        texto_completo = ""
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
                else:
                    print("Texto da página está vazio, interrompendo o processamento.")
                    break
        self.processar_pagina(texto_completo)

def extrair_texto(caminho_do_pdf):
    texto = ""
    try:
        # Primeiro tenta PyPDF2
        with open(caminho_do_pdf, 'rb') as arquivo_pdf:
            leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
            for pagina in leitor_pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto += texto_pagina.replace('\n', ' ') + ' '
        
        # Se PyPDF2 falhar ou retornar texto vazio, tenta pdfplumber
        if not texto.strip():
            with pdfplumber.open(caminho_do_pdf) as pdf:
                for pagina in pdf.pages:
                    texto_pagina = pagina.extract_text()
                    if texto_pagina:
                        texto += texto_pagina.replace('\n', ' ') + ' '
    
    except Exception as e:
        print(f"Erro ao abrir o PDF {caminho_do_pdf}: {e}")
    
    texto = texto.strip()
    if not texto:
        print(f"Nenhum texto foi extraído do PDF {caminho_do_pdf}.")
    return texto

def registro_existe(df, cnpj, data_inicio, data_fim, valor_total):
    return not df[(df['CNPJ'] == cnpj) & (df['Data Inicio'] == data_inicio) & (df['Data Fim'] == data_fim) & (df['Valor Total'] == valor_total)].empty

def adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo):
    try:
        df = pd.read_excel(caminho_planilha)
    except FileNotFoundError:
        print(f"O arquivo '{caminho_planilha}' não foi encontrado. Criando um novo.")
        df = pd.DataFrame(columns=['CNPJ', 'Valor Total', 'Volume Total', 'Data Emissão', 'Data Inicio', 'Data Fim', 'Número Fatura', 'Valor ICMS', 'Correção PCS', 'Distribuidora', 'Nome do Arquivo'])
    
    cnpj = informacoes.get('cnpj', '')
    data_inicio = informacoes.get('data_inicio', '')
    data_fim = informacoes.get('data_fim', '')
    valor_total = informacoes.get('valor_total', '')

    if registro_existe(df, cnpj, data_inicio, data_fim, valor_total):
        print(f"Registro duplicado encontrado para CNPJ: {cnpj}, Data Início: {data_inicio}, Data Fim: {data_fim}, Valor Total: {valor_total}. Não será inserido.")
        return False 

    # Converte os valores para numéricos
    valor_total = pd.to_numeric(str(informacoes.get('valor_total', '')).replace('.', '').replace(',', '.'), errors='coerce')
    volume_total = pd.to_numeric(str(informacoes.get('volume_total', '')).replace('.', '').replace(',', '.'), errors='coerce')
    valor_icms = pd.to_numeric(str(informacoes.get('valor_icms', '')).replace('.', '').replace(',', '.'), errors='coerce')
    
    nova_linha = pd.DataFrame([{
        'CNPJ': cnpj,
        'Valor Total': valor_total,
        'Volume Total': volume_total,
        'Data Emissão': informacoes.get('data_emissao', ''),
        'Data Inicio': data_inicio,
        'Data Fim': data_fim,
        'Número Fatura': informacoes.get('numero_fatura', ''),
        'Valor ICMS': valor_icms,
        'Correção PCS': '',
        'Distribuidora': DIST,
        'Nome do Arquivo': nome_arquivo  # Adiciona o nome do arquivo
    }])
    df = pd.concat([df, nova_linha], ignore_index=True)
    df.to_excel(caminho_planilha, index=False)
    return True  # Indica que o registro foi inserido

def mover_arquivo(origem, destino):
    shutil.move(origem, destino)
    print(f"Arquivo movido para {destino}")

def main(file_path, pdf_file, caminho_planilha):
    texto_pypdf = extrair_texto(pdf_file)
    if not texto_pypdf:
        print(f"Erro ao extrair texto do PDF: {pdf_file}")
        return

    extrator = ExtratorFaturas()
    informacoes = extrator.extrair_informacoes(texto_pypdf)
    
    # Verifica se todos os campos foram extraídos
    campos_necessarios = ['cnpj', 'valor_total', 'volume_total', 'data_emissao', 'data_inicio', 'data_fim', 'numero_fatura', 'valor_icms', 'correcao_pcs']
    campos_faltantes = [campo for campo in campos_necessarios if not informacoes.get(campo)]
    
    if campos_faltantes:
        print(f"Campos faltantes no PDF {pdf_file}: {', '.join(campos_faltantes)}")
        return

    nome_arquivo = os.path.basename(pdf_file)  # Extrai apenas o nome do arquivo
    inserido = adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo)
    print(informacoes)

    if inserido:
        destino = os.path.join(diretorio_destino, nome_arquivo)
        mover_arquivo(pdf_file, destino)
    else:
        print('Arquivo já foi inserido na planilha. Não será movido.')
    
# Exemplo de uso
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Comgás\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Comgás\Lidos'
caminho_planilha = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\COMGAS.xlsx'

for arquivo in os.listdir(file_path):
    if arquivo.endswith('.pdf') or arquivo.endswith('.PDF'):
        arquivo_full = os.path.join(file_path, arquivo)
        arquivo = os.path.basename(arquivo)

        main(arquivo, arquivo_full, caminho_planilha)