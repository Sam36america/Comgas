from PIL import Image
import pytesseract
import re
import os
import pandas as pd
import numpy as np
from config import corte_comgas, caminho_excel
from funcoes import *

#in/out 3100, 1300, 3800, 1450
#X-Y X-Y
correcao_pcs= '' 
dist= 'COMGÁS'
usuario_conectado = 'samuel.santos'
#usuario_conectado = 'caio.augusto'

def extrator_cnpj(imagem, cordenadas): #
        try:
                cnpj = imagem.crop(corte[cordenadas])
                #cnpj.show()
                cnpj = pdf_ocr(cnpj)
                cnpj = cnpj.replace(' ', '')
                cnpj = cnpj.replace (',','').replace('/','').replace('-','').replace('.', '')
                cnpj = re.findall(r'\s?\d{12}\-?\d{2}\s?', cnpj)
                return cnpj[0]
        except:
                return False

def extrator_valor_total(imagem, coordenadas): #
        try:        
                valor_total = imagem.crop(corte[coordenadas])
                #valor_total.show()
                valor_total = pdf_ocr(valor_total)
                if '§' in valor_total:
                    valor_total = valor_total.replace('§', '5')
                valor_total = re.findall(r'(\d{1,3}[\,\.]?\d{1,3}\.?\s?\,?\d{1,3}\.?\s?\,?\d{1,2})',valor_total)
                valor_total  = valor_total[0].strip()
                #volume_total = round(float(volume_total),4)
                valor_total = valor_total.replace('.','').replace(",",".")
                return valor_total
        except:
                return False
    
def extrator_volume_total(imagem, coordenadas):#
        try:        
                volume_total = imagem.crop(corte[coordenadas])
                #volume_total.show()
                volume_total = pdf_ocr(volume_total)
                volume_total = re.findall(r'(\d{1,3}[\,\.]?\d{1,3}\.?\s?\,?\d{1,3}\.?\s?\,?\d{1,2})',volume_total)
                volume_total = volume_total[0].strip()
                volume_total = volume_total.replace('.','').replace(",",".")
                #volume_total = round(float(volume_total),5)
                volume_total = str(volume_total)
                
                return volume_total
        
        except:
                return False
    
def extrator_data_emissao(imagem, coordenadas):#
        try:        
                data_emissao = imagem.crop(corte[coordenadas])
                #data_emissao.show()
                data_emissao = pdf_ocr(data_emissao)
                data_emissao = re.findall(r'\d{2}\.?\/?\d{2}\.?\/?\d{4}',data_emissao)
                data_emissao  = data_emissao[0].strip()
       
                return data_emissao
        except:
                return False
    
def extrator_data_inicio(imagem, coordenadas):#
    
        try:
                data_inicio = imagem.crop(corte[coordenadas])
                #data_inicio.show()
                data_inicio = pdf_ocr(data_inicio)
                data_inicio = re.findall(r'\d{2}\.?\/?\d{2}\.?\/?\d{4}',data_inicio)
                data_inicio  = data_inicio[0].strip()

                return data_inicio
        except:
                return False

def extrator_data_fim(imagem, coordenadas):#

        try:
                data_fim = imagem.crop(corte[coordenadas])
                #data_fim.show()
                data_fim = pdf_ocr(data_fim)
                data_fim = re.findall(r'[ate]?\s?(\d{2}\/?\.?\d{2}\/?\.?\d{2,4})\s?\;?',data_fim)
                data_fim = data_fim[0].strip()

                return data_fim
        except:
                return False
        
def extrator_numero_fatura(imagem, coordenadas):#
        
        try:
                numero_fatura = imagem.crop(corte[coordenadas])
                #numero_fatura.show()
                numero_fatura = pdf_ocr(numero_fatura)
                numero_fatura = re.findall(r'N?\.?\s?(\d+?\s?\d+\s?\.?\d+\.?\d+\.?\d+)',numero_fatura)
                numero_fatura  = numero_fatura[0].strip()
                
                return numero_fatura
        except:
                        
                return False
        
def extrator_valor_icms(imagem, coordenadas):
        
        try:
                valor_icms = imagem.crop(corte[coordenadas])
                #valor_icms.show()
                valor_icms = pdf_ocr(valor_icms)
                valor_icms = re.findall(r'(\d{1,3}[\,\.]?\d{1,3}\.?\s?\,?\d{1,3}\.?\s?\,?\d{1,2})', valor_icms)
                valor_icms  = valor_icms[0].strip()
                
                return valor_icms
        except:
                        
                return False

def main(file, pdf_file, page_number=1):
    images = convert_from_path(pdf_file, 500, poppler_path=r'C:\poppler-0.68.0\bin')
    num_pages = len(images)

    if page_number < 1 or page_number > num_pages:
        print(f"Número da página {page_number} está fora do intervalo. Processando a primeira página.")
        page_number = 1

    imagem = images[page_number - 1]  # Ajusta o índice da página

    campos_faltantes = []

    cnpj = extrator_cnpj(imagem, 'cnpj')
    if cnpj == False or len(cnpj) < 14:
        cnpj = extrator_cnpj(imagem, 'cnpj_ajustado')
        if cnpj == False:
            cnpj = extrator_cnpj(imagem, 'cnpj_ajustado2') # CRIAR OCR

    valor_total = extrator_valor_total(imagem, 'valor_total0')
    if valor_total == False:
        valor_total = extrator_valor_total(imagem, 'valor_total')
        if valor_total == False:
           valor_total = extrator_valor_total(imagem, 'valor_total_ajustado')
           if valor_total == False:
                 valor_total = extrator_valor_total(imagem, 'valor_total_ajustado2')
                 if valor_total == False:
                         valor_total = '---' 

    volume_total = extrator_volume_total(imagem, 'volume_total')
    if volume_total == False or len(volume_total) <= 4:
        volume_total = extrator_volume_total(imagem, 'volume_total_ajustado')
        if volume_total == False:
            volume_total = extrator_volume_total(imagem, 'volume_total_ajustado2')
            if volume_total == False:
                   volume_total = extrator_volume_total(imagem, 'volume_total_ajustado3')
                   if volume_total == False:
                          volume_total = '---'

    data_emissao = extrator_data_emissao(imagem, 'data_emissao')
    if data_emissao == False:
        data_emissao = extrator_data_emissao(imagem, 'data_emissao2')
        if data_emissao == False:
            data_emissao = extrator_data_emissao(imagem, 'data_emissao3')
    
    data_inicio = extrator_data_inicio(imagem, 'data_inicio')
    if data_inicio == False:
        data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado')               
        if data_inicio == False:
            data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado2')
            if data_inicio == False:
                data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado3')
                if data_inicio == False:
                    data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado4')
                    if data_inicio == False:
                        data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado5')
                        if data_inicio == False:
                              data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado6')
                         
    data_fim = extrator_data_fim(imagem, 'data_fim')
    if data_fim == False:
        data_fim = extrator_data_fim(imagem, 'data_fim_ajustado')
        if data_fim == False:
            data_fim = extrator_data_fim(imagem, 'data_fim_ajustado2')
            if data_fim == False:
                  data_fim = extrator_data_fim(imagem, 'data_fim_ajustado3')
                  if data_fim == False: 
                         data_fim = extrator_data_fim(imagem, 'data_fim_ajustado3')
                         if data_fim == False:
                                data_fim = extrator_data_fim(imagem, 'data_fim_ajustado4')
                                if data_fim == False:
                                    data_fim = extrator_data_fim(imagem, 'data_fim_ajustado5')
                                    if data_fim == False:
                                          data_fim = extrator_data_fim(imagem, 'data_fim_ajustado6')

    numero_fatura = extrator_numero_fatura(imagem, 'numero_fatura')
    if numero_fatura == False:
        numero_fatura = extrator_numero_fatura(imagem, 'numero_fatura_ajustado')
        if numero_fatura == False:
            numero_fatura = extrator_numero_fatura(imagem, 'numero_fatura_ajustado2')
            if numero_fatura == False:
                numero_fatura = extrator_numero_fatura(imagem, 'numero_fatura_ajustado3')
        
    valor_icms = extrator_valor_icms(imagem, 'valor_icms')
    if valor_icms == False:
        valor_icms = extrator_valor_icms(imagem, 'valor_icms_ajustado')
        if valor_icms == False:
            valor_icms = extrator_valor_icms(imagem, 'valor_icms_ajustado2')
            if valor_icms == False:
                valor_icms = extrator_valor_icms(imagem, 'valor_icms_ajustado3')      
        
    if not cnpj or not valor_total or not volume_total or not data_emissao or not data_inicio or not data_fim or not numero_fatura or not valor_icms:
        print('Fatura não movida devido a dados incompletos.')
    else: 
        verificar = verificar_download(cnpj, data_inicio, data_fim, valor_total, caminho_excel)
        if verificar:
            data_frame = dados_excel(cnpj, valor_total, volume_total, data_emissao, data_inicio, data_fim, numero_fatura, valor_icms, correcao_pcs, dist, arquivo)
            if adicionar_dados_excel(caminho_excel, data_frame):
                mover_arquivo(pdf_file, diretorio_destino)
            else: 
                print('Erro ao adicionar dados na planilha.')
                adicionar_dados_excel(caminho_excel, data_frame)

        else:
                print('Dados já inseridos!')

# Exemplo de uso
corte = corte_comgas()
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Comgás\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Comgás\Lidos'

for arquivo in os.listdir(file_path):
        if arquivo.endswith('.pdf') or arquivo.endswith('.PDF'):
                arquivo_full = rf'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Comgás\Faturas\{arquivo}' 
        main(arquivo, arquivo_full)

