# COORDENADAS PARA EXTRAÇÃO DE DADOS USANDO OCR
                    # A PRIMEIRA COORDENADA DE CADA VALOR, FOI FEITA ATRAVÉZ DA FATURA PADRÃO
def corte_comgas():
    corte = {
                                                 
    'cnpj': (2690, 350, 3310, 520),#                                                                
    'cnpj_ajustado': (2700, 320, 3370, 420),
    

    'valor_total': (1900, 2000, 2500, 2300),#

    'volume_total': (2385, 3585, 2565, 3700),          #  VOLUME TOTAL DANIFICADO
    
    'data_emissao': (800, 875, 1800, 1500),#

    'data_inicio': (1900, 850, 2500, 950),                

    'data_fim': (1900, 920, 2400, 1010),               

    'numero_fatura': (1100, 400, 1800, 800),

    'valor_icms': (3240, 2400, 4200, 2500)                        # PADRÃO

    }
    return corte

# CAMINHO DA PLANILHA

caminho_excel = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\COMGAS.xlsx'

'''while True
    for Player(get_moeda):
        if player(moeda >= 3):
            brilhar(dourado)
        else:
            brilhar(branco)'''


