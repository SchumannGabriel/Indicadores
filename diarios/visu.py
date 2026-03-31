import smartsheet
import pandas as pd

# ================================
# CONFIGURAÇÕES SMARTSHEET
# ================================
TOKEN = '32M5yHYGUMBRyOTkf3GstuBbJ36Q4T9TvefrX'
ID_PLANILHA = '3432321207193476'

def ver_colunas_real():
    try:
        smart = smartsheet.Smartsheet(TOKEN)
        sheet = smart.Sheets.get_sheet(ID_PLANILHA)
        
        print("\n" + "="*50)
        print("LISTA DE COLUNAS DETECTADAS (RAIO-X)")
        print("="*50)
        
        for col in sheet.columns:
            # repr() mostra caracteres ocultos como \n (quebra de linha)
            # Os colchetes [] ajudam a ver espaços vazios
            print(f"Nome: [{col.title}]  |  Tipo: {col.type}")
            
        print("="*50)
        print("\nCopie o nome exatamente como aparece dentro dos colchetes []")
        
    except Exception as e:
        print(f"Erro ao acessar: {e}")

if __name__ == "__main__":
    ver_colunas_real()