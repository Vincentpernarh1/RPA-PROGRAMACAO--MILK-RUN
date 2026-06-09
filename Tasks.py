import json
import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError, expect
import warnings
import pyxlsb
import csv
import math
import xlwings as xw
import time
from datetime import date, timedelta
import traceback
import shutil

# from App import load_config


caminho_base = os.getcwd()



def load_config():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    config_path = os.path.join(base_path, "config.json")

    if not os.path.exists(config_path):
        raise FileNotFoundError(f"config.json not found in: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)
   

CONFIG = load_config()

caminho_pasta_matriz = os.path.join(caminho_base, CONFIG['paths']['folders']['base_matriz'])
base_planilhas_recebidos = os.path.join(caminho_base, CONFIG['paths']['folders']['base_planilhas_recebidos'])
# --- ADJUST HELPER ---
def _check_name(filename, config_key):
    """
    Checks if a filename matches any search term in CONFIG['paths']['dynamic_files'].
    """
    # Handle single string or list of terms
    terms_from_config = CONFIG['paths']['dynamic_files'].get(config_key, [])
    if isinstance(terms_from_config, str):
        terms = [terms_from_config]
    else:
        terms = terms_from_config
    
    for term in terms:
        if term in filename:
            return True
    return False

def get_path_from_config(key, folder_key, q=None):
    """
    Gets a file path from config.
    key: The key from 'dynamic_files' (e.g., 'pfep_search_terms').
    folder_key: The key from 'folders' (e.g., 'base_matriz').
    """
    folder_path = os.path.join(caminho_base, CONFIG['paths']['folders'][folder_key])

    if not os.path.isdir(folder_path):
        if q: q.put(("status", f"AVISO: Pasta não encontrada: {folder_path}"))
        return None

    for nome in os.listdir(folder_path):
        if _check_name(nome, key) and not nome.startswith("~$") and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            return os.path.join(folder_path, nome)
    
    if q: q.put(("status", f"AVISO: Nenhum arquivo encontrado para '{key}' em '{folder_path}'"))
    return None
# --- END ADJUSTED HELPER ---




def download_Demanda(page, url_order, q, username, password):
   
    # Processar_Demandas(q)
    try:
        # --- 1. Login and Initial Navigation ---
        q.put(("status", "Navegando para a página de login..."))
        page.goto(url_order, timeout=60000)
        q.put(("progress", 10))

        q.put(("status", "Realizando login..."))
        page.get_by_role("textbox", name="User").fill(username)
        page.get_by_role("textbox", name="Password").fill(password)
        page.get_by_role("button", name="Log In").click()
        q.put(("status", "Login realizado com sucesso!"))
        q.put(("progress", 20))

     
        # --- 2. Navigate to the correct report section ---
        q.put(("status", "Navegando para a seção de relatórios..."))
        page.locator("#ID_button").click()
        page.get_by_text("ELOG - Importar A8 Automatica").nth(1).click()
        q.put(("progress", 30))

        # --- 3. Date-based Search Loop ---
        current_date = date.today()
        records_found = False
        
        current_date = date.today() 
        # current_date = date.today() - timedelta(days=2)  # this is to be use for other dates simulations
        records_found = False
        
        while not records_found:
            date_str = current_date.strftime('%m/%d/%Y')
            q.put(("status", f"Procurando registros em: {date_str}"))
            
            # Fill date fields and click search
            page.get_by_role("textbox", name="Data Inicial").fill(date_str)
            page.get_by_role("textbox", name="Data Final").fill(date_str)
            page.get_by_role("button", name="Pesquisar").click()

            try:
                # If it appears, the code continues. If not, it raises a TimeoutError.
                page.get_by_text("Total records:").click(timeout = 3000)
                
                # This code only runs if the expect() call above succeeds
                q.put(("status", f"Registros encontrados em {date_str}!"))
                records_found = True
                q.put(("progress", 45))

            except TimeoutError:
                # This code runs only if the locator was not visible after 60 seconds
                q.put(("status", f"Nenhum registro encontrado em {date_str}. Tentando dia anterior."))
                current_date -= timedelta(days=1)
                time.sleep(2) # Small delay before trying again

        # --- 4. Row-wise TXT File Download ---
        q.put(("status", "Iniciando download dos arquivos individuais..."))
        
        # Define the base directory for downloads for the found date
        download_path_base = os.path.join(caminho_base, CONFIG['paths']['folders']['base_demanda'])
        os.makedirs(download_path_base, exist_ok=True)
        
        # This selector targets rows within the table's body to avoid header rows.
        data_rows = page.get_by_role("cell", name="Download")

        row_count = data_rows.count()
        q.put(("status", f"Encontrados {row_count-1} arquivos para download."))

        for i in range(1,row_count):
            
            row = data_rows.nth(i)
           
            with page.expect_download() as download_info:
               download_link =row.click()
            
            download = download_info.value
            
            # Construct the full path and save the file
            file_path = os.path.join(download_path_base, f"{i}_{download.suggested_filename}")
            
            if os.path.exists(file_path):
                os.remove(file_path)
            download.save_as(file_path)
            q.put(("status", f"Baixado: {download.suggested_filename}"))

        q.put(("status", "Todos os downloads dos arquivos individuais estão completos."))

        q.put(("progress", 10))

        q.put(("status", "Deseja continuar com a transformação das bases?"))

        Processar_Demandas(q)
        

    except Exception as e:
        q.put(("status", f"❌ Ocorreu um erro: {e}"))
        # You might want to add more specific error handling here


def Processar_Demandas(q):
   
    caminho_pasta = os.path.join(caminho_base, CONFIG['paths']['folders']['base_demanda'])
    caminho_df_fornecedor = os.path.join(caminho_base, CONFIG['paths']['folders']['base_bases'], CONFIG['paths']['files']['db_fornecedores'])
    
    df_DB_fornecedor = pd.read_excel(caminho_df_fornecedor)
    df_DB_fornecedor = df_DB_fornecedor[["CODIMS", "CODSAP", "UF", "FANTAS"]]

    # Verifica se a pasta de demandas existe
    if not os.path.isdir(caminho_pasta):
        print(f"Aviso: A pasta '{caminho_pasta}' não foi encontrada.")
        return pd.DataFrame()

    # Lista para armazenar os DataFrames de cada arquivo processado
    lista_dfs = []
    df_temp = pd.DataFrame()

    # Percorre todos os arquivos na pasta de demandas
    q.put(("progress", 12))
    for nome_arquivo in os.listdir(caminho_pasta):
        caminho_completo_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        nome_arquivo_lower = nome_arquivo.lower()
       
        try:
            # --- MANTÉM A LÓGICA ORIGINAL PARA ARQUIVOS .TXT E .CSV ---
            if nome_arquivo_lower.endswith((".txt", ".csv")):
                dados_arquivo_atual = []
                with open(caminho_completo_arquivo, "r", encoding="utf-8", errors="ignore") as arquivo:
                    linhas_a_processar = arquivo.readlines()

                # Processa cada linha extraída do arquivo de texto
                for linha in linhas_a_processar:
                    if "AUTOMATIC" in linha:
                        continue

                    linha = linha.strip()

                    # A lógica de fatiamento requer um comprimento mínimo
                    if len(linha) >= 20:
                        try:
                            # Extrai os dados com base na posição dos caracteres
                            PN = linha[3:14]
                            SAP = linha[-20:-11]
                            quantidade = linha[-11:].replace("+", "")

                            # Adiciona os dados extraídos à lista deste arquivo
                            dados_arquivo_atual.append({
                                "PN": int(PN.strip()),
                                "SAP": int(SAP.strip()),
                                "QUANT": int(quantidade.strip()),
                            })
                        except (ValueError, IndexError):
                            # Ignora linhas que não seguem o formato esperado
                            continue

                # Se dados foram extraídos do arquivo, cria um DataFrame
                if dados_arquivo_atual:
                    df_temp = pd.DataFrame(dados_arquivo_atual)
                    
                    # --- NOVO: Adiciona o identificador como False para TXT/CSV ---
                    df_temp['__highlight_sap'] = False 
                    
                    lista_dfs.append(df_temp)

            # --- LÓGICA MODIFICADA PARA PROCESSAR ARQUIVOS EXCEL (.XLS, .XLSX) ---
            elif nome_arquivo_lower.endswith((".xls", ".xlsx")):

                # Mapeamento dos nomes de coluna do arquivo Excel para os nomes desejados
                colunas_mapeamento = {
                    'DESENHO': 'PN',
                    'COD ORIGEM': 'SAP',
                    'ENTREGA SOLICITADA': 'QUANT'
                }

                # Lê o arquivo Excel
                df_excel = pd.read_excel(caminho_completo_arquivo)

                # Pega a lista de colunas que precisamos do arquivo original
                colunas_originais_necessarias = list(colunas_mapeamento.keys())

                # Verifica se todas as colunas necessárias existem no arquivo
                if not all(coluna in df_excel.columns for coluna in colunas_originais_necessarias):
                    print(f"Aviso: O arquivo '{nome_arquivo}' não contém todas as colunas necessárias e será ignorado.")
                    continue

                # 1. Seleciona apenas as colunas que nos interessam
                df_temp = df_excel[colunas_originais_necessarias].copy()

                # 2. Renomeia as colunas para o padrão final
                df_temp.rename(columns=colunas_mapeamento, inplace=True)
                df_temp = df_temp[df_temp["QUANT"] > 0]

                # --- NOVO: Adiciona o identificador para linhas do Excel ---
                df_temp['__highlight_sap'] = True
                # --- FIM DA ADIÇÃO ---

                # 3. Adiciona o DataFrame processado à lista para concatenação posterior
                lista_dfs.append(df_temp)

        except Exception as e:
            print(f"Erro ao processar o arquivo '{nome_arquivo}': {e}")
            continue

    if not lista_dfs:
        print("Nenhum dado válido foi processado.")
        return pd.DataFrame()

    # Concatena todos os DataFrames da lista em um único DataFrame final
    df_final = pd.concat(lista_dfs, ignore_index=True)
    
    df_final['__highlight_sap'] = df_final['__highlight_sap'].fillna(False)

    colunas_numericas = ["PN", "SAP", "QUANT"]
    for col in colunas_numericas:
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    # Remove linhas onde a conversão numérica falhou (resultando em NaT/NaN)
    df_final.dropna(subset=colunas_numericas, inplace=True)

    df_unique_fornecedor = df_DB_fornecedor.drop_duplicates(subset=["CODSAP"], keep="first")

    # Convert the unique CODSAP column to integer and set as index
    codsap_map = df_unique_fornecedor.set_index("CODSAP")["FANTAS"]
    codsap_map_estado = df_unique_fornecedor.set_index("CODSAP")["UF"]

    # Map to df_final
    df_final["FORNECEDOR"] = df_final["SAP"].astype('Int64', errors='ignore').map(codsap_map)
    df_final["ESTADO"] = df_final["SAP"].astype('Int64', errors='ignore').map(codsap_map_estado)

    # Converte colunas para inteiro após remover os nulos
    for col in colunas_numericas:
        df_final[col] = df_final[col].astype(int)

    condicao_estado = df_final['ESTADO'] != 'MG'
    condicao_sap_excecao = df_final['SAP'].isin(CONFIG['business_logic']['sap_exclusion_list'])

    # Mantém se a linha NÃO for 'MG', OU se o SAP estiver na lista de exceção
    df_final = df_final[condicao_estado | condicao_sap_excecao].copy()      
    light_yellow = CONFIG['business_logic']['style_highlight_color'] 
    df_funilaria = pd.DataFrame(columns=['SAP', 'FORNECEDOR']) # Inicializa vazio

    q.put(("progress", 13))
    if '__highlight_sap' in df_final.columns:
        funilaria_mask = (df_final['__highlight_sap'] == True)
        df_funilaria = df_final[funilaria_mask][['SAP', 'FORNECEDOR']].copy()
        df_funilaria.drop_duplicates(inplace=True)

        DF_Horarios = le_arquivo_horario() # Chama sua função
        
        if not DF_Horarios.empty:
            # 2. Cria o mapa: [Supplier Code] -> [Horário de Janela]
            #    Usa drop_duplicates para garantir que cada Supplier Code seja único
            mapa_horario_pronto = DF_Horarios.drop_duplicates(
                subset=['Supplier Code']
            ).set_index('Supplier Code')['Horário de Janela']

            # 3. Aplica o mapa ao df_funilaria. É muito mais rápido que 'apply'
            df_funilaria['HORÁRIO'] = df_funilaria['SAP'].map(mapa_horario_pronto)
        else:
            print("Aviso: DF_Horarios está vazio. A coluna 'HORÁRIO' não será preenchida.")
            df_funilaria['HORÁRIO'] = pd.NaT # Ou None
        
        # Lógica original de máscara para estilização
        highlight_mask = funilaria_mask.tolist()
        
        # 3. Remove a coluna auxiliar. Ela não é mais necessária.
        df_final = df_final.drop(columns=['__highlight_sap'])
    else:
        # Se a coluna nunca foi criada (só arquivos TXT), cria uma máscara vazia (tudo False)
        highlight_mask = [False] * len(df_final)
        print("Nenhum dado de Excel processado, a aba 'funilaria' estará vazia.")
    
    sap_styles = [
        f'background-color: {light_yellow}' if mask else ''
        for mask in highlight_mask
    ]

   
    styler = df_final.style

    try:
        styler = styler.apply(lambda _: sap_styles, subset=['SAP'])
    except Exception as e:
        print(f"Não foi possível aplicar o estilo à coluna 'SAP': {e}")
        
    output_path = os.path.join(caminho_base, CONFIG['paths']['folders']['base_resultados'], CONFIG['paths']['files']['demandas_total_output'])
    demand_path = output_path
    q.put(("progress", 20))

    
    try:
        # Cria um ExcelWriter para salvar ambos os DataFrames
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Salva a aba principal estilizada
            styler.to_excel(writer, sheet_name='Demandas_Total', index=False)
            
            # Salva a nova aba 'funilaria' (sem estilo)
            df_funilaria.to_excel(writer, sheet_name='funilaria', index=False)
            
        print(f"Arquivo salvo com sucesso com abas 'Demandas_Total' e 'funilaria' em: {output_path}")
        Atualiza_PFEP(demand_path,q)
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel com múltiplas abas: {e}")

        try:
            df_final.to_excel(output_path, index=False)
            print(f"Arquivo salvo (APENAS ABA PRINCIPAL, sem estilo) em: {output_path}")
            Atualiza_PFEP(demand_path,q)
        except Exception as e_fallback:
            print(f"Erro fatal ao salvar o arquivo: {e_fallback}")
    q.put(("progress", 22))
    return

def le_arquivo_horario() :
   
    # Define as colunas de interesse
    colunas_horarios = ['Supplier Code', 'Horário de Janela']
    
    # Inicializa o DataFrame final como vazio. 
    DF_Horarios = pd.DataFrame(columns=colunas_horarios)
    
    # 1. Define o caminho para a pasta
    caminho_matriz_folder = os.path.join(caminho_base, CONFIG['paths']['folders']['base_matriz'])
    
    # 2. Verifica se a pasta existe
    if not os.path.isdir(caminho_matriz_folder):
        print(f"Aviso: A pasta de Matriz não foi encontrada em: {caminho_matriz_folder}")
    else:
        # 3. Encontra o nome do arquivo dinamicamente
        nome_arquivo_horarios_completo = None
        termo_busca = CONFIG['paths']['dynamic_files']['horarios_search_term']

        try:
            for nome_arquivo in os.listdir(caminho_matriz_folder):
                nome_lower = nome_arquivo.lower()
                # Verifica se o termo está no nome E se é um arquivo Excel
                if termo_busca in nome_lower and (nome_lower.endswith('.xlsx') or nome_lower.endswith('.xls')):
                    nome_arquivo_horarios_completo = os.path.join(caminho_matriz_folder, nome_arquivo)
                    print(f"Arquivo de horários encontrado: {nome_arquivo}")
                    break # Para no primeiro arquivo que encontrar
        except Exception as e:
            print(f"Erro ao listar arquivos na pasta Matriz: {e}")

        # 4. Se o arquivo foi encontrado, tenta ler os dados
        if nome_arquivo_horarios_completo:
            
            # --- CORREÇÃO: Define o nome literal da aba ---
            sheet_name_literal = CONFIG['paths']['sheet_names']['horarios_sheet']
            
            try:
                
                DF_Horarios = pd.read_excel(
                    nome_arquivo_horarios_completo, 
                    sheet_name=sheet_name_literal, # Passa o nome literal da aba
                    usecols=colunas_horarios
                )
                
                print(f"Horários carregados da aba: '{sheet_name_literal}'")
                
                # Remove linhas que possam ter vindo com NaNs (ex: linhas em branco)
                DF_Horarios.dropna(subset=colunas_horarios, inplace=True)
                # print(f"Total de {len(DF_Horarios)} registros de horários carregados.")

            except FileNotFoundError:
                print(f"Erro: O arquivo '{nome_arquivo_horarios_completo}' não foi encontrado.")
            except ValueError as e:
                # Erro comum se as colunas ou a aba não existirem
                print(f"Erro ao ler colunas/aba: {e}")
                print(f"Verifique se a aba '{sheet_name_literal}' E as colunas {colunas_horarios} existem.")
            except Exception as e:
                print(f"Erro inesperado ao ler o arquivo Excel de horários: {e}")
                # (Pode acontecer se a aba não for encontrada)
                if sheet_name_literal not in str(e): # Evita msg duplicada
                     print(f"Verifique se a aba '{sheet_name_literal}' existe no arquivo.")
        
        else:
            if os.path.isdir(caminho_matriz_folder):
                 print(f"Aviso: Nenhum arquivo contendo '{termo_busca}' foi encontrado em {caminho_matriz_folder}")

    return DF_Horarios


def Atualiza_PFEP(path_demandas,q):
    q.put(("status", "Iniciando atualização do PFEP..."))
    q.put(("progress", 30))
    caminho_pasta_pfep = os.path.join(caminho_base, CONFIG['paths']['folders']['base_matriz'])
    nome_pfep = None
    
    for nome in os.listdir(caminho_pasta_pfep):
        if _check_name(nome, 'pfep_search_terms') and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            nome_pfep = os.path.join(caminho_pasta_pfep, nome)
            break
    
    if not nome_pfep:
        print("⚠️ Arquivo PFEP não encontrado.")
        return

    demand_folder = os.path.dirname(path_demandas)
    demand_file = os.path.basename(path_demandas)

    app = None
    wb = None
    wb_demandas = None

    q.put(("status", "Abrindo PFEP e atualizando dados..."))
    try:
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        
        app.api.AskToUpdateLinks = False  # 🔒 Prevents Excel popup

        q.put(("status", "Abrindo arquivo de demandas..."))
        wb_demandas = app.books.open(path_demandas)

        # Open PFEP safely, no popup or freeze
        wb = app.books.open(nome_pfep, update_links=False)
        wb.app.api.EnableEvents = True
        ws = wb.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']]


        # Processar_programacao(wb,q)

        q.put(("status", "Atualizando dados do PFEP..."))

        q.put(("status", "Atualizando fórmulas do PFEP..."))
        formula = (
            f'=IF(OFFSET($EL$1,ROW()-1,0)="Fora do Escopo",0,'
            f'IF(OFFSET($CG$1,ROW()-1,0)="YES",0,'
            f'SUMIFS(\'{demand_folder}\\[{demand_file}]Demandas_Total\'!$C:$C,'
            f'\'{demand_folder}\\[{demand_file}]Demandas_Total\'!$B:$B,'
            f'LEFT(OFFSET($V$1,ROW()-1,0),9),'
            f'\'{demand_folder}\\[{demand_file}]Demandas_Total\'!$A:$A,'
            f'OFFSET($C$1,ROW()-1,0))))'
        )
        
        
        ws.range('P5').formula = formula
        
        q.put(("status", "Executando macro de recálculo do PFEP..."))
        wb.app.macro("Calcula_Todo_PFEP")()
        q.put(("status", "Macro de recálculo finalizada."))
                        
        # print("Salvando PFEP...")
        q.put(("status", "Salvando PFEP atualizado..."))
        wb.save()
        # print("💾 PFEP atualizado com sucesso e recálculo executado.")
        q.put(("status", "PFEP atualizado com sucesso!"))


        # if wb_demandas:
        #     wb_demandas.close()

        Processar_programacao(wb_demandas,wb,q)  # NEED TO FREE THIS FOR CODE TO CONTINUE
        
        
        q.put(("progress", 50))

    except Exception as e:
        print(f"❌ Erro inesperado durante a atualização do PFEP: {e}")

    finally:
        # q.put(("progress", 35))
        pass
        

def Processar_programacao(wb_demandas,pfep, q):

    q.put(("status", "Iniciando atualização da Programação FIASA..."))
    # Ensure 'caminho_base' is accessible
    global caminho_base 
    
    # caminho_pasta_matriz = os.path.join(caminho_base, '1 - MATRIZ')
    nome_prog_fiasa = None
    cargolift_sp_Supplier = None
    cargolift_sp_PFEP = None
    
    # --- Locate Programação FIASA file ---
    for nome in os.listdir(base_planilhas_recebidos):
        if _check_name(nome, 'fiasa_search_terms'):
            nome_prog_fiasa = os.path.join(base_planilhas_recebidos, nome)

    
    # --- Locate Programação FIASA file ---
    for nome in os.listdir(caminho_pasta_matriz):
        # if _check_name(nome, 'fiasa_search_terms'):
        #     nome_prog_fiasa = os.path.join(caminho_pasta_matriz, nome)

        if _check_name(nome, 'cargolift_pfep_terms'):
            cargolift_sp_PFEP = os.path.join(caminho_pasta_matriz, nome)

        if _check_name(nome, 'cargolift_supplier_terms'):
            cargolift_sp_Supplier = os.path.join(caminho_pasta_matriz, nome)
            
    
    if not nome_prog_fiasa:
        print("⚠️ Arquivo 'Programação FIASA - OFICIAL' não encontrado.")
        return

    # --- Sheets in PFEP workbook (already open) ---
    ws_pfep = pfep.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']]
    ws_supplier = pfep.sheets[CONFIG['paths']['sheet_names']['supplier_db_sheet']]
    q.put(("status", "Lendo dados do PFEP e Supplier DB..."))
    # --- Find last rows ---
    last_row_pfep = ws_pfep.range('C' + str(ws_pfep.cells.last_cell.row)).end('up').row
    last_row_supplier = ws_supplier.range('B' + str(ws_supplier.cells.last_cell.row)).end('up').row
    
    print(f"📊 PFEP last row: {last_row_pfep}")
    print(f"📊 Supplier DB last row: {last_row_supplier}")

    # --- Copy PFEP filtered data ---
    data_pfep = None
    q.put(("status", "Filtrando e copiando dados do PFEP..."))
    if ws_pfep.api.AutoFilterMode:
        ws_pfep.api.AutoFilterMode = False

        
    # print("Applying filter to PFEP (Column P <> 0,00)...") 
    q.put(("status", "Aplicando filtro no PFEP (Coluna P <> 0,00)..."))
    
    try:
        # 1️⃣ Define the xlwings range (includes header)
        filter_range = ws_pfep.range(f"A6:AZ{last_row_pfep}")

        # 2️⃣ Apply the filter
        filter_range.api.AutoFilter(Field:=16, Criteria1:="<>0,00")

        # 3️⃣ Get visible rows after filtering
        visible_cells = filter_range.api.SpecialCells(12)  # xlCellTypeVisible = 12

        # 4️⃣ Combine values from visible "areas", skipping header
        data_pfep = []
        for i, area in enumerate(visible_cells.Areas):
            area_range = ws_pfep.range(area.Address)
            values = area_range.value

            if isinstance(values, list) and isinstance(values[0], list):
                # Skip first row of first area (header)
                if i == 0:
                    data_pfep.extend(values[1:])
                else:
                    data_pfep.extend(values)
            else:
                # Single row case (no sublists)
                if i > 0:  # skip header only on first block
                    data_pfep.append(values)

        print(f"✅ Copied {len(data_pfep)} visible rows (header excluded) from PFEP.")
        
            
    except Exception as e:
        print(f"⚠️ No visible data found in PFEP after filtering.")
        q.put(("status", "⚠️ Nenhum dado visível encontrado no PFEP após o filtro."))


    q.put(("status", "Lendo dados do Supplier DB...")) 
    if ws_pfep.api.AutoFilterMode:
        ws_pfep.api.AutoFilterMode = False
    
    # --- Copy Supplier DB data ---
    if ws_supplier.api.AutoFilterMode:
        print("Clearing filters from Supplier DB sheet...")
        ws_supplier.api.AutoFilterMode = False
    q.put(("status", "Copiando dados do Supplier DB..."))   
    data_supplier = ws_supplier.range(f"A6:AT{last_row_supplier}").value
    if data_supplier:
        print("✅ Copied Supplier DB data.")

    # --- Open Programação FIASA workbook ---
    print("Opening Programação FIASA...")
    q.put(("status", "Abrindo Programação FIASA..."))
    app_prog_fiasa = xw.App(visible=True, add_book=False) 
    app_prog_fiasa.display_alerts = False
    app_prog_fiasa.api.AskToUpdateLinks = False
    wb_fiasa = app_prog_fiasa.books.open(nome_prog_fiasa,update_links=False,read_only=False)
    q.put(("status", "Programação FIASA aberta."))
    q.put(("progress", 60))
    try:
    
        ws_cola_pfep = wb_fiasa.sheets[CONFIG['paths']['sheet_names']['fiasa_cola_pfep']]
        ws_cola_supplier = wb_fiasa.sheets[CONFIG['paths']['sheet_names']['fiasa_cola_supplier']]
        
        ws_cola_pfep.range('A2:HB15000').clear_contents()
        ws_cola_supplier.range('A2:AT5000').clear_contents()

        if data_pfep:
            print("Pasting PFEP data...")
            # 1️⃣ Paste values
            dest_range = ws_cola_pfep.range('A2').resize(len(data_pfep), len(data_pfep[0]))
            dest_range.value = data_pfep

            # 2️⃣ Copy formatting from row 5
            template_row = ws_cola_pfep.range('5:5')  # Entire row 5 as template
            format_range = ws_cola_pfep.range('A2').resize(len(data_pfep), dest_range.columns.count)
            template_row.api.Copy()
            format_range.api.PasteSpecial(Paste=-4122)  # xlPasteFormats
            ws_cola_pfep.api.Application.CutCopyMode = False

        else:
            print("⚠️ No PFEP data to paste.")
        
        
        print("Removing the filters from Programação FIASA...")
        if ws_cola_pfep.api.AutoFilterMode:
            ws_cola_pfep.api.AutoFilterMode = False
            
       
        q.put(("status", "Colando dados do Supplier DB na Programação FIASA..."))
        if data_supplier:
            print("Pasting Supplier DB data...")
            
            dest_range = ws_cola_supplier.range('A2').resize(len(data_supplier), len(data_supplier[0]))
            dest_range.value = data_supplier

            template_row = ws_cola_supplier.range('5:5')  # Entire row 5 as template
            format_range = ws_cola_supplier.range('A2').resize(len(data_supplier), dest_range.columns.count)
            template_row.api.Copy()
            format_range.api.PasteSpecial(Paste=-4122)  # xlPasteFormats
            ws_cola_supplier.api.Application.CutCopyMode = False

        else:
            print("⚠️ No Supplier data to paste.")
        q.put(("status", "Executando recálculo da Programação FIASA..."))
        print("Recalculating, saving, and closing Programação FIASA...")
        wb_fiasa.app.api.CalculateFullRebuild()       

        q.put(("status", "Fechando PFEP de Arquivo de demandas..."))

        suppliers_carrier = CONFIG['business_logic']['carrier_mappings']['suppliers_carrier']
        suppliers_fiasa = CONFIG['business_logic']['carrier_mappings']['suppliers_fiasa']
        q.put(("status", "Atualizando Carrier na Programação FIASA..."))

        # Worksheet
        ws_Sup_db_corrier = wb_fiasa.sheets[CONFIG['paths']['sheet_names']['supplier_db_sheet']]
        if ws_Sup_db_corrier.api.AutoFilterMode:
            ws_Sup_db_corrier.api.AutoFilterMode = False

        # Find last used row in column C
        last_row = ws_Sup_db_corrier.range('C' + str(ws_Sup_db_corrier.cells.last_cell.row)).end('up').row
        
        
        q.put(("status", f"Processando {last_row - 1} linhas para atualização de Carrier..."))
        # Get data
        supplier_codes = ws_Sup_db_corrier.range(f'C2:C{last_row}').value
        fca_values = ws_Sup_db_corrier.range(f'D2:D{last_row}').value
        q.put(("progress", 65))
        
        # Loop through rows
        for i, code in enumerate(supplier_codes, start=2):
            if code is None:
                continue

            try:
                code_str = str(code)
               
            except:
                continue


            # 1️⃣ Check in main carrier dict
            if code_str in suppliers_carrier:
                carrier_value = suppliers_carrier[code_str]

                # 2️⃣ If in FIASA dict, only update if column D == "FCA"
                if code_str in suppliers_fiasa:
                    if "FCA" in str(fca_values[i - 2]).strip().upper() :
                        ws_Sup_db_corrier.range(f'AB{i}').value = carrier_value
                else:
                    ws_Sup_db_corrier.range(f'AB{i}').value = carrier_value
        q.put(("status", "Carrier atualizado na Programação FIASA."))
        
        
        q.put(("status", "Salvando Programação FIASA..."))
        wb_fiasa.save()
        
        progrma_cargolift (cargolift_sp_PFEP,cargolift_sp_Supplier,wb_fiasa,q,pfep,wb_demandas)

    finally:
        q.put(("status", "Finalizando atualização da Programação FIASA..."))
        wb_fiasa.close()
        app_prog_fiasa.quit()
        q.put(("status", "Programação FIASA atualizada com sucesso!"))


def progrma_cargolift(arquivo_cargolift_sp_PFEP, arquivo_cargolift_sp_Supplier, wb_fiasa, q,pfep,wb_demandas):
   
    
    q.put(("status", "Iniciando atualização da Programação Cargolift SP..."))
    q.put(("progress", 70))

    ws_pfep = wb_fiasa.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']]
    ws_supplier_db = wb_fiasa.sheets[CONFIG['paths']['sheet_names']['supplier_db_sheet']]

    # Remove any active filters before we start
    for ws in [ws_pfep, ws_supplier_db]:
        try:
            if ws.api.FilterMode:
                ws.api.ShowAllData()
        except:
            pass
        ws.api.AutoFilterMode = False

    # ========== PFEP ==========
    q.put(("status", "Filtrando PFEP por CARGOLIFT..."))
    last_row_pfep = ws_pfep.range('A' + str(ws_pfep.cells.last_cell.row)).end('up').row

    # PFEP columns go to AQ (43)
    filter_range_pfep = ws_pfep.range(f"A1:AQ{last_row_pfep}")
   
#    Need to verify one logic before anything more

    filter_range_pfep.api.AutoFilter(Field:=43, Criteria1:=["CARGOLIFT", "CARGOLIFT SUL","LINE HAUL"], Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
    
    # filter_range_pfep.api.AutoFilter(Field=43,Criteria1="<>OUT",Operator=xw.constants.AutoFilterOperator.xlAnd,
    #             Criteria2="<>#N/A"
    #         )


    # Collect visible data
    try:
        visible_cells_pfep = filter_range_pfep.api.SpecialCells(12)  # xlCellTypeVisible
        data_pfep = []
        for i, area in enumerate(visible_cells_pfep.Areas):
            area_range = ws_pfep.range(area.Address)
            values = area_range.value
            if isinstance(values, list) and isinstance(values[0], list):
                if i == 0:
                    data_pfep.extend(values[1:])  # skip header
                else:
                    data_pfep.extend(values)
            else:
                if i > 0:
                    data_pfep.append(values)
        print(f"✅ Copied {len(data_pfep)} visible rows from PFEP.")
    except Exception:
        data_pfep = []
        print("⚠️ Nenhum dado visível encontrado em PFEP com filtro 'CARGOLIFT'.")

    # Clear filter
    try:
        if ws_pfep.api.FilterMode:
            ws_pfep.api.ShowAllData()
    except:
        pass
    ws_pfep.api.AutoFilterMode = False

    # ========== SUPPLIER DB ==========
    q.put(("status", "Filtrando Supplier DB por CARGOLIFT..."))
    last_row_supplier = ws_supplier_db.range('A' + str(ws_supplier_db.cells.last_cell.row)).end('up').row

    # Supplier DB columns go to AJ (36)
    filter_range_supplier = ws_supplier_db.range(f"A1:AJ{last_row_supplier}")
    filter_range_supplier.api.AutoFilter(Field:=28, Criteria1:="CARGOLIFT")  # AB = 28
    q.put(("progress", 75))
    try:
        visible_cells_supplier = filter_range_supplier.api.SpecialCells(12)
        data_supplier = []
        for i, area in enumerate(visible_cells_supplier.Areas):
            area_range = ws_supplier_db.range(area.Address)
            values = area_range.value
            if isinstance(values, list) and isinstance(values[0], list):
                if i == 0:
                    data_supplier.extend(values[1:])
                else:
                    data_supplier.extend(values)
            else:
                if i > 0:
                    data_supplier.append(values)
        print(f"✅ Copied {len(data_supplier)} visible rows from Supplier DB.")
    except Exception:
        data_supplier = []
        print("⚠️ Nenhum dado visível encontrado em Supplier DB com filtro 'CARGOLIFT'.")

    try:
        if ws_supplier_db.api.FilterMode:
            ws_supplier_db.api.ShowAllData()
    except:
        pass
    ws_supplier_db.api.AutoFilterMode = False

    # ========== DESTINATION FILES ==========
    q.put(("status", "Abrindo arquivos de destino (Cargolift SP)..."))
    # ✅ Single Excel instance for both destination files
    app_cargolift = xw.App(visible=True, add_book=False)
    app_cargolift.display_alerts = False
    app_cargolift.api.AskToUpdateLinks = False

    wb_cargolift_sp_PFEP = app_cargolift.books.open(
        arquivo_cargolift_sp_PFEP, update_links=False, read_only=False
    )
    wb_cargolift_sp_Supplier = app_cargolift.books.open(
        arquivo_cargolift_sp_Supplier, update_links=False, read_only=False
    )

    
    ws_dest_pfep = wb_cargolift_sp_PFEP.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']]
    ws_dest_supplier = wb_cargolift_sp_Supplier.sheets[CONFIG['paths']['sheet_names']['cargolift_sp_supplier_sheet']]

    # Clear destination before pasting
    ws_dest_pfep.range('A3').expand().clear_contents()
    ws_dest_supplier.range('A3').expand().clear_contents()

    q.put(("progress", 80))
    # Paste PFEP
    if data_pfep:
        print("📋 Colando dados PFEP...")
        ws_dest_pfep.range('A3').value = data_pfep

    # Paste Supplier DB
    if data_supplier:
        print("📋 Colando dados Supplier DB...")
        ws_dest_supplier.range('A3').value = data_supplier
   
    
    # Save and close
    wb_cargolift_sp_PFEP.save()
    wb_cargolift_sp_Supplier.save()
    Corregir_peso_e_valor(wb=wb_cargolift_sp_Supplier,demandas_path =wb_demandas ,q = q,pfep_source = pfep)
    wb_cargolift_sp_PFEP.close()
    wb_cargolift_sp_Supplier.close()
    app_cargolift.quit()

    q.put(("status", "✅ Atualização da Programação Cargolift SP concluída!"))
    print("✅ Atualização concluída com sucesso!")


# --- 1. FUNÇÃO DE NORMALIZAÇÃO CORRIGIDA ---
def normalize_value(value):
    """
    Converte valor para string, remove espaços, coloca em maiúsculo
    e remove '.0' do final (crucial para PNs lidos como float).
    """
    if value is None:
        return ""
    # Converte para string, remove espaços
    s = str(value).strip().upper()
    
    # Remove ".0" de PNs que o Excel leu como float
    if s.endswith('.0'):
        return s[:-2] # Remove os dois últimos caracteres ('.0')
        
    return s




def Corregir_peso_e_valor(q, wb=None, demandas_path=None, pfep_source=None):
    
    # Renomeia variáveis para clareza
    wb_cargolift = wb
    wb_demandas = demandas_path
    wb_pfep = pfep_source
    
    global caminho_base 
    global normalize_value 

    q.put(("status",f"--- 🚀 Iniciando a função Corregir_peso_e_valor ---"))
    q.put(("progress", 82))
    q.put(("status",f"Workbook Alvo: {wb_cargolift.name}"))
    q.put(("status",f"Workbook Demandas: {wb_demandas.name}"))
    q.put(("status",f"Workbook PFEP: {wb_pfep.name}"))

    try:
        # --- 1. Obter dados do 'wb_demandas' (PN e SAP) ---
        q.put(("status", "Etapa 1: Lendo dados do 'Demandas'..."))
        sheet_demandas = wb_demandas.sheets.active
        df_demandas = sheet_demandas.range('A1').expand().options(pd.DataFrame, index=False, header=True).value
        
        if 'PN' not in df_demandas.columns or 'SAP' not in df_demandas.columns:
            q.put(("status", "ERRO: Colunas 'PN' ou 'SAP' não encontradas no arquivo Demandas."))
            return

        df_demandas['PN_norm'] = df_demandas['PN'].apply(normalize_value)
        df_demandas['SAP_norm'] = df_demandas['SAP'].apply(normalize_value)
        q.put(("status", f"Ok. Encontrados {len(df_demandas)} registros no 'Demandas'."))

        # --- 2. Obter dados do 'wb_pfep' (Part Number RTM) ---
        q.put(("status","Etapa 2: Lendo dados do 'PFEP'..."))
        sheet_pfep = wb_pfep.sheets.active
        header_range = sheet_pfep.range('A6').expand('right')
        header_values = header_range.value
        
        pn_col_index = None # Índice 0-based
        if isinstance(header_values, list):
            try:
                pn_col_index = header_values.index('Part Number RTM')
            except ValueError:
                q.put(("status", "ERRO: Coluna 'Part Number RTM' não encontrada na linha 6 do 'PFEP'."))
                return
        else:
            q.put(("status", "ERRO: Não foi possível ler os cabeçalhos da linha 6 do 'PFEP'."))
            return

        pn_col_letter = xw.utils.col_name(pn_col_index + 1)
        q.put(("status",f"Ok. Coluna 'Part Number RTM' encontrada em: {pn_col_letter}6 (Índice 0-based: {pn_col_index})"))
        
        last_row_pfep = sheet_pfep.range(f'{pn_col_letter}7').end('down').row
        pfep_pns_raw = sheet_pfep.range(f'{pn_col_letter}7:{pn_col_letter}{last_row_pfep}').value
        
        pfep_pns_set = {normalize_value(pn) for pn in pfep_pns_raw if pn is not None}
        q.put(("status",f"Ok. Encontrados {len(pfep_pns_set)} PNs únicos no 'PFEP'."))
        
        if pfep_pns_set:
             q.put(("status",f" 	-> Debug: Exemplo de PN normalizado do PFEP: '{next(iter(pfep_pns_set))}'"))

        # --- 3. "XLOOKUP": Encontrar PNs que estão no 'Demandas' mas não no 'PFEP' ---
        q.put(("status","Etapa 3: Cruzando dados (PNs Demandas x PFs PFEP)..."))
        
        missing_pns_mask = ~df_demandas['PN_norm'].isin(pfep_pns_set)
        df_missing_pns = df_demandas[missing_pns_mask]
        
        found_pns_mask = ~missing_pns_mask
        found_pns_count = found_pns_mask.sum()
        missing_pns_count = len(df_missing_pns)
        
        q.put(("status", f"Ok. PNs de 'Demandas' ENCONTRADOS no 'PFEP': {found_pns_count}"))
        q.put(("status", f"Ok. PNs de 'Demandas' NÃO ENCONTRADOS no 'PFEP': {missing_pns_count}"))

        if missing_pns_count > 0 and not df_missing_pns.empty:
            q.put(("status", f" 	-> Exemplo de PN NÃO encontrado: '{df_missing_pns.iloc[0]['PN_norm']}' (Original: '{df_missing_pns.iloc[0]['PN']}')"))
        if found_pns_count > 0:
            df_found = df_demandas[found_pns_mask]
            if not df_found.empty:
                q.put(("status", f" 	-> Exemplo de PN ENCONTRADO: '{df_found.iloc[0]['PN_norm']}' (Original: '{df_found.iloc[0]['PN']}')"))

        if df_missing_pns.empty:
            q.put(("status","Ok. Nenhum PN do 'Demandas' está faltando no 'PFEP'. Encerrando."))
            q.put(("status","--- ✅ Processo concluído (sem alterações) ---"))
            return

        if wb_demandas :
                wb_demandas.close()

        if  wb_pfep :
             wb_pfep.close()
            
        q.put(("status", f"Ok. {len(df_missing_pns)} PNs faltantes serão filtrados pelo SAP..."))

        # --- 4. Filtrar pelos SAPs do JSON ---
        q.put(("status", "Etapa 4: Filtrando PNs faltantes pelo JSON 'Forncedores_Responsavel.json'..."))
        json_path = os.path.join(caminho_base, CONFIG['paths']['folders']['base_bases'], CONFIG['paths']['files']['fornecedores_responsavel_json'])
        
        if not os.path.exists(json_path):
            q.put(("status", f"ERRO: Arquivo JSON '{json_path}' não encontrado."))
            return

        with open(json_path, 'r', encoding='utf-8') as f:
            fornecedores_data = json.load(f)
        
        valid_saps_from_json = {normalize_value(sap): data for sap, data in fornecedores_data.items()}
        q.put(("status",f"Ok. Carregados {len(valid_saps_from_json)} SAPs do arquivo JSON."))

        saps_to_update_mask = df_missing_pns['SAP_norm'].isin(valid_saps_from_json.keys())
        df_final_list = df_missing_pns[saps_to_update_mask]

        if df_final_list.empty:
            q.put(("status","Ok. Nenhum dos PNs faltantes corresponde a um SAP válido no JSON. Encerrando."))
            q.put(("status","--- ✅ Processo concluído (sem alterações) ---"))
            return
        
        unique_saps_to_update = df_final_list['SAP_norm'].unique()
        q.put(("status", f"Ok. Processando atualizações para {len(unique_saps_to_update)} SAP(s) único(s)."))
        q.put(("status", f" 	-> SAPs para atualizar: {list(unique_saps_to_update)}"))


        # --- 5 & 6. Atualizar o 'wb_cargolift' ---
        q.put(("status","Etapa 5/6: Iniciando atualizações no 'Cargolift'..."))
        
        sheet_target = None
        target_sheet_name_fragment = CONFIG['paths']['sheet_names']['cargolift_sp_supplier_sheet'].strip().upper()
        q.put(("progress", 86))
        try:
            sheet_names_list = [sheet.name for sheet in wb_cargolift.sheets]
            q.put(("status", f"Planilhas encontradas no '{wb_cargolift.name}': {sheet_names_list}"))

            for sheet in wb_cargolift.sheets:
                normalized_sheet_name = sheet.name.strip().upper() 
                if target_sheet_name_fragment in normalized_sheet_name:
                    sheet_target = sheet
                    q.put(("status", f"SUCESSO: Planilha de atualização encontrada: '{sheet.name}'"))
                    break 

            if sheet_target is None:
                q.put(("status", f"ERRO: Nenhuma planilha encontrada no workbook '{wb_cargolift.name}' que contenha o nome '{target_sheet_name_fragment}'."))
                return
                
        except Exception as e:
            q.put(("status", f"ERRO: Falha ao tentar encontrar a planilha de atualização. Detalhe: {e}"))
            return

        # --- ⬇️⬇️⬇️ LÓGICA DE DIAS REMOVIDA (Conforme solicitado) ⬇️⬇️⬇️ ---
        m3_cols = ['I', 'J', 'K', 'L', 'M', 'N'] # Mon, Tue, Wed, Thu, Fri, Sat
        kg_cols = ['P', 'Q', 'R', 'S', 'T', 'U'] # Mon, Tue, Wed, Thu, Fri, Sat
        # --- ⬆️⬆️⬆️ FIM DA REMOÇÃO ⬆️⬆️⬆️ ---


        q.put(("status",f"Mapeando 'Supplier Code' da planilha '{sheet_target.name}' para otimização..."))
        last_row_target = sheet_target.range('C1').end('down').row
        sap_codes_raw = sheet_target.range(f'C2:C{last_row_target}').value
        
        sap_to_row_map = {}
        if isinstance(sap_codes_raw, list):
            for i, sap in enumerate(sap_codes_raw):
                sap_norm = normalize_value(sap) 
                if sap_norm:
                    sap_to_row_map[sap_norm] = i + 2
        else: # Caso de uma única célula
            sap_norm = normalize_value(sap_codes_raw)
            if sap_norm:
                sap_to_row_map[sap_norm] = 2

        q.put(("status",f"Ok. {len(sap_to_row_map)} SAPs mapeados na planilha 'Cargolift'."))

        if sap_to_row_map:
            exemplo_sap_mapeado = next(iter(sap_to_row_map.keys()))
            q.put(("status", f" 	-> Debug: Exemplo de SAP normalizado do 'Cargolift' (Coluna C): '{exemplo_sap_mapeado}'"))

        updates_made = 0
        for sap_norm in unique_saps_to_update:
            if sap_norm in sap_to_row_map:
                row_number = sap_to_row_map[sap_norm]
                
                json_data_for_sap = valid_saps_from_json[sap_norm]
                m3_to_add = json_data_for_sap.get("M3", 0)
                kg_to_add = json_data_for_sap.get("Kg", 0)
                
                # --- ⬇️⬇️⬇️ NOVA LÓGICA DE ATUALIZAÇÃO ⬇️⬇️⬇️ ---
                target_m3_col = None
                target_kg_col = None
                current_m3 = 0
                current_kg = 0

                try:
                    # Lê os valores da semana inteira para esta linha
                    m3_values = sheet_target.range(f'I{row_number}:N{row_number}').value
                    kg_values = sheet_target.range(f'P{row_number}:U{row_number}').value
                except Exception as e:
                    q.put((f"ERRO: Falha ao ler dados da linha {row_number}. {e}"))
                    continue # Pula para o próximo SAP

                # Encontra a primeira coluna de M3 com valor > 0
                if isinstance(m3_values, list):
                    for i, val in enumerate(m3_values):
                        if val is not None:
                            try:
                                if float(val) > 0:
                                    target_m3_col = m3_cols[i] # Ex: 'I'
                                    current_m3 = float(val)
                                    break # Para no primeiro que encontrar
                            except ValueError:
                                continue # Ignora valores não numéricos (texto)
                
                # Encontra a primeira coluna de Kg com valor > 0
                if isinstance(kg_values, list):
                    for i, val in enumerate(kg_values):
                        if val is not None:
                            try:
                                if float(val) > 0:
                                    target_kg_col = kg_cols[i] # Ex: 'P'
                                    current_kg = float(val)
                                    break # Para no primeiro que encontrar
                            except ValueError:
                                continue # Ignora valores não numéricos

                # Se encontrou colunas de planejamento para M3 e Kg, atualiza
                if target_m3_col and target_kg_col:
                    new_m3 = current_m3 + m3_to_add
                    new_kg = current_kg + kg_to_add
                    
                    # Escreve de volta nas colunas específicas que encontrou
                    sheet_target.range(f'{target_m3_col}{row_number}').value = new_m3
                    sheet_target.range(f'{target_kg_col}{row_number}').value = new_kg
                    
                    q.put(("status", f" 	-> SUCESSO: SAP {sap_norm} (Linha {row_number}): Col M3 '{target_m3_col}' ({current_m3} + {m3_to_add} = {new_m3}). Col Kg '{target_kg_col}' ({current_kg} + {kg_to_add} = {new_kg})."))
                    updates_made += 1
                else:
                    # Se não encontrou nenhuma coluna com valor > 0
                    q.put(("status", f" 	-> AVISO: SAP '{sap_norm}' (Linha {row_number}): Nenhum valor de planejamento (>0) encontrado nas colunas I-N e P-U. Pulando."))
                # --- ⬆️⬆️⬆️ FIM DA NOVA LÓGICA ⬆️⬆️⬆️ ---
                
            else:
                q.put(("status", f" 	-> AVISO: SAP '{sap_norm}' (da lista de PNs faltantes) não foi encontrado na coluna 'Supplier Code' do 'Cargolift'. Pulando."))
        
        q.put(("status",f"--- ✅ Processo concluído. {updates_made} atualizações de SAP realizadas. ---"))


        # calling the cut and copy fucntion to cut and copy data form each programme to the cargo lift.
        # this is the final part of the fuction 

        q.put(("status","--REABRINDO OS ARQUIVOS PARA COMEÇAR COPIA E COLAR--"))
        q.put(("progress",90))
        
        Copiar_planejamentos_para_cargolift_Arquivos(q=q) 
        # wb_cargolift_sp_Supplier.close()
    except Exception as e:
        q.put(("status",f"--- ❌ ERRO CRÍTICO na função ---"))
        q.put(("status", f"Erro: {e}"))
        q.put(("status",f"Traceback: {traceback.format_exc()}"))



def Copiar_planejamentos_para_cargolift_Arquivos(wb_cargolift = None,q = None) :

    q.put(("status","Inicializando a funcção de copia e colar os dados"))
    q.put(("progress",92))

    # caminho_pasta_matriz = os.path.join(caminho_base, '1 - MATRIZ')
    Programacao_CKD = get_path_from_config('ckd_terms','base_planilhas_recebidos',q)
    cargolift_prog_FPT = get_path_from_config('fpt_bt_terms','base_planilhas_recebidos',q)
    
    nome_prog_fiasa = get_path_from_config('fiasa_search_terms', 'base_matriz', q)
    cargolift_sp_PFEP = get_path_from_config('cargolift_pfep_terms', 'base_matriz', q)
    Programacao_FPT_Sul = get_path_from_config('fpt_sul_terms', 'base_matriz', q)

    cargolift_sp_Supplier = get_path_from_config('cargolift_supplier_terms', 'base_matriz', q)
    
    
    
    # Check if files were found before opening
    if not cargolift_sp_PFEP:
        q.put(("status", "ERRO: Arquivo 'Cargolift SP - PFEP' não encontrado na Matriz."))
        return
    if not cargolift_sp_Supplier:
        q.put(("status", "ERRO: Arquivo 'Cargolift SP - Suppliers' não encontrado na Matriz."))
        return
        
    q.put(("status", "Abrindo arquivos mestre..."))
    app_cargolift_sp_PFEP = xw.App(visible=True, add_book=False) 
    app_cargolift_sp_PFEP.display_alerts = False
    app_cargolift_sp_PFEP.api.AskToUpdateLinks = False
    wb_cargolift_sp_PFEP = app_cargolift_sp_PFEP.books.open(cargolift_sp_PFEP,update_links=False,read_only=False)
    # sheet_wb_cargolift_sp_PFEP = wb_cargolift_sp_PFEP.sheets['PFEP']


    app_cargolift_sp_Supplier = xw.App(visible=True, add_book=False) 
    app_cargolift_sp_Supplier.display_alerts = False
    app_cargolift_sp_Supplier.api.AskToUpdateLinks = False
    wb_cargolift_sp_Supplier = app_cargolift_sp_Supplier.books.open(cargolift_sp_Supplier,update_links=False,read_only=False) # Use a app correta
    # sheet_wb_cargolift_sp_Supplier = wb_cargolift_sp_Supplier.sheets['Cargolift SP - Suppliers DB Wk ']

    # Check if FPT file was found
    if cargolift_prog_FPT:
        q.put(("status", "Arquivo 'FPT BT' encontrado, iniciando cópia..."))
        Copiar_planejamentos_para_FPT_BT(
            q=q, cargolift_prog_FPT = cargolift_prog_FPT,wb_cargolift_sp_PFEP = wb_cargolift_sp_PFEP,
            wb_cargolift_sp_Supplier = wb_cargolift_sp_Supplier, Programacao_FPT_Sul = Programacao_FPT_Sul,
            programacao_fiasa_path = nome_prog_fiasa, Programacao_CKD_path = Programacao_CKD)
    else:
        q.put(("status", "AVISO: Arquivo 'FPT BT' não encontrado em Planilhas_Recebidos."))
        
    # ================= FINAL FILTERING =================
    q.put(("status", "Aplicando filtros finais nos arquivos Cargolift SP..."))
    
    try:
        # 1. Filtro no Supplier DB
        sheet_supplier = wb_cargolift_sp_Supplier.sheets[CONFIG['paths']['sheet_names']['cargolift_sp_supplier_sheet']]
        
        if sheet_supplier.api.AutoFilterMode:
            sheet_supplier.api.AutoFilterMode = False
            
        # Scan Row 89, Columns I to N (9 to 14)
        target_col_sup = None
        for col_idx in range(9, 15):
            val = sheet_supplier.range((89, col_idx)).value
            if val is not None and str(val).strip() != "":
                target_col_sup = col_idx
                break
                
        if target_col_sup:
            q.put(("status", f"Aplicando filtro > 0 na coluna {target_col_sup} do Supplier DB..."))
            last_row_sup = sheet_supplier.range('A' + str(sheet_supplier.cells.last_cell.row)).end('up').row
            filter_range_sup = sheet_supplier.range((1, 1), (last_row_sup, 60)) 
            filter_range_sup.api.AutoFilter(Field:=target_col_sup, Criteria1:=">0")

        # 2. Filtro no PFEP
        sheet_pfep = wb_cargolift_sp_PFEP.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']]
        
        if sheet_pfep.api.AutoFilterMode:
            sheet_pfep.api.AutoFilterMode = False

        # Scan Row 89, Columns F to K (6 to 11)
        target_col_pfep = None
        for col_idx in range(6, 12):
            val = sheet_pfep.range((89, col_idx)).value
            if val is not None and str(val).strip() != "":
                target_col_pfep = col_idx
                break
        
        if target_col_pfep:
            q.put(("status", f"Aplicando filtro > 0 na coluna {target_col_pfep} do PFEP..."))
            last_row_pfep = sheet_pfep.range('A' + str(sheet_pfep.cells.last_cell.row)).end('up').row
            filter_range_pfep = sheet_pfep.range((1, 1), (last_row_pfep, 60))
            
            filter_range_pfep.api.AutoFilter(Field:=target_col_pfep, Criteria1:=">0")
            
            q.put(("status", "Aplicando filtro adicional na coluna E (Part Weight > 0) do PFEP..."))
            filter_range_pfep.api.AutoFilter(Field:=5, Criteria1:=">0")

    except Exception as e:
        q.put(("status", f"ERRO ao aplicar filtros finais: {e}"))
        print(f"ERRO ao aplicar filtros finais: {e}")
    # ===================================================

    # You might want to save and close the master files here or in the calling function
    q.put(("status", "Salvando arquivos finais..."))
    # try:
    #     wb_cargolift_sp_PFEP.save()
    #     wb_cargolift_sp_PFEP.close()
    #     app_cargolift_sp_PFEP.quit()
        
    #     wb_cargolift_sp_Supplier.save()
    #     wb_cargolift_sp_Supplier.close()
    #     app_cargolift_sp_Supplier.quit()
    # except Exception as e:
    #     q.put(("status", f"AVISO: Problema ao salvar os arquivos finais: {e}"))

    q.put(("status", "Processo FPT BT concluído."))

    if CONFIG['business_logic'].get('auto_limpeza_diaria', False):
        limpar_e_arquivar(q)
    else:
        q.put(("status", "Auto limpeza desativada. Finalizado."))



def Copiar_planejamentos_para_FPT_BT(q=None, cargolift_prog_FPT = None , wb_cargolift_sp_PFEP = None, wb_cargolift_sp_Supplier =  None, Programacao_FPT_Sul =  None, programacao_fiasa_path =  None ,Programacao_CKD_path = None):
    
    app_cargolift_prog_FPT = None
    wb_cargolift_prog_FPT = None
   
    
    try:
        q.put(("status", "Abrindo FPT BT..."))
        q.put(("progress",94))
        app_cargolift_prog_FPT = xw.App(visible=False, add_book=False) 
        app_cargolift_prog_FPT.display_alerts = False
        app_cargolift_prog_FPT.api.AskToUpdateLinks = False
        wb_cargolift_prog_FPT = app_cargolift_prog_FPT.books.open(cargolift_prog_FPT,update_links=False,read_only=False)

        sheet_cargolift_prog_FPT_PFEP = wb_cargolift_prog_FPT.sheets[CONFIG['paths']['sheet_names']['fpt_bt_pfep']]
        sheet_cargolift_prog_FPT__Suppliers_DB  = wb_cargolift_prog_FPT.sheets[CONFIG['paths']['sheet_names']['fpt_bt_supplier']]
        
        sheet_wb_cargolift_sp_Supplier = wb_cargolift_sp_Supplier.sheets[CONFIG['paths']['sheet_names']['cargolift_sp_supplier_sheet']]
        sheet_wb_cargolift_sp_PFEP = wb_cargolift_sp_PFEP.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']] # Assuming 'PFEP'

        filter_criteria = CONFIG['business_logic']['fpt_filter_criteria_sp']

        xlCellTypeVisible = 12 # VBA Constant for SpecialCells

        # --- 1. Process Suppliers DB ---
        q.put(("status", "Processando FPT Suppliers DB..."))
        sheet_sup = sheet_cargolift_prog_FPT__Suppliers_DB
        sheet_sup.api.AutoFilterMode = False # Clear existing filters
        
        last_row_sup = sheet_sup.range('A' + str(sheet_sup.cells.last_cell.row)).end('up').row
        
        if last_row_sup > 2:
            filter_range_sup = sheet_sup.range(f'A2:X{last_row_sup}')
            data_range_sup = sheet_sup.range(f'A3:X{last_row_sup}')
            
            # --- Apply filter for SP/SUL ---
            filter_range_sup.api.AutoFilter(Field:=24, Criteria1:=filter_criteria, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
            
            try:
                visible_cells_sup = data_range_sup.api.SpecialCells(xlCellTypeVisible)
                data_sup = []
                for area in visible_cells_sup.Areas:
                    values = sheet_sup.range(area.Address).value
                    if isinstance(values, list) and isinstance(values[0], list): 
                        data_sup.extend(values)
                    else: 
                        data_sup.append(values)
                
                if not data_sup:
                    q.put(("status", "Nenhum dado visível em FPT Suppliers."))
                else:
                    
                    next_row_sup = sheet_wb_cargolift_sp_Supplier.range('A' + str(sheet_wb_cargolift_sp_Supplier.cells.last_cell.row)).end('up').row + 1
                    sheet_wb_cargolift_sp_Supplier.range(f'A{next_row_sup}').value = data_sup
                    q.put(("status", f"{len(data_sup)} linhas coladas em 'Cargolift SP - Suppliers DB Wk '"))

            except Exception as e:
                q.put(("status", f"Nenhum dado encontrado no filtro FPT Suppliers: {e}"))
            
            sheet_sup.api.AutoFilterMode = False # Clear filter
        else:
            q.put(("status", "Nenhum dado para processar em FPT Suppliers DB."))

        # --- 2. Process PFEP ---
        q.put(("status", "Processando FPT PFEP..."))
        sheet_pfep = sheet_cargolift_prog_FPT_PFEP
        sheet_pfep.api.AutoFilterMode = False # Clear existing filters
        
        last_row_pfep = sheet_pfep.range('AX' + str(sheet_pfep.cells.last_cell.row)).end('up').row
        
        if last_row_pfep > 3:
            filter_range_pfep = sheet_pfep.range(f'A3:AX{last_row_pfep}')
            data_range_pfep = sheet_pfep.range(f'M4:BI{last_row_pfep}')
            
            # --- Apply filter for SP/SUL ---
            filter_range_pfep.api.AutoFilter(Field:=50, Criteria1:=filter_criteria, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
            
            try:
                visible_cells_pfep = data_range_pfep.api.SpecialCells(xlCellTypeVisible)
                data_pfep = []
                for area in visible_cells_pfep.Areas:
                    values = sheet_pfep.range(area.Address).value
                    if isinstance(values, list) and isinstance(values[0], list): 
                        data_pfep.extend(values)
                    else: 
                        data_pfep.append(values)
                
                if not data_pfep:
                    q.put(("status", "Nenhum dado visível em FPT PFEP."))
                else:

                    next_row_pfep = sheet_wb_cargolift_sp_PFEP.range('A' + str(sheet_wb_cargolift_sp_PFEP.cells.last_cell.row)).end('up').row + 1
                    sheet_wb_cargolift_sp_PFEP.range(f'A{next_row_pfep}').value = data_pfep
                    q.put(("status", f"{len(data_pfep)} linhas coladas em 'PFEP'"))
                    
            except Exception as e:
                q.put(("status", f"Nenhum dado encontrado no filtro FPT PFEP: {e}"))
                
            sheet_pfep.api.AutoFilterMode = False # Clear filter
        else:
            q.put(("status", "Nenhum dado para processar em FPT PFEP."))
        

       
    except Exception as e:
        q.put(("status", f"ERRO ao processar FPT_BT: {e}"))
        print(f"ERRO ao processar FPT_BT: {e}")
    
    finally:
        # Close the local FPT file
        if wb_cargolift_prog_FPT :
            wb_cargolift_prog_FPT.close()
        if app_cargolift_prog_FPT:
            app_cargolift_prog_FPT.quit()
        q.put(("status", "Arquivo FPT BT fechado."))

    
    # --- Call function to process Mopar and CKD ---
    q.put(("progress",96))
    print("processing Mopar and CKD")
    data_ckd_pfep_sp, data_ckd_supplier_sp = _ler_dados_ckd(q, Programacao_CKD_path)
    copiar_e_colar_SP(q= q , data_ckd_pfep_sp = data_ckd_pfep_sp, data_ckd_supplier_sp =  data_ckd_supplier_sp, wb_cargolift_sp_PFEP = wb_cargolift_sp_PFEP,
                                            wb_cargolift_sp_Supplier = wb_cargolift_sp_Supplier)
    print("Done processing Mopar and CKD")



xlCellTypeVisible = 12 
last_col_format_index = 400


def _ler_dados_ckd(q, Programacao_CKD_path):
    """
    Abre o arquivo CKD, lê os dados (SUL e SP combinados) e fecha o arquivo.
    Retorna (data_ckd_pfep_sp, data_ckd_supplier_sp).
    """
    data_ckd_pfep_sp = []
    data_ckd_supplier_sp = []
    
    app_programacao_ckd = None
    wb_programacao_ckd = None
    
    # Combined filter for both SP and SUL
    filter_criteria = ["CARGOLIFT", "CARGOLIFT SUL"]

    if not Programacao_CKD_path:
        q.put(("status", "AVISO: Caminho para 'Programacao_CKD' não fornecido. Pulando esta etapa."))
        return data_ckd_pfep_sp, data_ckd_supplier_sp

    try:
        q.put(("status", f"Abrindo arquivo CKD para ler dados: {Programacao_CKD_path}"))
        app_programacao_ckd = xw.App(visible=False, add_book=False)
        app_programacao_ckd.display_alerts = False
        app_programacao_ckd.api.AskToUpdateLinks = False
        wb_programacao_ckd = app_programacao_ckd.books.open(Programacao_CKD_path, update_links=False, read_only=True)

        # --- Read CKD PFEP ---
        q.put(("status", "Lendo dados CKD PFEP..."))
        try:
            sheet_ckd_pfep = wb_programacao_ckd.sheets[CONFIG['paths']['sheet_names']['ckd_pfep']]
            sheet_ckd_pfep.api.AutoFilterMode = False
            last_row_ckd_pfep = sheet_ckd_pfep.range('W' + str(sheet_ckd_pfep.cells.last_cell.row)).end('up').row

            if last_row_ckd_pfep > 1:
                filter_range_ckd_pfep = sheet_ckd_pfep.range(f'A1:W{last_row_ckd_pfep}')
                data_range_ckd_pfep = sheet_ckd_pfep.range(f'A2:CC{last_row_ckd_pfep}') # Read A:CC (81 cols)
                
                # --- LER DADOS COMBINADOS ---
                filter_range_ckd_pfep.api.AutoFilter(Field:=23, Criteria1:=filter_criteria, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
                try:
                    visible_cells_sp = data_range_ckd_pfep.api.SpecialCells(xlCellTypeVisible)
                    raw_data_pfep_sp = []
                    for area in visible_cells_sp.Areas:
                        values = sheet_ckd_pfep.range(area.Address).value
                        if isinstance(values, list) and isinstance(values[0], list):
                            raw_data_pfep_sp.extend(values)
                        else:
                            raw_data_pfep_sp.append(values)

                    q.put(("status", "Mapeando colunas CKD PFEP..."))
                    for read_row in raw_data_pfep_sp:
                        try:
                            # --- 1. PERFORM MAPPING FIRST ---
                            read_row[54] = read_row[13] 
                            
                            # --- 2. NOW, COPY THE FINAL DATA ---
                            new_row = [None] * last_col_format_index
                            base_data = read_row[41:81] 
                            new_row[0:len(base_data)] = base_data
                            data_ckd_pfep_sp.append(new_row)
                        except IndexError:
                            q.put(("status", "ERRO: A linha no CKD PFEP é muito curta para o mapeamento."))
                except Exception as e:
                    q.put(("status", f"Nenhum dado 'CARGOLIFT' em CKD PFEP: {e}"))

                sheet_ckd_pfep.api.AutoFilterMode = False 
                
        except Exception as e:
            q.put(("status", f"ERRO ao ler sheet CKD PFEP: {e}"))

        # --- Read CKD SUPPLIER ---
        q.put(("status", "Lendo dados CKD Suppliers..."))
        try:
            sheet_ckd_sup = wb_programacao_ckd.sheets[CONFIG['paths']['sheet_names']['ckd_supplier']]
            sheet_ckd_sup.api.AutoFilterMode = False
            last_row_ckd_sup = sheet_ckd_sup.range('L' + str(sheet_ckd_sup.cells.last_cell.row)).end('up').row

            if last_row_ckd_sup > 1: 
                filter_range_ckd_sup = sheet_ckd_sup.range(f'A1:L{last_row_ckd_sup}')
                data_range_ckd_sup = sheet_ckd_sup.range(f'A2:BD{last_row_ckd_sup}') # Read A:BD (56 cols)
                
                # --- LER DADOS COMBINADOS ---
                filter_range_ckd_sup.api.AutoFilter(Field:=12, Criteria1:=filter_criteria, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
                try:
                    visible_cells_sup_sp = data_range_ckd_sup.api.SpecialCells(xlCellTypeVisible)
                    raw_data_sup_sp = []
                    for area in visible_cells_sup_sp.Areas:
                        values = sheet_ckd_sup.range(area.Address).value
                        if isinstance(values, list) and isinstance(values[0], list):
                            raw_data_sup_sp.extend(values)
                        else:
                            raw_data_sup_sp.append(values)

                    q.put(("status", "Mapeando colunas CKD Supplier..."))
                    for read_row in raw_data_sup_sp:
                        try:
                            # --- 1. PERFORM MAPPING FIRST ---
                            read_row[32] = read_row[2] 
                            read_row[33] = read_row[3] 
                            read_row[34] = read_row[4] 
                            read_row[35] = read_row[5] 
                            read_row[36] = read_row[6] 
                            
                            # --- 2. NOW, COPY THE FINAL DATA ---
                            new_row = [None] * last_col_format_index
                            base_data = read_row[29:55] 
                            new_row[0:len(base_data)] = base_data
                            
                            data_ckd_supplier_sp.append(new_row)
                        except IndexError:
                            q.put(("status", "ERRO: A linha no CKD é muito curta para o mapeamento."))
                        
                except Exception as e:
                    q.put(("status", f"Nenhum dado 'CARGOLIFT' em CKD Suppliers: {e}"))
                
                sheet_ckd_sup.api.AutoFilterMode = False 
                
        except Exception as e:
            q.put(("status", f"ERRO ao ler sheet CKD Supplier: {e}"))

    except Exception as e:
        q.put(("status", f"ERRO ao ler arquivo CKD: {e}"))
    finally:
        if wb_programacao_ckd:
            wb_programacao_ckd.close()
        if app_programacao_ckd:
            app_programacao_ckd.quit()
        q.put(("status", "Arquivo Leitura (CKD) fechado."))
    
    return data_ckd_pfep_sp, data_ckd_supplier_sp




# -----------------------------------------------------------------------------
# FUNÇÕES AUXILIARES DE LEITURA (Refatoradas)
# -----------------------------------------------------------------------------

def _safe_open_workbook(q, app, path):
    """Abre um workbook com segurança, desabilitando alertas e atualizações."""
    try:
        if not path or not os.path.exists(path):
            if q: q.put(("status", f"AVISO: Arquivo não encontrado: {path}"))
            return None
            
        if q: q.put(("status", f"Abrindo arquivo: {os.path.basename(path)}"))
        app.display_alerts = False
        app.api.AskToUpdateLinks = False
        wb = app.books.open(path, update_links=False, read_only=True)
        return wb
    except Exception as e:
        if q: q.put(("status", f"ERRO ao abrir {path}: {e}"))
        return None

def _get_sheet_by_name_or_index(q, wb, sheet_name_list):
    """Tenta encontrar uma planilha por uma lista de nomes; se falhar, pega a primeira."""
    for name in sheet_name_list:
        try:
            return wb.sheets[name]
        except:
            continue # Tenta o próximo nome
            
    # Se nenhum nome foi encontrado, pega a primeira planilha
    try:
        sheet = wb.sheets[0]
        if q: q.put(("status", f"AVISO: Nenhuma planilha com nome {sheet_name_list} encontrada. Usando a primeira: {sheet.name}"))
        return sheet
    except Exception as e:
        if q: q.put(("status", f"ERRO: Workbook {wb.name} parece estar vazio ou corrupto. {e}"))
        return None

def _read_data_from_range(q, sheet, start_cell, end_col):
    """Lê dados de um range, começando em start_cell até a última linha e coluna especificada."""
    try:
        last_row = sheet.range(start_cell).end('down').row
        if last_row == sheet.cells.last_cell.row: # Caso de coluna vazia
            last_row = sheet.range(start_cell).end('up').row
        
        # Correção se a planilha tiver apenas cabeçalho (ou 1 linha)
        if sheet.range(start_cell).row > last_row:
             if q: q.put(("status", f"Nenhum dado (>= {start_cell}) encontrado na planilha {sheet.name}."))
             return []

        # Define o range de A2 até a última linha e coluna
        range_str = f'{start_cell}:{end_col}{last_row}'
        data = sheet.range(range_str).value
        
        # Garantir que seja sempre uma lista de listas (2D)
        if last_row == sheet.range(start_cell).row and not isinstance(data[0], list):
            data = [data] # Transforma 1D em 2D
            
        if q: q.put(("status", f"Lidas {len(data)} linhas ({range_str}) de {sheet.name}."))
        return data
    except Exception as e:
        if q: q.put(("status", f"ERRO ao ler dados do range {range_str} em {sheet.name}: {e}"))
        return []




def _find_filter_column(q, sheet, start_col_char, end_col_char):
    """
    [VERSÃO FINAL v10]
    Esta é a correção. Em vez de checar a Linha 2,
    usa COUNTA (Cont.Valores) para encontrar a primeira coluna
    que não está vazia no range de DADOS.
    """
    try:
        start_idx = sheet.range(f'{start_col_char}1').column
        end_idx = sheet.range(f'{end_col_char}1').column
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        if last_row <= 1:
            q.put(("status", "AVISO: Planilha só tem cabeçalho, não é possível encontrar a coluna de filtro."))
            return None

        q.put(("status", f"Buscando coluna com dados (Cont.Valores > 0) na Linha 2 até {last_row}..."))

        for col_idx in range(start_idx, end_idx + 1):
            col_letter = sheet.range((1, col_idx)).address.split('$')[1]
            
            # Define o range de DADOS para esta coluna
            data_range = sheet.range(f'{col_letter}2:{col_letter}{last_row}')
            
            # Usa a função COUNTA (Cont.Valores) do Excel para ver se há algum valor
            count = 0
            try:
                # O CountA é a forma mais rápida de verificar se há *qualquer* valor
                count = sheet.api.Application.WorksheetFunction.CountA(data_range.api)
            except Exception as e_counta:
                # Plano B se o CountA falhar (raro)
                q.put(("status", f"Aviso: CountA falhou para coluna {col_letter}. Verificando manualmente."))
                values = data_range.value
                if isinstance(values, list): # Múltiplas linhas
                    if any(v is not None and v != "" for v in values):
                        count = 1
                elif values is not None and values != "": # Linha única
                    count = 1

            q.put(("status", f"Verificando {col_letter} (Cont.Valores={count})"))

            # Se achou uma coluna com dados, essa é a nossa coluna!
            if count > 0:
                q.put(("status", f"Coluna de filtro encontrada (primeira com dados): {col_letter}"))
                return col_idx # Retorna o índice 1-based
        
        q.put(("status", f"AVISO: Nenhuma coluna com dados (Cont.Valores > 0) encontrada em {sheet.name} entre {start_col_char}-{end_col_char}."))
        return None
    except Exception as e:
        if q: q.put(("status", f"ERRO ao procurar coluna de filtro com CountA: {e}"))
        return None



def _read_filtered_data(q, sheet, filter_col_start, filter_col_end, copy_range_str):
    """
    [VERSÃO FINAL v13]
    Esta é a correção final, implementando sua lógica.
    Nós usamos 'Areas' para encontrar as *linhas* visíveis,
    mas então lemos o *range de colunas completo* (ex: A:AP)
    para aquelas linhas, como você especificou.
    """
    data = []
    
    # 1. Encontrar a coluna para filtrar
    filter_col_idx = _find_filter_column(q, sheet, filter_col_start, filter_col_end)
    if not filter_col_idx:
        return []

    try:
        # 2. Limpar qualquer filtro antigo
        sheet.api.AutoFilterMode = False

        # 3. Definir os ranges
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        if last_row <= 1:
            if q: q.put(("status", f"Nenhum dado para filtrar em {sheet.name} (apenas cabeçalho)."))
            return []
            
        # Pega o range de colunas completo (ex: 'A' e 'AP')
        start_col_letter = copy_range_str.split(':')[0]
        last_col_letter = copy_range_str.split(':')[-1]
        
        filter_range = sheet.range(f'A1:{last_col_letter}{last_row}')
        data_range = sheet.range(f'A2:{last_col_letter}{last_row}')

        # 4. Aplicar o filtro
        xlAnd = 1 
        filter_range.api.AutoFilter(
            Field:=filter_col_idx,
            Criteria1:=">0",
            Operator:=xlAnd,
            Criteria2:="<>#N/A"
        )

        # 5. Copiar os dados visíveis (USANDO 'Areas' DA FORMA CORRETA)
        try:
            xlCellTypeVisible = 12
            visible_cells = data_range.api.SpecialCells(xlCellTypeVisible)
            
            # 1. Sim, nós usamos 'Areas', como você disse.
            for area in visible_cells.Areas:
                # 2. Descobrimos os limites da linha da área visível
                start_row = area.row
                end_row = start_row + area.rows.count - 1

                # 3. Construímos o range para ler, usando o range de colunas COMPLETO
                #    que você especificou (ex: 'A' até 'AP')
                correct_range_str = f"{start_col_letter}{start_row}:{last_col_letter}{end_row}"
                
                # 4. Lemos o valor do range de largura total
                values = sheet.range(correct_range_str).value
                
                if not values:
                    continue
                
                # 5. Anexamos (com a correção do 'tuple')
                if isinstance(values, list) and isinstance(values[0], list):
                    data.extend(values) # Bloco 2D (múltiplas linhas)
                else:
                    data.append(list(values)) # Bloco 1D (linha única)
                    
        except Exception as e_visible:
            if q: q.put(("status", f"Nenhum dado visível em {sheet.name} após o filtro (Coluna {filter_col_idx})."))
            data = []
        
        if data:
            q.put(("status", f"Lidas {len(data)} linhas filtradas de {sheet.name} (Coluna {filter_col_idx})."))

    except Exception as e:
        if q: q.put(("status", f"ERRO ao filtrar ou copiar dados de {sheet.name}: {e}"))
    finally:
        # 6. Limpar o filtro
        try:
            sheet.api.AutoFilterMode = False
        except Exception as e_clear:
            if q: q.put(("status", f"Aviso: Não foi possível limpar o filtro em {sheet.name}. {e_clear}"))
            
    return data



def _read_filtered_data(q, sheet, filter_col_start, filter_col_end, copy_range_str):
    data = []
    
    filter_col_idx = _find_filter_column(q, sheet, filter_col_start, filter_col_end)
    if not filter_col_idx:
        return []

    try:
        sheet.api.AutoFilterMode = False

        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        if last_row <= 1:
            if q: q.put(("status", f"Nenhum dado para filtrar em {sheet.name} (apenas cabeçalho)."))
            return []
            
        last_col_letter = copy_range_str.split(':')[-1]
        filter_range = sheet.range(f'A1:{last_col_letter}{last_row}')
        data_range = sheet.range(f'A2:{last_col_letter}{last_row}')

        xlAnd = 1 
        filter_range.api.AutoFilter(
            Field:=filter_col_idx,
            Criteria1:=">0",
            Operator:=xlAnd,
            Criteria2:="<>#N/A"
        )

        try:
            xlCellTypeVisible = 12
            visible_cells = data_range.api.SpecialCells(xlCellTypeVisible)
            
            for area in visible_cells.Areas:
                values = sheet.range(area.Address).value
                if not values:
                    continue
                if isinstance(values, list) and isinstance(values[0], list):
                    data.extend(values)
                else:
                    data.append(values)
                    
        except Exception:
            if q: q.put(("status", f"Nenhum dado visível em {sheet.name} após o filtro.")) 
            data = []
        
        if data:
            # 🔧 Normalizar as linhas antes de retornar
            max_len = max(len(r) if isinstance(r, (list, tuple)) else 1 for r in data)
            normalized = []
            for row in data:
                if isinstance(row, (list, tuple)):
                    row = list(row)
                else:
                    row = [row]
                row += [None] * (max_len - len(row))
                normalized.append(row)
            data = normalized

            # 🧹 Remover linhas cujo primeiro valor está vazio
            before = len(data)
            data = [row for row in data if str(row[0]).strip() not in ("", "None", "nan")]
            removed = before - len(data)
            if removed:
                q.put(("status", f"Removidas {removed} linhas com primeira coluna vazia de {sheet.name}."))

            q.put(("status", f"[DEBUG] Dados normalizados ({len(data)} linhas, {max_len} colunas) antes do retorno."))
            q.put(("status", f"Lidas {len(data)} linhas filtradas de {sheet.name}."))

    except Exception as e:
        if q: q.put(("status", f"ERRO ao filtrar ou copiar dados de {sheet.name}: {e}"))
    finally:
        try:
            sheet.api.AutoFilterMode = False
        except Exception as e_clear:
            if q: q.put(("status", f"Aviso: Não foi possível limpar o filtro em {sheet.name}. {e_clear}"))
            
    return data




def _read_mopar_data(q):
    """Lê os dados dos arquivos MOPAR Supplier e PFEP."""
    if q: q.put(("status", "Lendo arquivos MOPAR..."))
    data_mopar_sup = []
    data_mopar_pfep = []
    app = None
    wb_sup = None
    wb_pfep = None

    try:
        cargolift_prog_MOPAR_path = get_path_from_config("mopar_supplier_terms", "base_planilhas_recebidos", q)
        cargolift_sp_Part_Number_MOPAR_path = get_path_from_config("mopar_pn_terms", "base_planilhas_recebidos", q)


        app = xw.App(visible=True, add_book=False)

        # --- Ler MOPAR Supplier ---
        wb_sup = _safe_open_workbook(q, app, cargolift_prog_MOPAR_path)
        if wb_sup:
            sheet_sup = _get_sheet_by_name_or_index(q, wb_sup, ['Suppliers'])
            if sheet_sup:
                # Copia A2 até AG(última linha)
                data_mopar_sup = _read_data_from_range(q, sheet_sup, 'A2', 'AG')

        # --- Ler MOPAR PFEP ---
        wb_pfep = _safe_open_workbook(q, app, cargolift_sp_Part_Number_MOPAR_path)
        if wb_pfep:
            sheet_pfep = _get_sheet_by_name_or_index(q, wb_pfep, ['PFEP'])
            if sheet_pfep:
                # Copia A2 até AU(última linha)
                data_mopar_pfep = _read_data_from_range(q, sheet_pfep, 'A2', 'AU')

    except Exception as e:
        if q: q.put(("status", f"ERRO ao ler arquivos MOPAR: {e}"))
    finally:
        if wb_sup: wb_sup.close()
        if wb_pfep: wb_pfep.close()
        if app: app.quit()
        if q: q.put(("status", "Arquivos MOPAR fechados."))
        
    return data_mopar_sup, data_mopar_pfep


def _read_porto_real_data(q):
    """Lê os dados dos arquivos Porto Real Supplier e PFEP."""
    if q: q.put(("status", "Lendo arquivos Porto Real..."))
    data_pr_sup = []
    data_pr_pfep = []
    app = None
    wb = None
    
    try:
        Cargolift_prog_PR_path = get_path_from_config("porto_real_terms", "base_planilhas_recebidos", q) # Assumindo que está em 'recebidos'
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        
        app.api.AskToUpdateLinks = False  # 🔒 Prevents Excel popup

        
        wb = _safe_open_workbook(q, app, Cargolift_prog_PR_path)
        if wb:
            # --- Ler Porto Real Supplier ---
            # Tenta 'Suppliers', depois 'supplier'
            sheet_sup = _get_sheet_by_name_or_index(q, wb, ['supplier', 'Suppliers'])
            if sheet_sup:
                # Copia B2 até AM(última linha)
                data_pr_sup = _read_data_from_range(q, sheet_sup, 'B2', 'AM')
            
            # --- Ler Porto Real PFEP ---
            sheet_pfep = _get_sheet_by_name_or_index(q, wb, ['PFEP'])
            if sheet_pfep:
                # Copia B2 até AV(última linha)
                data_pr_pfep = _read_data_from_range(q, sheet_pfep, 'B2', 'AV')
                
    except Exception as e:
        if q: q.put(("status", f"ERRO ao ler arquivo Porto Real: {e}"))
    finally:
        if wb: wb.close()
        if app: app.quit()
        if q: q.put(("status", "Arquivo Porto Real fechado."))
        
    return data_pr_sup, data_pr_pfep


def _read_fiape_data(q):
    """Lê os dados filtrados dos arquivos FIAPE."""
    if q: q.put(("status", "Lendo arquivos FIAPE..."))
    data_fiape_sup = []
    data_fiape_pfep = []
    app = None
    wb = None
    
    try:
        Cargolift_prog_FIAPE_path = get_path_from_config("fiape_terms", "base_planilhas_recebidos", q) # Assumindo que está em 'recebidos'
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        
        app.api.AskToUpdateLinks = False  # 🔒 Prevents Excel popup

        
        wb = _safe_open_workbook(q, app, Cargolift_prog_FIAPE_path)
        if wb:
            # --- Ler FIAPE PFEP ---
            sheet_pfep = _get_sheet_by_name_or_index(q, wb, ['PFEP'])
            if sheet_pfep:
                # Filtra F-L (>0 e não #N/A), copia A:BC
                data_fiape_pfep = _read_filtered_data(q, sheet_pfep, 'F', 'L', 'A:BC')
            
            # --- Ler FIAPE Suppliers DB ---
            sheet_sup = _get_sheet_by_name_or_index(q, wb, ['Suppliers DB'])
            if sheet_sup:
                # Filtra L-N (>0 e não #N/A), copia A:AP
                data_fiape_sup = _read_filtered_data(q, sheet_sup, 'I', 'N', 'A:AG')

    except Exception as e:
        if q: q.put(("status", f"ERRO ao ler arquivo FIAPE: {e}"))
    finally:
        if wb: wb.close()
        if app: app.quit()
        if q: q.put(("status", "Arquivo FIAPE fechado."))
        
    return data_fiape_sup, data_fiape_pfep


def _paste_data_to_sheet(q, sheet, data_blocks, source_name, last_col_format_index=40):
    """
    Cola múltiplos blocos de dados em uma planilha de destino e aplica formatação.
    """
    if not sheet:
        if q: q.put(("status", f"ERRO: Planilha de destino para {source_name} é inválida."))
        return
        
    try:
        if q: q.put(("status", f"Iniciando colagem em {source_name}..."))
        
        # Encontra a primeira linha vazia para colar
        start_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
        if start_row == 2 and sheet.range('A1').value is None:
            start_row = 1
        if sheet.range('A1').value is not None and start_row == 1:
            start_row = 2
        if start_row == 2 and sheet.range('A2').value is not None:
            start_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
            
        current_row = start_row
        total_pasted = 0
        
        # Itera e cola cada bloco de dados
        for i, data in enumerate(data_blocks):
            if not data:
                if q: q.put(("status", f"Bloco {i+1} de {source_name} está vazio. Pulando."))
                continue
            
            # Garante que os dados sejam 2D, se for uma única linha
            if not isinstance(data[0], list):
                data = [data]
            
            rows_to_paste = len(data)
            if rows_to_paste > 0:
                sheet.range(f'A{current_row}').value = data
                current_row += rows_to_paste
                total_pasted += rows_to_paste
                if q: q.put(("status", f"Coladas {rows_to_paste} linhas do bloco {i+1} em {source_name}."))

        # Aplicar formatação em TODO O BLOCO colado (se algo foi colado)
        if total_pasted > 0:
            end_row = current_row - 1
            # Tenta copiar o formato da linha 3 (como no original)
            try:
                source_format_range = sheet.range((3, 1), (3, last_col_format_index))
                dest_format_range = sheet.range((start_row, 1), (end_row, last_col_format_index))
                
                source_format_range.copy()
                dest_format_range.paste(paste='formats')
                
                sheet.book.app.api.CutCopyMode = False
                if q: q.put(("status", f"Formatação aplicada a {total_pasted} linhas de {source_name}."))
            except Exception as e_format:
                if q: q.put(("status", f"AVISO: Não foi possível aplicar formatação em {source_name}. {e_format}"))
        else:
            if q: q.put(("status", f"Nenhum dado para colar em {source_name}."))

    except Exception as e:
        if q: q.put(("status", f"ERRO ao colar dados em {source_name}: {e}"))


def copiar_e_colar_SP(q=None, data_ckd_pfep_sp=None, data_ckd_supplier_sp=None,
                      wb_cargolift_sp_PFEP=None, wb_cargolift_sp_Supplier=None):
    """
    Orquestra a leitura de dados (MOPAR, FIAPE, Porto Real) e os cola,
    juntamente com os dados CKD, nas planilhas de destino SP.
    """
    q.put(("status", "Iniciando processo de colagem SP..."))
    # --- 1. Definir Destinos ---
    try:
        sheet_wb_cargolift_sp_Supplier = wb_cargolift_sp_Supplier.sheets[CONFIG['paths']['sheet_names']['cargolift_sp_supplier_sheet']]
        sheet_wb_cargolift_sp_PFEP = wb_cargolift_sp_PFEP.sheets[CONFIG['paths']['sheet_names']['pfep_main_sheet']]
    except Exception as e:
        if q: q.put(("status", f"ERRO FATAL: Não foi possível encontrar as planilhas de destino. {e}"))
        return

    # --- 2. Ler todos os dados das fontes ---
    # (Os arquivos são abertos e fechados dentro de cada função)
    data_mopar_sup, data_mopar_pfep = _read_mopar_data(q)
    data_fiape_sup, data_fiape_pfep = _read_fiape_data(q)
    data_pr_sup, data_pr_pfep = _read_porto_real_data(q)
    q.put(("status", "Leitura de todas as fontes concluída."))
    # --- 3. Agrupar dados para colagem ---
    # A ordem de colagem será: MOPAR, FIAPE, Porto Real, CKD
    
    supplier_data_blocks = [
        data_mopar_sup,
        data_fiape_sup,
        data_pr_sup,
        data_ckd_supplier_sp 
    ]
    
    q.put(("status", "Dados Supplier agrupados para colagem."))
    pfep_data_blocks = [
        data_mopar_pfep,
        data_fiape_pfep,
        data_pr_pfep,
        data_ckd_pfep_sp
    ]

    q.put(("status", "Dados PFEP agrupados para colagem."))

    # --- 4. Colar blocos de dados ---
    # Assume que a formatação vai até a coluna 50 (AU)
    _paste_data_to_sheet(q, sheet_wb_cargolift_sp_Supplier, supplier_data_blocks, "Supplier SP", 50) 
    _paste_data_to_sheet(q, sheet_wb_cargolift_sp_PFEP, pfep_data_blocks, "PFEP SP", 50) 
    q.put(("status", "Colagem de todos os dados SP concluída."))
    if q: q.put(("status", "Processo de colagem SP concluído."))




def limpar_e_arquivar(q=None):
    """
    Cria uma pasta diária no Historico_Resultados, copia os arquivos finais para lá,
    e move os arquivos de entrada (Demanda, Planilhas_Recebidos) para uma subpasta 'Entradas'.
    """
    if q: q.put(("status", "Iniciando arquivamento e limpeza diária..."))
    
    try:
        agora = datetime.now()
        nome_pasta_diaria = agora.strftime('%Y-%m-%d_%H-%M')
        
        caminho_historico_base = os.path.join(caminho_base, CONFIG['paths']['folders'].get('base_historico', 'Historico_Resultados'))
        caminho_pasta_hoje = os.path.join(caminho_historico_base, nome_pasta_diaria)
        caminho_entradas_hoje = os.path.join(caminho_pasta_hoje, 'Entradas')
        
        os.makedirs(caminho_entradas_hoje, exist_ok=True)
        
        # 1. Copiar Resultados Finais para o Histórico
        if q: q.put(("status", "Copiando arquivos finais para o histórico..."))
        
        # Demandas_Total
        demandas_path = os.path.join(caminho_base, CONFIG['paths']['folders']['base_resultados'], CONFIG['paths']['files']['demandas_total_output'])
        if os.path.exists(demandas_path):
            shutil.copy2(demandas_path, caminho_pasta_hoje)
            
        # Programação FIASA - OFICIAL
        fiasa_path = get_path_from_config('fiasa_search_terms', 'base_matriz', q)
        if fiasa_path and os.path.exists(fiasa_path):
            shutil.copy2(fiasa_path, caminho_pasta_hoje)
            
        # Cargolift SP - PFEP
        cargolift_pfep_path = get_path_from_config('cargolift_pfep_terms', 'base_matriz', q)
        if cargolift_pfep_path and os.path.exists(cargolift_pfep_path):
            shutil.copy2(cargolift_pfep_path, caminho_pasta_hoje)
            
        # Cargolift SP - Suppliers
        cargolift_sup_path = get_path_from_config('cargolift_supplier_terms', 'base_matriz', q)
        if cargolift_sup_path and os.path.exists(cargolift_sup_path):
            shutil.copy2(cargolift_sup_path, caminho_pasta_hoje)
            
        # 2. Mover arquivos da Demanda para Entradas (limpeza)
        if q: q.put(("status", "Limpando pasta Demanda..."))
        pasta_demanda = os.path.join(caminho_base, CONFIG['paths']['folders']['base_demanda'])
        if os.path.exists(pasta_demanda):
            for nome_arquivo in os.listdir(pasta_demanda):
                caminho_arquivo = os.path.join(pasta_demanda, nome_arquivo)
                if os.path.isfile(caminho_arquivo):
                    shutil.move(caminho_arquivo, os.path.join(caminho_entradas_hoje, nome_arquivo))
                    
        # 3. Mover arquivos de Planilhas_Recebidos para Entradas (limpeza)
        if q: q.put(("status", "Limpando pasta Planilhas_Recebidos..."))
        pasta_recebidos = os.path.join(caminho_base, CONFIG['paths']['folders']['base_planilhas_recebidos'])
        if os.path.exists(pasta_recebidos):
            for nome_arquivo in os.listdir(pasta_recebidos):
                caminho_arquivo = os.path.join(pasta_recebidos, nome_arquivo)
                if os.path.isfile(caminho_arquivo) and not nome_arquivo.startswith('~$'):
                    shutil.move(caminho_arquivo, os.path.join(caminho_entradas_hoje, nome_arquivo))
                    
        if q: q.put(("status", f"Arquivamento concluído com sucesso em: {nome_pasta_diaria}"))
        
    except Exception as e:
        if q: q.put(("status", f"ERRO durante o arquivamento/limpeza: {e}"))
        print(f"Erro no limpar_e_arquivar: {e}")



def Processar_Embalagens() : 
    print("getting to the later paert of the code...........")
















