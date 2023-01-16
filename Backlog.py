# import os
import time
# import getpass
# import winsound
# import pandas as pd
# from datetime import datetime, timedelta
from selenium import webdriver
import selenium.webdriver.chrome.options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

chrome_options = selenium.webdriver.chrome.options.Options()
chrome_options.add_experimental_option( "debuggerAddress","127.0.0.1:9222" )
chrome_driver = r'C:\Atlas\cfg\chromedriver.exe'
driver = webdriver.Chrome( chrome_driver,chrome_options = chrome_options )
wait = WebDriverWait( driver,10 )

## cd C:\Program Files\Google\Chrome\Application

## cd C:\Program Files (x86)\Google\Chrome\Application
## chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\Py\ChromeProfile
## https://chromedriver.chromium.org/downloads


print( 'Nome da guia:',driver.title )
guia_poss = driver.window_handles[ 0 ]

# REFERENTE AO PORTAL OSS
poss_xp_clica_filtro = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/div/select'
poss_xp_filtro_op4 = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/div/select/option[4]'
poss_xp_filtro_op11 = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/div/select/option[11]'
poss_xp_col_abertura = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/app-ofrgrid2/div/div[2]/div[1]/div[1]/app-ofrgrid2-column[16]/div/div'
poss_xp_col_abertura_ordenacrescente = '/html/body/div[2]/div[2]/div/div/div/div/button[3]/span'
poss_css_primeira_linha = 'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div > div > div > div'
poss_css_col_primeiro_inc = 'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child(1) > div > div > div > div:nth-child(18)'
poss_css_col_outage = 'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child(1) > div > div > div > div:nth-child(23)'
poss_css_col_tipo = 'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child(1) > div > div > div > div:nth-child(29)'
poss_css_col_status_tarefa = 'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child(1) > div > div > div:nth-child(1) > div:nth-child(3)'
poss_test = 'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child(1) > div > div > div'

# REFERENTE AO INC ABERTO - categorização basica
inc_xp_principal_ct3_produto = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[8]/ofr-field/div/mat-form-field[4]/div/div[1]/div/mat-select/div/div[2]'
inc_xp_principal_ct3_produto_node = '/html/body/div[2]/div[2]/div/div/mat-option[2]/span'
inc_xp_principal_grupo = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[14]/ofr-field/div/mat-form-field[2]/div/div[1]/div/mat-select/div/div[2]'
inc_xp_principal_grupo_preenche = '/html/body/div[2]/div[2]/div/div/mat-option[5]/span'
inc_xp_principal_usuario = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[14]/ofr-field/div/mat-form-field[3]/div/div[1]/div/mat-select/div/div[2]'
inc_xp_principal_usuario_preenche = '/html/body/div[2]/div[2]/div/div/mat-option[1]/span'

# REFERENTE AO INC ABERTO - fechamento
inc_xp_abre_status_para_fechar = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[2]/ofr-field[2]/mat-form-field/div/div[1]/div/mat-select/div/div[2]'
inc_xp_abre_status_para_resolvido = '/html/body/div[2]/div[2]/div/div/mat-option[7]/span'
inc_xp_aba_fechamento = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/mat-tab-header/div[2]/div/div/div[6]'
inc_xp_aba_fechamento_input_data = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[6]/div/div/div[1]/ofr-field/app-ofrfield-date-time/mat-form-field/div/div[1]/div[1]/input'
inc_xp_aba_principal = '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/mat-tab-header/div[2]/div/div/div[1]'

# REFERENTE AO NEWMONITOR
otg_xp_ntciket = '//*[@id="divTicket"]/table/tbody/tr[2]/td[2]'
otg_xp_status = '//*[@id="divTicket"]/table/tbody/tr[11]/td[2]'
otg_xp_data_abertura = '//*[@id="divInf"]/table/tbody/tr[3]/td[2]'
otg_xp_data_fechamento = '//*[@id="divInf"]/table/tbody/tr[4]/td[2]'
otg_xp_botao_fechamento = '//*[@id="divDados"]/table/tbody/tr[1]/td[4]/a'
otg_xp_solucao = '//*[@id="divComp"]/table/tbody/tr[4]/td[2]'
otg_xp_msg_fechamento = '//*[@id="divComp"]/table/tbody/tr[6]/td[2]'

# PEGAR VALORES DE CATEGORIZAÇÃO E INFOS PARA SEGUIR COM FECHAMENTO DO INC - TODOS EM CSS SELECTOR
poss_status_tarefa = ''
poss_tarefa_principal = ''

inc_linha = 1
inc_sub_linha = 1

# INC, N-OTG, TIPO, ABERTURA, NODE
poss_col_base = (18,23,29,32,33)
poss_col_base_coletada = [ ]
# STATUS, PRINCIPAL - CATEGORIZAÇÃO DO FECHAMENTO
poss_col_ref_tarefa = (3,4,8,9,10,11,12,13,14,15,16)
poss_col_ref_tarefa_coletada = [ ]
# PARTE DE BAIXO - CATEGORIZAÇÃO DO INC
poss_col_ref_inc = (20,21,22,25,26,27,28,29,30,31)
poss_col_ref_inc_coletado = [ ]

# DICIONARIO PARA FECHAMENTOS DO OUTAGE:
#
dicionario_otg_fechamento = {'AGRUPAMENTO DE OUTAGES': 'CANCELAR',
                             'ABERTURA EM DUPLICIDADE': 'CANCELAR',
                             'ERRO DE ABERTURA': 'CANCELAR',
                             'FALTA DE ENERGIA NAS RESIDÊNCIAS': (
                             'RESIDENCIAL','REDE COAXIAL','FONTE','CONCESSIONARIA','INTERROMPIDO','QUEDA DE ENERGIA',
                             'FORNECIMENTO RESTABELECIDO','N/A','N/A'),
                             'QUEDA DE ENERGIA - FORNECIMENTO REESTABELECIDO': (
                             'RESIDENCIAL','REDE COAXIAL','FONTE','CONCESSIONARIA','INTERROMPIDO','QUEDA DE ENERGIA',
                             'FORNECIMENTO RESTABELECIDO','N/A','N/A'),
                             'PROBLEMAS DE SEGURANÇA PÚBLICA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','REDE EXTERNA','CHECKLIST','CHECKLIST','REALIZADO',
                             'N/A','N/A'),
                             'ATIVO EQUALIZAÇÃO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','ATIVO','DEFEITO','NIVEIS FORA DO PADRAO','EQUALIZADO',
                             'N/A','N/A'),
                             'PROBLEMA INTERNO - CLIENTE': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','INFRA DE REDE','CHECKLIST','CHECKLIST','REALIZADO',
                             'N/A','N/A'),
                             'ATIVO MODULO COM DEFEITO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','ATIVO','DEFEITO','MODULO DE ATIVO','SUBSTITUIDO',
                             'N/A','N/A'),
                             'PASSIVO COM ÁGUA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','PASSIVO','DANIFICADO','COM AGUA','SUBSTITUIDO',
                             'PASSIVO','N/A'),
                             'TX / RX COM DEFEITO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'INGRESSO DE RUÍDO - NORMALIZADO DURANTE IDENTIFICAÇÃO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','RUIDO','OFENSOR NAO IDENTIFICADO','N/A',
                             'NORMALIZADO DURANTE IDENTIFICACAO','N/A','N/A'),
                             'DISJUNTOR DESARMADO': (
                             'RESIDENCIAL','REDE COAXIAL','FONTE','DISJUNTOR','DESARMADO','N/A','SUBSTITUIDO','N/A',
                             'N/A'),
                             'ADEQUAÇÃO DE REDE COAXIAL': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'INGRESSO DE RUÍDO - LOGRADOURO FILTRADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ATIVO MODULO DE RETORNO COM DEFEITO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'PASSIVO DANFICADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'LIMPEZA DE RUÍDO - MANUTENÇÃO REDE COAXIAL': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO OPTICO - DANO PARCIAL': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'LPI DANIFICADA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'CABO COAXIAL DANIFICADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'LIMPEZA DE RUÍDO - LOGRADOURO IDENTIFICADO SEM FILTRO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ATIVO QUEIMADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','ATIVO','DEFEITO','QUEIMADO','SUBSTITUIDO','ATIVO',
                             'N/A'),
                             'OLT HW SUBSTITUIDO PARTE COM DEFEITO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'CAIXA DE EMENDA DANIFICADA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'PASSIVO QUEIMADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'INGRESSO DE RUÍDO - NORMALIZADO SEM INTERVENÇÃO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'EQUALIZACAO DE RF/QAM': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'REFEITA CONEXÃO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ATIVO COM MAU CONTATO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'CABOS DROP ROMPIDO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'INGRESSO DE RUÍDO - LOGRADOURO ATENUADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'FECHAMENTO DE FUSÕES': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ATIVO COM ÁGUA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'QUEDA DE ENERGIA - ACIONADO GERADOR': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO OPTICO - PODA DE ARVORE': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'FUSÍVEL DE SAÍDA DA FONTE QUEIMADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'CMTS AP SATURAÇÃO DE PORTADORA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'CMTS AP ALTERADO CONFIGURAÇÃO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'FUSÍVEL QUEIMADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO  COAXIAL - OBRAS PÚBLICAS': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'NORMALIZADO SEM INTERVENÇÃO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'MODULO FONTE DANIFICADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'CABO DROP': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'SPI INTERNA FONTE DANIFICADA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'FIBRA OPTICA ATENUADA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO  COAXIAL - PODE DE ÁRVORE': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'RECEPTOR ÓPTICO PLACA OIB (GS7000)': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO COAXIAL - ACIDENTE DE TRANSITO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'SUBSTITUIÇÃO DE NAP': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'RECEPTOR ÓPTICO CABO RF  (GS7000)': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'INGRESSO DE RUÍDO - ALTERADO FREQ/MODULAÇÃ	': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'MANOBRA ÓPTICA - MELHORIA DE REDE': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'MODULO INVERSOR DANIFICADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO  COAXIAL - CABO BANDOLADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'QUEDA DE ENERGIA - JANELA DE SINCRONISMO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'SRV AP ALTERADO CONFIGURAÇÃO DHCP': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO OPTICO - VANDALISMO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO OPTICO - CABO QUEIMADO': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'MANOBRA ÓPTICA - TESTE  COMUTAÇÃO COLETORES': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO OPTICO - CARGA ALTA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO ÓPTICO - TROCA DE POSTE / EMERGENCIAL': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),
                             'ROMPIMENTO CABO COAXIAL - CARGA ALTA': (
                             'RESIDENCIAL','REDE COAXIAL','NODE','XXXXX','XXXXXX','XXXXXX','XXXXXX','N/A','N/A'),

                             }


# FUNÇÃO PARA ABRIR FILTRO DO PORTAL E "ATUALIZAR"
def func_atualizalista():
    #
    # JÁ DEIXA EM ABERTO AS "VISUALIZAÇÕES"
    #
    # OK ABRE A LISTA DE FILTRO, CLICA NO FILTRO BACKLOG E FECHA A LISTA
    WebDriverWait( driver,20 ).until( EC.presence_of_element_located( (By.XPATH,poss_xp_clica_filtro) ) ).click()
    WebDriverWait( driver,20 ).until( EC.presence_of_element_located( (By.XPATH,poss_xp_filtro_op4) ) ).click()
    WebDriverWait( driver,20 ).until( EC.presence_of_element_located( (By.XPATH,poss_xp_clica_filtro) ) ).click()
    time.sleep( 2 )
    WebDriverWait( driver,20 ).until( EC.presence_of_element_located( (By.XPATH,poss_xp_clica_filtro) ) ).click()
    WebDriverWait( driver,20 ).until( EC.presence_of_element_located( (By.XPATH,poss_xp_filtro_op11) ) ).click()
    WebDriverWait( driver,20 ).until( EC.presence_of_element_located( (By.XPATH,poss_xp_clica_filtro) ) ).click()

    val_primeira_linha = False
    cont_val_primeira_linha = 0
    # AGUARDAR ATUALIZAÇÃO - ESPERA APARECER A PRIMEIRA LINHA
    while val_primeira_linha == False:
        time.sleep( 1 )
        cont_val_primeira_linha = cont_val_primeira_linha + 1
        try:
            val_primeira_linha = driver.find_element( By.CSS_SELECTOR,
                                                      'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div > div > div > div' ).is_enabled
            print( 'Aguardando atualizar',cont_val_primeira_linha )

        except:
            print( 'Aguardando atualizar',cont_val_primeira_linha )
            if cont_val_primeira_linha == 3:
                cont_val_primeira_linha = 0
                print( 'Demorou muito, puxando filtro novamente' )
                func_atualizalista()


# FUNÇÃO PARA ABRIR INCIDENTE EM NOVA GUIA E TROCAR PRA ELA
#
def func_abre_inc_em_nova_guia( inc_abrir ):
    global guia_inc
    ActionChains( driver ) \
        .key_down( Keys.CONTROL ) \
        .click( inc_abrir ) \
        .key_up( Keys.CONTROL ) \
        .perform()
    time.sleep( 2 )
    guia_inc = driver.window_handles[ 1 ]
    driver.switch_to.window( guia_inc )


##OK RETORNA A QTD DE SUB LINHA, DE UMA DETERMINADA LINHA
def func_check_qtd_tarefas( inc_sub_linha,inc_linha ):
    try:
        while driver.find_element( By.XPATH,
                                   f'/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/app-ofrgrid2/div/div[2]/app-ofrgrid2-tabular-rows/div/div[{inc_linha}]/div/div/div[{inc_sub_linha}]' ).is_enabled() == True:
            inc_sub_linha = inc_sub_linha + 1  # /html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/app-ofrgrid2/div/div[2]/app-ofrgrid2-tabular-rows/div/div[]/div/div/div[]
    except:
        return inc_sub_linha - 1


# FUNÇÃO QUE COLETA AS INFOS DAS COLUNAS E RETORNA EM UMA LISTA - TBM SCROLL ATE AS COLUNAS
def func_coleta_infos( colunas ):
    global inc_linha,inc_sub_linha
    coleta = [ ]
    for col in colunas:
        scroll_col = driver.find_element( By.CSS_SELECTOR,
                                          f'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child({inc_linha}) > div > div > div:nth-child({inc_sub_linha}) > div:nth-child({col})' )
        scroll_col.location_once_scrolled_into_view
        col_coletada = driver.find_element( By.CSS_SELECTOR,
                                            f'body > app-root > mat-drawer-container > mat-drawer-content > div > app-ofrview > div > app-ofrgrid2 > div > div.grid2content > app-ofrgrid2-tabular-rows > div > div:nth-child({inc_linha}) > div > div > div:nth-child({inc_sub_linha}) > div:nth-child({col})' ).text
        if col_coletada == '':
            col_coletada = 'VAZIO'
        coleta.append( col_coletada )
    return coleta


# FUNÇÃO PARA VERIFICAR SE TAREFAS ESTÃO FECHADAS
def func_verifica_se_tem_tarefa_em_aberto( qtd_sub_linha ):
    check_status_tarefa = 1
    print( 'Verificando se há tarefas em aberto' )
    while check_status_tarefa <= qtd_sub_linha:
        check_tarefas = driver.find_element( By.XPATH,
                                             f'/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/app-ofrgrid2/div/div[2]/app-ofrgrid2-tabular-rows/div/div[{inc_linha}]/div/div/div[{check_status_tarefa}]/div[3]' ).text
        print( 'Verificando tarefa',check_status_tarefa,check_tarefas )
        check_status_tarefa = check_status_tarefa + 1
        if check_tarefas in [ "Cancelado","Fechado","" ]:
            print( 'Não consta tarefas em aberto' )
        else:
            print( 'Consta tarefas em aberto, verificar próximo incidente????' )
            return False
    print( 'Fim da verificação de tarefas em aberto' )
    return True


# FUNÇÃO PARA VERIFICAR SE INC ABRIU EM NOVA GUIA
def func_verifica_se_inc_abriu():
    n = 0
    print( 'Verificando se nova guia do incidente abriu corretamente' )
    while True:
        time.sleep( 2 )
        n = n + 1
        try:
            if driver.find_element( By.XPATH,
                                    '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/mat-tab-header/div[2]/div/div/div[1]' ).is_enabled() == True:
                print( 'Nova guia do incidente abriu corretamente!!' )
                return False
        except:
            if n == 5:
                driver.refresh()
                print( 'Possivelmente não carregou, atualizando pagina',n )
                n = 0
            else:
                pass


# OK FUNÇÃO PRA ABRIR O OTG EM NOVA GUIA
def func_abrir_otg_em_novaguia( n_otg ):
    poss_verificar_otg_do_inc = 'link da ferramenta' + n_otg
    print( '\nVerificando ticket:',n_otg )
    driver.execute_script( f"window.open('{poss_verificar_otg_do_inc}','_blank');" )
    guia_otg = driver.window_handles[ 2 ]
    driver.switch_to.window( guia_otg )

    time.sleep( 1 )


# OK FUNÇÃO QUE VERIFICA SE ABRIU O OTG MESMO
def func_verifica_se_guia_otg_carregou( numeroticket ):
    n = 0
    print( 'Verificando se nova guia do newmonitor carregou completamente' )
    while True:
        time.sleep( 2 )
        n = n + 1
        try:
            if driver.find_element( By.XPATH,otg_xp_ntciket ).text == numeroticket:
                print( 'Nova guia do newmonitor carregou corretamente!!' )
                return False
        except:
            if n == 5:
                print( 'Possivelmente não carregou, atualizando pagina',n )
                driver.refresh()
                if driver.title == 'Erro de privacidade':
                    driver.find_element( By.XPATH,'/html/body/div/div[2]/button[3]' ).click()
                    WebDriverWait( driver,5 ).until(
                        EC.presence_of_element_located( (By.XPATH,'/html/body/div/div[3]/p[2]/a') ) ).click()

                n = 0
            else:
                pass


# PRIMEIRO O 1 (SUBLINHA, DEPOIS A LINHA)


while True:
    func_atualizalista()
    inc_linha = 1
    inc_sub_linha = 1

    while inc_linha <= 20:

        print( 'Atualizou!!!' )

        # PEGA AS INFORMAÇÕES BASE
        poss_col_base_coletada = func_coleta_infos( poss_col_base )
        print( 'Eis a coleta base do incidente:',poss_col_base_coletada )

        # ISSO AQUI SAI DEPOIS
        poss_col_ref_tarefa_coletada = func_coleta_infos( poss_col_ref_tarefa )
        print( 'Eis a coleta da tarefa',poss_col_ref_tarefa_coletada )

        if poss_col_base_coletada[ 2 ] == 'OUTAGE' and poss_col_base_coletada[ 4 ] != 'VAZIO':
            # VERIFICAR SE TEM TAREFA EM ABERTO
            qtd_sub_linha = func_check_qtd_tarefas( inc_sub_linha,inc_linha )
            print( 'Eis a quantidade de tarefa do incidente:',qtd_sub_linha )
            scroll_col_status_tarefas = driver.find_element( By.CSS_SELECTOR,poss_css_col_status_tarefa )
            scroll_col_status_tarefas.location_once_scrolled_into_view

            # SE NÃO TIVER TAREFA EM ABERTO: ABRIR INC EM NOVA GUIA, TAMBÉM COLETA AS INFOS PARA CORRIGIR CATEGORIZAÇÃO
            if func_verifica_se_tem_tarefa_em_aberto( qtd_sub_linha ) == True:
                print( 'Tarefas fechadas ou canceladas, abrindo incidente e verificando outage' )
                # PEGAR AS INFOS DA CATEGORIZAÇÃO E DATA
                poss_col_ref_inc_coletado = func_coleta_infos( poss_col_ref_inc )
                print( 'eis as infos do inc',poss_col_ref_inc_coletado )
                # ABRE INC EM NOVA GUIA
                func_abre_inc_em_nova_guia( driver.find_element( By.XPATH,
                                                                 f'/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/app-ofrgrid2/div/div[2]/app-ofrgrid2-tabular-rows/div/div[{inc_linha}]' ) )
                func_verifica_se_inc_abriu()
                time.sleep( 1 )
                # VERIFICA CATEGORIZAÇÃO NODE E TOTAL:
                # VERIFICA CATEGORIZAÇÃO NODE E TOTAL:
                # VERIFICA CATEGORIZAÇÃO NODE E TOTAL:
                # VERIFICA CATEGORIZAÇÃO NODE E TOTAL:
                # VERIFICA CATEGORIZAÇÃO NODE E TOTAL: (OBS A COLUNA NODE VEM VAZIA, SERIA INTERESSANTE FUTURAMENTE FAZER UMA FUNÇÃO)
                driver.find_element( By.XPATH,
                                     '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[8]/ofr-field/div/mat-form-field[7]/div/div[1]/div/mat-select/div/div[1]' ).location_once_scrolled_into_view
                if driver.find_element( By.XPATH,
                                        '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[8]/ofr-field/div/mat-form-field[7]/div/div[1]/div/mat-select/div/div[1]/span' ).text == 'Categorização Operacional 3':
                    time.sleep( 0.5 )
                    driver.find_element( By.XPATH,
                                         '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[8]/ofr-field/div/mat-form-field[4]/div/div[1]/div/mat-select/div/div[2]' ).click()
                    WebDriverWait( driver,10 ).until( EC.presence_of_element_located(
                        (By.XPATH,'/html/body/div[2]/div[2]/div/div/mat-option[2]/span') ) ).click()

                if driver.find_element( By.XPATH,
                                        '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[8]/ofr-field/div/mat-form-field[7]/div/div[1]/div/mat-select/div/div[1]/span' ).text == '':
                    print( 'Acho que não consegui corrigir a categorização do outage que ficou pendente' )

                print( 'eis se corrigou a categorizacao 3',driver.find_element( By.XPATH,
                                                                                '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[8]/ofr-field/div/mat-form-field[7]/div/div[1]/div/mat-select/div/div[1]/span' ).text )

                if poss_col_ref_inc_coletado[ 1 ] == 'VAZIO':
                    print( 'Categorização de "Grupo" pendente (aba principal do incidente)' )
                    driver.find_element( By.XPATH,inc_xp_principal_grupo ).location_once_scrolled_into_view
                    driver.find_element( By.XPATH,inc_xp_principal_grupo ).click()
                    time.sleep( 0.5 )
                    driver.find_element( By.XPATH,inc_xp_principal_grupo_preenche ).click()
                if poss_col_ref_inc_coletado[ 2 ] == 'VAZIO':
                    print( 'Categorização de "Usuario" pendente (aba principal do incidente)' )
                    driver.find_element( By.XPATH,inc_xp_principal_usuario ).location_once_scrolled_into_view
                    driver.find_element( By.XPATH,inc_xp_principal_usuario ).click()
                    time.sleep( 0.5 )
                    driver.find_element( By.XPATH,inc_xp_principal_usuario_preenche ).click()

                func_abrir_otg_em_novaguia( poss_col_base_coletada[ 1 ] )
                # FUNÇÃO PRA VERIFICAR SE OTG ABRIU, MANDA O NUMERO DO TICKET
                func_verifica_se_guia_otg_carregou( poss_col_base_coletada[ 1 ] )

                # COLETA AS INFOS PARA FECHAMENTO
                if driver.find_element( By.XPATH,otg_xp_status ).text == 'Fechado':
                    print( 'Outage com Status FECHADO, coletando as informações e posteriomente fechando a guia' )
                    driver.find_element( By.XPATH,otg_xp_botao_fechamento ).click()
                    time.sleep( 1 )
                    otg_solucao = driver.find_element( By.XPATH,otg_xp_solucao ).text
                    otg_msg_fechamento = driver.find_element( By.XPATH,otg_xp_msg_fechamento ).text
                    driver.find_element( By.XPATH,otg_xp_data_fechamento ).location_once_scrolled_into_view
                    otg_data_abertura = driver.find_element( By.XPATH,otg_xp_data_abertura ).text
                    otg_data_fechamento = driver.find_element( By.XPATH,otg_xp_data_fechamento ).text
                    print( 'eis a coleta do otg',otg_data_abertura,otg_data_fechamento,otg_msg_fechamento,otg_solucao )
                    print( 'Eis o possivel fechamento',dicionario_otg_fechamento[ otg_solucao ] )
                    driver.close()
                    driver.switch_to.window( guia_inc )

                    if dicionario_otg_fechamento[ otg_solucao ] == 'CANCELAR':
                        driver.find_element( By.XPATH,
                                             '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[2]/ofr-field[2]/mat-form-field/div/div[1]/div/mat-select/div/div[1]' ).location_once_scrolled_into_view
                        driver.find_element( By.XPATH,
                                             '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[1]/div/div/div[2]/ofr-field[2]/mat-form-field/div/div[1]/div/mat-select/div/div[2]' ).click()
                        time.sleep( 0.5 )
                        driver.find_element( By.XPATH,'/html/body/div[2]/div[2]/div/div/mat-option[2]/span' ).click()

                    if len( dicionario_otg_fechamento[ otg_solucao ] ) == 9:
                        print( 'deu certo o len',len( dicionario_otg_fechamento[ otg_solucao ] ) )
                        WebDriverWait( driver,10 ).until( EC.presence_of_element_located(
                            (By.XPATH,inc_xp_aba_principal) ) ).location_once_scrolled_into_view
                        WebDriverWait( driver,10 ).until(
                            EC.presence_of_element_located( (By.XPATH,inc_xp_abre_status_para_fechar) ) ).click()
                        WebDriverWait( driver,10 ).until(
                            EC.presence_of_element_located( (By.XPATH,inc_xp_abre_status_para_resolvido) ) ).click()
                        WebDriverWait( driver,10 ).until(
                            EC.presence_of_element_located( (By.XPATH,inc_xp_aba_fechamento) ) ).click()
                        lista_fechamento = 1
                        for item in dicionario_otg_fechamento[ otg_solucao ]:
                            print( 'Eis cada item dentro do dicionario de fechamento:',item,
                                   'Eis o n da lista fechamento',lista_fechamento )
                            time.sleep( 1 )
                            while driver.find_element( By.XPATH,
                                                       f'/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[6]/div/div/div[4]/ofr-field/div/mat-form-field[{lista_fechamento}]/div/div[1]/div/mat-select/div/div[1]/span' ).text != item:
                                try:
                                    print( 'passou no try, cima',lista_fechamento,item )
                                    driver.find_element( By.XPATH,
                                                         f'/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[6]/div/div/div[4]/ofr-field/div/mat-form-field[{lista_fechamento}]/div/div[1]/div/mat-select/div/div[2]' ).click()
                                    print( 'passou no try, baixo',lista_fechamento,item )
                                    WebDriverWait( driver,10 ).until( EC.presence_of_element_located( (By.XPATH,
                                                                                                       f"/html/body/div[2]/div[2]/div/div/mat-option/span[normalize-space()= '{item}']") ) ).click()
                                    print( 'baixo do web' )
                                    time.sleep( 0.5 )
                                except:
                                    print( 'passou no escept, cima',lista_fechamento,item )
                                    driver.find_element( By.XPATH,
                                                         f"/html/body/div[2]/div[2]/div/div/mat-option/span[normalize-space()= '{item}']" ).click()

                                    print( 'passou no escept, baixo',lista_fechamento,item )
                                    time.sleep( 0.5 )

                            if lista_fechamento == 9:
                                lista_fechamento = 1
                                break
                            lista_fechamento = lista_fechamento + 1

                    # SALVAR E FECHAR
                    driver.find_element( By.XPATH,
                                         '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-toolbar/button[5]' ).location_once_scrolled_into_view
                    time.sleep( 0.5 )
                    driver.find_element( By.XPATH,
                                         '/html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-toolbar/button[5]' ).click()
                    time.sleep( 4 )
                    driver.close()
                    driver.switch_to.window( guia_poss )



            else:
                print( 'Consta tarefas em aberto, verificando proximo...' )

        else:
            print( 'por alguma motivo não foi possivel resolver esse INC, verificando o proximo' )

            # COLOCAR NESSE ELSE UMA FORMA DE SALVAR O INC EM UMA LISTA DE INC's ABERTOS PARA NÃO VERIFICAR TÃO CEDO

            # ABRIR OUTAGE EM NOVA GUIA
            # VERIFICAR SE OUTAGE ESTA EM ABERTO
            # SE ESTIVER FECHADO, PEGAR AS INFOS DE FECHAMENTO - DATA/HORARIO - FECHAMENTOS - (TALVEZ AS NOTAS DO TICKET)
            # FECHAR GUIA DO OUTAGE
            # IR PARA GUIA DO INC QUE FOI ABERTO
            # VERIFICAR SE A CARACTERIZAÇÃO ESTÁ OK E CORRIGIR
            # MUDAR PARA RESOLVIDO E IR ATÉ O FECHAMENTO
            # PREENCHER O FECHAMENTO E SALVAR E FECHAR A GUIA, E SEGUE O BAILE
            print( 'Tipo Outage' )

        # CHAMA A FUNÇÃO PRA ABRIR O INC EM NOVA GUIA
        # func_abre_inc_em_nova_guia(driver.find_element(By.CSS_SELECTOR, poss_test),1)
        # time.sleep(2)

        # chama a função pra abrir otg em nova GUIA
        # func_abrir_otg_em_novaguia(poss_n_outage)
        # guia_otg = driver.window_handles[2]
        # driver.switch_to.window(guia_otg)

        ## referente ao botao de abrir a lista de fechamento, muda apenas a antepenultima coluna, que vai do 1 ao 9                                                                       #essa \/
        # /html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[6]/div/div/div[4]/ofr-field/div/mat-form-field[1]/div/div[1]/div/mat-select/div/div[2]
        # /html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrdataeditor/div/div/app-ofrform/div/mat-tab-group/div/mat-tab-body[6]/div/div/div[4]/ofr-field/div/mat-form-field[2]/div/div[1]/div/mat-select/div/div[2]

        # possinc = driver.window_handles[1]

        # driver.switch_to.window(possinc)

        # time.sleep(2)
        # driver.switch_to.window(poss)
        # time.sleep(2)
        # abre_inc_em_nova_guia(inc_em_nova_guia)
        # time.sleep(2)
        # driver.switch_to.window(possinc)

        # OK - não precisa, pois, a cfg do filtro já ordena sozinho - SCROLL ATÉ A COLUNA 16, CLICA NA COLUNA ABERTURA E ORDENA EM CRESCENTE
        # scroll_col_16 = driver.find_element(By.XPATH, xp_col_abertura)
        # scroll_col_16.location_once_scrolled_into_view
        # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xp_col_abertura))).click()
        # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xp_col_abertura_ordenacrescente))).click()
        # OK - SCROLL ATÉ A COLUNA DOS INC's

        print( 'fim' )
        inc_linha = inc_linha + 1

        # REF. A LISTA                          #REF. A LINHA    #REF. SUB LINHA   #REF COLUNA
        # /html/body/app-root/mat-drawer-container/mat-drawer-content/div/app-ofrview/div/app-ofrgrid2/div/div[2]/app-ofrgrid2-tabular-rows/div/div[8]/div/div/div[1]/div[18]

        # A IDEIA É JOGAR DENTRO DA FUNÇÃO QUE COLETA AS INFOS A FUNÇÃO QUE VERIFICA SE TEM TAREFA EM ABERTO E SUBLINHAS, COM BASE EM LISTAS E COLETAR COLUNA POR COLUNA
        # talvez mudar da coluna otg pra quando a coluna ticket nao for vazia??