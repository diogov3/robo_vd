from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
from datetime import date
from time import sleep
from pathlib import Path
import os

#VARIÁVEIS
chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--disable-dev-shm-usage')
navegador = webdriver.Chrome(executable_path=r"C:\Users\v3c.suporte\Downloads\chromedriver.exe",chrome_options=chrome_options)
w = WebDriverWait(navegador, 30)
w2 = WebDriverWait(navegador, 1000)
url = 'https://sgi.e-boticario.com.br/Paginas/Acesso/Entrar.aspx?ReturnUrl=%2fDefault.aspx'
usuario = '543161'
senha = 'quasar10'
oldAdress = 'C:/Users/v3c.suporte/Downloads' #pasta origem
newAdress = 'C:/Users/v3c.suporte/Grupo Enseada/Comercial - Documentos/General/1.2. Gestão Comercial/Business Intelligence/Venda Direta/Planilhas VD robo/' #pasta destino
codEstrutura = ['2285','10821','4097']
uf = ['_RJ','_SF','_ES']

#CICLO ATUAL
hoje = date.today()
path  = r"C:\Users\v3c.suporte\Grupo Enseada\Comercial - Documentos\General\1.2. Gestão Comercial\Business Intelligence\Venda Direta\Outros\Ciclos.xlsx"
data = pd.read_excel(path,sheet_name="Data")
df = pd.DataFrame(data)
tabela = df[df['Data'] == str(hoje)]
valor = tabela.values[0][1]
ciclo = valor[6:]
ciclo_ponto = valor.replace('/','.')

#TROCAR E FECHAR SEGUNDA JANELA
def janela():
    handles = []
    handles = navegador.window_handles
    for handle in handles:
        print(handle)
    newHandle = handles[1]
    oldHandle = handles[0]
    navegador.switch_to.window(newHandle)
    navegador.close()
    navegador.switch_to.window(oldHandle)

#FAZER LOGIN
navegador.get(url)
navegador.maximize_window()
w2.until(EC.element_to_be_clickable((By.ID,"username"))).send_keys(usuario)
w2.until(EC.element_to_be_clickable((By.ID,"password"))).send_keys(senha)
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="btnLogin"]/span[1]'))).click()

# INPUT BASE
for x in codEstrutura:
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menu-cod-6"]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-6"]/div/div[1]/ul/li[1]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-6"]/div/div[1]/ul/li[1]/ul/li[1]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_cmsSituacaoComercial_Tb1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_cmsSituacaoComercial_D0"]/div/ul/li[1]/a/label'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_cmsSituacaoComercial_D0"]/div/ul/li[2]/a/label'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_txtEstruturaComercialCodigo_Tb1"]'))).send_keys(x)
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlCicloAtualInicial_Ddl1"]'))).send_keys(ciclo)
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlCicloAtualFinal_Ddl1"]'))).send_keys(ciclo)
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_btnBuscar_B1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_btnExportar_B1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_exportarPanel"]/div/div/div[2]/div[1]/div/div/div[2]/ul/li[2]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_colunasTab"]/div[1]/div/div/label'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_okButton_B1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="popupOkButton"]'))).click()
    sleep(7)
    janela()

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
for y in uf:
    lista_arquivos = os.listdir(oldAdress)
    data_criacao = lambda f: f.stat().st_ctime
    directory = Path(oldAdress)
    files = directory.glob('*.xls')
    sorted_files = sorted(files, key=data_criacao, reverse=True)
    ultimo_arquivo = sorted_files[0]
    novo_nome = 'C:/Users/v3c.suporte/Downloads/'+ciclo_ponto+y+'_Base.xls'
    novo_nome_dir = newAdress+ciclo_ponto+y+'_Base.xls'
    os.rename(ultimo_arquivo,novo_nome)
    if os.path.isfile(novo_nome_dir):
        os.replace(novo_nome, novo_nome_dir) 
    else:
        os.rename(novo_nome, novo_nome_dir)

#INPUT LOJA
for x in codEstrutura:
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menu-cod-1"]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-1"]/div/div[1]/ul/li[1]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-1"]/div/div[1]/ul/li[1]/ul/li[2]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_txtDataInicio_Tb1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="bodyControl"]/div[2]/table[5]/tbody/tr/td[2]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_txtEstruturaComercialCodigo_Tb1"]'))).send_keys(x)
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_ddlCicloInicio_Ddl1"]'))).send_keys(ciclo)
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_ddlCicloFim_Ddl1"]'))).send_keys(ciclo)
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_cmsSituacaoFiscal_Tb1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_cmsSituacaoFiscal_D0"]/div/ul/li[2]/a/label'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_btnBuscar_B1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ControleBuscaPedido_btnExportar_B1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_exportarPanel"]/div/div/div[2]/div[1]/div/div/div[2]/ul/li[2]/a'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_colunasTab"]/div[1]/div/div/label'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_okButton_B1"]'))).click()
    w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="popupOkButton"]'))).click()
    sleep(7)
    janela()

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
for y in uf:
    lista_arquivos = os.listdir(oldAdress)
    data_criacao = lambda f: f.stat().st_ctime
    directory = Path(oldAdress)
    files = directory.glob('*.xls')
    sorted_files = sorted(files, key=data_criacao, reverse=True)
    ultimo_arquivo = sorted_files[0]    
    novo_nome = oldAdress+ciclo_ponto+y+'_Loja.xls'
    novo_nome_dir = newAdress+ciclo_ponto+y+'_Loja.xls'
    os.rename(ultimo_arquivo,novo_nome)
    if os.path.isfile(novo_nome_dir):
        os.replace(novo_nome, novo_nome_dir) 
    else:
        os.rename(novo_nome, novo_nome_dir)

# INICIO E REINICIO

# EXPORTA INICIO ES
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menu-cod-4"]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-4"]/div/div[1]/ul/li[1]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-4"]/div/div[1]/ul/li[1]/ul/li[3]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_assunto_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_assunto_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_assunto_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_assunto_d1"]/option[4]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_repetidorLinha_repetidorColuna_0_medidorIndicador_4_divSemMeta2_4"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]/option[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_pesquisarButton_btn"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_botaoExportar_btn"]'))).click()
sleep(5)
janela()
sleep(2)

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_Inicio_ES.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_Inicio_ES.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)
    

# EXPORTA REINICIO ES
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_painelDetalhes_listBoxIndicador"]/option[12]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]/option[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_pesquisarButton_btn"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_botaoExportar_btn"]'))).click()
sleep(5)
janela()
sleep(2)

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_Reinicio_ES.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_Reinicio_ES.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# VOLTA E TROCA A ESTRUTURA
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_painelDetalhes_botaoVoltar_btn"]'))).click()

w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_seletorEstrutura_botao"]'))).click()

# SELECIONA O RJ
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_seletorEstrutura_arvore"]/ul/li[2]/span/a/span'))).click()
sleep(1)

# EXPORTA O INICIO RJ
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_repetidorLinha_repetidorColuna_0_medidorIndicador_4_divSemMeta2_4"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]/option[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_pesquisarButton_btn"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_botaoExportar_btn"]'))).click()
sleep(5)
janela()
sleep(2)

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_Inicio_RJ.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_Inicio_RJ.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# EXPORTA O REINICIO RJ
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_painelDetalhes_listBoxIndicador"]/option[12]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]/option[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_pesquisarButton_btn"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_botaoExportar_btn"]'))).click()
sleep(5)
janela()
sleep(2)

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_Reinicio_RJ.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_Reinicio_RJ.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# VOLTA E TROCA A ESTRUTURA
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_painelDetalhes_botaoVoltar_btn"]'))).click()

w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_seletorEstrutura_botao"]'))).click()

# SELECIONA O SF
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_seletorEstrutura_arvore"]/ul/li[3]/span/a/span'))).click()
sleep(1)

# EXPORTA O INICIO SF
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_repetidorLinha_repetidorColuna_0_medidorIndicador_4_divSemMeta2_4"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]/option[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_pesquisarButton_btn"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_botaoExportar_btn"]'))).click()
sleep(5)
janela()
sleep(2)

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_Inicio_SF.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_Inicio_SF.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# EXPORTA O REINICIO SF
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_painelDetalhes_listBoxIndicador"]/option[12]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ddlBaseConsulta_d1"]/option[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_pesquisarButton_btn"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_botaoExportar_btn"]'))).click()
sleep(5)
janela()
sleep(2)

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_Reinicio_SF.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_Reinicio_SF.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# EXPORTAR QDB ES
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menu-cod-1"]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-1"]/div/div[1]/ul/li[5]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="submenu-cod-1"]/div/div[1]/ul/li[5]/ul/li[2]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_txtCodigoCD_Tb1"]'))).send_keys('850')
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_ddlSituacaoFiscal_Tb1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_ddlSituacaoFiscal_D0"]/div/ul/li[2]/a/label'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_txtDataInicialCaptacao_Tb1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="bodyControl"]/div[4]/table[5]/tbody/tr/td[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_txtDataFinalCaptacao_Tb1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="bodyControl"]/div[5]/table[5]/tbody/tr/td[2]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_ddlCicloInicialCaptacao_Ddl1"]'))).send_keys(ciclo)
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_ddlCicloFinalCaptacao_Ddl1"]'))).send_keys(ciclo)
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_btnBuscar_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_btnExportar_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_exportarPanel"]/div/div/div[2]/div[1]/div/div/div[2]/ul/li[2]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_colunasTab"]/div[1]/div/div/label'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_okButton_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="popupOkButton"]'))).click()
sleep(7)
janela()

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_QDB_ES.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_QDB_ES.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# EXPORTAR QDB RJ
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_txtCodigoCD_Tb1"]'))).send_keys('924')
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_btnBuscar_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_btnExportar_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_exportarPanel"]/div/div/div[2]/div[1]/div/div/div[2]/ul/li[2]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_colunasTab"]/div[1]/div/div/label'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_okButton_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="popupOkButton"]'))).click()
sleep(7)
janela()

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_QDB_RJ.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_QDB_RJ.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

# EXPORTAR QDB SF
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_txtCodigoCD_Tb1"]'))).send_keys('1166')
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_btnBuscar_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_ctlRelatorioItensPorVendedor_btnExportar_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_exportarPanel"]/div/div/div[2]/div[1]/div/div/div[2]/ul/li[2]/a'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_colunasTab"]/div[1]/div/div/label'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="agendamentoExportacao_okButton_B1"]'))).click()
w2.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="popupOkButton"]'))).click()
sleep(7)
janela()

# IDENTIFICAR O ULTIMO ARQUIVO, TROCAR DE NOME E JOGAR NA PASTA
lista_arquivos = os.listdir(oldAdress)
data_criacao = lambda f: f.stat().st_ctime
directory = Path(oldAdress)
files = directory.glob('*.xls')
sorted_files = sorted(files, key=data_criacao, reverse=True)
ultimo_arquivo = sorted_files[0]

novo_nome = oldAdress+ciclo_ponto+'_QDB_SF.xls'
novo_nome_dir = newAdress+ciclo_ponto+'_QDB_SF.xls'

os.rename(ultimo_arquivo,novo_nome)

if os.path.isfile(novo_nome_dir):
    os.replace(novo_nome, novo_nome_dir) 
else:
    os.rename(novo_nome, novo_nome_dir)

navegador.quit()