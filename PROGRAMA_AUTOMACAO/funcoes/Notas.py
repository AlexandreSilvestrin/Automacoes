import re
import tabula
import os
import pandas as pd
from io import StringIO
import openpyxl
from unidecode import unidecode
import PyPDF2
import warnings

# Ignorar o FutureWarning específico
warnings.filterwarnings("ignore", category=FutureWarning, message="errors='ignore' is deprecated")
def pegardata():
    global NOMETXT
    nome = NOMETXT.replace('.txt', '')
    return nome[-6:-4], nome[-4:]

def criar_pasta(localsalvar):
    if not os.path.exists(f'{localsalvar}/TUDO.xlsx'):
        teste = pd.DataFrame(columns=['Data', 'Número', 'CNPJ/CPF' , 'Vazia1', 'Vazia2','Valor', 'NF Nome', 'Tipo'])
        teste.to_excel(f'{localsalvar}/TUDO.xlsx', index=False)

def listar_pastas(local):
    return os.listdir(local)

def organizerdf(df1, df2):
    for indice, linha in df2.iterrows():
        numero_correspondente = linha['Número']
        
        # Localizar a linha no DataFrame principal com o número correspondente
        indice_inserir = df1[df1['Número'] == numero_correspondente].index[0]
        
        # Inserir a linha do DataFrame secundário na linha seguinte ao DataFrame principal
        df1 = pd.concat([df1.loc[:indice_inserir], linha.to_frame().transpose(), df1.loc[indice_inserir+1:]]).reset_index(drop=True)

    return df1

def verifica_tipo_pdf(pasta, arquivo, local):
    try:
        with open(f'{local}/{pasta}/Serviços Tomados/{arquivo}', 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page_num in range(len(reader.pages)):
                text += reader.pages[page_num].extract_text()
    except:
        with open(f'{local}/{pasta}/{arquivo}', 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page_num in range(len(reader.pages)):
                text += reader.pages[page_num].extract_text()
    if 'S E M     M O V I M E N T O' in text:
        return 'Tipo 3'
    elif 'Notas Fiscais de Serviços' in text:
        return 'Tipo 2'
    elif 'Notas de Entradas de Serviços' in text:
        return 'Tipo 1'
    else:
        return 'Tipo não identificado'

def lertomadostxt(pasta, arquivo_tomados, local):
    # le arquivo'
    try:
        try:
            with open(f'{local}/{pasta}/Serviços Tomados/{arquivo_tomados}', 'r') as entrada:
                dados = entrada.read()
                dados = dados.replace(',', '')
                dados = dados.replace('.', '')
        except:
            try:
                with open(f'{local}/{pasta}/Serv Tomados/{arquivo_tomados}', 'r') as entrada:
                    dados = entrada.read()
                    dados = dados.replace(',', '')
                    dados = dados.replace('.', '')
            except:
                with open(f'{local}/{pasta}/{arquivo_tomados}', 'r') as entrada:
                    dados = entrada.read()
                    dados = dados.replace(',', '')
                    dados = dados.replace('.', '')
    except:
        dadosF = pd.DataFrame(columns=['Data', 'Número', 'Valor', 'ISS Retido', 'CNPJ/CPF'])
        linha_cabecalho = pd.DataFrame({'Data': [pasta], 'Número': ['00'], 'Valor': ['00'], 'CNPJ/CPF': ['00'], 'Tipo': ['00']})
        dadosF = pd.concat([linha_cabecalho, dadosF], ignore_index=True)
        return dadosF
    else:
        # r'-{13}.*?-{13}'  - oque sera removido  {13} quantidade . qualquer caracter * zero ou mais aparicoes ? pega sempre as primeiras aparicoes// re.DOTALL o . detecta a quebra de linha
        linhas_organizadas = re.sub(r'\n {30}', '', dados, flags=re.DOTALL).split('\n')
        linhas_selecionadas = []

        padrao_regex = re.compile(r'empresa(.*?)folha', re.DOTALL | re.IGNORECASE)
        empresa = padrao_regex.search(dados).group(0).replace('Folha', '').strip()
        #linhas_selecionadas.append(empresa)

        padrao_regex = r"\|\s\d+\|\s\w\w\w"
        for linha in linhas_organizadas:
            if re.search(padrao_regex, linha):
                linhas_selecionadas.append(linha)

        linhas_selecionadas =  '\n'.join(linhas_selecionadas)
        dados_io = StringIO(linhas_selecionadas)
        dadosF = pd.read_csv(dados_io, sep='|', header=None, dtype=str)
        dadosF = dadosF[[1,4,5,9,10]]
        dadosF.columns= ['Data', 'Número', 'Valor' , 'ISS Retido', 'CNPJ/CPF']
        # Criar uma linha de cabeçalho vazia
        dadosF['ISS Retido'] = dadosF['ISS Retido'].astype(int)
        dfiss = dadosF[dadosF['ISS Retido'] > 0].copy() 
        if not dfiss.empty:
            dfiss['Tipo'] = 'ISS'
            dfiss = dfiss[['Data', 'Número', 'ISS Retido', 'CNPJ/CPF', 'Tipo']].reset_index(drop=True)
            dfiss.rename(columns={'ISS Retido': 'Valor'}, inplace=True)

            dadosF['Tipo'] = 'Total'
            dadosF = dadosF[['Data', 'Número', 'Valor', 'CNPJ/CPF', 'Tipo']]
            dadosF = organizerdf(dadosF, dfiss)

        linha_cabecalho = pd.DataFrame({'Data': [empresa], 'Número': ['00'], 'Valor': ['00'], 'CNPJ/CPF': ['00'], 'Tipo': ['00']})

        # Concatenar os dois DataFrames
        dadosF = pd.concat([linha_cabecalho, dadosF], ignore_index=True)

        return dadosF
        # salva arquivo

def lerpdf(pasta, arquivo, local):
    tipo = verifica_tipo_pdf(pasta, arquivo, local)
    print(tipo)
    if 'Tipo 3' == tipo:
        dfFINAL = pd.DataFrame()
        return dfFINAL
    elif tipo == 'Tipo 1':
        # Specify the area coordinates (left, top, right, bottom) for extraction
        area = [120.285,12.375,573.705,765.765]

        # Read the PDF and extract data from the specified area on the specified page
        try:
            df_list = tabula.read_pdf(f'{local}/{pasta}/Serviços Tomados/{arquivo}', pages=1, area=area,lattice=True)
        except:
            df_list = tabula.read_pdf(f'{local}/{pasta}/{arquivo}', pages=1, area=area,lattice=True)
        tabela = pd.concat(df_list, ignore_index=True)
        # Print the extracted data
        
        tabela.rename(columns={'Seg\rSocial': 'ValorSS'}, inplace=True)
        tabela.rename(columns={'Unnamed: 0': 'Data'}, inplace=True)
        tabelaF = tabela[['Data', 'Número', 'CNPJ/CPF', 'PIS Retido',  'COFINS Retida', 'CSLL retida', 'IRRF', 'ValorSS']]
        tabelaF = tabelaF.replace(r'\r', ' ', regex=True)
        tabelaF = tabelaF.replace(r',', '', regex=True)
        tabelaF = tabelaF.replace(r'.', '')
        print(pasta)
        tabelaF[['PIS Retido', 'COFINS Retida', 'CSLL retida']] = tabelaF[['PIS Retido', 'COFINS Retida', 'CSLL retida']].apply(lambda x: x.str.split())
        tabelaF['PIS Retido'] =tabelaF['PIS Retido'].apply(lambda x: x[0])
        tabelaF['COFINS Retida'] =tabelaF['COFINS Retida'].apply(lambda x: x[0])
        tabelaF['CSLL retida'] =tabelaF['CSLL retida'].apply(lambda x: x[0])
        tabelaF['CNPJ/CPF'] = tabelaF['CNPJ/CPF'].str.replace(r'[^\d]', '', regex=True)
        tabelaF[['PIS Retido', 'COFINS Retida', 'CSLL retida']] =tabelaF[['PIS Retido', 'COFINS Retida', 'CSLL retida']].apply(lambda x: pd.to_numeric(x))
        tabelaF['Valor'] = tabelaF[['PIS Retido', 'COFINS Retida', 'CSLL retida']].sum(axis=1)
        tabelaF['IRRF'] = tabelaF['IRRF'].astype(int)
        tabelaF['ValorSS'] = tabelaF['ValorSS'].astype(int)
        tabelaF['Data'] = pd.to_datetime(tabelaF['Data'], format='%d/%m/%Y')
        tabelaF['Data'] = tabelaF['Data'].dt.day.astype(str)
        dadosIRRF = tabelaF[tabelaF['IRRF']>0].copy()
        dadosIRRF['Tipo'] = 'IRRF'
        dadosIRRF = dadosIRRF[['Data', 'Número', 'CNPJ/CPF', 'Tipo', 'IRRF',]].reset_index(drop=True)
        dadosIRRF.rename(columns={'IRRF': 'Valor'}, inplace=True)
        dadosRS = tabelaF.copy()
        dadosRS['Tipo'] = 'Retencao Social'
        dadosRS = dadosRS[['Data', 'Número', 'CNPJ/CPF', 'Tipo', 'Valor',]].reset_index(drop=True)
        dadosINSS = tabelaF[tabelaF['ValorSS']>0].copy()
        dadosINSS['Tipo'] = 'INSS'
        dadosINSS = dadosINSS[['Data', 'Número', 'CNPJ/CPF', 'Tipo', 'ValorSS']].reset_index(drop=True)
        dadosINSS.rename(columns={'ValorSS': 'Valor'}, inplace=True)
        dfFINAL = pd.concat([dadosRS, dadosIRRF, dadosINSS], ignore_index=True)
        dfFINAL['Número'] = dfFINAL['Número'].astype(str).str.zfill(10)
    elif tipo == 'Tipo 2':
        area = [123.672,31.476,562.523,779.733]
        try:
            df_list = tabula.read_pdf(f'{local}/{pasta}/Serviços Tomados/{arquivo}', pages=1, area=area,lattice=True)
        except:
            df_list = tabula.read_pdf(f'{local}/{pasta}/{arquivo}', pages=1, area=area,lattice=True)
        tabela = pd.concat(df_list, ignore_index=True)

        tabela.columns = ['Data', 'nada', 'Número', 'CNPJ/CPF', 'nada', 'nadaa', 'asdf', 'PIS', 'COFINS', 'CSLL' ,'IRRF', 'INSS']
        tabela = tabela[['Data', 'Número', 'CNPJ/CPF', 'PIS', 'COFINS', 'CSLL' ,'IRRF', 'INSS']]
        tabela = tabela.replace(r'\r', ' ', regex=True)
        tabela = tabela.replace(r',', '', regex=True)
        tabela = tabela.replace(r'.', '')

        tabela[['PIS', 'COFINS', 'CSLL', 'IRRF', 'INSS']] = tabela[['PIS', 'COFINS', 'CSLL', 'IRRF', 'INSS']].apply(lambda x: x.str.split())
        tabela['PIS'] = tabela['PIS'].apply(lambda x: x[0])
        tabela['COFINS'] = tabela['COFINS'].apply(lambda x: x[0])
        tabela['CSLL'] = tabela['CSLL'].apply(lambda x: x[0])
        tabela['IRRF'] = tabela['IRRF'].apply(lambda x: x[0])
        tabela['INSS'] = tabela['INSS'].apply(lambda x: x[0])
        tabela[['PIS', 'COFINS', 'CSLL', 'IRRF', 'INSS']] = tabela[['PIS', 'COFINS', 'CSLL', 'IRRF', 'INSS']].apply(lambda x: pd.to_numeric(x))
        tabela['SOCIAL'] = tabela[['PIS', 'COFINS', 'CSLL']].sum(axis=1)
        tabela.drop(columns=['PIS', 'COFINS', 'CSLL'], inplace=True)
        tabelaINSS = tabela[tabela['INSS']> 0 ].copy()
        tabelaINSS.rename(columns={'INSS': 'Valor'}, inplace=True)
        if not tabelaINSS.empty:
            tabelaINSS['Tipo'] = 'INSS'
        tabelaIRRF = tabela[tabela['IRRF']> 0 ].copy()
        tabelaIRRF.rename(columns={'IRRF': 'Valor'}, inplace=True)
        if not tabelaIRRF.empty:
            tabelaIRRF['Tipo'] = 'IRRF'
        tabelaSOCIAL = tabela[tabela['SOCIAL'] > 0 ].copy()
        tabelaSOCIAL.rename(columns={'SOCIAL': 'Valor'}, inplace=True)
        if not tabelaSOCIAL.empty:
            tabelaSOCIAL['Tipo'] = 'Retencao Social'
        dfFINAL = pd.concat([tabelaIRRF, tabelaSOCIAL, tabelaINSS])
        dfFINAL.drop(columns=['IRRF', 'INSS', 'SOCIAL'], inplace=True)
        dfFINAL['CNPJ/CPF'] = dfFINAL['CNPJ/CPF'].str.replace(r'[-./]', '', regex=True)
        dfFINAL['Número'] = dfFINAL['Número'].astype(str).str.zfill(10)
    else:
        return 'faill'

    return dfFINAL

def lerentradatxt(pasta, arquivo_entrada, local):
    if os.path.exists(f'{local}/{pasta}/Entradas'):
        with open(f'{local}/{pasta}/Entradas/{arquivo_entrada}', 'r', encoding='latin1') as entrada:
            dados = entrada.read().split('\n')

        dadosF = []
        padrao_regex = re.compile(r"\|\d{2}/\d{2}/\d{4}\|")
        for e, linha in enumerate(dados):
            if re.search(padrao_regex, linha):
                dadosF.append(linha)
                dadosF.append(dados[e+1])

        dadosF =  '\n'.join(dadosF)
        dadosF = dadosF.replace('.', '')
        dadosF = dadosF.replace(',', '')
        dadosF = dadosF.replace('  |            |  |    |    | |           |   |          |          |', '')
        dadosF = dadosF.replace('''   |          |          |\n|''', '')
        dados_io = StringIO(dadosF)
        dadosF = pd.read_csv(dados_io, sep='|', header=None, dtype=str)
        dadosF = dadosF[[1,4,6,8,14]]
        dadosF.columns= ['Data', 'Número', 'CNPJ/CPF', 'Valor', 'Nome']
        dadosF['Data'] = pd.to_datetime(dadosF['Data'], format='%d/%m/%Y')
        dadosF['Data'] = dadosF['Data'].dt.day
        dadosF['CNPJ/CPF'] = dadosF['CNPJ/CPF'].astype(str).replace('/', '', regex=True)
        dadosF['Nome'] = dadosF['Nome'].apply(lambda x: re.sub(r'\d{5,}', '', x))
        return True, dadosF
    else:
        return False, ''

def juntartomadospdf(df1, dfpdf):
    for index, row in dfpdf.iterrows():
        numero_atual = row['Número']
        cnpj_atual = row['CNPJ/CPF']
        # Encontrar o índice da linha correspondente no df1
        idx_df1 = df1[(df1['Número'] == numero_atual)  & (df1['CNPJ/CPF'] == cnpj_atual)].index
        # Inserir a linha do dfpdf abaixo da linha correspondente no df1
        if not idx_df1.empty:
            idx_df1 = idx_df1[0]  # Usar o primeiro índice se houver mais de um
            df1 = pd.concat([df1.iloc[:idx_df1 + 1], pd.DataFrame(row).T, df1.iloc[idx_df1 + 1:]]).reset_index(drop=True)
        else:
            df1 = pd.concat([df1, pd.DataFrame(row).T]).reset_index(drop=True)
    return df1

def lerbanco():
    dfbanco = pd.read_excel('BANCOCNPJ.xlsx')
    dfbanco.columns = ('CNPJ/CPF', 'Nome')
    dfbanco['CNPJ/CPF'] = dfbanco['CNPJ/CPF'].apply(lambda x: str(x).zfill(14))
    mapa_cnpj_nome = dict(zip(dfbanco['CNPJ/CPF'], dfbanco['Nome']))
    return mapa_cnpj_nome

def obter_arquivo_com_padrao(pasta, local):
    try:
        for nome_arquivo in os.listdir(f'{local}/{pasta}/Serviços Tomados'):
            if any(padrao.lower() in unidecode(nome_arquivo).lower() for padrao in ('retencoes' , 'retencao', 'Retençao')):
                return True, nome_arquivo
    except:
        for nome_arquivo in os.listdir(f'{local}/{pasta}'):
            if any(padrao.lower() in unidecode(nome_arquivo).lower() for padrao in ('retencoes' , 'retencao')):
                return True, nome_arquivo
    return False, ''

def alterarnome(df1):
    df1['NF Nome'] = 'NF '+ df1['Número'].fillna('').astype(str) + ' ' + df1['Nome'].fillna('')
    
    def adicionar_texto(row):
        if row['Tipo'] == 'IRRF':
            return 'IRRF RETIDO CF. ' + row['NF Nome']
        elif row['Tipo'] == 'Retencao Social':
            return 'RETENÇÃO SOCIAL CF. ' + row['NF Nome']
        elif row['Tipo'] == 'INSS':
            return 'INSS RETIDO CF. ' + row['NF Nome']
        elif row['Tipo'] == 'ISS':
            return 'ISS RETIDO CF. ' + row['NF Nome']
        else:
            return row['NF Nome']

    df1['NF Nome'] = df1.apply(adicionar_texto, axis=1)
    df1['NF Nome'] = df1['NF Nome'].replace('NF 00 ', '')
    df1 = df1.drop(columns=['Nome'])
    return df1

def salvar_em_excel(df, nome_arquivo, localsalvar):
    df = df.drop(0)
    # Obtendo o mês e o ano
    mes, ano = pegardata()
    df['Ano'] = ano
    df['Mês'] = mes
    df['Dia'] = df['Data']
    df['Data'] = df['Dia'].astype(str) + '/' + df['Mês'].astype(str) + '/' + df['Ano'].astype(str)
    df['Data'] = df['Data'].str.strip()
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
    # Removendo as colunas de ano, mês e dia após criar a coluna de data
    df = df.drop(['Dia', 'Mês', 'Ano'], axis=1)
    # Salva o DataFrame em um arquivo Excel
    df['Vazia1'] = ''
    df['Vazia2'] = ''
    df['Vazia3'] = ''
    df['Vazia4'] = ''
    df['Vazia5'] = ''
    df['Valor'] = df['Valor'].astype(float).astype(int)
    df = df[['Vazia1','Vazia2','Vazia3', 'Vazia4', 'Valor', 'Data', 'Vazia5', 'NF Nome']]
    df.to_excel(f'{localsalvar}/{nome_arquivo}', index=False, header=False)
    print(f'DataFrame salvo em "{nome_arquivo}".')

def salvar_junto(df, localsalvar):
    tudo = pd.read_excel(f'{localsalvar}/TUDO.xlsx')
    linha_vazia = pd.DataFrame({'Data': [''], 'Número': [''], 'Valor': [''], 'CNPJ/CPF': ['']})
    fim = pd.concat([tudo, linha_vazia,linha_vazia, df]).reset_index(drop=True)
    fim.to_excel(f'{localsalvar}/TUDO.xlsx', index=False)

def attbanco(df, banco):
    for index, row in df.iterrows():
        cnpj = row['CNPJ/CPF']
        nome = row['Nome']
# Verifica se o CNPJ já está no dicionário
        if cnpj not in banco:
            banco[cnpj] = nome
            df_resultado = pd.DataFrame(list(banco.items()), columns=['CNPJ', 'Nome'])
            df_resultado['CNPJ'] = df_resultado['CNPJ'].astype(float)
            df_resultado.to_excel('BANCOCNPJ.xlsx', index=False)
    return banco

def pradronizarxl(localsalvar):
    for arq in os.listdir(localsalvar):
        # Carregar o arquivo Excel
        workbook = openpyxl.load_workbook(f'{localsalvar}/{arq}')

        # Selecionar a planilha desejada (substitua 'Sheet1' pelo nome da sua planilha)
        sheet = workbook['Sheet1']

        # Definir o tamanho da coluna A para 20 (substitua 'A' pelo identificador da sua coluna)
        sheet.column_dimensions['A'].width = 5
        sheet.column_dimensions['B'].width = 18
        sheet.column_dimensions['C'].width = 18
        sheet.column_dimensions['D'].width = 5
        sheet.column_dimensions['E'].width = 12
        sheet.column_dimensions['F'].width = 11.14
        sheet.column_dimensions['G'].width = 6
        sheet.column_dimensions['H'].width = 85

        #Salvar as alterações no arquivo
        workbook.save(f'{localsalvar}/{arq}')

def gerarNotas(local, localsalvar, txt, txt2):
    global NOMETXT
    NOMETXT = txt
    dfbanco = lerbanco()
    lista= listar_pastas(local)
    criar_pasta(localsalvar)
    for pasta in lista:
        dftomados = lertomadostxt(pasta, txt, local)
        condeentrada ,dfentrada = lerentradatxt(pasta, txt2, local)
        cond, nome = obter_arquivo_com_padrao(pasta, local)

        if cond:
            dfpdf = lerpdf(pasta, nome, local)
            dftomados = juntartomadospdf(dftomados, dfpdf)

        if condeentrada:
            attbanco(dfentrada[['CNPJ/CPF', 'Nome']], dfbanco)

        dftomados['Nome'] = dftomados['CNPJ/CPF'].map(dfbanco)

        if condeentrada:
            linha_vazia = pd.DataFrame({'Data': ['00'], 'Número': ['00'], 'Valor': ['00'], 'CNPJ/CPF': ['']})
            
            dftomados = pd.concat([dftomados, linha_vazia, dfentrada]).reset_index(drop=True)
            
        dftomados['Número'] = dftomados['Número'].astype(float).astype(int).apply(lambda x: f'{x:02}')
        nome = dftomados.iloc[0, 0].replace('Empresa: ', '') + '.xlsx'
        dftomados = alterarnome(dftomados)
        dftomados['Vazia1'] = ''
        dftomados['Vazia2'] = ''
        salvar_junto(dftomados[['Data', 'Número', 'CNPJ/CPF' , 'Vazia1', 'Vazia2','Valor', 'NF Nome', 'Tipo']], localsalvar)
        salvar_em_excel(dftomados[[ 'Valor', 'Data', 'NF Nome', 'CNPJ/CPF']], nome, localsalvar)
        pradronizarxl(localsalvar)
        print('################################################')

if __name__ == "__main__":
    gerarNotas(r'C:/Users/Alexandre/Downloads/drive-download-20240409T175707Z-001/ARQUIVOS LBR/TXT LBR', r"C:\Users\Alexandre\Documents\Programacao\Python\Programa pyqt\arquivos saida teste\testeess", 'I56032024.txt', 'E032024.txt')