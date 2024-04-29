import pandas as pd
import os



def transformar(arquivo, nomearq, salvar):
    def campo01(valor=''):
        valor = str(valor)
        quant = 5 - len(valor)
        return f'{quant*' '}{valor}'

    def campo02(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 18 - len(valor)
        return f'{quant*' '}{valor}'

    def campo03(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 18 - len(valor)
        return f'{quant*' '}{valor}'

    def campo04(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 4 - len(valor)
        return f'{quant*' '}{valor}'

    def campo05(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 1 - len(valor)
        return f'{quant*' '}{valor}'

    def campo06(valor=''):
        print(valor)
        if valor != '':
            try:
                valor = round(valor)
            except:
                pass
            valor = str(valor)
            quant = 12 - len(valor)
            return f'{quant*'0'}{valor}'
        else:
            quant = 12 - len(valor)
            return f'{quant*' '}{valor}'


    def campo07(valor=''):
        try:
            valor = valor.strftime('%d/%m/%Y')
        except:
            pass
        valor = str(valor)
        quant = 10 - len(valor)
        return f'{quant*' '}{valor}'

    def campo08(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 6 - len(valor)
        return f'{quant*' '}{valor}'

    def campo09(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 143 - len(valor)
        if quant < 0:
            return f'{valor[:143]}'
        return f'{valor}{quant*' '}'

    def campo10(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 20 - len(valor)
        return f'{valor}{quant*' '}'

    def campo11(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 20 - len(valor)
        return f'{valor}{quant*' '}'

    def campo12(valor=''):
        try:
            valor = int(valor)
            valor = str(valor)
            valor =  valor[:1]+ '.' + valor[1:3] + '.' + valor[-3:]
        except:
            pass
        quant = 20 - len(valor)
        return f'{quant*' '}{valor}'

    def campo13(valor=''):
        if valor != '':
            try:
                valor = round(valor)
            except:
                pass
            valor = str(valor)
            quant = 15 - len(valor)
            return f'{quant*'0'}{valor}'
        else:
            quant = 15 - len(valor)
            return f'{quant*' '}{valor}'

    def campo14(valor=''):
        try:
            valor = int(valor)
            valor = str(valor)
            valor =  valor[:1]+ '.' + valor[1:3] + '.' + valor[-3:]
        except:
            pass
        valor = str(valor)
        quant = 20 - len(valor)
        return f'{quant*' '}{valor}'

    def campo15(valor=''):
        if valor != '':
            try:
                valor = int(valor)
            except:
                pass
            valor = str(valor)
            quant = 15 - len(valor)
            return f'{quant*'0'}{valor}'
        else:
            quant = 15 - len(valor)
            return f'{quant*' '}{valor}'

    def campo16(valor=''):
        valor = str(valor)
        quant = 1 - len(valor)
        return f'{quant*' '}{valor}'

    def campo17(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 4 - len(valor)
        return f'{quant*' '}{valor}'

    def campo18(valor=''):
        try:
            valor = int(valor)
        except:
            pass
        valor = str(valor)
        quant = 10 - len(valor)
        return f'{quant*' '}{valor}'

    print('vai ler')
    df = pd.read_excel(arquivo, header=None)
    df = df.fillna('')
    print(df)
    novacoluna = 'novacoluna'
    num = 1
    while len(df.columns) < 15:
        df[f'{novacoluna} {num}'] = ''
        num+=1

    print(f'adc {num} colunas')

    df.columns = ['campo1','codigo debito', 'codigo credito', 'codigo historico', 'valor', 'data', 'campo8' , 'nome', 'campo10', 'campo11', 'centro', 'valor1', 'centroC', 'valorC', 'letra']
    df['data'] = df['data'].apply(lambda x: '' if pd.isnull(x) else x)
    print('leu')
    texto = ''

    for indice, linha in df.iterrows():
        lista_da_linha = linha.values.tolist()
        campoo1, codigo_debito, codigo_credito, codigo_historico, valor, data, campoo8 , nome, campoo10, campoo11, centroD, valorD, centroC, valorC , letra= lista_da_linha
        texto = texto+f'{campo01(campoo1)}{campo02(codigo_debito)}{campo03(codigo_credito)}{campo04(codigo_historico)}{campo05()}{campo06(valor)}{campo07(data)}{campo08()}{campo09(nome)}{campo10(campoo10)}{campo11(campoo11)}{campo12(centroD)}{campo13(valorD)}{campo14(centroC)}{campo15(valorC)}{campo16(letra)}{campo17()}{campo18()}\n'


    with open(f'{salvar}/{nomearq}.prn', 'w') as arq:
        arq.write(texto)   
    
if __name__ == "__main__":
    transformar(r"C:\Users\Alexandre\Desktop\Banco_do_Brasil.xlsx", 'teste1', r"C:\Users\Alexandre\Documents\Programacao\Python\Programa pyqt")