import pandas as pd
import os

class PRN:
    def __init__(self, caminho, salvar) -> None:
        self.caminho, self.salvar= caminho, salvar

    def printarInformacoes(self, conteudo):
        print(conteudo)

    def verificar(self):
        caminho = self.caminho
        if os.path.isfile(caminho):
            nome = os.path.splitext(os.path.basename(caminho))[0]
            self.transformar(self.caminho, nome, self.salvar)
        elif os.path.isdir(caminho):
            for arquivo in os.listdir(caminho):
                caminho_completo = os.path.join(caminho, arquivo)
                if arquivo.endswith('.xlsx'):
                    arq = os.path.splitext(arquivo)[0]
                    self.transformar(caminho_completo, arq, self.salvar)
        else:
            self.printarInformacoes("O caminho fornecido não é válido.")

    def transformar(self, arquivo, nomearq, salvar):
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
                    valor = round(valor)
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

        try:
            df = pd.read_excel(arquivo, header=None)
            df = df.fillna('')
            novacoluna = 'novacoluna'
            num = 1
            while len(df.columns) < 15:
                df[f'{novacoluna} {num}'] = ''
                num+=1


            df.columns = ['campo1','codigo debito', 'codigo credito', 'codigo historico', 'valor', 'data', 'campo8' , 'nome', 'campo10', 'campo11', 'centro', 'valor1', 'centroC', 'valorC', 'letra']
            df['data'] = df['data'].apply(lambda x: '' if pd.isnull(x) else x)
            texto = ''

            for indice, linha in df.iterrows():
                lista_da_linha = linha.values.tolist()
                campoo1, codigo_debito, codigo_credito, codigo_historico, valor, data, campoo8 , nome, campoo10, campoo11, centroD, valorD, centroC, valorC , letra= lista_da_linha
                texto = texto+f'{campo01(campoo1)}{campo02(codigo_debito)}{campo03(codigo_credito)}{campo04(codigo_historico)}{campo05()}{campo06(valor)}{campo07(data)}{campo08()}{campo09(nome)}{campo10(campoo10)}{campo11(campoo11)}{campo12(centroD)}{campo13(valorD)}{campo14(centroC)}{campo15(valorC)}{campo16(letra)}{campo17()}{campo18()}\n'


            with open(f'{salvar}/{nomearq}.prn', 'w') as arq:
                arq.write(texto)   
            
            self.printarInformacoes(f'NOME: {nomearq}.prn GERADO')
        except Exception as e:
            self.printarInformacoes(f'NOME: {nomearq}.prn ## OUVE UM ERRO NESSE ARQUIVO ## {e}')

class PRNui(PRN):
    def __init__(self, caminho, salvar, ui) -> None:
        super().__init__(caminho, salvar)
        self.ui = ui

    def printarInformacoes(self, conteudo):
        self.ui.printPRN(conteudo)



if __name__ == "__main__": 
    prn = PRN(r"C:\Users\Alexandre\Downloads\Faturamento Com CC 03-2024 - Correta (1).xlsx", r"C:\Users\Alexandre\Desktop\Nova pasta")
    prn.verificar()