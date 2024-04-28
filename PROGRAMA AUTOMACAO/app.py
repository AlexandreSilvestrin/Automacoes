import pandas as pd
from PyQt5 import uic, QtCore
from PyQt5.QtWidgets import *
import os
from funcoes.transformarPRN2 import PRNui
from funcoes.NOTAS2 import NotasUI, PesquisaAPIThread

class JanelaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("interface/mainW.ui", self)
        self.configurarUI()
    
    def configurarUI(self):
        self.textoPRNarq.setReadOnly(True)
        self.textoPRNlocal.setReadOnly(True)

        self.btnVoltar.clicked.connect(lambda: self.mostrar_pagina(0, cond=True))
        self.btnVoltar.setVisible(False)

        ### configurar FUNCOES DAS NOTAS

        self.btnNotas.clicked.connect(lambda: self.mostrar_pagina(1))
        self.btnNotaarq.clicked.connect(self.localizararqN)
        self.btnNotalocal.clicked.connect(self.localizarlocalN)
        self.btnNotatransformar.clicked.connect(self.transformarNotas)
        self.btnCNPJ.clicked.connect(self.abrir_segunda_janela)
        self.btn_exportar.clicked.connect(self.exportarbanco)

        ######################

        self.btnAbrPastaNota.clicked.connect(self.abrirpastaN)
        self.btnAbrPastaNota.setVisible(False)
        ##
        self.btnPRN.clicked.connect(lambda: self.mostrar_pagina(2))
        self.btnPRNarq.clicked.connect(self.localizararq)
        self.btnPRNpasta.clicked.connect(self.localizarpasta)
        self.btnPRNlocal.clicked.connect(self.localizarlocal)
        
        self.btnAbrPasta.clicked.connect(self.abrirpasta)
        self.btnAbrPasta.setVisible(False)

        ##
        self.btnRazao.clicked.connect(lambda: self.mostrar_pagina(3))
        self.btnRateio.clicked.connect(lambda: self.mostrar_pagina(4))
        self.btnPRNtransformar.clicked.connect(self.transformarprn)

    def verificarCampos(self):
        localNotas = self.localNotas.text()
        localNotasSalvar = self.localNotasSalvar.text()
        txtTomados = self.txtTomados.text()
        txtEntrada = self.txtEntrada.text()

        if os.path.exists(localNotas) and os.path.exists(localNotasSalvar) and txtTomados.strip() != '' and txtEntrada.strip() != '':
            print('sim')
            return True, localNotas,  localNotasSalvar, txtTomados, txtEntrada
        else:
            print('nao')
            return False, localNotas,  localNotasSalvar, txtTomados, txtEntrada

    def transformarNotas(self):
        try:
            verificacao ,localNotas,  localNotasSalvar, txtTomados, txtEntrada = self.verificarCampos()
            print('aq')
            if verificacao:
                Cnotas = NotasUI(localNotas, localNotasSalvar, txtTomados, txtEntrada, self)
                Cnotas.gerarNotas()
            else:
                QMessageBox.critical(self, "Erro", "O caminho do arquivo não é válido ou não existe.")
        except Exception as e: 
            print(e)

    def printNotas(self, conteudo):
        if 'LIMPAR' == conteudo:
            self.txtinfoNota.setText(f'##############')
        else:
            conteudoo = self.txtinfoNota.toPlainText()
            self.txtinfoNota.setText(f'{conteudo}\n{conteudoo}')
        QApplication.processEvents()

    def printPRN(self, conteudo):
        if 'LIMPAR' == conteudo:
            self.txtinfoPRN.setText(f'##############')
        else:
            conteudoo = self.txtinfoPRN.toPlainText()
            self.txtinfoPRN.setText(f'{conteudo}\n{conteudoo}')
        QApplication.processEvents()

    def abrirpastaN(self):
        caminho_da_pasta = self.localNotasSalvar.text()
        os.startfile(caminho_da_pasta)

    def abrirpasta(self):
        caminho_da_pasta = self.textoPRNlocal.text()
        os.startfile(caminho_da_pasta)

    def sumirbtn(self, botao, cond=False):
        if cond:
            self.btnNotas.setVisible(True)
            self.btnPRN.setVisible(True)
            self.btnRazao.setVisible(True)
            self.btnRateio.setVisible(True)
            
            self.btnVoltar.setVisible(False)
            botao.setVisible(False)
        else:
            self.btnNotas.setVisible(False)
            self.btnPRN.setVisible(False)
            self.btnRazao.setVisible(False)
            self.btnRateio.setVisible(False)

            self.btnVoltar.setVisible(True)
            botao.setVisible(True)

    def mostrar_pagina(self,index, cond=False):
        sender_button = self.sender()
        self.sumirbtn(sender_button, cond)
        self.stackedWidget.setCurrentIndex(index)

    def localizarpasta(self):
        options = QFileDialog.Options()
        file_name = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", options=options)
        self.textoPRNarq.setText(file_name)

    def localizararq(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo Excel", "", "Arquivos Excel (*.xlsx)", options=options)
        self.textoPRNarq.setText(file_name)

    def localizarlocal(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", options=options)
        self.textoPRNlocal.setText(folder_path)
    
    def localizararqN(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", options=options)
        self.localNotas.setText(folder_path)

    def localizarlocalN(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", options=options)
        self.localNotasSalvar.setText(folder_path)

    def transformarprn(self):
        try:
            localarq = self.textoPRNarq.text()
            localsalvar = self.textoPRNlocal.text()
            prn = PRNui(localarq, localsalvar, self)
            prn.verificar()
            conteudoo = self.txtinfoPRN.toPlainText()
            self.txtinfoPRN.setText(f'{conteudoo} \n FINALIZADO, arquivo(s) salvo(s)')
            self.btnAbrPasta.setVisible(True)
        except Exception as e:
            self.txtinfoPRN.setText(f'{conteudoo} \n {e}\n OUVE UM ERRO ao transformar')

    def abrir_segunda_janela(self):
        verificacao ,localNotas,  localNotasSalvar, txtTomados, txtEntrada = self.verificarCampos()
        print('aq')
        if verificacao:
            Cnotas = NotasUI(localNotas, localNotasSalvar, txtTomados, txtEntrada, self)
            self.segunda_janela = SegundaJanela(Cnotas)
            self.segunda_janela.show()
        else:
            QMessageBox.critical(self, "Erro", "O caminho do arquivo não é válido ou não existe.")

    def exportarbanco(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", options=options)
        print(folder_path)
        if folder_path == '':
            pass
        else:
            dfbanco = pd.read_excel('BANCOCNPJ.xlsx')
            dfbanco.columns = ('CNPJ/CPF', 'Nome')
            dfbanco['CNPJ/CPF'] = dfbanco['CNPJ/CPF'].apply(lambda x: str(x).zfill(14))
            dfbanco.to_excel(f'{folder_path}/BANCOCNPJ.xlsx', index=False)
            QMessageBox.information(self, "SALVO", f"Banco de dados salvo em: \n {folder_path}")

class SegundaJanela(QMainWindow):
    def __init__(self, Cnotas):
        super().__init__()
        uic.loadUi("interface/secW.ui", self)
        self.setWindowTitle("Segunda Janela")
        self.Cnotas = Cnotas
        self.df = self.Cnotas.pegarCNPJS()
        self.preencher_tabela(self.df)
        self.btnPesqAPI.clicked.connect(self.iniciar_pesquisa)
        # Preencher a tabela com os dados iniciais
        self.thread = PesquisaAPIThread()
        self.thread.resultado_encontrado.connect(self.atualizar)
        self.btnSalvarCNPJS.clicked.connect(self.salvarCNPJS)
        self.tableWidget.setColumnWidth(1, 550)
        self.tableWidget.setColumnWidth(0, 150)
        self.btn_salvar_banco.clicked.connect(self.juntarCNPJSbanco)
        self.btnPARAR.clicked.connect(self.pararAPI)
        self.salvo = True
        self.btn_salvar_banco.setVisible(False) 

    def closeEvent(self, event):
        if self.salvo:
            self.thread.terminate()  # Encerrando a thread ao fechar a janela
            event.accept()  
        else:
            resposta = QMessageBox.question(self, 'Fechar', 'Deseja fechar sem salvar? O progresso sera perdido', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if resposta == QMessageBox.Yes:
                self.thread.terminate()  # Encerrando a thread ao fechar a janela
                event.accept()  # Aceitando o evento de fechar a janela
            else:
                event.ignore()  # Ignorando o evento de fechar a janela # Aceitando o evento de fechar a janela

    def pararAPI(self):
        self.btn_salvar_banco.setVisible(True) 
        self.thread.terminate()
        self.labelP.setText(f"PARADO: {self.progress} de {len(self.df)} concluidos")
        self.labelP.setStyleSheet("background-color: red")

    def juntarCNPJSbanco(self):
        df = self.df[self.df['Nome'].notna()]
        caixa_dialogo = QMessageBox.question(self, 'JUNTAR CNPJS E BANCO', "Você quer salvar os CNPJS no banco de DADOS?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if caixa_dialogo == QMessageBox.Yes:
            self.salvo = True
            self.Cnotas.atualizarBANCOCNPJ(df)
            QMessageBox.information(self, "SALVO", "BANCO DE DADOS PREENCHIDO")
            self.df = self.df[self.df['Nome'].isna()].reset_index(drop=True)
            self.preencher_tabela(self.df)
        else:
            print("Usuário escolheu 'Não'")

    def salvarCNPJS(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", options=options)
        self.df.to_excel(f'{folder_path}/CNPJSnovos.xlsx', index=False)
        QMessageBox.information(self, "SALVO", "ARQUIVO EXCEL SALVO.")

    def iniciar_pesquisa(self):
        self.podesalvar = False
        if self.salvo:
            self.btn_salvar_banco.setVisible(False) 
            self.salvo = False
            self.thread.definirparametro(self.df)
            self.thread.start()
        else:
            QMessageBox.information(self, 'Iniciar', 'Se inicicar novamente antes de salvar o progresso ira recomeçar')

    def atualizar(self, i, nome, quant):

        df = self.df.copy()
        df.at[i, 'Nome'] = nome
        self.df = df.copy()


        # Configurar número de linhas e colunas
        self.tableWidget.setRowCount(len(df.index))
        self.tableWidget.setColumnCount(len(df.columns))
        self.progress += 1
        if len(self.df) == self.progress:
            self.labelP.setText(f"FINALIZADO: {self.progress} de {len(self.df)} concluidos")
            self.labelP.setStyleSheet("background-color: green")
        else:
            self.labelP.setText(f"AGUARDE: {self.progress} de {len(self.df)} concluidos")
            self.labelP.setStyleSheet("background-color: yellow")
        total_minutos_necessarios = (len(self.df)-self.progress) * 1 / 3
        horas = int(total_minutos_necessarios // 60)
        minutos = int(total_minutos_necessarios % 60)
        segundos = int((total_minutos_necessarios % 1) * 60)
        self.labelP_2.setText(f"Tempo estimado: {horas}:{minutos}:{segundos}")
        QApplication.processEvents()
        # Preencher a tabela com os dados do DataFrame
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.tableWidget.setItem(i, j, item)

    def preencher_tabela(self,df):
        self.progress = 0
        self.labelP.setText(f"Parado: {self.progress} de {len(self.df)} concluidos")
        self.labelP.setStyleSheet("")

        # Configurar número de linhas e colunas
        self.tableWidget.setRowCount(len(df.index))
        self.tableWidget.setColumnCount(len(df.columns))
        
        # Preencher a tabela com os dados do DataFrame
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.tableWidget.setItem(i, j, item)



if __name__ == "__main__":
    app = QApplication([])
    window = JanelaPrincipal()
    window.show()
    app.exec_()
