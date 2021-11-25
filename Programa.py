import sys
from numpy import empty, string_
from numpy.core.fromnumeric import choose
from numpy.lib.twodim_base import diag
import pandas as pd
import re
import docx
from PyQt5 import uic, QtWidgets, QtGui
from PyQt5.QtCore import QDate
from pandas.core.frame import DataFrame


def gerarWordRelTrimestral():

    caminhoArquivo = pickArchive()
    print(caminhoArquivo[0])

    dataframe = pd.read_excel(
        r""+caminhoArquivo[0], sheet_name="Teste")

    listaUnidades = []
    listaArea = []

    doc = docx.Document()

    # Carrega informações de Unidades
    for index, row in dataframe.iterrows():
        listaUnidades.append(row["Unidade"])

    # Exclui repetidos
    listaUnidades = list(dict.fromkeys(listaUnidades))

    for unidade in listaUnidades:

        df_maskUnidade = dataframe["Unidade"] == unidade
        filtradoUnidade = dataframe[df_maskUnidade]
        listaServicos = []
        # Carrega informações de Serviço
        for index, row in filtradoUnidade.iterrows():
            listaServicos.append(row["Serviço"])

        # Exclui repetidos
        listaServicos = list(dict.fromkeys(listaServicos))

        doc.add_paragraph().add_run(str(unidade))

        for servicos in listaServicos:
            texto = servicos + " ("
            listaArea = []
            df_maskServico = filtradoUnidade["Serviço"] == servicos
            filtradoServico = filtradoUnidade[df_maskServico]
            for index, row in filtradoServico.iterrows():
                listaArea.append(row["Área"])

        # FUNCIONANDO

            if len(listaArea) == 1:
                texto = texto + str(listaArea[0])
            else:
                for i in range(len(listaArea)):
                    texto = texto + str(listaArea[i])
                    if (i+1) != len(listaArea):
                        texto = texto + str("/")
            texto = texto + str(")")
            doc.add_paragraph().add_run(re.sub(' +', ' ', texto))

    caminho = str(chooseSavePath())
    if not caminho:
        doc.save(r""+r"C:\Users\pedror\Documents\Text.docx")
    else:
        doc.save(r""+caminho+"/Text.docx")

    msg = QtWidgets.QMessageBox()
    msg.information(
        None, "Sucesso", "Realizado com sucesso", QtWidgets.QMessageBox.Ok)


def analiseRel20():

    caminhoArquivo = pickArchive()

    dataFiltro = formulario.dataFiltroFS.text()

    dataframe = pd.read_excel(
        r""+caminhoArquivo[0], sheet_name="Novo Relatório")

    filtradofinal = dataframe.query(
        'DT_EXECUCAO == "" and STATUS == "OK" and SIGLA != "1.1." and DT_PROGRAMACAO <= "' + dataFiltro + '"')

    listaDisciplinas = filtradofinal['DISC_NOME'].unique().tolist()

    path = chooseSavePath()

    for i in listaDisciplinas:
        filtradofinal.query('DISC_NOME == "' + i + '"').to_excel(r"" + path +
                                                                 "/"+i+" até "+dataFiltro.replace("/", "-")+".xlsx", index=False)
    
    msg = QtWidgets.QMessageBox()
    msg.information(None, "Sucesso", "Realizado com sucesso", QtWidgets.QMessageBox.Ok)


def pickArchiveNome():
    dialog = QtWidgets.QFileDialog()
    folder_path = dialog.getSaveFileName(
        None, "Select Excel file to import", "", "Excel (*.xls *.xlsx)")
    return folder_path


# escolha .xls* do arquivo pra importar
def pickArchive():
    dialog = QtWidgets.QFileDialog()
    folder_path = dialog.getOpenFileName(
        None, "Select Excel file to import", "", "Excel (*.xls *.xlsx)")
    return folder_path

# do arquivo .docx


def chooseSavePath():
    dialog = QtWidgets.QFileDialog()
    folder_path = str(dialog.getExistingDirectory(
        None, "Escolha a porra da pasta onde salvar..."))
    return folder_path


app = QtWidgets.QApplication([])
formulario = uic.loadUi("telaPrincipal.ui")

# Rotas dos botões
formulario.btnRodar.clicked.connect(gerarWordRelTrimestral)
formulario.btnAnalisar.clicked.connect(analiseRel20)

# Alterando o campo de data para o Current date
now = QDate.currentDate()
formulario.dataFiltroFS.setDate(now)

formulario.show()
app.exec()
