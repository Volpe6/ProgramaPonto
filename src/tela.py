from tkinter import *
from ponto import Ponto

class Tela(object):

    def __init__(self, master=None):

        self.master = master
        self.ponto = Ponto(
            'Ponha seu nome aqui',
            'Ponha sua modalidade de bolsa aqui',
            []#Ponha as observações aqui
        )
        self.ponto.processa()

        self.container_superior = Frame(self.master)
        self.container_inferior = Frame(self.master)

        #componentes do container superior
        self.t_label = Label(self.container_superior, text=f'Caminho: {self.ponto.DIR_PADRAO}')

        #componentes do container inferiot
        self.btn        = Button(self.container_inferior, text='iniciar')
        self.btn_salvar = Button(self.container_inferior, text='Salvar') 

        self.btn['command']        = self.acao_btn_iniciar
        self.btn_salvar['command'] = self.acao_btn_salvar

        #container superior
        self.t_label.pack(side=TOP)
        self.container_superior.pack(side=TOP)

        #container inferior
        self.btn.pack(side=LEFT)
        self.btn_salvar.pack(side=RIGHT)
        self.container_inferior.pack(side=BOTTOM)



    

    def acao_btn_iniciar(self):
        if self.btn['text'] == 'iniciar':
            self.btn['text'] = 'finalizar'
            self.t_label['text'] = 'ponto iniciado'
            self.ponto.add_hora_entrada()
        else:
            self.ponto.add_hora_saida()
            self.t_label['text'] = 'ponto finalizado'


    def acao_btn_salvar(self):
        self.ponto.salva_arq()
        self.t_label['text'] = f'Arq. salvo no local:{self.ponto.DIR_PADRAO}'
