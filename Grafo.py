# -*- coding: utf-8 -*-
from graphviz import Graph
import networkx as bbt_grafo
import xlrd as excel


class Grafo:
    def __init__(self):
        self.meuGrafo = bbt_grafo.Graph()

    def __str__(self):
        A = Graph()
        for aresta_A, aresta_B, dados_aresta in self.meuGrafo.edges_iter(data=True):
            peso_aresta = dict(('label', str(peso)) for _, peso in dados_aresta.items())
            A.edge(aresta_A.encode("utf8"), aresta_B.encode("utf8"), **peso_aresta)
        return A.source

    def adicionar_vertice(self, vertice):
        self.meuGrafo.add_node(vertice)

    def adicionar_aresta(self, aresta_A, aresta_B, peso):
        self.meuGrafo.add_edge(aresta_A, aresta_B, weight=peso)

    def ler_excel(self, arquivo):
        planilha = excel.open_workbook(arquivo, formatting_info=True)
        primeiraTabela = planilha.sheet_by_index(0)
        for i in range(1, primeiraTabela.nrows):
            for j in range(1, primeiraTabela.ncols):
                if primeiraTabela.cell(i, j).value != '':
                    self.adicionar_aresta(primeiraTabela.cell(i, 0).value, primeiraTabela.cell(0, j).value,
                                          primeiraTabela.cell(i, j).value)
        return True

    def renderiza_grafo(self, lugar_para_gerar, nome_grafo):
        A = Graph(comment=nome_grafo, filename=(lugar_para_gerar + 'Grafo.gv'), engine='dot')
        for aresta_A, aresta_B, dados_aresta in self.meuGrafo.edges_iter(data=True):
            peso_aresta = dict(('label', str(peso)) for _, peso in dados_aresta.items())
            A.edge(aresta_A, aresta_B, **peso_aresta)
        A.view()
