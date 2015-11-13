# -*- coding: utf-8 -*-
from graphviz import Graph
import networkx as bbt_grafo
import xlrd as excel


class Grafo:
    def __init__(self, grafo=None):
        if grafo is not None:
            self.meuGrafo = grafo
        else:
            self.meuGrafo = bbt_grafo.Graph()

    def __str__(self):
        A = Graph()
        for vertice in self.meuGrafo.nodes_iter(data=False):
            A.node(vertice.encode("utf8"))
        for aresta_A, aresta_B, dados_aresta in self.meuGrafo.edges_iter(data=True):
            peso_aresta = dict((chave, str(valor)) for chave, valor in dados_aresta.items())
            peso_aresta['label'] = peso_aresta['weight']
            del peso_aresta['weight']
            A.edge(aresta_A.encode("utf8"), aresta_B.encode("utf8"), **peso_aresta)
        return A.source

    def adicionar_vertice(self, vertice):
        self.meuGrafo.add_node(vertice)

    def adicionar_aresta(self, aresta_A, aresta_B, peso, rod=None):
        self.meuGrafo.add_edge(aresta_A, aresta_B, weight=peso, rodoviaria=rod)

    def ler_excel(self, arquivo):
        planilha = excel.open_workbook(arquivo, formatting_info=True)
        primeiraTabela = planilha.sheet_by_index(0)
        for i in range(1, primeiraTabela.nrows):
            for j in range(1, primeiraTabela.ncols):
                xfx = primeiraTabela.cell_xf_index(i, j)
                xf = planilha.xf_list[xfx]
                corFundoCelula = xf.background.pattern_colour_index
                if primeiraTabela.cell(i, j).value != '':
                    if (corFundoCelula == 13):
                        self.adicionar_aresta(primeiraTabela.cell(i, 0).value, primeiraTabela.cell(0, j).value,
                                              primeiraTabela.cell(i, j).value, "Nao")
                    elif (corFundoCelula == 64):
                        self.adicionar_aresta(primeiraTabela.cell(i, 0).value, primeiraTabela.cell(0, j).value,
                                              primeiraTabela.cell(i, j).value, "Sim")
        return True

    def renderiza_grafo(self, lugar_para_gerar, nome_grafo):
        A = Graph(comment=nome_grafo, filename=(lugar_para_gerar + '/Grafo.gv'), engine='dot')
        for vertice in self.meuGrafo.nodes_iter(data=False):
            A.node(vertice.encode("utf8"))
        for aresta_A, aresta_B, dados_aresta in self.meuGrafo.edges_iter(data=True):
            peso_aresta = dict((chave, str(valor)) for chave, valor in dados_aresta.items())
            peso_aresta['label'] = peso_aresta['weight']
            del peso_aresta['weight']
            A.edge(aresta_A, aresta_B, **peso_aresta)
        A.view()

    def arvore_geradora_minima(self):
        return Grafo(bbt_grafo.minimum_spanning_tree(self.meuGrafo))

    def numero_componentes_limitado(self, L=0.0):
        grafo = bbt_grafo.Graph()
        for aresta_A, aresta_B, dados_aresta in self.meuGrafo.edges_iter(data=True):
            valores_aresta = dict((chave, valor) for chave, valor in dados_aresta.items())
            peso = valores_aresta['weight']
            rodoviario = valores_aresta['rodoviaria']
            if (peso != 0) and (peso <= L) and rodoviario == 'Sim':
                grafo.add_edge(aresta_A, aresta_B, weight=peso)
            else:
                grafo.add_node(aresta_A)
                grafo.add_node(aresta_B)
        return bbt_grafo.number_connected_components(grafo), Grafo(grafo)

    def hamilton(self):
        grafo = bbt_grafo.Graph()
        for aresta_A, aresta_B, dados_aresta in self.meuGrafo.edges_iter(data=True):
            valores_aresta = dict((chave, valor) for chave, valor in dados_aresta.items())
            peso = valores_aresta['weight']
            rodoviario = valores_aresta['rodoviaria']
            if (aresta_A != aresta_B) and (rodoviario == 'Sim') and not not grafo.has_edge(aresta_A, aresta_B):
                grafo.add_edge(aresta_A, aresta_B, weight=peso)
            elif (aresta_A != aresta_B) and (not grafo.has_edge(aresta_B, aresta_A)):
                grafo.add_edge(aresta_A, aresta_B, weight=peso)
        print Grafo(grafo)
        F = [(grafo,[grafo.nodes()[0]])]
        n = grafo.number_of_nodes()
        while F:
            graph, path = F.pop()
            confs = []
            for vertice in graph.neighbors(path[-1]):
                conf_p = path[:]
                conf_p.append(vertice)
                conf_g = bbt_grafo.Graph(graph)
                conf_g.remove_node(path[-1])
                confs.append((conf_g,conf_p))
            for g,p in confs:
                if len(p) == n:
                    p = [x.encode("utf8") for x in p]
                    return p
                else:
                    F.append((g,p))
        return None

