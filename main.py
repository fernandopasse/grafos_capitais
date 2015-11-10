# -*- coding: utf-8 -*-
import xlrd
from graphviz import Graph
import networkx as nx

def geraGrafo(N):
    A = Graph('G', filename='Grafo.gv', engine='dot')
    # add nodes
    # loop over edges
    """
    if N.is_multigraph():
        for u, v, key, edgedata in N.edges_iter(data=True, keys=True):
            str_edgedata = dict(('label', str(v)) for k, v in edgedata.items())
            A.edge(u, v, key=str(key), **str_edgedata)
    else:
    """
    for u, v, edgedata in N.edges_iter(data=True):
        str_edgedata = dict(('label', str(v)) for k, v in edgedata.items())
        A.edge(u, v, **str_edgedata)
    A.view()
"""
Função <abrir_arquivo>: faz a leitura do arquivo xls com dados
"""
def abrirArquivo(local):
    planilha = xlrd.open_workbook(local, formatting_info=True)
    primeira_tabela = planilha.sheet_by_index(0)
    numero_colunas = primeira_tabela.ncols
    numero_linhas = primeira_tabela.nrows
    G = nx.Graph()  # Cria um novo grafo
    for cel_linha in range(1, numero_linhas):
        for cel_coluna in range(1, numero_colunas):
            cel_nome_linha = primeira_tabela.cell(cel_linha, 0)
            cel_nome_coluna = primeira_tabela.cell(0, cel_coluna)
            cel_valor = primeira_tabela.cell(cel_linha, cel_coluna)
            cel_id = primeira_tabela.cell_xf_index(cel_linha, cel_coluna)
            cel_obj = planilha.xf_list[cel_id]
            cel_fundo = cel_obj.background.pattern_colour_index
            if cel_valor.value != '':
                # print('(%s, %s, %s, %s)' % (cel_nome_linha.value, cel_nome_coluna.value, cel_valor.value, cel_fundo))
                G.add_edge(cel_nome_linha.value, cel_nome_coluna.value,
                           weight=cel_valor.value)  # adiciona arestas ao grafo
    return G


def algoritmo_prim(grafo):
    return nx.minimum_spanning_tree(grafo)


if __name__ == "__main__":
    local = "dados.xls"
    grafo = abrirArquivo(local)
    arvGeradoraMinima = algoritmo_prim(grafo)
    print(sorted(arvGeradoraMinima.edges(data=True)))
    geraGrafo(arvGeradoraMinima)
