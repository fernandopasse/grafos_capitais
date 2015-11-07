# -*- coding: utf-8 -*-
import xlrd
import graphviz as Graph
import networkx as nx


#def geraGrafo(grafo, local):
#    nx.write_dot(grafo, 'aaa.gv')

"""
Função <abrir_arquivo>: faz a leitura do arquivo xls com dados
"""
def abrirArquivo(local):
    planilha = xlrd.open_workbook(local, formatting_info = True)
    primeira_tabela = planilha.sheet_by_index(0)
    numero_colunas = primeira_tabela.ncols
    numero_linhas = primeira_tabela.nrows
    G = nx.Graph() # Cria um novo grafo
    for cel_linha in range(1, numero_linhas):
    	for cel_coluna in range(1, numero_colunas):
    		cel_nome_linha = primeira_tabela.cell(cel_linha, 0)
    		cel_nome_coluna = primeira_tabela.cell(0, cel_coluna)
    		cel_valor = primeira_tabela.cell(cel_linha, cel_coluna)
    		cel_id = primeira_tabela.cell_xf_index(cel_linha, cel_coluna)
    		cel_obj = planilha.xf_list[cel_id]
    		cel_fundo = cel_obj.background.pattern_colour_index
    		if cel_valor.value != '':
        		#print('(%s, %s, %s, %s)' % (cel_nome_linha.value, cel_nome_coluna.value, cel_valor.value, cel_fundo))
                    G.add_edge(cel_nome_linha.value, cel_nome_coluna.value, weight = cel_valor.value) # adiciona arestas ao grafo
    return G

def algoritmo_prim(grafo):
    return nx.minimum_spanning_tree(grafo)
    
if __name__ == "__main__":
    local = "dados.xls"
    grafo = abrirArquivo(local)
    arvGeradoraMinima = algoritmo_prim(grafo)
    print(sorted(arvGeradoraMinima.edges(data=True)))
#   geraGrafo(arvGeradoraMinima, '')