# -*- coding: utf-8 -*-
from Grafo import Grafo

if __name__ == "__main__":
    local = "dados.xls"
    grafo = Grafo()
    grafo.ler_excel(local)
    grafo.renderiza_grafo("","Teste")