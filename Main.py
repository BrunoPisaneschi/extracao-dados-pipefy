# -*- coding:utf-8 -*-
#!/usr/bin/python3

from flow.FlowMaster import pipefy

if __name__ == '__main__':
    caminho_db = ""
    nome_db = "Consolidado_MotivoRecusa_PrimeiroContato"
    caminho_excel = ""
    nome_excel = "Consolidado_MotivoRecusa_PrimeiroContato_2"
    token = ""
    pipefy = pipefy(token, caminho_db, nome_db, caminho_excel, nome_excel)
    pipefy.extract_datas()