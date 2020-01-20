#!/usr/bin/env python
import preprocessing

excelColumn = {'Projeto':1,
                'VP':3,
                'Grupo Solucionador':4,
                'Programa': 5,
                'Porte':6,
                'Status':7,
                'Data de Início do Projeto':11,
                'Inicio EF':13, 
                'Termino EF':14, 
                'Complitude EF':15,
                'Início ET':26,
                'Término ET':27,
                'Complitude ET':28,
                'Início do Desenv.':29,
                'Término do Desenv.':30,
                'Complitude Desenv':31,
                'Início Smoke Test':32,
                'Término Smoke Test':33,
                'Complitude Test':34,
                'Início Testes Integrados':35,
                'Término Testes Integrados':36,
                'Complitude Integrados':37,
                'Início da HML':38,
                'Término da HML':39,
                'Complitude HML':40,
                'Implantação (Go Live) (Início)':41,
                'Implantação (Go Live) (Término)':42,
                'Complitude Implantação':43,
                'Descrição':47,
                'A.N.':48,
                'GP':49,
                'LT':50,
                'Líder de Testes':51,
                'Gerente Responsável':52,
                'Patrocinador do Projeto (GD em INFRA)':60,
                'Data de Término Planejada':62,
                'Gestão de Demandas':63,
                'Prioridade':64,
                'Objetivo - Macro':84,
                'Funding':87,
                'Solicitante':95,
                'Project Name':96,
                'Fórmula Status Fase':109,
                'Fórmula Status Projeto':110,
                'Forecast Status da Entrega':165}


"""
Pasta onde estão os arquivos excel de projetos"""
origemArquivosExcel = 'pendentes'
"""
Linha inicial do arquivo excel 
"""
linhaInicial = 2