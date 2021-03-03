from excel_datas import Extract_excel

excel = Extract_excel()

'''
    caminho_arquivo = 'c:\\users\example\arquivo.xlsx'
    nome_planilha = 'planilha_1'
    linha = 1
    coluna = 0

    Linha e Coluna sempre colocar para ignorar o index, se meus dados começam a partir
    da linha 5 no arquivo.xlsx, no codigo eu devo colocar que ele comece da linha 4,
    lembrando que a posição 0 em programaçao conta, ou seja 0-1-2-3-4 sendo 0 index
    ou nome das colunas, e 4 onde começa os dados. A mesma logica se aplica para colunas.
'''
todos_os_dados = excel.todos_dados('caminho_arquivo','nome_planilha','linha','coluna')

'''
    Return todos_os_dados = [ [linha1], [linha2], [linha3] ]
'''

dado_unico = excel.dado_unico('caminho_arquivo','nome_planilha','linha','coluna')

'''
    Todas as LINHAS de uma unica coluna.
    Return dado_unico = [ coluna0_linha0, coluna0_linha1, coluna0_linha2 ]
'''

dado_unico_linha = excel.dado_unico_line('caminho_arquivo','nome_planilha','linha','coluna')

'''
    Todas as COLUNAS de uma unica linha.
    Return dado_unico_linha = [ coluna0_linha0, coluna1_linha0, coluna2_linha0 ]
'''