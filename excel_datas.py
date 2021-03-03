import xlrd, datetime
from datetime import time

class Extract_excel:
  ## Todas as linhas(menos index), todas colunas(menos index)
  # Pegar linha a linha do arquivo

  def todos_dados(self,path,sheet,row_start,col_start):
    #path: Caminho do arquivo excel xlsx/xlsm
    #sheet: folha do arquivo excel (sheets)
    #row -> linha de inicio
    #col -> coluna de inicio

    #>-> Sempre pegar a primeira linha (ignorando o index)
    if row_start == '' or row_start == 0:
      row_start = 1
    
    row = row_start
    col = col_start
  
    book = xlrd.open_workbook(path)
    worksheet = book.sheet_by_name(sheet)

    total_linhas = worksheet.nrows
    total_cols = worksheet.ncols

    resultados = []
    result = []

    while row < total_linhas:
        while col < total_cols:
            data_type = worksheet.cell_type(row,col)
            if data_type == 3: #Data Cell
                try:
                    data = worksheet.cell_value(row,col)
                    data_tuple = datetime.datetime(*xlrd.xldate_as_tuple(data, book.datemode))
                    datafeita = data_tuple.date()
                    result.append(datafeita)
                    col += 1
                except:
                    #Exceção pra quando valor for horário.
                    data = int(data * 24 * 3600) # convert to number of seconds
                    data_time = time(data//3600, (data%3600)//60, data%60) # hours, minutes, seconds
                    resultadodata = data_time
                    result.append(resultadodata)
                    col  += 1
            else:
              try:
                result.append(int(worksheet.cell_value(row,col)))
                col += 1
              except:
                result.append(worksheet.cell_value(row,col))
                col += 1
        
        resultados.append(result)

        result = []
        col = col_start
        row += 1

    return resultados
  

  def dado_unico(self,path,sheet,row_start,col_start):
    if row_start == '' or row_start == 0:
      row_start = 1
    
    row = row_start
    col = col_start

    book = xlrd.open_workbook(path)
    worksheet = book.sheet_by_name(sheet)

    total_linhas = worksheet.nrows

    dados = []

    while row < total_linhas:
      try:
        dados.append(int(worksheet.cell_value(row,col)))
        row += 1
      except:
        try:
          dados.append(str(worksheet.cell_value(row,col)))
          row += 1
        except Exception as error:
          return "DEU RUIM MANO",error
    
    return dados
  
  def dado_unico_line(self,path,sheet,row_start,col_start):
    if row_start == '' or row_start == 0:
      row_start = 1
    
    row = row_start
    col = col_start

    book = xlrd.open_workbook(path)
    worksheet = book.sheet_by_name(sheet)

    total_linhas = worksheet.nrows
    total_cols = worksheet.ncols

    dados = []

    while col < total_cols:
      data_type = worksheet.cell_type(row,col)
      if data_type == 3: #Data Cell
        try:
          data = worksheet.cell_value(row,col)
          data_tuple = datetime.datetime(*xlrd.xldate_as_tuple(data, book.datemode))
          datafeita = data_tuple.date()
          dados.append(datafeita)
          col += 1
        except Exception as erro:
          # print(erro)
          #Exceção pra quando valor for horário.
          data = int(data * 24 * 3600) # convert to number of seconds
          data_time = time(data//3600, (data%3600)//60, data%60) # hours, minutes, seconds
          resultadodata = data_time
          dados.append(resultadodata)
          col  += 1
      else:
        try:
          dados.append(int(worksheet.cell_value(row,col)))
          col += 1
        except:
          try:
            dados.append(str(worksheet.cell_value(row,col)))
            col += 1
          except Exception as error:
            return "DEU RUIM MANO",error
    
    return dados