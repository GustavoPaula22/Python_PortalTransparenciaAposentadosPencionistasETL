import requirements as rq
import os
import time
import pandas as pd
import pyodbc
from datetime import datetime


def validador():
    # Localiza arquivo de leitura
    caminho_completo = os.path.join(rq.diretorio_downloads, rq.nome_arquivo)
    print(caminho_completo)

    # Valida existencia do arquivo
    contador = 1
    while True:
        if contador <= 100:
            print('while - 04')
            if os.path.exists(caminho_completo):
                print('402 - sucesso')
                libera = 'sim'
                break
            else:
                contador = contador + 1
                print('404 - nao encontrado')
                time.sleep(5)
        else:
            libera = 'nao'
            break
    return libera


def dados():
    global cursor, conn
    # Ignore Warning
    import sys
    import warnings
    if not sys.warnoptions:
        warnings.simplefilter("ignore")

    # Localiza arquivo de leitura
    caminho_completo = os.path.join(rq.diretorio_downloads, rq.nome_arquivo)
    print(caminho_completo)

    # Inicio de leitura do excel e atribuindo a primeira linha como nome das colunas
    df = pd.read_csv(caminho_completo, header=0, sep=',')
    df['Carga Horária Mensal'].fillna(0, inplace=True)

    # Convertendo conlunas necessarias
    df['Carga Horária Mensal'] = df['Carga Horária Mensal'].astype(int)
    df['Remuneração Total (R$)'] = df['Remuneração Total (R$)'].astype(float)
    df['Abono de férias / Férias CLT (R$)'] = df['Abono de férias / Férias CLT (R$)'].astype(float)
    df['Valor 13º (R$)'] = df['Valor 13º (R$)'].astype(float)
    df['Remuneração do Mês (R$)'] = df['Remuneração do Mês (R$)'].str.replace("""R$""", '').str.replace(',', '.').astype(float)
    df['Valor Corte Teto (R$)'] = df['Valor Corte Teto (R$)'].astype(float)
    df['Demais Descontos (R$) *'] = df['Demais Descontos (R$) *'].astype(float)
    df['Valor Líquido (R$)'] = df['Valor Líquido (R$)'].astype(float)
    print(df)

    # Inicia Conexao com o banco de dados
    conn_str = f"Driver={{ODBC Driver 17 for SQL Server}};Server={rq.server};Database={rq.database};UID={rq.user};PWD={rq.senha}"

    # Testando a conexão ao banco
    try:
        conn = pyodbc.connect(conn_str)
        print("Conexão bem-sucedida!")
        cursor = conn.cursor()
    except pyodbc.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")

    # Inserir dados na tabela
    count = 1
    for index, row in df.iterrows():
        print(f'Inserindo registro numero: {count}')
        try:
            cursor.execute(
                """
                INSERT INTO [dbo].[f_aposentados]
                ([DeMesAno]
                ,[DeOrgao]
                ,[DeUnidade]
                ,[DeCpf]
                ,[NmFuncionario]
                ,[DeDepartamento]
                ,[NmCargo]
                ,[DeFuncao]
                ,[DeVinculo]
                ,[DtAdmissao]
                ,[DeCargaHorario]
                ,[VlRemuneracaoTotal]
                ,[VlAbono]
                ,[VlDecimoTerceiro]
                ,[VlRemuneracaoMes]
                ,[VlCorteTeto]
                ,[VlDescontos]
                ,[VlLiquido])
                VALUES
                (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    row['Mês/Ano'],
                    row['Órgão'],
                    row['Unidade'],
                    row['CPF'],
                    row['Nome do Servidor'],
                    row['Departamento'],
                    row['Nome do cargo efetivo, comissionada e temporário'],
                    row['Função'],
                    row['Tipo Vínculo'],
                    row['Data Admissão'],
                    row['Carga Horária Mensal'],
                    row['Remuneração Total (R$)'],
                    row['Abono de férias / Férias CLT (R$)'],
                    row['Valor 13º (R$)'],
                    row['Remuneração do Mês (R$)'],
                    row['Valor Corte Teto (R$)'],
                    row['Demais Descontos (R$) *'],
                    row['Valor Líquido (R$)']
                )
            )
            conn.commit()
        except pyodbc.Error as e:
            print(f"Erro ao inserir dados: {e}")
        count = count + 1

    os.remove(caminho_completo)


def valida_exe():
    global cursor, conn, conn_log, cursor_log

    # Inicia Conexao com o banco de dados
    conn_str = f"Driver={{ODBC Driver 17 for SQL Server}};Server={rq.server};Database={rq.database};UID={rq.user};PWD={rq.senha}"
    conn_str_log = f"Driver={{ODBC Driver 17 for SQL Server}};Server={rq.server};Database={rq.database};UID={rq.user};PWD={rq.senha}"

    # Testando a conexão ao banco
    try:
        conn = pyodbc.connect(conn_str)
        print("Conexão bem-sucedida! - 1")
        cursor = conn.cursor()
    except pyodbc.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")

    try:
        cursor.execute(
            """
            SELECT EOMONTH(GETDATE()) AS UltimoDiaDoMesAtual;
            """
        )
        resultado = cursor.fetchone()
        ultimo_dia_do_mes_atual = resultado[0]
        conn.commit()

        data_atual = datetime.now().date()
        print("Data atual:", data_atual)
        print("Data ultimo_dia_do_mes_atual:", ultimo_dia_do_mes_atual)

        # _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
        # ultimo_dia_do_mes_atual = datetime.now().date()  # remover depois dos testes
        # _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_

        if data_atual == ultimo_dia_do_mes_atual:
            # Testando a conexão ao banco
            try:
                conn_log = pyodbc.connect(conn_str_log)
                print("Conexão bem-sucedida! - 2")
                cursor_log = conn_log.cursor()
            except pyodbc.Error as e:
                print(f"Erro ao conectar ao banco de dados: {e}")

            try:
                data_inicio = datetime.now()
                cursor_log.execute(
                    """
                    INSERT INTO [dbo].[Log]
                    (
                    [DataInicio],
                    [Descricao],
                    [Status]
                    )
                    VALUES
                    (
                    ?,
                    ?,
                    ?
                    )
                    """,
                    (
                        data_inicio,
                        'Extração de dados Portal Transparencia dados PTAP-Pencionistas e Aposentados'
                        'Sucesso',
                    )
                )
                cursor_log.commit()
            except pyodbc.Error as e:
                data_inicio = datetime.now()
                cursor_log.execute(
                    """
                    INSERT INTO [dbo].[Log]
                    (
                    [DataInicio],
                    [Descricao],
                    [Status]
                    )
                    VALUES
                    (
                    ?,
                    ?,
                    ?
                    )
                    """,
                    (
                        data_inicio,
                        'Extração de dados Portal Transparencia dados PTAP-Pencionistas e Aposentados'
                        'Falha',
                    )
                )
                cursor_log.commit()
                print(f"Erro ao conectar ao banco de dados: {e}")

            return 'Sim'
        else:
            return 'Nao'

    except pyodbc.Error as e:
        print(f"Erro ao inserir dados: {e}")
