#18/04/2023
#@PLima

import pandas as pd
import pyodbc
import xlsxwriter
import oracledb
import openpyxl

# Configurações do banco de dados
server = '10.20.0.129'
database = 'pronep'
username = 'SYSTEM'
password = 'Pronasis1508'

connection = oracledb.connect( user="SYSTEM", password="Pronasis1508", dsn="10.20.0.129/pronep")
connection.current_schema = "IW_PROD_RJ"

descricoes = []
valores = []



class Conect_bd:
    
    def v2_connection_es_lista():
        print(f'connection.current_schema = "IW_PROD_ES"')
        print("\n============================== v2_connection_es_lista ========================")
        print(f"connection: {connection}\nSuccessfully connected to Oracle Database")
        cursor = connection.cursor()
        query = """
                        select 
                            A.ID,
                            A.DATAT0 AS DT_INICIO ,
                            A.DATATF - 1 AS DT_FIM
                        from IW_PROD_ES.TD_ICW_FECH_EST_C8 A
                        where A.id>0 
                        --and (A.DATATF >=  to_date( '01/01/2024','dd/mm/yyyy hh24:mi:ss' ) 
                        --and A.DATATF <=  to_date( '01/05/2024','dd/mm/yyyy hh24:mi:ss' )) 
                        order by A.ID desc
        
                """
        results = cursor.execute(query)
        data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
        data_frame.to_excel('arquivos\IW_PROD_ES_Lista.xlsx' , index=False)
        cursor.close()

    def v2_connection_rj_lista():
        print(f'connection.current_schema = "IW_PROD_RJ"')
        print("\n============================== v2_connection_rj_lista ========================")
        print(f"connection: {connection}\nSuccessfully connected to Oracle Database")
        cursor = connection.cursor()
        query = """
                        select 
                            A.ID,
                            A.DATAT0 AS DT_INICIO ,
                            A.DATATF - 1 AS DT_FIM
                        from IW_PROD_RJ.TD_ICW_FECH_EST_C8 A
                        where A.id>0 
                        --and (A.DATATF >=  to_date( '01/01/2024','dd/mm/yyyy hh24:mi:ss' ) 
                        --and A.DATATF <=  to_date( '01/05/2024','dd/mm/yyyy hh24:mi:ss' )) 
                        order by A.ID desc
        
                """
        results = cursor.execute(query)
        data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
        data_frame.to_excel('arquivos\IW_PROD_RJ_Lista.xlsx' , index=False)
        cursor.close()
        
        
    def v2_connection_sp_lista():
        print(f'connection.current_schema = "IW_PROD_SP"')
        print("\n============================== v2_connection_sp_lista ========================")
        print(f"connection: {connection}\nSuccessfully connected to Oracle Database")
        cursor = connection.cursor()
        query = """
                        select 
                            A.ID,
                            A.DATAT0 AS DT_INICIO ,
                            A.DATATF - 1 AS DT_FIM
                        from IW_PROD_SP.TD_ICW_FECH_EST_C8 A
                        where A.id>0 
                        --and (A.DATATF >=  to_date( '01/01/2024','dd/mm/yyyy hh24:mi:ss' ) 
                        --and A.DATATF <=  to_date( '01/05/2024','dd/mm/yyyy hh24:mi:ss' )) 
                        order by A.ID desc
        
                """
        results = cursor.execute(query)
        data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
        data_frame.to_excel('arquivos\IW_PROD_SP_Lista.xlsx' , index=False)
        cursor.close()
        

    def v2_connection_es_detalhado(id):
        print(f'connection.current_schema = "IW_PROD_RJ"\ncom id:{id}')
        print("\n============================== v2_connection_rj_detalhado ========================")
        print(f"connection: {connection}\nSuccessfully connected to Oracle Database")
        cursor = connection.cursor()
        query ="""

                    SELECT
                        --add manualmente a essa query o ID
                        B.ID
                        --, A.CD
                        , F.NAME AS FILIALNAME
                        , TO_CHAR( B.DATAT0, 'DD/MM/YYYY') AS DATAT0
                        , TO_CHAR( (B.DATATF - 1 ) , 'DD/MM/YYYY') AS DATATF
                        , E.Name AS TIPOMATERIAL
                        , A.ScMaterial AS COD_MATERIAL
                        , C.CodeName AS NOME_MATERIAL
                        , D.MU AS UM
                        , (CASE
                            WHEN A.saldo_qt_t0 >= 0
                            THEN A.saldo_qt_t0
                            ELSE 0 END) AS Qtde_Saldo_Inic
                        , (CASE
                            WHEN A.saldo_qt_tf >= 0
                            THEN A.saldo_qt_tf
                            ELSE 0 END) AS Qtde_Saldo_Final
                        , (CASE
                            WHEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3) >= 0
                            THEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3)
                            ELSE 0 END ) AS Qtde_E_Total
                        , (CASE
                            WHEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3) >= 0
                            THEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3)
                            ELSE 0 END ) AS Qtde_S_TOTAL
                        , (CASE
                            WHEN A.ENTR_QT_T0_NF >= 0
                            THEN A.ENTR_QT_T0_NF
                            ELSE 0 END ) AS Qtde_E_NF
                        , (CASE
                            WHEN A.SAIDA_QT_T0_PAC >= 0
                            THEN A.SAIDA_QT_T0_PAC
                            ELSE 0 END) AS Qtde_S_Pac
                        , (CASE
                            WHEN A.ENTR_QT_T0_INICSD >= 0
                            THEN A.ENTR_QT_T0_INICSD
                            ELSE 0 END) AS Qtde_E_Ini_Sd
                        , (CASE
                            WHEN A.SAIDA_QT_T0_CCNPER >= 0
                            THEN A.SAIDA_QT_T0_CCNPER
                            ELSE 0 END) AS Qtde_S_Outr
                        , (CASE
                            WHEN A.ENTR_QT_T0_DEVOL >= 0
                            THEN A.ENTR_QT_T0_DEVOL
                            ELSE 0 END) AS Qtde_E_Devol
                        , (CASE
                            WHEN A.SAIDA_QT_T0_CCPERD >= 0
                            THEN A.SAIDA_QT_T0_CCPERD
                            ELSE 0 END) AS Qtde_S_Perda
                        , (CASE
                            WHEN A.ENTR_QT_T0_INVENT >= 0
                            THEN A.ENTR_QT_T0_INVENT
                            ELSE 0 END) AS Qtde_E_Invent
                        , (CASE
                            WHEN A.SAIDA_QT_T0_INVENT >= 0
                            THEN A.SAIDA_QT_T0_INVENT
                            ELSE 0 END) AS Qtde_S_Invent
                        , (CASE
                            WHEN A.ENTR_QT_T0_EMP3 >= 0
                            THEN A.ENTR_QT_T0_EMP3
                            ELSE 0 END) AS Qtde_E_3os
                        , (CASE
                            WHEN A.SAIDA_QT_T0_DEVOL >= 0
                            THEN A.SAIDA_QT_T0_DEVOL
                            ELSE 0 END) AS Qtde_S_Devol
                        , (CASE
                            WHEN A.SAIDA_QT_T0_EMP3 >= 0
                            THEN A.SAIDA_QT_T0_EMP3
                            ELSE 0 END) AS Qtde_S_3os
                        , (CASE
                            WHEN A.saldo_vl_t0 >= 0
                            THEN A.saldo_vl_t0
                            ELSE 0 END) AS Valor_Inic
                        , (CASE
                            WHEN A.saldo_vl_tf >= 0
                            THEN A.saldo_vl_tf
                            ELSE 0 END) AS Valor_Final
                        , (CASE
                            WHEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3) >=0
                            THEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3)
                            ELSE 0 END) AS Valor_E_Total
                        , (CASE
                            WHEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3) >= 0
                            THEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3)
                            ELSE 0 END) AS Valor_S_Total
                        , (CASE
                            WHEN A.ENTR_VL_NF_ENTRADA >= 0
                            THEN A.ENTR_VL_NF_ENTRADA
                            ELSE 0 END) AS Valor_E_NF
                        , (CASE
                            WHEN A.VLSAI_QT_T0_PAC >= 0
                            THEN A.VLSAI_QT_T0_PAC
                            ELSE 0 END) AS Valor_S_Pac
                        , (CASE
                            WHEN A.ENTR_VL_INI_SD >= 0
                            THEN A.ENTR_VL_INI_SD
                            ELSE 0 END) AS Valor_E_Ini_Sd
                        , (CASE
                            WHEN A.VLSAI_QT_T0_CCNPER >= 0
                            THEN A.VLSAI_QT_T0_CCNPER
                            ELSE 0 END) AS Valor_S_Outr
                        , (CASE
                            WHEN A.ENTR_VL_DEVOL >= 0
                            THEN A.ENTR_VL_DEVOL
                            ELSE 0 END) AS Valor_E_Devol
                        , (CASE
                            WHEN A.VLSAI_QT_T0_CCPERD >= 0
                            THEN A.VLSAI_QT_T0_CCPERD
                            ELSE 0 END) AS Valor_S_Perda
                        , (CASE
                            WHEN A.ENTR_VL_INVENT >= 0
                            THEN A.ENTR_VL_INVENT
                            ELSE 0 END) AS Valor_E_Invent
                        , (CASE
                            WHEN A.VLSAI_QT_T0_INVENT >= 0
                            THEN A.VLSAI_QT_T0_INVENT
                            ELSE 0 END) AS Valor_S_Invent
                        , (CASE
                            WHEN A.ENTR_VL_T0_EMP3 >= 0
                            THEN A.ENTR_VL_T0_EMP3
                            ELSE 0 END) AS Valor_E_3os
                        , (CASE
                            WHEN A.VLSAI_QT_T0_DEVOL >= 0
                            THEN A.VLSAI_QT_T0_DEVOL
                            ELSE 0 END) AS Valor_S_Devol
                        , (CASE
                            WHEN A.VLSAI_QT_T0_EMP3 >= 0
                            THEN A.VLSAI_QT_T0_EMP3
                            ELSE 0 END) AS Valor_S_3os
                    FROM   
                        IW_PROD_ES.TD_ICW_FECH_ESTOQ8 A,
                        IW_PROD_ES.TD_ICW_FECH_EST_C8 B,
                        IW_PROD_ES.SccCode C,
                        IW_PROD_ES.MatMaterialType D,
                        IW_PROD_ES.SccTable E,
                        IW_PROD_ES.MatDispensingArea F
                    WHERE  A.ScMaterial = C.id
                        AND A.ScMaterial = D.ScMaterial
                        AND A.idlot = B.id
                        AND C.idtable = E.ID
                        AND B.ID = :ID
                        AND A.CD = F.ID
                        AND ( A.entr_vl <> 0
                                OR A.saida_vl <> 0
                                OR A.saldo_m <> 0
                                OR A.saldo_qt_tf <> 0
                                OR A.saldo_qt_t0 <> 0
                                OR A.saldo_vl_tf <> 0
                                OR A.saldo_vl_t0 <> 0
                                OR A.entr_qt_t0 <> 0
                                OR A.saida_qt_t0 <> 0
                                OR A.perdas_qt_t0 <> 0
                                OR A.perdas_vl <> 0 )
                        AND ENTR_QT_T0_NF IS NOT NULL
                        AND F.ID = 1 
                    ORDER BY f.name, e.name
        """
        id = id
        print(f'************************** Id: {id}')
        results = cursor.execute(query, [id])
        data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
        print('\nAqui esta o data frame:\n')
        
        #salvando arquivo xlsx:
        #print(f"\n============================= IW_PROD_ES Salvando arquivo xslx =============================")
        #data_frame.to_excel('arquivos\IW_PROD_ES_Resultado.xlsx' , index=False , sheet_name='Analitico')
        
        #print(f"\n============================= v2_exibicao_es_detalhado ")
        #ler o arquivo xlsx;
        #data_frame = pd.read_excel('arquivos\IW_PROD_ES_Resultado.xlsx')
        
        #retornar em data frame o xlsx sintetico:
        #resultados inseridos num array:
        sub_totais = {}
        print(f"\nAqui esta o SUB_TOTAIS:\n")
        sub_totais = {
                        'IW ES - FECHAMENTO': id,
                        'Sub total dietas inicial' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                        'Sub total dietas final' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                        'Sub total Mat. Enfermagem inicial' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                        'Sub total Mat. Enfermagem final' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                        'Sub total Materiais Diversos inicial' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                        'Sub total Materiais Diversos final' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                        'Sub total Medicamentos inicial' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                        'Sub total Medicamentos final' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                        'Total estoque inicial' : data_frame["VALOR_INIC"].sum(),
                        'Total estoque final' : data_frame["VALOR_FINAL"].sum()
                    }
        print(sub_totais)
        
        #Extraindo Chaves e Valores do Dicionário:
        for chave, valor in sub_totais.items():
            descricoes.append(chave)
            valores.append(valor)
        #Criando o DataFrame:
        df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
        # Nomeando as colunas
        df_subtotais.columns = ["Descrição", "Valor"] 
        
        with pd.ExcelWriter('arquivos\IW_PROD_ES_Resultado.xlsx', engine='openpyxl', mode='w') as writer:
            data_frame.to_excel(writer, sheet_name='Analitico', index=False, header=True)
            df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=False)
        return df_subtotais

        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    
    def v2_connection_rj_detalhado(id):
        print(f'connection.current_schema = "IW_PROD_RJ"\ncom id:{id}')
        print("\n============================== v2_connection_rj_detalhado ========================")
        print(f"connection: {connection}\nSuccessfully connected to Oracle Database")
        # Consulta SQL
        cursor = connection.cursor()
        query = """
                            SELECT
                                --add manualmente a essa query o ID
                                B.ID
                                --, A.CD
                                , F.NAME AS FILIALNAME
                                , TO_CHAR( B.DATAT0, 'DD/MM/YYYY') AS DATAT0
                                , TO_CHAR( (B.DATATF - 1 ) , 'DD/MM/YYYY') AS DATATF
                                , E.Name AS TIPOMATERIAL
                                , A.ScMaterial AS COD_MATERIAL
                                , C.CodeName AS NOME_MATERIAL
                                , D.MU AS UM
                                , (CASE
                                    WHEN A.saldo_qt_t0 >= 0
                                    THEN A.saldo_qt_t0
                                    ELSE 0 END) AS Qtde_Saldo_Inic
                                , (CASE
                                    WHEN A.saldo_qt_tf >= 0
                                    THEN A.saldo_qt_tf
                                    ELSE 0 END) AS Qtde_Saldo_Final
                                , (CASE
                                    WHEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3) >= 0
                                    THEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3)
                                    ELSE 0 END ) AS Qtde_E_Total
                                , (CASE
                                    WHEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3) >= 0
                                    THEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3)
                                    ELSE 0 END ) AS Qtde_S_TOTAL
                                , (CASE
                                    WHEN A.ENTR_QT_T0_NF >= 0
                                    THEN A.ENTR_QT_T0_NF
                                    ELSE 0 END ) AS Qtde_E_NF
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_PAC >= 0
                                    THEN A.SAIDA_QT_T0_PAC
                                    ELSE 0 END) AS Qtde_S_Pac
                                , (CASE
                                    WHEN A.ENTR_QT_T0_INICSD >= 0
                                    THEN A.ENTR_QT_T0_INICSD
                                    ELSE 0 END) AS Qtde_E_Ini_Sd
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_CCNPER >= 0
                                    THEN A.SAIDA_QT_T0_CCNPER
                                    ELSE 0 END) AS Qtde_S_Outr
                                , (CASE
                                    WHEN A.ENTR_QT_T0_DEVOL >= 0
                                    THEN A.ENTR_QT_T0_DEVOL
                                    ELSE 0 END) AS Qtde_E_Devol
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_CCPERD >= 0
                                    THEN A.SAIDA_QT_T0_CCPERD
                                    ELSE 0 END) AS Qtde_S_Perda
                                , (CASE
                                    WHEN A.ENTR_QT_T0_INVENT >= 0
                                    THEN A.ENTR_QT_T0_INVENT
                                    ELSE 0 END) AS Qtde_E_Invent
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_INVENT >= 0
                                    THEN A.SAIDA_QT_T0_INVENT
                                    ELSE 0 END) AS Qtde_S_Invent
                                , (CASE
                                    WHEN A.ENTR_QT_T0_EMP3 >= 0
                                    THEN A.ENTR_QT_T0_EMP3
                                    ELSE 0 END) AS Qtde_E_3os
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_DEVOL >= 0
                                    THEN A.SAIDA_QT_T0_DEVOL
                                    ELSE 0 END) AS Qtde_S_Devol
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_EMP3 >= 0
                                    THEN A.SAIDA_QT_T0_EMP3
                                    ELSE 0 END) AS Qtde_S_3os
                                , (CASE
                                    WHEN A.saldo_vl_t0 >= 0
                                    THEN A.saldo_vl_t0
                                    ELSE 0 END) AS Valor_Inic
                                , (CASE
                                    WHEN A.saldo_vl_tf >= 0
                                    THEN A.saldo_vl_tf
                                    ELSE 0 END) AS Valor_Final
                                , (CASE
                                    WHEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3) >=0
                                    THEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3)
                                    ELSE 0 END) AS Valor_E_Total
                                , (CASE
                                    WHEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3) >= 0
                                    THEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3)
                                    ELSE 0 END) AS Valor_S_Total
                                , (CASE
                                    WHEN A.ENTR_VL_NF_ENTRADA >= 0
                                    THEN A.ENTR_VL_NF_ENTRADA
                                    ELSE 0 END) AS Valor_E_NF
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_PAC >= 0
                                    THEN A.VLSAI_QT_T0_PAC
                                    ELSE 0 END) AS Valor_S_Pac
                                , (CASE
                                    WHEN A.ENTR_VL_INI_SD >= 0
                                    THEN A.ENTR_VL_INI_SD
                                    ELSE 0 END) AS Valor_E_Ini_Sd
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_CCNPER >= 0
                                    THEN A.VLSAI_QT_T0_CCNPER
                                    ELSE 0 END) AS Valor_S_Outr
                                , (CASE
                                    WHEN A.ENTR_VL_DEVOL >= 0
                                    THEN A.ENTR_VL_DEVOL
                                    ELSE 0 END) AS Valor_E_Devol
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_CCPERD >= 0
                                    THEN A.VLSAI_QT_T0_CCPERD
                                    ELSE 0 END) AS Valor_S_Perda
                                , (CASE
                                    WHEN A.ENTR_VL_INVENT >= 0
                                    THEN A.ENTR_VL_INVENT
                                    ELSE 0 END) AS Valor_E_Invent
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_INVENT >= 0
                                    THEN A.VLSAI_QT_T0_INVENT
                                    ELSE 0 END) AS Valor_S_Invent
                                , (CASE
                                    WHEN A.ENTR_VL_T0_EMP3 >= 0
                                    THEN A.ENTR_VL_T0_EMP3
                                    ELSE 0 END) AS Valor_E_3os
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_DEVOL >= 0
                                    THEN A.VLSAI_QT_T0_DEVOL
                                    ELSE 0 END) AS Valor_S_Devol
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_EMP3 >= 0
                                    THEN A.VLSAI_QT_T0_EMP3
                                    ELSE 0 END) AS Valor_S_3os
                            FROM   
                                IW_PROD_RJ.TD_ICW_FECH_ESTOQ8 A,
                                IW_PROD_RJ.TD_ICW_FECH_EST_C8 B,
                                IW_PROD_RJ.SccCode C,
                                IW_PROD_RJ.MatMaterialType D,
                                IW_PROD_RJ.SccTable E,
                                IW_PROD_RJ.MatDispensingArea F
                            WHERE  A.ScMaterial = C.id
                                AND A.ScMaterial = D.ScMaterial
                                AND A.idlot = B.id
                                AND C.idtable = E.ID
                                AND B.ID = :ID
                                AND A.CD = F.ID
                                AND ( A.entr_vl <> 0
                                        OR A.saida_vl <> 0
                                        OR A.saldo_m <> 0
                                        OR A.saldo_qt_tf <> 0
                                        OR A.saldo_qt_t0 <> 0
                                        OR A.saldo_vl_tf <> 0
                                        OR A.saldo_vl_t0 <> 0
                                        OR A.entr_qt_t0 <> 0
                                        OR A.saida_qt_t0 <> 0
                                        OR A.perdas_qt_t0 <> 0
                                        OR A.perdas_vl <> 0 )
                                AND ENTR_QT_T0_NF IS NOT NULL
                                AND F.ID = 1 
                            ORDER BY f.name, e.name
        """
        id = id
        print(f'************************** Id: {id}')
        results = cursor.execute(query, [id])
        data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
        
        #salvando arquivo xlsx:
        print(f"\n============================= IW_PROD_RJ Salvando arquivo xslx =============================")
        data_frame.to_excel('arquivos\IW_PROD_RJ_Resultado.xlsx' , index=False , sheet_name='Analitico')
        
        print(f"\n============================= v2_exibicao_rj_detalhado ")
        #ler o arquivo xlsx;
        data_frame = pd.read_excel('arquivos\IW_PROD_RJ_Resultado.xlsx')
        
        #retornar em data frame o xlsx sintetico:
        #resultados inseridos num array:
        sub_totais = {}
        print(f"\nAqui esta o SUB_TOTAIS:\n")
        sub_totais = {
                        'Sub total dietas inicial' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                        'Sub total dietas final' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                        'Sub total Mat. Enfermagem inicial' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                        'Sub total Mat. Enfermagem final' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                        'Sub total Materiais Diversos inicial' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                        'Sub total Materiais Diversos final' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                        'Sub total Medicamentos inicial' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                        'Sub total Medicamentos final' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                        'Total estoque inicial' : data_frame["VALOR_INIC"].sum(),
                        'Total estoque final' : data_frame["VALOR_FINAL"].sum()
                    }
        
        #Extraindo Chaves e Valores do Dicionário:
        for chave, valor in sub_totais.items():
            descricoes.append(chave)
            valores.append(valor)
        #Criando o DataFrame:
        df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
        # Nomeando as colunas
        df_subtotais.columns = ["Descrição", "Valor"] 
        
        with pd.ExcelWriter('arquivos\IW_PROD_RJ_Resultado.xlsx', engine='openpyxl', mode='w') as writer:
            data_frame.to_excel(writer, sheet_name='Analitico', index=False, header=None)
            df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=None)

        return df_subtotais
        
    def v2_connection_sp_detalhado(id):
        print(f'connection.current_schema = "IW_PROD_SP"\ncom id:{id}')
        print("\n============================== v2_connection_sp_detalhado ========================")
        print(f"connection: {connection}\nSuccessfully connected to Oracle Database")
        # Consulta SQL
        cursor = connection.cursor()
        query = """
                            SELECT
                                --add manualmente a essa query o ID
                                B.ID
                                --, A.CD
                                , F.NAME AS FILIALNAME
                                , TO_CHAR( B.DATAT0, 'DD/MM/YYYY') AS DATAT0
                                , TO_CHAR( (B.DATATF - 1 ) , 'DD/MM/YYYY') AS DATATF
                                , E.Name AS TIPOMATERIAL
                                , A.ScMaterial AS COD_MATERIAL
                                , C.CodeName AS NOME_MATERIAL
                                , D.MU AS UM
                                , (CASE
                                    WHEN A.saldo_qt_t0 >= 0
                                    THEN A.saldo_qt_t0
                                    ELSE 0 END) AS Qtde_Saldo_Inic
                                , (CASE
                                    WHEN A.saldo_qt_tf >= 0
                                    THEN A.saldo_qt_tf
                                    ELSE 0 END) AS Qtde_Saldo_Final
                                , (CASE
                                    WHEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3) >= 0
                                    THEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3)
                                    ELSE 0 END ) AS Qtde_E_Total
                                , (CASE
                                    WHEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3) >= 0
                                    THEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3)
                                    ELSE 0 END ) AS Qtde_S_TOTAL
                                , (CASE
                                    WHEN A.ENTR_QT_T0_NF >= 0
                                    THEN A.ENTR_QT_T0_NF
                                    ELSE 0 END ) AS Qtde_E_NF
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_PAC >= 0
                                    THEN A.SAIDA_QT_T0_PAC
                                    ELSE 0 END) AS Qtde_S_Pac
                                , (CASE
                                    WHEN A.ENTR_QT_T0_INICSD >= 0
                                    THEN A.ENTR_QT_T0_INICSD
                                    ELSE 0 END) AS Qtde_E_Ini_Sd
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_CCNPER >= 0
                                    THEN A.SAIDA_QT_T0_CCNPER
                                    ELSE 0 END) AS Qtde_S_Outr
                                , (CASE
                                    WHEN A.ENTR_QT_T0_DEVOL >= 0
                                    THEN A.ENTR_QT_T0_DEVOL
                                    ELSE 0 END) AS Qtde_E_Devol
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_CCPERD >= 0
                                    THEN A.SAIDA_QT_T0_CCPERD
                                    ELSE 0 END) AS Qtde_S_Perda
                                , (CASE
                                    WHEN A.ENTR_QT_T0_INVENT >= 0
                                    THEN A.ENTR_QT_T0_INVENT
                                    ELSE 0 END) AS Qtde_E_Invent
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_INVENT >= 0
                                    THEN A.SAIDA_QT_T0_INVENT
                                    ELSE 0 END) AS Qtde_S_Invent
                                , (CASE
                                    WHEN A.ENTR_QT_T0_EMP3 >= 0
                                    THEN A.ENTR_QT_T0_EMP3
                                    ELSE 0 END) AS Qtde_E_3os
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_DEVOL >= 0
                                    THEN A.SAIDA_QT_T0_DEVOL
                                    ELSE 0 END) AS Qtde_S_Devol
                                , (CASE
                                    WHEN A.SAIDA_QT_T0_EMP3 >= 0
                                    THEN A.SAIDA_QT_T0_EMP3
                                    ELSE 0 END) AS Qtde_S_3os
                                , (CASE
                                    WHEN A.saldo_vl_t0 >= 0
                                    THEN A.saldo_vl_t0
                                    ELSE 0 END) AS Valor_Inic
                                , (CASE
                                    WHEN A.saldo_vl_tf >= 0
                                    THEN A.saldo_vl_tf
                                    ELSE 0 END) AS Valor_Final
                                , (CASE
                                    WHEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3) >=0
                                    THEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3)
                                    ELSE 0 END) AS Valor_E_Total
                                , (CASE
                                    WHEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3) >= 0
                                    THEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3)
                                    ELSE 0 END) AS Valor_S_Total
                                , (CASE
                                    WHEN A.ENTR_VL_NF_ENTRADA >= 0
                                    THEN A.ENTR_VL_NF_ENTRADA
                                    ELSE 0 END) AS Valor_E_NF
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_PAC >= 0
                                    THEN A.VLSAI_QT_T0_PAC
                                    ELSE 0 END) AS Valor_S_Pac
                                , (CASE
                                    WHEN A.ENTR_VL_INI_SD >= 0
                                    THEN A.ENTR_VL_INI_SD
                                    ELSE 0 END) AS Valor_E_Ini_Sd
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_CCNPER >= 0
                                    THEN A.VLSAI_QT_T0_CCNPER
                                    ELSE 0 END) AS Valor_S_Outr
                                , (CASE
                                    WHEN A.ENTR_VL_DEVOL >= 0
                                    THEN A.ENTR_VL_DEVOL
                                    ELSE 0 END) AS Valor_E_Devol
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_CCPERD >= 0
                                    THEN A.VLSAI_QT_T0_CCPERD
                                    ELSE 0 END) AS Valor_S_Perda
                                , (CASE
                                    WHEN A.ENTR_VL_INVENT >= 0
                                    THEN A.ENTR_VL_INVENT
                                    ELSE 0 END) AS Valor_E_Invent
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_INVENT >= 0
                                    THEN A.VLSAI_QT_T0_INVENT
                                    ELSE 0 END) AS Valor_S_Invent
                                , (CASE
                                    WHEN A.ENTR_VL_T0_EMP3 >= 0
                                    THEN A.ENTR_VL_T0_EMP3
                                    ELSE 0 END) AS Valor_E_3os
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_DEVOL >= 0
                                    THEN A.VLSAI_QT_T0_DEVOL
                                    ELSE 0 END) AS Valor_S_Devol
                                , (CASE
                                    WHEN A.VLSAI_QT_T0_EMP3 >= 0
                                    THEN A.VLSAI_QT_T0_EMP3
                                    ELSE 0 END) AS Valor_S_3os
                            FROM   
                                IW_PROD_SP.TD_ICW_FECH_ESTOQ8 A,
                                IW_PROD_SP.TD_ICW_FECH_EST_C8 B,
                                IW_PROD_SP.SccCode C,
                                IW_PROD_SP.MatMaterialType D,
                                IW_PROD_SP.SccTable E,
                                IW_PROD_SP.MatDispensingArea F
                            WHERE  A.ScMaterial = C.id
                                AND A.ScMaterial = D.ScMaterial
                                AND A.idlot = B.id
                                AND C.idtable = E.ID
                                AND B.ID = :ID
                                AND A.CD = F.ID
                                AND ( A.entr_vl <> 0
                                        OR A.saida_vl <> 0
                                        OR A.saldo_m <> 0
                                        OR A.saldo_qt_tf <> 0
                                        OR A.saldo_qt_t0 <> 0
                                        OR A.saldo_vl_tf <> 0
                                        OR A.saldo_vl_t0 <> 0
                                        OR A.entr_qt_t0 <> 0
                                        OR A.saida_qt_t0 <> 0
                                        OR A.perdas_qt_t0 <> 0
                                        OR A.perdas_vl <> 0 )
                                AND ENTR_QT_T0_NF IS NOT NULL
                                AND F.ID = 1 
                            ORDER BY f.name, e.name
        """
        id = id
        print(f'************************** {id}')
        results = cursor.execute(query, [id])
        data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
        #salvando arquivo xlsx:
        print(f"\n============================= IW_PROD_SP Salvando arquivo xslx =============================")
        data_frame.to_excel('arquivos\IW_PROD_SP_Resultado.xlsx' , index=False , sheet_name='Analitico')
        
        print(f"\n============================= v2_exibicao_sp_detalhado ")
        #ler o arquivo xlsx;
        data_frame = pd.read_excel('arquivos\IW_PROD_SP_Resultado.xlsx')
        
        #retornar em data frame o xlsx sintetico:
        #resultados inseridos num array:
        sub_totais = {}
        print(f"\nAqui esta o SUB_TOTAIS:\n")
        sub_totais = {
                        'Sub total dietas inicial' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                        'Sub total dietas final' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                        'Sub total Mat. Enfermagem inicial' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                        'Sub total Mat. Enfermagem final' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                        'Sub total Materiais Diversos inicial' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                        'Sub total Materiais Diversos final' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                        'Sub total Medicamentos inicial' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                        'Sub total Medicamentos final' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                        'Total estoque inicial' : data_frame["VALOR_INIC"].sum(),
                        'Total estoque final' : data_frame["VALOR_FINAL"].sum()
                    }
        
        #Extraindo Chaves e Valores do Dicionário:
        for chave, valor in sub_totais.items():
            descricoes.append(chave)
            valores.append(valor)
        #Criando o DataFrame:
        df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
        # Nomeando as colunas
        #df_subtotais.columns = ["Descrição", "Valor"] 
        
        with pd.ExcelWriter('arquivos\IW_PROD_SP_Resultado.xlsx', engine='openpyxl', mode='w') as writer:
            data_frame.to_excel(writer, sheet_name='Analitico', index=False, header=None)
            df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=None)

        return df_subtotais       