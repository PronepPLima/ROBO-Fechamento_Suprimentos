ALTER SESSION SET CURRENT_SCHEMA = IW_PROD_RJ;

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
FROM   TD_ICW_FECH_ESTOQ8 A,
       TD_ICW_FECH_EST_C8 B,
       SccCode C,
       MatMaterialType D,
       SccTable E,
       MatDispensingArea F
WHERE  A.ScMaterial = C.id
       AND A.ScMaterial = D.ScMaterial
       AND A.idlot = B.id
       AND C.idtable = E.ID
       --codigo que aparece na tela1675, filtrar por data, para depois selecionar o ID e botao direito do mouse e etc: 
       AND B.ID = 44 --$P!{ID}
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
       AND F.ID = 1 --$P!{Filial}
ORDER BY f.name, e.name