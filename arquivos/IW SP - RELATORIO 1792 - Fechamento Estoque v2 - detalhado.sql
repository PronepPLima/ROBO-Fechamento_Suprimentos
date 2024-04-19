
ALTER SESSION SET CURRENT_SCHEMA = IW_PROD_RJ;

SELECT a.scmaterial,
        a.saldo_qt_tf  AS SALDOQTTF      
       ,a.saldo_qt_t0  AS SALDOQTT0           
        ,a.saldo_vl_t0  AS SALDOVLT0
       ,a.saldo_vl_tf  AS SALDOVLTF
       ,a.saldo_m  AS SALDOM
       ,a.entr_vl  AS ENTRVL
       ,a.saida_vl  AS SAIDAVL
       ,a.entr_qt_tf  AS ENTRQTTF
       ,a.entr_qt_t0  AS ENTRQTT0
       ,a.saida_qt_t0  AS SAIDAQTT0
       ,a.saida_qt_tf  AS SAIDAQTTF
       ,a.perdas_vl  AS PERDASVL
       ,a.perdas_qt_t0  AS PERDASQTT0
       ,a.perdas_qt_tf  AS PERDASQTTF
       ,a.cons_m_diario  AS CONSMDIARIO
       ,a.entr_vl_t0_outr  AS ENTRVLT0OUTR
       ,a.entr_qt_t0_outr  AS ENTRQTT0OUTR
       ,D.mu,
       C.codename
       ,E.name   AS TIPOMATERIAL
       ,d.mu 
       , a.cd
       ,f.name as NOMECD
       ,' ' as TODOSCD
       ,A.ENTR_QT_T0_NF  AS ENTR_QT_T0_NF 
       ,A.ENTR_QT_T0_INICSD  AS ENTR_QT_T0_INICSD       
       ,A.ENTR_QT_T0_DEVOL  AS ENTR_QT_T0_DEVOL         
       ,A.ENTR_QT_T0_INVENT  AS ENTR_QT_T0_INVENT     
       ,A.ENTR_QT_T0_EMP3  AS ENTR_QT_T0_EMP3 
       ,A.ENTR_VL_INI_SD  AS ENTR_VL_INI_SD 
       ,A.ENTR_VL_NF_ENTRADA  AS ENTR_VL_NF_ENTRADA 
       ,A.ENTR_VL_INVENT  AS ENTR_VL_INVENT
       ,A.ENTR_VL_DEVOL  AS ENTR_VL_DEVOL
       ,A.ENTR_VL_T0_EMP3  AS ENTR_VL_T0_EMP3
       ,A.SAIDA_QT_T0_PAC  AS SAIDA_QT_T0_PAC
       ,A.SAIDA_QT_T0_CCNPER  AS SAIDA_QT_T0_CCNPER
       ,A.SAIDA_QT_T0_CCPERD  AS SAIDA_QT_T0_CCPERD
       ,A.SAIDA_QT_T0_INVENT  AS SAIDA_QT_T0_INVENT
       ,A.SAIDA_QT_T0_DEVOL  AS SAIDA_QT_T0_DEVOL
       ,A.SAIDA_QT_T0_EMP3  AS SAIDA_QT_T0_EMP3
       ,A.VLSAI_QT_T0_PAC  AS VLSAI_QT_T0_PAC
       ,A.VLSAI_QT_T0_CCPERD  AS VLSAI_QT_T0_CCPERD
       ,A.VLSAI_QT_T0_CCNPER  AS VLSAI_QT_T0_CCNPER
       ,A.VLSAI_QT_T0_INVENT  AS VLSAI_QT_T0_INVENT
       ,A.VLSAI_QT_T0_DEVOL  AS VLSAI_QT_T0_DEVOL
       ,A.VLSAI_QT_T0_EMP3  AS VLSAI_QT_T0_EMP3
	, TO_CHAR(B.DATAT0, 'DD/MM/YYYY') AS DATAT0
	, TO_CHAR(  (B.DATATF - 1 ) , 'DD/MM/YYYY') AS DATATF
	, F.NAME AS FILIALNAME
	, ( SELECT SUM(round(saldo_vl_t0,2)) FROM  TD_ICW_FECH_ESTOQ8 WHERE  IDLOT = 1 /*$P!{ID}*/ AND saldo_vl_t0 >0 ) AS TOTALVALORT0
	, ( SELECT SUM(round(saldo_vl_tF,2)) FROM  TD_ICW_FECH_ESTOQ8 WHERE  IDLOT = 1 /*$P!{ID}*/ AND saldo_vl_tf >0 ) AS TOTALVALORTF
FROM   TD_ICW_FECH_ESTOQ8 A,
       TD_ICW_FECH_EST_C8 B,
       scccode C,
       matmaterialtype D,
       scctable E,
       matdispensingarea F
WHERE  A.scmaterial = C.id
       AND A.scmaterial = D.scmaterial
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