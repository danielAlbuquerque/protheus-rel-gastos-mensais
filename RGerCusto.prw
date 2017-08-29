#include 'protheus.ch'

/*/{Protheus.doc} RGerCusto
// Relatório de custos gerênciais
@author Daniel_Albuquerque
@since 24/08/2017
@type function
/*/
User Function RGerCusto()
	Local cTime := Time()
	Local cArquivo := "RELGERENCIALCUSTO"+ SubStr( cTime, 1, 2 ) + SubStr( cTime, 4, 2 ) + SubStr( cTime, 7, 2 ) +".XLS"
	Local oExcelApp := NIL
	Local cPath := "C:\Windows\Temp\"
	Local nTotal := 0
	Local cCentroCusto
	
	Private oExcel
	
	// Verifica se o excel está instalado
	If !ApOleClient("MSExcel")
		MsgAlert("Microsoft Excel não instalado!")
		Return
	EndIf
	
	// Params
	If !(Pergunte("RGERCUSTO", .T.))
		Return
	Endif
	
	cCentroCusto := GetDescCC(mv_par02)
	
	aColunas := {}
	aLocais := {}
	oBrush1 := TBrush():New(, RGB(193,205,205))
	
	oExcel := FWMSExcel():New()
	cAba := "Gastos Departamentais - " + mv_par01
	cTabela := "RELATÓRIO DE GASTAOS DEPARTAMENTAIS - " + mv_par01
	
	oExcel:AddWorkSheet(cAba)
	oExcel:AddTable(cAba, cTabela)
	
	oExcel:AddColumn(cAba,cTabela,cCentroCusto ,1, 1, .F.) 
	oExcel:AddColumn(cAba,cTabela,"JANEIRO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"FEVEREIRO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"MARÇO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"ABRIL" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"MAIO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"JUNHO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"JULHO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"AGOSTO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"SETEMBRO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"OUTUBRO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"NOVEMBRO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"DEZEMBRO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"YTD REAL" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"YTD PREVISTO" ,2, 3, .F.)
	oExcel:AddColumn(cAba,cTabela,"YTD VARIAÇÃO" ,2, 3, .F.)
	
	MsAguarde({|lFim| FillExcel(alltrim(mv_par01), alltrim(mv_par02), @lFim)},"Processamento","Aguarde a finalização do processamento...")
	
	
    // GERA O EXCEL
    If !Empty(oExcel:aWorkSheet)

    	oExcel:Activate()
    	oExcel:GetXMLFile(cArquivo)
 
    	CpyS2T("\SYSTEM\"+cArquivo, cPath)

    	oExcelApp := MsExcel():New()
    	oExcelApp:WorkBooks:Open(cPath+cArquivo) 
    	oExcelApp:SetVisible(.T.)

    EndIf
   
Return

Static Function GetDescCC(cCC)
	Local aArea := GetArea()
	Local cDesc := NIL
	
	DbSelectArea("CTT")
	CTT->(DbSetOrder(1))
	
	If CTT->(DbSeek(xFilial("CTT")+cCC))
		cDesc := alltrim(cCC) + " - " + CTT->CTT_DESC01
	Else
		Alert("Centro de custo não encontrado")
	EndIf
	
	DbCloseArea("CTT")
	
	RestArea(aArea)
Return(cDesc)


Static Function FillExcel(cAno, cCCusto, lFim)
	Local aArea := GetArea()
	Local cMesAtual := MONTH(DATE())
	
	
	
	If cMesAtual < 10
		cMesAtual := '0' + cValToChar(cMesAtual)
	Else
		cMesAtual := cValToChar(cMesAtual)
	EndIf
	
	BeginSql Alias "SQL_CUSTO"
		SELECT 
			ct1.ct1_desc01 Desc_Custo,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '01' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS JANEIRO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '02' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS FEVEREIRO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '03' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS MARCO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '04' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS ABRIL,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '05' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS MAIO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '06' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS JUNHO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '07' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS JULHO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '08' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS AGOSTO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '09' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS SETEMBRO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '10' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS OUTUBRO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '11' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS NOVEMBRO,
			(SELECT VI.Real FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = '12' AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS DEZEMBRO,
			
			(SELECT VI.Real_acumulado FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = %exp:cMesAtual% AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS YTD_REAL,
			(SELECT VI.Previsto_Acumulado FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = %exp:cMesAtual% AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS YTD_PREVISTO,
			(SELECT VI.Variacao_Acumulado FROM V_AMC_REL_CUS_GER_MAT VI WHERE VI.ANO = %exp:cAno% AND VI.Mes = %exp:cMesAtual% AND substr(CT1.CT1_CONTA,3,7) = substr(VI.ord_conta,1,7) and vi.conta = %exp:cCCusto%) AS YTD_VARIACAO
		FROM %table:CT1% ct1
		WHERE
			ct1.%notDel% and
			substr(ct1.ct1_conta,1,2) = '34'
	EndSql

	While !SQL_CUSTO->(EOF())
		ConOut("# SQL_CUSTO: " + cValToChar(SQL_CUSTO->JANEIRO))
		oExcel:AddRow(cAba,cTabela, { SQL_CUSTO->Desc_Custo ,;
                                  SQL_CUSTO->JANEIRO ,; 
                                  SQL_CUSTO->FEVEREIRO ,; 
                                  SQL_CUSTO->MARCO ,; 
                                  SQL_CUSTO->ABRIL ,;
                                  SQL_CUSTO->MAIO ,;
                                  SQL_CUSTO->JUNHO ,;
                                  SQL_CUSTO->JULHO ,;
                                  SQL_CUSTO->AGOSTO ,;
                                  SQL_CUSTO->SETEMBRO ,;
                                  SQL_CUSTO->OUTUBRO ,;
                                  SQL_CUSTO->NOVEMBRO ,;
                                  SQL_CUSTO->DEZEMBRO ,;
                                  SQL_CUSTO->YTD_REAL ,;
                                  SQL_CUSTO->YTD_PREVISTO ,;
                                  SQL_CUSTO->YTD_VARIACAO })
                                  
        If lFim
        	MsgInfo("Cancelado!","Fim")
			Exit
        Endif
        MsProcTxt("Processando... ")
        
		SQL_CUSTO->(DbSkip())
	EndDo
	
	SQL_CUSTO->(DbCloseArea())
	
	RestArea(aArea)
	
	
Return(Nil)