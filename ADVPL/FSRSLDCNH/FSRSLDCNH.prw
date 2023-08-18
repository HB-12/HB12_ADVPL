#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TBICONN.CH" 
#INCLUDE "REPORT.CH"
 
/*/{Protheus.doc} FSRSldCNH
    
    RelatÃ³rio de Saldo de Produto CNH.
    
    @type  Function
    @author user
    @since 02/03/2020
  
  Executa a consulta no banco de dados e retorna area com dados para impressÃ£o.

    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
    /*/
User Function FSRSldCNH()

Local aAreOld       := {GetArea()} 
Private dDataIni  	:= ''
Private dDataFim  	:= ''
Private aSelFil     := {}

If FPergunte('FSRSldCNH')
    dDataIni	:= MV_PAR01
    dDataFim	:= MV_PAR02

    //GESTAO - inicio
    If MV_PAR03 == 1
        If Empty(aSelFil)
            aSelFil := AdmGetFil(.F., .F., "SD1")
            If Empty(aSelFil)
                AAdd(aSelFil, cFilAnt)
            EndIf
        EndIf
    Else
        AAdd(aSelFil, cFilAnt)
    EndIf	

    Processa({|| reportDef() })

  EndIf 

aEval(aAreOld, {|xAux| RestArea(xAux)})
    
Return Nil


//-------------------------------------------------------------------
/*/{Protheus.doc} reportDef
Funcao para chamada das definicoes do relatorio TReport

@protected
@author    Alex Teixeira de Souza
@since     23/07/2020
@obs      
@param	    	
            			
Alteracoes Realizadas desde a Estruturacao Inicial
Data       Programador     Motivo
/*/
//------------------------------------------------------------------- 
static function reportDef()
	Local oReport
	Local oSection1
	Local cTitulo	:= "Relatorio Saldo CNH "
    Local cPictVal  := "@E 999,999,999.99"

    Private cAliasQry   		:= GetNextAlias()

	oReport := TReport():New('FSRSldCNH', cTitulo, , {|oReport| PrintReport(oReport)},"Impressão do Relatório")
	oreport:nfontbody	:= 7
    oReport:DisableOrientation(.T.)
	oReport:SetLandscape()
	oReport:SetTotalInLine(.F.)
	oReport:ShowHeader()	

    oSection1 := TRSection():New(oReport,"SALDO CNH",cAliasQry) 
    oSection1:LAUTOSIZE := .T.

    TRCell():New(oSection1,"D1_FILIAL",,"Filial",,TamSx3("D1_FILIAL")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"D1_DOC",,"NF",,TamSx3("D1_DOC")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"B1_ZCODCNH",,"Cod CNH",,50,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"B1_COD",,"Cod Protheus",,TamSx3("B1_COD")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"B1_DESC",,"Descricao",,TamSx3("B1_DESC")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"QTDENT",,"Qtd Entrada",cPictVal,TamSx3("D1_QUANT")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"QTDSAI",,"Qtd Saida",cPictVal,TamSx3("D2_QUANT")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"QTDSALDO",,"Saldo",cPictVal,TamSx3("D2_QUANT")[1],/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )

 	oReport:PrintDialog()
	
	If (Select(cAliasQry)!= 0)
		dbSelectArea(cAliasQry)
		dbCloseArea()
	
		If File(cAliasQry+GetDBExtension())
			FErase(cAliasQry+GetDBExtension())
		EndIf
	EndIf


return (oReport)

Static Function PrintReport(oReport)
	Local oSection1   
    Local lProc			:= .f.
	Local nQtdRec		:= 0

    oSection1   	:= 	oReport:Section(1)

	DBSelectArea("SD1")
	SD1->(DBSetOrder(1))

	DBSelectArea("SD2")
	SD2->(DBSetOrder(1))

    Processa({|| lProc := FSeleDados()})//,"Aguarde...","Coletando Dados..."

	If lProc

		//Define a mensagem apresentada durante a geração do relatório.
		oReport:SetMsgPrint(Space(30)+"Lendo Registros")

		//Seta o contador da regua   
		(cAliasQry)->(dbGoTop())          
		(cAliasQry)->(DBEval({||  nQtdRec++})) 

		oReport:SetMeter(nQtdRec)

		//seleciono o arquivo de trabalho gerado pela query e coloco no inicio
		dbSelectArea(cAliasQry)
		(cAliasQry)->(dbGoTop())          

		oSection1:Init()	

		Do While !(cAliasQry)->(Eof()) .and. !oReport:Cancel()

			If oReport:Cancel()
				Exit
			EndIf

			oReport:IncMeter()

            oSection1:Cell("D1_FILIAL"):Setvalue( (cAliasQry)->D1_FILIAL )
            oSection1:Cell("D1_DOC"):Setvalue( (cAliasQry)->D1_DOC )
            oSection1:Cell("B1_ZCODCNH"):Setvalue( (cAliasQry)->B1_ZCODCNH )
			oSection1:Cell("B1_COD"):Setvalue( (cAliasQry)->D1_COD )
          	oSection1:Cell("B1_DESC"):Setvalue( (cAliasQry)->B1_DESC )
          	oSection1:Cell("QTDENT"):Setvalue( (cAliasQry)->D1_QUANT )
          	oSection1:Cell("QTDSAI"):Setvalue( (cAliasQry)->D2_QUANT )
			oSection1:Cell("QTDSALDO"):Setvalue( (cAliasQry)->D1_QUANT - (cAliasQry)->D2_QUANT )

            oSection1:PrintLine()

			(cAliasQry)->(DBSkip())	

		EndDo   

		oSection1:Finish() 

	Endif

Return


Static Function FSeleDados()
	/***********************************************
	*  Seleciona os titulos marcados pelo usuario. *
	***********************************************/
	Local cQuery		:= ""
	Local lRet  		:= .t.
	Local cTmpSD1Fil    := ""

	If Select(cAliasQry) <> 0
		dbSelectArea(cAliasQry)
		dbCloseArea()
	EndIf

	cQuery := " SELECT  D1_FILIAL, B1_ZCODCNH, D1_DOC, D1_COD,  B1_DESC,  D1_QUANT, SUM(D2_QUANT) D2_QUANT FROM ( "+CRLF
	cQuery += " SELECT D1_FILIAL, B1_ZCODCNH, D1_DOC, D1_COD,  B1_DESC, D1_QUANT, ISNULL(D2_QUANT,0) D2_QUANT FROM " + RetSQLName("SD1") + " SD1  " +CRLF
	cQuery += " INNER JOIN " + RetSQLName("SB1") + " SB1 ON SD1.D1_FILIAL = SB1.B1_FILIAL AND SD1.D1_COD = SB1.B1_COD AND SB1.D_E_L_E_T_ <> '*'  " +CRLF
	cQuery += " LEFT OUTER JOIN " + RetSQLName("SD2") + " SD2 ON SD1.D1_FILIAL = SD2.D2_FILIAL AND SD1.D1_DOC = SD2.D2_NFORI AND SD1.D1_SERIE = SD2.D2_SERIORI AND SD1.D1_COD = SD2.D2_COD AND SD2.D_E_L_E_T_ <> '*'  " +CRLF
	cQuery += " WHERE  " +CRLF
	Do Case
		Case !Empty(MV_PAR01) .AND. SB1->(FieldPos("B1_ZCODCNH")) > 0
			cQuery += "SB1.B1_ZCODCNH = '"+Alltrim(MV_PAR01)+"' AND "+CRLF 
		Case !Empty(MV_PAR02)
			cQuery += "SB1.B1_COD = '"+Alltrim(MV_PAR02)+"' AND "+CRLF 
		OtherWise
			cQuery += " '3' = '2' AND "+CRLF 
	EndCase	
	cQuery += "SD1.D1_FILIAL " + GetRngFil( aSelFil, "SD1", .T., @cTmpSD1Fil ) + "  AND "+CRLF 
	cQuery += "SD1.D_E_L_E_T_ = ' ' "
	cQuery += ") DADOS "+CRLF 
	cQuery += "GROUP BY D1_FILIAL, B1_ZCODCNH, D1_DOC, D1_COD,  B1_DESC, D1_QUANT "+CRLF        
	cQuery += "ORDER BY D1_FILIAL, D1_DOC, D1_COD  "
	cQuery := ChangeQuery(cQuery)

	dbUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),cAliasQry,.T.,.T.)     

	If (cAliasQry)->(Eof())
		lRet := .f.
	Endif		

Return lRet


/*/{Protheus.doc} FPergunte
Perguntas da rotina
@author rafaelalmeida
@since 10/06/2019
@version 1.0
@return ${return}, ${return_description}

@param lView, logical, descricao
@param lEdit, logical, descricao
@type function
/*/
Static Function FPergunte(cNomRot,lView,lEdit)

Local aParambox	:= {}
Local aRet 		:= {}
Local cProdCNH	:= Space(50)
Local cProduto	:= Space(TamSx3("B1_COD")[1])

Local lRet		:= .T.
Local nX		:= 0
Local nSelFil	:= 1

Default cNomRot	:= "FSRSLDCNH"
Default lView	:= .T.
Default lEdit	:= .T.

Private lWhen	:= lEdit
     //01
aAdd( aParambox ,{1,"Cod CNH"			,cProdCNH,"",".T.",,".T.",80,.F.})       //01
aAdd( aParambox ,{1,"Cod Protheus"		,cProduto,"",".T.","SB1",".T.",80,.F.})       //01
aAdd( aParambox, {2,"Seleciona Filial"	,nSelFil, {"1=Sim","2=Nao"},  70, ".T.", .T.})

//Carrega o array com os valores utilizados na Ãºltima tela ou valores default de cada campo.
For nX := 1 To Len(aParamBox)
	aParamBox[nX][3] := ParamLoad(cNomRot,aParamBox,nX,aParamBox[nX][3])
Next nX

//Define se ira apresentar tela de perguntas
If lView
	lRet := ParamBox(aParamBox,"Parametros",aRet,{|| .T.},{},.T.,Nil,Nil,Nil,cNomRot,.F.,.F.)
Else
	For nX := 1 To Len(aParamBox)
		Aadd(aRet, aParamBox[nX][3])
	Next nX
EndIf

If lRet
	//Carrega perguntas em variaveis usadas no programa
	If ValType(aRet) == "A" .And. Len(aRet) == Len(aParamBox)
		For nX := 1 to Len(aParamBox)
			If aParamBox[nX][1] == 2 .And. ValType(aRet[nX]) == "C"
				&("Mv_Par"+StrZero(nX,2)) := aScan(aParamBox[nX][4],{|x| Alltrim(x) == aRet[nX]})
			ElseIf aParamBox[nX][1] == 2 .And. ValType(aRet[nX]) == "N"
				&("Mv_Par"+StrZero(nX,2)) := aRet[nX]
			Else
				&("Mv_Par"+StrZero(nX,2)) := aRet[nX]
			Endif
		Next nX
	EndIf

	If lEdit
		//Salva parametros
		ParamSave(cNomRot,aParamBox,"1")
	EndIf
EndIf

Return(lRet)
