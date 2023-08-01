#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TBICONN.CH" 
#INCLUDE "REPORT.CH"
 
/*/{Protheus.doc} FSRFisFin
    
    RelatÃ³rio Fiscal / Financeiro
    
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
User Function FSRFisFin()

Local aAreOld       := {GetArea()} 
Private dDataIni  	:= ''
Private dDataFim  	:= ''
Private aSelFil     := {}

If FPergunte('FSRFINFIS')
    dDataIni	:= MV_PAR01
    dDataFim	:= MV_PAR02

    //GESTAO - inicio
    If MV_PAR03 == 1
        If Empty(aSelFil)
            aSelFil := AdmGetFil(.F., .F., "SE1")
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
	Local cTitulo	:= "Relação de Titulos  "+Dtoc(dDataIni)+" a "+Dtoc(dDataFim)
    Local cPictVal  := "@E 999,999,999.99"

    Private cAliasQry   		:= GetNextAlias()

	oReport := TReport():New('FSRFISFIN', cTitulo, , {|oReport| PrintReport(oReport)},"Impressão do Relatório")
	oreport:nfontbody	:= 7
    oReport:DisableOrientation(.T.)
	oReport:SetLandscape()
	oReport:SetTotalInLine(.F.)
	oReport:ShowHeader()	

    oSection1 := TRSection():New(oReport,"Relação de Titulos",cAliasQry) 
    //oSection1:LAUTOSIZE := .T.

    TRCell():New(oSection1, "FILIAL"			,, "Filial",, TamSX3("C7_FILIAL")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "PREFIXO"			,, "Prefixo",, TamSX3("E1_PREFIXO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "NUM"				,, "Num Titulo",, TamSX3("E1_NUM")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "TIPO"				,, "TP",, TamSX3("E1_TIPO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "CLIFOR"			,, "Cliente",, TamSX3("E1_CLIENTE")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "LOJA"				,, "Loja",, TamSX3("E1_LOJA")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "CNPJ"				,, "CNPJ",, 20           , .F.)        //"Filial
    TRCell():New(oSection1, "NOME"				,, "Nome",, TamSX3("A1_NOME")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "NATUREZA"			,, "Natureza",, TamSX3("ED_CODIGO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "DESCNAT"    		,, "Descricao",, TamSX3("ED_DESCRIC")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "EMISSAO"    		,, "Emissao",, TamSX3("E1_EMISSAO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "VENCTO"    		,, "Vencto",, TamSX3("E1_VENCREA")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "E5_HISTOR"			,, "Historico",, TamSX3("E5_HISTOR")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "E5_DATA"			,, "Dt Baixa",, TamSX3("E5_DATA")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "VALORIG"			,, "Valor Original",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "JURMULTA"			,, "Jur/Multa",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "CORRECAO"			,, "Correcao",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "DESCONTOS"			,, "Descontos",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "PIS"				,, "Pis",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "COFINS"			,, "Cofins",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "CSLL"				,, "Csll",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "IRF"				,, "IRF",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "ISS"				,, "ISS",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "INSS"				,, "INSS",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "VALACESS"			,, "Valor Acessorio",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "TOTBAIXADO"		,, "Total Baixado",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "BCO"				,, "Banco",, 5           , .F.)        //"Filial
    TRCell():New(oSection1, "DTDIGIT"			,, "Dt Digit",, TamSX3("E5_DATA")[1]       , .F.)        //"Filial
    TRCell():New(oSection1, "MOT"				,, "Motivo",, 5           , .F.)        //"Filial
	
 
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

	DBSelectArea("SE1")
	SE1->(DBSetOrder(1))

    Processa({|| lProc := FSeleDados()})//,"Aguarde...","Coletando Dados..."

	If lProc

		//Define a mensagem apresentada durante a geração do relatório.
		//oReport:SetMsgPrint(Space(30)+"Lendo Registros")

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

			aDados := {0,"",ctod(""),"",""}

			SE5->(DBGoto((cAliasQry)->E5RECNO))


			oSection1:Cell("FILIAL"):SetValue(SE5->E5_FILIAL )
			oSection1:Cell("PREFIXO"):SetValue(SE5->E5_PREFIXO )
			oSection1:Cell("NUM"):SetValue(SE5->E5_NUMERO )
			oSection1:Cell("TIPO"):SetValue(SE5->E5_TIPO )
			oSection1:Cell("CLIFOR"):SetValue(SE5->E5_CLIFOR )
			oSection1:Cell("LOJA"):SetValue(SE5->E5_LOJA )


			cNomCli := Posicione("SA1",1,FwxFilial("SA1",SE5->E5_FILIAL)+SE5->E5_CLIFOR+SE5->E5_LOJA,"A1_NOME")
			cCNPJ	:= Posicione("SA1",1,FwxFilial("SA1",SE5->E5_FILIAL)+SE5->E5_CLIFOR+SE5->E5_LOJA,"A1_CGC")
			oSection1:Cell("NOME"):SetValue(cNomCli )	
			oSection1:Cell("CNPJ"):SetValue(cCNPJ )		

			oSection1:Cell("NATUREZA"):SetValue(SE5->E5_NATUREZ )
			oSection1:Cell("DESCNAT"):SetValue(Posicione("SED",1,xFilial("SED")+SE5->E5_NATUREZ,"ED_DESCRIC"))
			oSection1:Cell("E5_DATA"):SetValue(SE5->E5_DATA )
			oSection1:Cell("TOTBAIXADO"):SetValue(SE5->E5_VALOR )
			oSection1:Cell("BCO"):SetValue(SE5->E5_BANCO )
			oSection1:Cell("DTDIGIT"):SetValue( SE5->E5_DTDIGIT )
			oSection1:Cell("MOT"):SetValue( SE5->E5_MOTBX )
			oSection1:Cell("E5_HISTOR"):SetValue( SE5->E5_HISTOR )

			If (nRecSE1 := BuscaSE1()) > 0

				SE1->(DBGoto(nRecSE1))

				oSection1:Cell("VENCTO"):SetValue(DtoC(SE1->E1_VENCREA) )
				oSection1:Cell("EMISSAO"):SetValue(DtoC(SE1->E1_EMISSAO) )
				oSection1:Cell("VALORIG"):SetValue(SE1->E1_VALOR )
				oSection1:Cell("PIS"):SetValue(SE1->E1_PIS )
				oSection1:Cell("COFINS"):SetValue(SE1->E1_COFINS )
				oSection1:Cell("CSLL"):SetValue(SE1->E1_CSLL )
				oSection1:Cell("IRF"):SetValue(SE1->E1_IRRF )
				oSection1:Cell("ISS"):SetValue(SE1->E1_ISS )
				oSection1:Cell("INSS"):SetValue(SE1->E1_INSS )
				oSection1:Cell("VALACESS"):SetValue(0)	
				oSection1:Cell("JURMULTA"):SetValue(SE1->E1_MULTA )			
				oSection1:Cell("CORRECAO"):SetValue(SE1->E1_ACRESC )
			Else				
				oSection1:Cell("VENCTO"):SetValue( Nil )
				oSection1:Cell("EMISSAO"):SetValue(Nil )
				oSection1:Cell("VALORIG"):SetValue( Nil )
				oSection1:Cell("PIS"):SetValue( Nil )
				oSection1:Cell("COFINS"):SetValue( Nil )
				oSection1:Cell("CSLL"):SetValue( Nil )
				oSection1:Cell("IRF"):SetValue( Nil )
				oSection1:Cell("ISS"):SetValue( Nil )
				oSection1:Cell("INSS"):SetValue( Nil )
				oSection1:Cell("VALACESS"):SetValue( Nil )	
				oSection1:Cell("JURMULTA"):SetValue( Nil )			
				oSection1:Cell("CORRECAO"):SetValue( Nil )


			Endif

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
	Local cQryFIN		:= ""
	Local lRet  		:= .t.
	Local cTmpSE5Fil    := ""

	If Select(cAliasQry) <> 0
		dbSelectArea(cAliasQry)
		dbCloseArea()
	EndIf

	cQryFIN += "  SELECT E5.R_E_C_N_O_ E5RECNO "+CRLF			
	cQryFIN += "      FROM "+RetSqlName("SE5")+" E5 "+CRLF
	cQryFIN += "      WHERE "+CRLF
	cQryFIN += "      E5.D_E_L_E_T_ <> '*' "+CRLF
	cQryFIN += "	  AND E5.E5_SITUACA <> 'C' "+CRLF
	cQryFIN += "	  AND E5.E5_TIPODOC <> 'ES' "+CRLF				
	cQryFIN += "	  AND E5.E5_TIPODOC <> 'JR' "+CRLF
	cQryFIN += "      AND E5_DATA BETWEEN '"+DTOS(MV_PAR01)+"' AND '"+DTOS(MV_PAR02)+"' "
	cQryFIN += "      AND E5_MOTBX <> '0'  "+CRLF
	cQryFIN += "      AND E5_FILIAL " + GetRngFil( aSelFil, "SE5", .T., @cTmpSE5Fil ) + " "
	cQryFIN += "      AND E5_RECPAG = 'R' "+CRLF
	cQryFIN += "      AND E5_MOTBX <> 'FAT' "+CRLF
	cQryFIN += "      AND E5_TIPO NOT IN " + FormatIn( MVISS +"|"+ MVIRF+"|"+ MVTAXA +"|"+ MVTXA +"|"+ MVINSS +"|"+ 'SES' +"|"+ 'CID' + "|"+ 'INA', "|") + "  "
	cQryFIN += "	  AND NOT EXISTS (SELECT  CAN.E5_NUMERO FROM "+RetSqlName("SE5")+" CAN "+CRLF
	cQryFIN += "			WHERE   CAN.E5_FILIAL = E5.E5_FILIAL AND "+CRLF
	cQryFIN += "			CAN.E5_NUMERO = E5.E5_NUMERO AND "+CRLF
	cQryFIN += "			CAN.E5_PARCELA = E5.E5_PARCELA AND "+CRLF
	cQryFIN += "			CAN.E5_TIPO = E5.E5_TIPO AND "+CRLF			
	cQryFIN += "			CAN.E5_CLIFOR = E5.E5_CLIFOR AND "+CRLF
	cQryFIN += "			CAN.E5_LOJA = E5.E5_LOJA AND "+CRLF
	cQryFIN += "			CAN.E5_SEQ = E5.E5_SEQ AND "+CRLF
	cQryFIN += "			CAN.E5_TIPODOC = 'ES' AND CAN.D_E_L_E_T_ = E5.D_E_L_E_T_ ) "+CRLF			

	cQryFIN := ChangeQuery(cQryFIN)

	dbUseArea(.T.,"TOPCONN",TCGENQRY(,,cQryFIN),cAliasQry,.T.,.T.)     

	If (cAliasQry)->(Eof())
		lRet := .f.
	Endif		

Return lRet

Static Function BuscaSE1()
	Local cAliasQry 	:= GetNextAlias()
	Local cQuery        := ""
	Local nRecSE1		:= 0

	cQuery := " SELECT SE1.R_E_C_N_O_ SE1RENO "
	cQuery += " FROM "+RetSqlName("SE1")+ " SE1 "
	cQuery += " WHERE SE1.E1_FILIAL =  '"+SE5->E5_FILIAL+"' "
	cQuery += " AND SE1.E1_PREFIXO = '"+SE5->E5_PREFIXO+"' "
	cQuery += " AND SE1.E1_NUM = '"+SE5->E5_NUMERO+"' "
	cQuery += " AND SE1.E1_PARCELA = '"+SE5->E5_PARCELA+"' "
	cQuery += " AND SE1.E1_TIPO = '"+SE5->E5_TIPO+"' "
	cQuery += " AND SE1.E1_CLIENTE = '"+SE5->E5_CLIFOR+"' "
	cQuery += " AND SE1.E1_LOJA = '"+SE5->E5_LOJA+"' "
	cQuery += " ORDER BY R_E_C_N_O_ DESC "

	cQuery := ChangeQuery(cQuery)

	dbUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),cAliasQry,.T.,.T.)

	If  !(cAliasQry)->(Eof())
		nRecSE1 := (cAliasQry)->SE1RENO
	Endif

	(cAliasQry)->(dbCloseArea())

Return nRecSE1

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

Local lRet		:= .T.
Local nX		:= 0
Local nSelFil	:= 1

Default cNomRot	:= "FSRFINFIS"
Default lView	:= .T.
Default lEdit	:= .T.

Private lWhen	:= lEdit

aAdd( aParambox ,{1,"Baixa de  : "      , Ctod(Space(8)),,".T.",,"lWhen",50,.T.})	
aAdd( aParambox ,{1,"Baixa Ate : "      , Ctod(Space(8)),,".T.",,"lWhen",50,.T.})	
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
				&("Mv_Par"+StrZero(nX,2)) := aRet[nX]
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

If ValType(MV_PAR03) != "N"
    MV_PAR03 := vAL(MV_PAR03)
Endif   


Return(lRet)
