#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TBICONN.CH" 
#INCLUDE "REPORT.CH"
 
/*/{Protheus.doc} FSRPedCom
    
    RelatÃ³rio de Pedido de Compras
    
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
User Function FSRPedCom()

Local aAreOld       := {GetArea()} 
Private dDataIni  	:= ''
Private dDataFim  	:= ''
Private dDigitIni  	:= ''
Private dDigitFim  	:= ''
Private aSelFil     := {}

If FPergunte('FSRPEDCOM')
    dDataIni	:= MV_PAR01
    dDataFim	:= MV_PAR02
    dDigitIni	:= MV_PAR03
    dDigitFim	:= MV_PAR04
    //GESTAO - inicio
    If MV_PAR05 == 1
        If Empty(aSelFil)
            aSelFil := AdmGetFil(.F., .F., "SC7")
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
	Local cTitulo	:= "Relatorio Pedido de Compras "
    Local cPictVal  := "@E 999,999,999.99"

    Private cAliasQry   		:= GetNextAlias()

	oReport := TReport():New('FSRPEDCOM', cTitulo, , {|oReport| PrintReport(oReport)},"Impressão do Relatório")
	oreport:nfontbody	:= 7
    oReport:DisableOrientation(.T.)
	oReport:SetLandscape()
	oReport:SetTotalInLine(.F.)
	oReport:ShowHeader()	

    oSection1 := TRSection():New(oReport,"Relatorio Pedido de Compras ",cAliasQry) 
    //oSection1:LAUTOSIZE := .T.

    TRCell():New(oSection1, "C7_FILIAL"		,, "Filial",, TamSX3("C7_FILIAL")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_PRODUTO"	,, "Produto",, TamSX3("C7_PRODUTO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_DESCRI"		,, "Descricao",, TamSX3("C7_DESCRI")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "B1_CONTA"		,, "Conta 1",, TamSX3("B1_CONTA")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "CT1_DESC1"		,, "Desc. Conta 1",, TamSX3("CT1_DESC01")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "B1_ZCONT2"		,, "Conta 2",, TamSX3("B1_ZCONT2")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "CT1_DESC2"		,, "Desc. Conta 2",, TamSX3("CT1_DESC01")[1]           , .F.)        //"Filial
   	TRCell():New(oSection1, "B1_GRUPO"		,, "Grupo",, TamSX3("B1_GRUPO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "BM_DESC"		,, "Desc. Grupo",, TamSX3("BM_DESC")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "A2_CGC"		,, "Fornecedor",, TamSX3("A2_CGC")[1]           , .F.)        //"Filial
   	TRCell():New(oSection1, "A2_NOME"		,, "Nome",, TamSX3("A2_NOME")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_UM"			,, "UM",, TamSX3("C7_UM")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_QUANT"		,, "Quantidade",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_PRECO"		,, "Prc Unitario",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_TOTAL"		,, "Vlr Total",cPictVal, 15           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_CC"			,, "C.Custo",, TamSX3("C7_CC")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "CTT_DESC01"	,, "Desricao",, TamSX3("CTT_DESC01")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "CTT_ZGESTO"	,, "Gestor",, TamSX3("CTT_ZGESTO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_ZPLACA"		,, "Placa",, TamSX3("C7_ZPLACA")[1]           , .F.)        //"Placa
    TRCell():New(oSection1, "C7_ZTMNT"		,, "Tipo Manut",, 20           , .F.)        //"Placa
    TRCell():New(oSection1, "C7_ZVRREEM"	,, "Reembolso",cPictVal, 15           , .F.)        //"Placa
    TRCell():New(oSection1, "C7_OBS"		,, "OBS",, TamSX3("C7_OBS")[1]           , .F.)        //"Placa
    TRCell():New(oSection1, "C7_NUM"		,, "Numero PC",, TamSX3("C7_NUM")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "C7_EMISSAO" 	,, "Data Pedido",, TamSX3("C7_EMISSAO")[1]           , .F.)        //"Filial
  	TRCell():New(oSection1, "USUPED"      	,, "Usuario PC"        ,, 15                                ,.F.)                           //"Mot"
  	TRCell():New(oSection1, "NOMEPED"     	,, "Nome Usuario"        ,, 50                                ,.F.)                           //"Mot"  
  	TRCell():New(oSection1, "USUARIO"      	,, "Usuario"        ,, 15                                ,.F.)                           //"Mot"
  	TRCell():New(oSection1, "NOME"         	,, "Nome Completo"        ,, 50                                ,.F.)                           //"Mot"
  	TRCell():New(oSection1, "DATALIB"      	,, "Data Liberacao"        ,, TamSX3("D1_EMISSAO")[1]     ,.F.)                           //"Mot"
    TRCell():New(oSection1, "D1DOC" 	   	,, "Num NF",, TamSX3("D1_DOC")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "D1EMISSAO"    	,, "Emissao NF",, TamSX3("D1_EMISSAO")[1]           , .F.)        //"Filial
    TRCell():New(oSection1, "D1DTDIGIT"    	,, "Digitacao NF",, TamSX3("D1_DTDIGIT")[1]           , .F.)        //"Filial
 
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

	DBSelectArea("SB1")
	SB1->(DBSetOrder(1))

	DBSelectArea("SD1")
	SD1->(DBSetOrder(1))

	DBSelectArea("SC7")
	SC7->(DBSetOrder(1))
	
	DBSelectArea("SCR")
	SCR->(DBSetOrder(1))

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

			SC7->(DBGoto((cAliasQry)->SC7RECNO))

			oSection1:Cell("A2_CGC"):SetValue(Posicione("SA2",1,xFilial("SA2")+SC7->C7_FORNECE+SC7->C7_LOJA,"A2_CGC" ))
			oSection1:Cell("A2_NOME"):SetValue(Posicione("SA2",1,xFilial("SA2")+SC7->C7_FORNECE+SC7->C7_LOJA,"A2_NOME" ))
			oSection1:Cell("CTT_DESC01"):SetValue(Posicione("CTT",1,xFilial("CTT")+SC7->C7_CC,"CTT_DESC01" ))
			oSection1:Cell("CTT_ZGESTO"):SetValue(Posicione("CTT",1,xFilial("CTT")+SC7->C7_CC,"CTT_ZGESTO" ))		
			
			oSection1:Cell("C7_ZPLACA"):SetValue(SC7->C7_ZPLACA )

			cTpManut := ""
			Do Case 
				Case SC7->C7_ZTMNT == "1"
					cTpManut := "Corretiva"	
				Case SC7->C7_ZTMNT == "2"
					cTpManut := "Preventiva"	
				Case SC7->C7_ZTMNT == "3"
					cTpManut := "Colisao"	
				Case SC7->C7_ZTMNT == "4"
					cTpManut := "Estoque"	
				Case SC7->C7_ZTMNT == "5"
					cTpManut := "Acessorios"	
				Case SC7->C7_ZTMNT == "6"
					cTpManut := "Sinistros"	
				Case SC7->C7_ZTMNT == "7"
					cTpManut := "Outros"	
				Case SC7->C7_ZTMNT == "8"
					cTpManut := "N MNT"	
				Case SC7->C7_ZTMNT == "9"
					cTpManut := "Pneus"	
				Case SC7->C7_ZTMNT == "A"
					cTpManut := "Implementos"	
			EndCase

			oSection1:Cell("C7_ZTMNT"):SetValue(cTpManut)
			oSection1:Cell("C7_ZVRREEM"):SetValue(SC7->C7_ZVRREEM )
			oSection1:Cell("C7_OBS"):SetValue(Alltrim(SC7->C7_OBS) )

			cDscSrv := Alltrim(SC7->C7_DESCRI)
			If Empty(cDscSrv)
				cDscSrv := Alltrim(Posicione("SB1",1,xFilial("SB1")+SC7->C7_PRODUTO,"B1_DESC"))
			Endif	

			oSection1:Cell("C7_FILIAL"):SetValue(SC7->C7_FILIAL )
			oSection1:Cell("C7_PRODUTO"):SetValue(SC7->C7_PRODUTO )

			If SB1->(DBSeek(xFilial("SB1")+SC7->C7_PRODUTO ))
				oSection1:Cell("B1_CONTA"):SetValue(SB1->B1_CONTA )
				oSection1:Cell("CT1_DESC1"):SetValue(Posicione("CT1",1,xFilial("CT1")+SB1->B1_CONTA,"CT1_DESC01") )
				oSection1:Cell("B1_ZCONT2"):SetValue(SB1->B1_ZCONT2 )
				oSection1:Cell("CT1_DESC2"):SetValue(Posicione("CT1",1,xFilial("CT1")+SB1->B1_ZCONT2,"CT1_DESC01") )
				oSection1:Cell("B1_GRUPO"):SetValue(SB1->B1_GRUPO )
				oSection1:Cell("BM_DESC"):SetValue(Posicione("SBM",1,xFilial("SBM")+SB1->B1_GRUPO,"BM_DESC") )
			Else
				oSection1:Cell("B1_CONTA"):SetValue(Nil )
				oSection1:Cell("CT1_DESC1"):SetValue(Nil )
				oSection1:Cell("B1_ZCONT2"):SetValue(Nil )
				oSection1:Cell("CT1_DESC2"):SetValue(Nil )
				oSection1:Cell("B1_GRUPO"):SetValue(Nil )
				oSection1:Cell("BM_DESC"):SetValue(Nil )
			Endif

			oSection1:Cell("C7_DESCRI"):SetValue(cDscSrv )
			oSection1:Cell("C7_UM"):SetValue(SC7->C7_UM )
			oSection1:Cell("C7_QUANT"):SetValue(SC7->C7_QUANT )
			oSection1:Cell("C7_PRECO"):SetValue(SC7->C7_PRECO )
			oSection1:Cell("C7_TOTAL"):SetValue(SC7->C7_TOTAL )
			oSection1:Cell("C7_NUM"):SetValue(SC7->C7_NUM )
			oSection1:Cell("C7_EMISSAO"):SetValue(DtoC(SC7->C7_EMISSAO) )
			oSection1:Cell("USUPED"):SetValue(SC7->C7_USER )
			cUsuario := Alltrim(UsrFullName(SC7->C7_USER))
			oSection1:Cell("NOMEPED"):SetValue(cUsuario)
			oSection1:Cell("C7_CC"):SetValue(SC7->C7_CC )
  
			nRecSCR := FAprovador(SC7->C7_FILIAL,SC7->C7_NUM)

			iF nRecSCR <> 0 	
				SCR->(DBGoto(nRecSCR))

				cUsuario := Alltrim(UsrFullName(SCR->CR_USERLIB))

				oSection1:Cell("USUARIO"):SetValue(SCR->CR_USERLIB )
				oSection1:Cell("DATALIB"):SetValue(DtoC(SCR->CR_DATALIB) )
				oSection1:Cell("NOME"):SetValue(cUsuario )
			Else
				oSection1:Cell("USUARIO"):SetValue(NIL )
				oSection1:Cell("DATALIB"):SetValue(NIL )
				oSection1:Cell("NOME"):SetValue(NIL )
			Endif

			If (cAliasQry)->SD1RECNO <> 0 

				SD1->(DBGoto((cAliasQry)->SD1RECNO))

				oSection1:Cell("D1DOC"):SetValue(SD1->D1_DOC )
				oSection1:Cell("D1EMISSAO"):SetValue(DtoC(SD1->D1_EMISSAO) )
				oSection1:Cell("D1DTDIGIT"):SetValue(DtoC(SD1->D1_DTDIGIT) )
			Else
				oSection1:Cell("D1DOC"):SetValue(Nil )
				oSection1:Cell("D1EMISSAO"):SetValue(Nil )
				oSection1:Cell("D1DTDIGIT"):SetValue(Nil )
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
	Local cQrySC7		:= ""
	Local lRet  		:= .t.
	Local cTmpSE2Fil    := ""

	If Select(cAliasQry) <> 0
		dbSelectArea(cAliasQry)
		dbCloseArea()
	EndIf

	cQrySC7 += "SELECT SC7.R_E_C_N_O_ SC7RECNO, ISNULL(SD1.R_E_C_N_O_,0) SD1RECNO "+CRLF
	cQrySC7 += "FROM " + RetSqlName("SC7") + " SC7 "+CRLF
	cQrySC7 += "LEFT OUTER JOIN " + RetSqlName("SD1") + " SD1 ON SC7.C7_FILIAL = SD1.D1_FILIAL AND SC7.C7_NUM = SD1.D1_PEDIDO AND SC7.C7_ITEM = SD1.D1_ITEMPC  AND SD1.D_E_L_E_T_ <> '*' "+CRLF
	cQrySC7 += "WHERE 
	cQrySC7 += " SC7.C7_FILIAL " + GetRngFil( aSelFil, "SC7", .T., @cTmpSE2Fil ) + " AND "
	If !Empty(dDataIni) .and. !Empty(dDataFim)
		cQrySC7 += " SC7.C7_EMISSAO  >= '"+DTOS(dDataIni)+"' AND  SC7.C7_EMISSAO  <= '"+DTOS(dDataFim)+"' AND "+CRLF		
	Endif
	If !Empty(dDigitIni) .and. !Empty(dDigitFim)
		cQrySC7 += " SD1.D1_DTDIGIT  >= '"+DTOS(dDigitIni)+"' AND  SD1.D1_DTDIGIT  <= '"+DTOS(dDigitFim)+"' AND "+CRLF		
	Endif

	If (Empty(dDataIni) .or. Empty(dDataFim)) .and. (Empty(dDigitIni) .or. Empty(dDigitFim))
		cQrySC7 += " '1' = '2' AND "
	Endif

	cQrySC7 += " SC7.D_E_L_E_T_ <> '*' "

	cQrySC7 += "ORDER BY SC7.C7_FILIAL, SC7.C7_NUM "+ CRLF

	cQrySC7 := ChangeQuery(cQrySC7)

	dbUseArea(.T.,"TOPCONN",TCGENQRY(,,cQrySC7),cAliasQry,.T.,.T.)     

	If (cAliasQry)->(Eof())
		lRet := .f.
	Endif		

Return lRet


Static Function FAprovador(cFil,cPedido)
	/***********************************************
	*  Seleciona os titulos marcados pelo usuario. *
	***********************************************/
	Local cQrySRC		:= ""
	Local nRecno  		:= 0
	Local cAliasTMP		:= GetNextAlias()

	cQrySRC += "SELECT  SCR.R_E_C_N_O_ SCRRECNO "+CRLF
	cQrySRC += "FROM " + RetSqlName("SCR") + " SCR "+CRLF
	cQrySRC += "WHERE 
	cQrySRC += " SCR.CR_FILIAL = '"+cFil+"' AND "	
	cQrySRC += " SCR.CR_NUM = '"+cPedido+"' AND "	
	cQrySRC += " SCR.CR_TIPO = 'PC' AND  "	
	cQrySRC += " SCR.CR_STATUS = '03' AND "
	cQrySRC += " SCR.CR_USERLIB <> '' AND SCR.D_E_L_E_T_ <> '*' "

	cQrySRC := ChangeQuery(cQrySRC)

	dbUseArea(.T.,"TOPCONN",TCGENQRY(,,cQrySRC),cAliasTMP,.T.,.T.)     

	If !(cAliasTMP)->(Eof())
		nRecno :=  (cAliasTMP)->SCRRECNO
	Endif		
	
	(cAliasTMP)->(dbCloseArea())

Return nRecno



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

Default cNomRot	:= "FSRPEDCOM"
Default lView	:= .T.
Default lEdit	:= .T.

Private lWhen	:= lEdit

aAdd( aParambox ,{1,"Emissao Inicial : "      , Ctod(Space(8)),,".T.",,"lWhen",50,.F.})	
aAdd( aParambox ,{1,"Emissao Final   : "      , Ctod(Space(8)),,".T.",,"lWhen",50,.F.})	
aAdd( aParambox ,{1,"Digitacao Inicial : "     , Ctod(Space(8)),,".T.",,"lWhen",50,.F.})	
aAdd( aParambox ,{1,"Digitacao Final   : "     , Ctod(Space(8)),,".T.",,"lWhen",50,.F.})	
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

If ValType(MV_PAR05) != "N"
    MV_PAR05 := vAL(MV_PAR05)
Endif   

Return(lRet)
