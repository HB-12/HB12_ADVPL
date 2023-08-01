#INCLUDE "protheus.ch"
#INCLUDE "rwmake.ch"
#INCLUDE "FWMVCDEF.CH"
#INCLUDE "totvs.ch"

#DEFINE _SEPARADOR ";"
#DEFINE _BUFFER 16384

//-------------------------------------------------------------------
/*/{Protheus.doc} FSMNT001
Importacao KM

Parametros da rotina

Tabelas para atulizar

@author Alex Teixeira
@since 18/05/2022
@version 1.0
/*/
//-------------------------------------------------------------------
User Function FSMNT001()
	Local   aSay      := {}
	Local   aButton   := {}
	Local   cTitulo   := "IMPORTAÇÃO DE KM - Filial:"+ SM0->M0_CODIGO + " " + SM0->M0_NOME
	Local   cDesc1    := "Esta rotina tem como função realizar a importação do arquivo de KM"
	Local   cDesc2    := "Este processo ira gerar ao final a atualizacao de contadores referente as linha"
	Local   cDesc3    := "Após o processamento dever ser verificado os KM atualizados "
	Local   lOk       := .F.

	Private oProcess  := NIL
	Private _cArquiv  := ""
	Private _dataIni  := ddatabase
	Private _dataFim  := ddatabase
	Private cNomeArq  := ""
	Private _nCabec	:= 1
	Private lAchou    := .F.
	Private dDTLEIV       := dDataBase
	Private nREGST9       := 0
	Private cHORA1V       := Space(5)
	Private cHORA2V       := Space(5)
	Private nCONT1V       := 0
	Private nCONT2V       := 0
	Private dDTLEIV       := dDataBase
	Private cBemInF       := space(len(ST9->T9_CODBEM))
	Private cGERAPREV     := AllTrim(GETMv("MV_NGGERPR"))
	Private aLog		    := {}
	Private nINDST9		:= 0

	// Mensagens de Tela Inicial
	aAdd( aSay, cDesc1 )
	aAdd( aSay, cDesc2 )
	aAdd( aSay, cDesc3 )
//aAdd( aSay, cDesc4 )
//aAdd( aSay, cDesc5 )
//aAdd( aSay, cDesc6 )
//aAdd( aSay, cDesc7 )

// Botoes Tela Inicial
	aAdd(  aButton, {  5, .T.,{|| FPergunte() } } )
	aAdd(  aButton, {  1, .T., { || lOk := .T., FechaBatch() } } )
	aAdd(  aButton, {  2, .T., { || lOk := .F., FechaBatch() } } )

	FormBatch( cTitulo, aSay, aButton )

	If lOk

		If MsgNoYes( "Confirma o processamento ?", cTitulo )
			oProcess := MsNewProcess():New( { | lEnd | lOk := FSTProc( @lEnd,  ) }, "Atualizando", "Aguarde, atualizando ...", .F. )
			oProcess:Activate()
		Endif
	Endif

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} FSTProc
Processamento da Fatura

@author Alex Teixeira
@since 14/07/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function FSTProc( lEnd)
	Local cFile := _cArquiv
	Local aStru := {}

	Private oTmpTable

	IF  File(cFile) .And. ! Empty(cFile)

		AADD(aStru,{ "PLACA" ,"C"  , 8,  0})
		AADD(aStru,{ "DATAKM", "D",   8, 0})
		AADD(aStru,{ "KM", "N",   12, 2})

		oTmpTable := FWTemporaryTable():New('TRD')
		oTmpTable:SetFields( aStru )
		//crea indice
		oTmpTable:AddIndex('T1ORD', {'PLACA'})
		//Creacion de la tabla
		oTmpTable:Create()

		If LeArquivo(cFile,"TRD")

			If lAchou
				ImportaKM("TRD")
			else
				ApMsgInfo("Nenhum registro importado.","ATENÇÃO")
			EndIf

		Endif
	Endif

	If oTmpTable <> Nil
		oTmpTable:Delete()
		oTmpTable := Nil
	Endif

Return .t.


//-------------------------------------------------------------------
/*/{Protheus.doc} ImportaKM
Rotina para importacao KM 

@author Alex Teixeira
@since 15/03/2021
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function ImportaKM(cAlias)
	Local lRet 		:= .t.
	Local nCount 	:= 0
	Local cMsg      := ""
	
	Private cPlaca	:= ""

	dbSelectArea("ST9")
	ST9->(dbSetOrder(1))

	nINDST9       := ST9->(IndexOrd())

	(cAlias)->(dbeval({|| nCount++}))

	dbSelectArea(cAlias)
	(cAlias)->(dbGoTop())

	oProcess:SetRegua2(nCount)

	Do While ! (cAlias)->(Eof())

		cPlaca 	:= Alltrim((cAlias)->PLACA)
		cPlaca  := StrTran(cPlaca,chr(13),"")
		cPlaca  := StrTran(cPlaca,chr(10),"")

		aVeic 	:= BuscaVeic(cPlaca)

		If aVeic[2] > 0

			cFilAnt := aVeic[1]
			ST9->(DBGoto(aVeic[2] ))

			cMsg   	:= ""
			nCont1 	:= (cAlias)->KM
			nCont2 	:= 0
			dDtCont := (cAlias)->DATAKM
			cHora  	:= Substr(Time(),1,5)
			lImport := .t.


			If lImport
				If !FSMNT830ACE(nCont1,nCont2,cHora, dDtCont, @cMsg)
					Aadd(aLog, {.f.,cFilAnt, "Placa: "+cPlaca+" "+cMsg})
				Endif
			Endif
		Else
			Aadd(aLog, {.f.,cFilAnt, "Placa: "+cPlaca+" Veiculo não encontrado no sistema"})
		Endif
		(cAlias)->(dbSkip())
	EndDo

	MostraLog(aLog)

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} LeArquivo
Rotina de leitura do arquivo .txt para alimentação da tabela temporaria

@author Alex Teixeira
@since 18/05/2022
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function LeArquivo(cFile,cAlias)

	Local cBuffer := ""
	Local cArqExt := ""
	Local lRet    := .T.
	Local aLinha  := {}
	Local aLay	  := {}
	Local aLayout := LerLayout()
	Local nLinha  := 0
	Local nFor	  := 0
	Local nMaxCol := 0
	Local nXi
	Local oFile
	Local nHandle

	aLay    := Separa(aLayout[len(aLayout),2],_SEPARADOR)
	nMaxCol := Val(aLay[1])
	cFile   := AllTrim(cFile)
	cArqExt := Substr(cFile,1,(Len(cFile) - 4))

	// --- Montar o nome do arquivo
	// ----------------------------
	cNomeArq := ""

	For nXi := Len(cArqExt) to 1 Step -1
		If SubStr(cArqExt,nXi,1) == "/" .or. SubStr(cArqExt,nXi,1) == "\"
			exit
		EndIf

		cNomeArq := SubStr(cArqExt,nXi,1) + cNomeArq
	Next

	// --- Validar se o arquivo já foi importado
	// -----------------------------------------

	dbSelectArea(cAlias)
	(cAlias)->(dbGoTop())

	nHandle := FT_FUse(cFile)

	if nHandle = -1
		MsgAlert("Arquivo não pode ser aberto")
		Return .f.
	endif

	FT_FGoTop()

	nFor := FT_FLastRec()
	nFor := nFor*4

	// Fecha o Arquivo
	FT_FUSE()

	oFile := FSReadTXT():New(cFile,Chr(13))

	If ! oFile:Open()
		MsgAlert("Arquivo não pode ser aberto")
		Return .F.
	EndIf

	oProcess:SetRegua1(nFor)

	cBuffer := ""

	While oFile:ReadLine(@cBuffer)
		oProcess:IncRegua1( "Lendo Arquivo..." )

		cBuffer := Alltrim(cBuffer)
		aLinha  := {}
		nLinha++
		aLinha := separa(cBuffer,_SEPARADOR)

		if _nCabec == 1 .and. nLinha == 1
			cBuffer := ""
			Loop
		Else

			lAchou := .T.

			(cAlias)->( DBAppend( .F. ) )

			For nXi := 1 To (cAlias)->(FCount())
				If len(aLinha) != nMaxCol
					MsgAlert("Importacao Cancelada. Linha "+ Alltrim(Str(nLinha))+" do arquivo com problema na estrutura")
					Return .F.
				Else
					cCampo := TRD->(Field(nXi))
					xValor := DadosCampo(aLayout,aLinha,cCampo)

					(cAlias)->(FieldPut(nXi,xValor))
				Endif
			Next
			(cAlias)->( DBCommit() )
		Endif

	enddo

	(cAlias)->(dbGoTop())

	// Fecha o Arquivo
	oFile:Close()

Return lRet



//-------------------------------------------------------------------
/*/{Protheus.doc} LerLayout
Leitura do layout de importacao do Arquivo

@author Alex Teixeira
@since 24/07/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function LerLayout()
	Local aLayout   := {}

	aadd(aLayout,{"PLACA","1;C;8;0"})
	aadd(aLayout,{"DATAKM","2;D;8;0"})
	aadd(aLayout,{"KM","3;N;12;2"})

Return aLayout

//-------------------------------------------------------------------
/*/{Protheus.doc} DadosCampo;
Leitura do layout de importacao do Arquivo

@author Alex Teixeira
@since 24/07/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function DadosCampo(aLayout,aLinha,cCampo)
	Local xRet
	Local nPos
	Local aLay := {}
	Local nPosLinha
	Local cTipo
	Local nTam
	Local nDec

	If (nPos := Ascan(aLayout,{|aEle| Upper(Alltrim(aEle[1])) == Upper(Alltrim(cCampo)) })) > 0
		aLay := Separa(aLayout[nPos,2],_SEPARADOR)

		nPosLinha := Val(aLay[1])
		cTipo     := aLay[2]
		nTam      := Val(aLay[3])

		If len(aLay) >= 4
			nDec := Val(aLay[4])
		else
			nDec := 0
		Endif

		Do Case
		Case cTipo == "C" // Campo Caracter
			xRet := Substr(aLinha[nPosLinha],1,nTam)

		Case cTipo == "D" // Campo Data
			xRet := CtoD(aLinha[nPosLinha])

		Case cTipo == "N"  .and. nDec <= 0 // Campo Numero sem decimal
			xRet := Int(Val(aLinha[nPosLinha]))

		Case cTipo == "N"  .and. nDec > 0  // Campo Numero sem decimal
			cNum     := StrTran(aLinha[nPosLinha],".","")
			aNum     := Separa(aLinha[nPosLinha],",")
			cNovoNum := aNum[1]

			If len(aNum) >= 2
				cNovoNum += "."+aNum[2]
			Endif

			xRet := Val(cNovoNum)
		ENDCASE
	Endif

Return xRet

/*/{Protheus.doc} FPergunte
Perguntas da rotina
@author Alex Teixeira
@since 23/07/2020
@version 1.0
@return ${return}, ${return_description}

@param lView, logical, descricao
@param lEdit, logical, descricao
@type function
/*/
Static Function FPergunte()
	Local aParambox	:= {}
	Local aRet 		:= {}
	Local lRet		:= .T.
	Local cArq		:= Space(120)
	Local cNomRot   := "FSMNT001"
	Local nCabec	:= 1

	aAdd(aParambox ,{1,"Arquivo: "		,cArq,"",".T.","DIR",".T.",120,.T.})       //01
	aAdd(aParambox, {2,"Cabecalho 1a linha", nCabec, {"1=SIM","2=NAO"},     70, ".T.", .T.})
//Define se ira apresentar tela de perguntas
	lRet := ParamBox(aParamBox,"Parametros",aRet,{|| .T.},{},.T.,Nil,Nil,Nil,cNomRot,.F.,.F.)

// Atualiza variaves Private
	_cArquiv  := MV_PAR01
	If ValType(MV_PAR02) == "N"
		_nCabec	  := MV_PAR02
	Else
		_nCabec	  := Val(MV_PAR02)
	Endif

Return(lRet)

//-------------------------------------------------------------------
/*/{Protheus.doc} MostraLog
Mostra Log de Processamento
Commit Function
@author Alex Teixeira
@since 14/07/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function MostraLog(aProcess)

	Local oOK 		:= LoadBitmap(GetResources(),'br_verde')
	Local oNO 		:= LoadBitmap(GetResources(),'br_vermelho')
	Local aDados 	:= aClone(aProcess)

	DEFINE DIALOG oDlg TITLE "Log de Processamento" FROM 180,180 TO 500,700 PIXEL

	oBrowse := TCBrowse():New( 01 , 01, 260,140,,{'','Filial','Log'},{20,30,500},oDlg,,,,,{||},,,,,,,.F.,,.T.,,.F.,,, )
	oBrowse:SetArray(aProcess)
	oBrowse:bLine := {||{If(aProcess[oBrowse:nAt,01],oOK,oNO),aProcess[oBrowse:nAt,02],;
		aProcess[oBrowse:nAt,03] } }

	DEFINE SBUTTON FROM 145,200 TYPE 6  ENABLE OF oDlg Action(ReportDef(aDados))
	DEFINE SBUTTON FROM 145,230 TYPE 1  ENABLE OF oDlg Action(oDlg:End())

	ACTIVATE DIALOG oDlg CENTERED

Return

//-------------------------------------------------------------------------------
/*/{Protheus.doc} ReportDef
Estrutura do relatorio
        
@author 	TOTVS
@since 		01/07/2012
@version 	1.0
@project	0000010314_EF_001

@return	nil
/*/ 
//-------------------------------------------------------------------------------
Static Function ReportDef(aDados)
	Local oReport
	Local oSection1
	Local oFont12 	:= TFont():New('Courier new',,-12,.T.)

	Private cCarga	:= ""
	Private cCliente	:= ""
	Private cItem		:= ""
	Private nTotQtd	:= 0

	oReport:= TReport():New("MOSTRALOG","Relatório Log",, {|oReport| ReportPrint(oReport,aDados)},"Este programa imprime o Relatorio de Log.")
	oReport:SetLandsCape()  //SetPortrait() //SetLandsCape()
	oReport:nFontBody 		:= 11
	oReport:oFontBody 		:= oFont12
	oReport:oFontHeader 	:= oFont12
	oReport:SetLeftMargin(5)
	oReport:SetTotalInLine(.F.)
	oReport:ShowHeader()

	oSection1 := TRSection():New(oReport,"Log",,)

	//Monta as celulas do relatório (Sessao,Campo,Alias,Titulo do campo,Picture,Tamanho)
	TRCell():new(oSection1, "FILIAL"	,"TMP", "Filial"	,/*Picture*/, 10 /*Tamanho*/,/*lPixel*/, /*{|| code-block de impressao }*/, "LEFT" ,.T.)
	TRCell():new(oSection1, "STATUS"	,"TMP", "Status"	,/*Picture*/, 10 /*Tamanho*/,/*lPixel*/, /*{|| code-block de impressao }*/, "LEFT" ,.T.)
	TRCell():new(oSection1, "MSG"		,"TMP", "Mensagem"	,/*Picture*/, 100 /*Tamanho*/,/*lPixel*/, /*{|| code-block de impressao }*/, "LEFT" ,.T.)

	oReport:PrintDialog()

Return Nil

//-------------------------------------------------------------------------------
/*/{Protheus.doc} ReportPrint
Chamada Principal do relatorio
        
@author 	TOTVS
@since 		01/07/2012
@version 	1.0
@project	0000010314_EF_001

@return	nil
/*/ 
//-------------------------------------------------------------------------------
Static Function ReportPrint(oReport,aDados)
	Local oSection1 	:= oReport:Section(1)
	Local nXj			:= 0

	oSection1:Init()
	For nXj := 1 to len(aDados)
		oSection1:Cell("FILIAL"):SetValue( aDados[nXj,2] )
		oSection1:Cell("MSG"):SetValue( aDados[nXj,3] )

		If aDados[nXj,1]
			oSection1:Cell("STATUS"):SetValue( "OK" )
		Else
			oSection1:Cell("STATUS"):SetValue( "Erro" )
		Endif
		oSection1:PrintLine()
	Next

	oSection1:Finish()

Return Nil


//-------------------------------------------------------------------
/*/{Protheus.doc} BuscaVeic
Busca veiculos cadastrados na tabela ST9
Commit Function
@author Alex Teixeira
@since 14/07/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function BuscaVeic(cPlaca)
	Local cQuery     := ""
	Local cAliasST9  := GetNextAlias()
	Local nRecno     := 0
	Local cFilST9	 := ""

	cQuery := " SELECT ST9.T9_FILIAL, ST9.R_E_C_N_O_ AS ST9RECNO "
	cQuery += " FROM "+RetSQLName("ST9")+" ST9 "
	cQuery += " WHERE ST9.T9_PLACA = '" + Alltrim(cPlaca) + "'"
	cQuery += " AND ST9.D_E_L_E_T_ <> '*' "

	cQuery   := ChangeQuery(cQuery)
	dbUseArea( .T., "TOPCONN", TcGenQry( ,, cQuery ), cAliasST9, .F., .T. )

	If  !(cAliasST9)->(Eof())
		cFilST9 := (cAliasST9)->T9_FILIAL
		nRecno  :=  (cAliasST9)->ST9RECNO
	Endif

	(cAliasST9)->( dbCloseArea() )

Return {cFilST9,nRecno}


//---------------------------------------------------------------------
/*/{Protheus.doc} MNTA830ACE
Programa de inclusao de Manutencao
@author Inácio Luiz Kolling
@since
@version P12
@return
/*/
//---------------------------------------------------------------------
Static Function FSMNT830ACE(nCont1,nCont2,cHora, dDtCont, cMsg)
	Local i
	Local nReg       := 0
	Local lRet       := .t.
	Local dtBasBkp	 := ddatabase

	Private cBEMVIR  := ST9->T9_CODBEM
	Private cNOMBEM  := ST9->T9_NOME
	Private cCOMP
	Private lTRB     := .F.
	Private lPRIMC   := .F.
	Private nCONT1AN := 0
	Private nCONT2AN := 0
	Private cTRB32	:= GetNextAlias()
	Private cTRB31	:= GetNextAlias()
	Private cTRB33	:= GetNextAlias()

	Default cMSg     := ""

	If ST9->T9_SITBEM = 'I'
		cMsg += "Situacao do Bem: Intaivo"
		lRet := .f.
	EndIf

	If ST9->T9_SITBEM = 'T'
		cMSg += "Este Bem foi TRANSFERIDO."
		lRet := .f.
	EndIf

	If lRet
		nREGST9 := ST9->( Recno() )
		nReg := nREGST9
		//+--------------------------------------+
		//| Cria Arquivo de Trabalho             |
		//+--------------------------------------+
		aCampos  := {}
		Aadd(aCampos,{"CODBEM", "C", 16, 0})
		Aadd(aCampos,{"VARDIA", "N", 06, 0})
		Aadd(aCampos,{"CONTAC", "N", 12, 0})
		Aadd(aCampos,{"DTULTA", "D", 08, 0})

		//Intancia classe FWTemporaryTable
		oTmpTab1  := FWTemporaryTable():New( cTRB32,aCAMPOS)
		oTmpTab1:AddIndex("Ind01",{"CODBEM"} )
		oTmpTab1:Create()

		aDBF := {{"CODBEM" , "C", 16, 0} ,;
			{"NOMEBEM", "C", 40, 0} ,;
			{"SERVICO", "C", 06, 0} ,;
			{"NOMESER", "C", 40, 0} ,;
			{"SEQRELA", "C", 03, 0} ,;
			{"DATAPRO", "D", 10, 0} ,;
			{"PARECER", "C", 09, 0}}

		//Intancia classe FWTemporaryTable 2
		oTmpTbl2  := FWTemporaryTable():New(cTRB31,aDBF)
		oTmpTbl2:AddIndex( "Ind01" , {"CODBEM","SERVICO","SEQRELA"}  )
		oTmpTbl2:Create()

		aRR := {}
		aRR := NGCOMPEST(ST9->T9_CODBEM,'B',.T.)

		If Len(aRR) > 0
			For i := 1 To Len(aRR)
				dbSelectArea("ST9")
				Dbsetorder(1)
				If !DbSeek(xFilial("ST9")+aRR[i])
					(cTRB32)->(DbAppend())
					(cTRB32)->CODBEM := aRR[i]
					(cTRB32)->VARDIA := ST9->T9_VARDIA
					(cTRB32)->CONTAC := ST9->T9_CONTACU
					(cTRB32)->DTULTA := ST9->T9_DTULTAC
				EndIf
			Next i
		EndIf

		//+--------------------------------------+
		//| Cria Arquivo de Trabalho             |
		//+--------------------------------------+
		aDBF  := STC->(DbStruct())
		Aadd(aDBF,{"TC_NREPASS","N",01,0})

		//Intancia classe FWTemporaryTable (ARQ. 3)
		oTmpTbl3 := FWTemporaryTable():New(cTRB33,aDBF)
		oTmpTbl3:AddIndex("Ind01", {"TC_CODBEM","TC_COMPONE","TC_SEQRELA"})
		oTmpTbl3:AddIndex("Ind02", {"TC_COMPONE","TC_CODBEM","TC_SEQRELA"})
		oTmpTbl3:Create()

		cPAI := cBEMVIR
		dbSelectArea("STC")
		dbSeek(xFilial('STC')+cPAI)
		While !Eof() .And. STC->TC_FILIAL = xFilial('STC') .And. STC->TC_CODBEM = cPAI
			nREC  := RECNO()
			cCOMP := STC->TC_COMPONE
			lTRB  := .T.
			(cTRB33)->(DbAppend())
			For i := 1 TO FCOUNT()
				(cTRB33)->(FieldPut(i, STC->(FIELDGET(i)) ))
			Next i
			(cTRB33)->TC_NREPASS := 1

			dbSelectArea("STC")
			If dbSeek(xFilial('STC')+cCOMP)
				MNTA830TRB(cCOMP)
			EndIf
			Dbgoto(nREC)
			Dbskip()
		EndDo

		dbSelectArea('ST9')
		Dbsetorder(nINDST9)
		Dbgoto(nREGST9)

		cBemInF		:= ST9->T9_CODBEM
		nCONT1V 	:= nCont1
		nCONT2V		:= nCont2
		cHORA1V		:= cHora
		cHORA2V		:= cHora
		dDTLEIV		:= dDtCont

		aValid := VldContador()

		If 	aValid[1]
			dDATAMOV  := dDTLEIV
			nCONTADOR := nCONT1V
			nCONTADAU := nCONT1V
			nCONT2    := nCONT2V
			nCONTAD2  := nCONT2V
			_PROXDT   := CTOD("  /  /  ")

			// CALCULA A DATA DA PROXIMA MANUTENCAO
			nDias := 0

			dbSelectArea(cTRB32)
			DbGotop()
			While !Eof()
				STF->(DbSeek(xFilial('STF')+(cTRB32)->CODBEM))
				While !Eof() .and. xFilial("STF") == STF->TF_FILIAL .And.;
						STF->TF_CODBEM == (cTRB32)->CODBEM

					If STF->TF_ATIVO == "N"
						dbSelectArea("STF")
						dbSkip()
						Loop
					EndIf

					If STF->TF_PERIODO == "E"
						dbSelectArea("STF")
						dbSkip()
						Loop
					EndIf

					_PROXDT := CTOD("  /  /  ")
					_PROXDT := NGXPROXMAN((cTRB32)->CODBEM)


					xParecer := space(09)

					If _PROXDT < dDatabase
						xParecer := 'ATRASADA'
					Elseif _PROXDT = dDatabase
						xParecer := 'PARA HOJE'
					EndIf
					If nDias > 0
						If _PROXDT > dDatabase
							If dDataBase + nDias > _PROXDT
								xParecer := 'A REALIZAR'
							EndIf
						EndIf
					EndIf

					If !Empty(xParecer)
						xdata := substr(dtoc(_PROXDT),1,6) + substr(dtoc(_PROXDT),9,2)
						dbSelectArea('ST4')
						dbsetorder(1)
						dbseek(xFilial('ST4')+STF->TF_SERVICO)

						(cTRB31)->(DbAppend())
						(cTRB31)->CODBEM  := ST9->T9_CODBEM
						(cTRB31)->NOMEBEM := Substr(ST9->T9_NOME,1,40)
						(cTRB31)->SERVICO := STF->TF_SERVICO
						(cTRB31)->NOMESER := SubStr(ST4->T4_NOME,1,40)
						(cTRB31)->SEQRELA := STF->TF_SEQRELA
						(cTRB31)->DATAPRO := _PROXDT
						(cTRB31)->PARECER := xParecer
					EndIf

					dbSelectArea('STF')
					Dbskip()
				EndDo
				dbSelectArea(cTRB32)
				DbSkip()
			EndDo


			//GERAR O.S AUTOMATICA POR CONTADOR
			If (cGERAPREV = "S" .Or. cGERAPREV = "C") .And. (!Empty(nCONT1V) .Or. !Empty(nCONT2V))
				//If NGCONFOSAUT(cGERAPREV)

				dbSelectArea('ST9')
				Dbsetorder(nINDST9)
				Dbgoto(nREGST9)

				ddatabase := dDATAMOV

				CALVARDIA(cBemInF,dDATAMOV,nCONT1V)

				Aadd(aLog, {.T.,cFilAnt, "Placa: "+cPlaca+" Contador atualizado "})

				Processa({|lEnd| FSGOSAUT(cBemInF,xFilial("STJ"),dDATAMOV,nCONT1V ) })
				//NGGEROSAUT(cBemInF,If(!Empty(nCONT1V),nCONT1V,nCONT2V),xFilial("STJ"))


				ddatabase :=  dtBasBkp

				//EndIf
			EndIf

			If ExistBlock("MNTA8302")
				ExecBlock("MNTA8302",.F.,.F.)
			EndIf
		Else
			cMsg += aValid[2]
			lRet := .f.
		EndIf
	Endif

	ddatabase :=  dtBasBkp	  

	// Deleta as tabelas temporarias
	oTmpTbl2:Delete()
	oTmpTab1:Delete()
	oTmpTbl3:Delete()

Return lRet

Static Function CALVARDIA(cCODBEM,dDATALEI,nCONTADOR)
	Local nVARDIA := 0, aVARDIA := {}

	nMAX := GETMV("MV_VARDIA")

	dbSelectArea('STP')
	dbSeTOrder(5)
	dbSeek(xFilial('STP')+cCODBEM+DTOS(dDATALEI),.T.)
	If EoF() .Or. STP->TP_CODBEM <> ST9->T9_CODBEM
		dbSkip(-1)
	EndIf

	While !BoF() .And. STP->TP_CODBEM == cCODBEM .And. nMAX > 0
		If !Empty(STP->TP_DTLEITU)
			aadd(aVARDIA,{STP->TP_DTLEITU,STP->TP_ACUMCON,STP->TP_VIRACON})
			nMAX--
		EndIf
		dbSkip(-1)
	End

	dbSelectArea('STW')
	dbSeTOrder(5)
	dbSeek(xFilial('STW')+cCODBEM+DTOS(dDATALEI),.T.)
	If EoF() .Or. STW->TW_CODBEM <> ST9->T9_CODBEM
		dbSkip(-1)
	EndIf

	While !bof() .And. STW->TW_CODBEM == cCODBEM .And. nMAX > 0
		If !Empty(STW->TW_DTLEITU)
			aadd(aVARDIA ,{STW->TW_DTLEITU,STW->TW_ACUMCON,STW->TW_VIRACON})
			nMAX--
		EndIf
		dbSkip(-1)
	End

	If Len(aVARDIA) > 0

		aVARDIA := aSort(aVARDIA,,,{|x,y| x[1] < y[1]})

		nmaiorvlr := If (aVARDIA[Len(aVARDIA)][3] = 0, aVARDIA[Len(aVARDIA)][2] , aVARDIA[Len(aVARDIA)][2] + 999999999 * aVARDIA[Len(aVARDIA)][3])
		nmenorvlr := If (aVARDIA[1][3] = 0, aVARDIA[1][2] , aVARDIA[1][2] + 999999999 * aVARDIA[1][3] )

		nVARDIA := Round ((nmaiorvlr - nmenorvlr ) / (aVARDIA[Len(aVARDIA)][1] - aVARDIA[1][1]),0)

	EndIf

// ATUALIZA A NOVA POSICAO DO CONTADOR E ACUMULADOR

	dbSelectArea('ST9')
	dbSeTOrder(1)
	If dbSeek(xFilial('ST9')+cCODBEM)

		RecLock("ST9",.F.)
		If nVARDIA > 0
			ST9->T9_VARDIA  := nVARDIA
		EndIf

		If ST9->T9_UNGARAN == 'K'
			nINCREM := ST9->T9_PRGARAN / ST9->T9_VARDIA
			ST9->T9_DTGARAN := NGPROXMDT(ST9->T9_DTCOMPR,'D',nINCREM)
		EndIf

		ST9->T9_DTULTAC := dDATALEI
		ST9->T9_POSCONT := nCONTADOR

		nacumul := nCONTADOR

		If nacumul >= 999999999
			ST9->T9_VIRADAS := ST9->T9_VIRADAS + 1
			ST9->T9_CONTACU := nacumul - ST9->T9_CONTACU
			ST9->T9_POSCONT := nacumul - ST9->T9_CONTACU

		Else
			ST9->T9_CONTACU := nCONTADOR
			ST9->T9_POSCONT := nCONTADOR

		EndIf
		MsUnLock("ST9")
	EndIf

Return nVARDIA


Static Function VldContador()
	Local lRet 		:= .T.
	Local cError	:= ""
	Local cQuery	:= ""
	Local TIPOACOM  := IIf(ST9->T9_TEMCONT = "S",.T.,.F.)
	Local TIPOACOM2 := .f.
	Local cAliasQry := ""

	cQuery := " SELECT COUNT(TP_CODBEM) AS TOTAL FROM " + RetSQLName("STP") + " STP"
	cQuery += " WHERE STP.TP_FILIAL = '" + xFilial("STP")+"' "
	cQuery += "   AND STP.TP_CODBEM = '" + cBemInF+"' "
	cQuery += "   AND STP.TP_DTLEITU || STP.TP_HORA > " + ValToSql(DTOS(dDTLEIV) + cHORA1V)+" "
	cQuery += "   AND STP.D_E_L_E_T_ <> '*'"
	cQuery := ChangeQuery(cQuery)


	cAliasQry := GetNextAlias()
	dbUseArea( .T., "TOPCONN", TCGENQRY(,,cQuery),cAliasQry, .F., .T.)

	If (cAliasQry)->TOTAL >= 1
		lRet := .F.
		cError := 'Há registros posteriores ao lançamento da quebra (contador).'
	EndIf
	(cAliasQry)->(dbCloseArea())

	If !lRet

		oSTP:= MNTCounter():New()
		oSTP:setOperation(3)

		If TIPOACOM .And. nCONT1V > 0
			oSTP:setValue("TP_CODBEM" ,cBemInF)
			oSTP:setValue("TP_POSCONT",nCONT1V)
			oSTP:setValue("TP_DTLEITU",dDTLEIV)
			oSTP:setValue("TP_HORA"   ,cHORA1V)
		EndIf

		If TIPOACOM2 .And. nCONT2V > 0
			oSTP:setValue("TPP_CODBEM",cBemInF)
			oSTP:setValue("TPP_POSCON",nCONT2V)
			oSTP:setValue("TPP_DTLEIT",dDTLEIV)
			oSTP:setValue("TPP_HORA"  ,cHORA2V)
		EndIf

		lRet := oSTP:Turn()
		If !lRet
			cError := oSTP:getErrorList()[1]
		EndIf
	Endif


	If lRet
		If ST9->T9_POSCONT >= nCONT1V
			lRet := .F.
			cError := 'Contador atual maior/igual ao da importacao.'
		Endif
	Endif

	If lRet
		If dDataBase < dDTLEIV
			lRet := .F.
			cError := 'Data registro superior a data do sistema.'
		Endif
	Endif	

Return {lRet,cError}

Static Function FSGOSAUT(cBEM,PFIL,dDtMov,nCont)
	Local cFilOST9   := NGTROCAFILI("ST9",PFIL)
	Local cFilOSTF   := NGTROCAFILI("STF",PFIL)
	Local cFilOTPE   := NGTROCAFILI("TPE",PFIL)
	Local cFilOSTJ   := NGTROCAFILI("STJ",PFIL)
	Local vOSABER    := {}
	Local nULTCOMAN  := 0
	Local nPCONTFIXO := GetMV("MV_NGCOFIX") //Percentual para calcular o contador fixo da manutencao
	Local nPERFIXO   := nPCONTFIXO / 100
	Local nRetTole   := GetNewPar("MV_GALVTOL",0)
	Local lTolera    := NGCADICBASE("TF_TOLECON","D","STF",.F.),nMargem := 0
	Local lTolConE   := If(NGCADICBASE("TF_MARGEM","A","STF",.F.),.t.,.f.)
	Local aNGGERAOS  := {}
	Local aCounter   := {} //Informações de último contador do bem
	Local cTContacu  := ''
	Local dTDtUltac  := Ctod( '  /  /  ' )
	Local nTVardia   := 0
	Local lCONUS1VEZ := .t.
	Local cVERGEROS  := AllTrim(GETMv("MV_NGVEROS"))


	Private lJaApag  := .F.
	Private nNUMOSGE := 0
	Private dMENOR   := Ctod('  /  /  '),dMAIOR := dMENOR
	Private cMEORD   := Space(Len(stj->tj_ordem)), cMAORD := cMEORD
	Private dDPROXM


	cBemAut := ST9->T9_CODBEM

	//Dados do primeiro contador
	cTContacu  := nCont
	dTDtUltac  := dDtMov
	nTVardia   := 0

	If ST9->T9_SITMAN <> "I" .AND. ST9->T9_SITBEM = "A"

		If ExistBlock("SUBOSAUT")
			ExecBlock("SUBOSAUT",.F.,.F.)
		Else
			dbSelectArea("STF")
			dbSetOrder(01)
			If dbSeek(cFilOSTF+cBEM)
				ProcRegua(LastRec())
				While !Eof() .And. STF->TF_FILIAL == cFilOSTF .And. STF->TF_CODBEM == cBEM

					IncProc()
					If STF->TF_ATIVO == "N"
						DbSkip()
						Loop
					EndIf

					If STF->TF_PERIODO == "E"  //manutencoes eventuais nao geram OS automatica
						DbSkip()
						Loop
					EndIf

					dbSelectArea("STF")
					nRegStf  := Recno()
					nTolCont := If(!lTolConE,STF->TF_TOLECON,STF->TF_TOLERA * nTVardia )

					If STF->TF_TIPACOM <> "T"   // .And. STF->TF_TIPACOM <> "A"

						If STF->TF_TIPACOM == "S"

							dbSelectArea("TPE")
							dbSetOrder(01)
							If dbSeek(cFilOTPE+cBEM)

								//Verifica os aberta para mesmo bem+servico+sequencia
								If (cVERGEROS == "V" .And. !lCONUS1VEZ) .Or. (cVERGEROS == "C")
									vOSABER := NGPROCOSAB(PFIL,"B",STF->TF_CODBEM,STF->TF_SERVICO,STF->TF_SEQRELA)
									If !vOSABER[1]
										dbSelectArea("STF")
										dbSkip()
										Loop
									EndIf
								ElseIf cVERGEROS == "S"
									vOSABER := NGPROCOSAB(PFIL,"B",STF->TF_CODBEM,STF->TF_SERVICO,STF->TF_SEQRELA)
									If !vOSABER[1]

										If  GetRemoteType() > -1 .And. !MsgYesNo("Já existe OS aberta para o mesmo Bem+Serviço+Sequência:"+chr(13)+chr(13); //"Já existe OS aberta para o mesmo Bem+Serviço+Sequência:"
											+"Bem"+": "+STF->TF_CODBEM+chr(13); //"Bem"
											+"Serviço"+": "+STF->TF_SERVICO+chr(13); //"Serviço"
											+"Sequência: "+ STF->TF_SEQRELA+chr(13)+chr(13); //"Sequência: "
											+"Deseja gerar OS automática por contador mesmo com OS já aberta ?"+chr(13)+chr(13); //"Deseja gerar OS automática por contador mesmo com OS já aberta ?"
											+"Confirma (Sim/Não)", "ATENÇÃO") //"Confirma (Sim/Não)"# "ATENÇÃO"
											dbSelectArea("STF")
											dbSkip()
											Loop
										EndIf
									EndIf
								EndIf

								dDPROXM := dTDtUltac

								/*
								dDPROXM := NGPROXMAN( dTDtUltac,"C",STF->TF_TEENMAN,;
									STF->TF_UNENMAN,STF->TF_CONMANU,STF->TF_INENMAN,;
									cTContacu, nTVardia )
								*/

								aNGGERAOS := NGGERAOS('P',dDPROXM ,cBEM,STF->TF_SERVICO,STF->TF_SEQRELA,'S','S','S',PFIL)

								If aNGGERAOS[1][1] = 'S'

									Aadd(aLog, {.T.,cFilAnt, "Placa: "+cPlaca+" OS gerada "})

									nNUMOSGE += 1
									dbSelectArea("STJ")
									dbSetOrder(01)
									If dbSeek(cFilOSTJ+aNGGERAOS[1][3]+"000001")
										RecLock("STJ",.f.)
										STJ->TJ_OBSERVA := "Importacao KM"
										STJ->(MsUnlock())
										If Empty(dMENOR)
											dMENOR := stj->tj_dtmpini
											dMAIOR := dMENOR
											cMEORD := aNGGERAOS[1,3]
											cMAORD := cMEORD
										Else
											dMENOR := Min(dMENOR,stj->tj_dtmpini)
											dMAIOR := Max(dMAIOR,stj->tj_dtmpini)
											cMAORD := aNGGERAOS[1,3]
										Endif
									Endif
								Else
									Aadd(aLog, {.F.,cFilAnt, "Placa: "+cPlaca+" OS:"+aNGGERAOS[1][2]})
								EndIf

							EndIf

						Else

							If STF->TF_TIPACOM = "F"

								If STF->( FieldPos("TF_CONPREV") ) > 0
									nULTCOMAN := STF->TF_CONPREV
								Else
									nULTCOMAN := STF->TF_CONMANU
								EndIf

								dbSelectArea("CTT")
								If CTT->(FieldPos("CTT_PERCFI")) > 0  //Campo utilizado pela Galvao Engenharia
									cCENTROCB := NGSEEK("ST9",STF->TF_CODBEM,1,"T9_CCUSTO")
									nPERCCUST := NGSEEK("CTT",cCENTROCB,1,"CTT_PERCFI")
									If nPERCCUST  > 0
										nPERCCUST := nPERCCUST / 100
										nINCPERC  := STF->TF_INENMAN * nPERCCUST
									Else
										nINCPERC := STF->TF_INENMAN * nPERFIXO  // Incremento da manutencao com percentual
									EndIf
								Else
									nINCPERC := STF->TF_INENMAN * nPERFIXO     // Incremento da manutencao com percentual
								EndIf

								nVEZMANU := Int(nULTCOMAN / STF->TF_INENMAN) // Numero de vezes que foi feito a manutencao
								nCONTFIX := IF(nVEZMANU==0, STF->TF_INENMAN, nVEZMANU * STF->TF_INENMAN) // Contador fixo exato
								nCONTPAS := nULTCOMAN - nCONTFIX             // Quantidade que passou da manutenção fixa

								If nCONTPAS < nINCPERC .Or. nINCPERC == 0
									If nCONTPAS < 0
										nCONTPAS := nCONTPAS * -1
									EndIf
									If nVEZMANU == 0 .And. nCONTPAS > nINCPERC
										nULTCOMAN := 0
									Else
										nULTCOMAN := nCONTFIX
									EndIf
								Else
									nULTCOMAN := nCONTFIX + STF->TF_INENMAN
								EndIf
							Else
								nULTCOMAN := STF->TF_CONMANU
							EndIf

							//Tolerancia da manutencao - Galvao Engenharia
							dbSelectArea("ST9")
							If ST9->(FieldPos("T9_GALTOLE")) > 0
								dbSelectARea("ST9")
								dbSetOrder(01)
								If dbSeek(xFilial("ST9")+STF->TF_CODBEM)
									If ST9->T9_GALTOLE > 0
										nULTCOMAN :=  nULTCOMAN - ST9->T9_GALTOLE
									Else
										nTOLCCUST := NGSEEK("CTT",ST9->T9_CCUSTO,1,"CTT_TOLERG")
										If nTOLCCUST > 0
											nULTCOMAN :=  nULTCOMAN - nTOLCCUST
										ElseIf nRetTole > 0
											nULTCOMAN := nULTCOMAN - nRetTole
										EndIf
									EndIf
								EndIf
							EndIf

							//dDPROXM := dTDtUltac

							dDPROXM := NGPROXMAN( STF->TF_DTULTMA, STF->TF_TIPACOM, STF->TF_TEENMAN,;
								STF->TF_UNENMAN, nULTCOMAN, STF->TF_INENMAN, cTContacu, nTVardia, STF->TF_DTULTMA )


							//Verifica os aberta para mesmo bem+servico+sequencia
							If (cVERGEROS == "V" .And. !lCONUS1VEZ) .Or. (cVERGEROS == "C")
								vOSABER := NGPROCOSAB(PFIL,"B",STF->TF_CODBEM,STF->TF_SERVICO,STF->TF_SEQRELA)
								If !vOSABER[1]
									dbSelectArea("STF")
									dbSkip()
									Loop
								EndIf
							ElseIf cVERGEROS == "S"
								vOSABER := NGPROCOSAB(PFIL,"B",STF->TF_CODBEM,STF->TF_SERVICO,STF->TF_SEQRELA)
								If  GetRemoteType() > -1 .And. !vOSABER[1] //-1 = Job, Web ou Working Thread (Sem remote)
									If !MsgYesNo("Já existe OS aberta para o mesmo Bem+Serviço+Sequência:"+chr(13)+chr(13); //"Já existe OS aberta para o mesmo Bem+Serviço+Sequência:"
										+"Bem"+": "+STF->TF_CODBEM+chr(13); //"Bem"
										+"Serviço"+": "+STF->TF_SERVICO+chr(13); //"Serviço"
										+"Sequência: "+ STF->TF_SEQRELA+chr(13)+chr(13); //"Sequência: "
										+"Deseja gerar OS automática por contador mesmo com OS já aberta ?"+chr(13)+chr(13); //"Deseja gerar OS automática por contador mesmo com OS já aberta ?"
										+"Confirma (Sim/Não)","ATENÇÃO") //"Confirma (Sim/Não)"# "ATENÇÃO"
										dbSelectArea("STF")
										dbSkip()
										Loop
									EndIf
								EndIf
							EndIf

							aNGGERAOS := NGGERAOS('P',dDataBase ,cBEM,STF->TF_SERVICO,STF->TF_SEQRELA,'S','S','S',PFIL)

							If aNGGERAOS[1][1] = 'S'

								Aadd(aLog, {.T.,cFilAnt, "Placa: "+cPlaca+" OS gerada "})

								nNUMOSGE += 1
								dbSelectArea("STJ")
								dbSetOrder(01)
								If dbSeek(cFilOSTJ+aNGGERAOS[1][3]+"000001")
									RecLock("STJ",.f.)
									STJ->TJ_OBSERVA := "Importacao KM "
									STJ->TJ_DTORIGI := dDPROXM
									STJ->(MsUnlock())
									If Empty(dMENOR)
										dMENOR := stj->tj_dtmpini
										dMAIOR := dMENOR
										cMEORD := aNGGERAOS[1,3]
										cMAORD := cMEORD
									Else
										dMENOR := Min(dMENOR,stj->tj_dtmpini)
										dMAIOR := Max(dMAIOR,stj->tj_dtmpini)
										cMAORD := aNGGERAOS[1,3]
									Endif
								Endif
							else
								Aadd(aLog, {.F.,cFilAnt, "Placa: "+cPlaca+" OS:"+aNGGERAOS[1][2]})
							EndIf
						EndIf
					EndIf
					dbSelectArea("STF")
					dbgoto(nRegStf)
					dbSkip()
				End
			EndIf
		EndIf
	EndIf

Return .T.
