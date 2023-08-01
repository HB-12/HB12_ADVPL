#INCLUDE "TOTVS.CH"
#INCLUDE "PARMTYPE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "RWMAKE.CH"

#DEFINE _OPC_cGETFILE ( GETF_RETDIRECTORY + GETF_LOCALHARD )

/*/{Protheus.doc} GerExcel

A rotina de cadastro das querys é a QryExcel

@type function
@author Cesar Lopes
@since 06/04/2023
@version P11,P12
@database MSSQL,Oracle

@table SZ6

@param aParam, Array, Vetor de Parâmetros para o JOB, contendo {cEmpAnt,cFilAnt}

@see QryExcel 
/*/
User Function GerExcel()
	Local oCbx
	Local cVar			:= ""
	Local aDados 		:=  GetItenSZ6()
    Local cCadastro     := "Exporta Excel por Query"
	
	Private oDlg
	Private cArq 		:= ""
	Private cDir 		:= GetSrvProfString("Startpath","")
	Private cWorkSheet 	:= ""
	Private cTable 		:= ""
	Private cDirTmp 	:= GetTempPath()
	Private cQuery  	:= ""
	Private cAliasT		:= ""
	Private cNomeArq 	:= ""

	cVar := aDados[1]

	//+----------------------------------------------------------------------------
	//| Definição da janela e seus conteúdos
	//+----------------------------------------------------------------------------
	DEFINE MSDIALOG oDlg TITLE "Seleção do Relatório" FROM 0,0 TO 280,552 OF oDlg PIXEL

	@ 06,06 TO 130,271 LABEL "Selecione" OF oDlg PIXEL

	@ 15, 15 SAY   "Relatório" SIZE 100,100 PIXEL OF oDlg
	@ 30, 15  MSCOMBOBOX oCbx VAR cVar ITEMS aDados 		SIZE 200, 200 OF oDlg PIXEL
	//+----------------------------------------------------------------------------
	//| Botoes da MSDialog
	//+----------------------------------------------------------------------------

	//@ 093,235 BUTTON "&Ok"       SIZE 36,16 PIXEL ACTION (GetPerg(cVar), oDlg:End())
	@ 093,235 BUTTON "&Ok"       SIZE 36,16 PIXEL ACTION Processa( {|| GetPerg(cVar) }, cCadastro, "Processando arquivo, aguarde...", .F. )
	@ 113,235 BUTTON "&Cancelar" SIZE 36,16 PIXEL ACTION oDlg:End()

	ACTIVATE MSDIALOG oDlg CENTER

Return

/*
	0 - Função Main 
*/
Static Function GeraRel(cVar)

	//Define a régua de processamento
	CursorWait() //Mostra Ampulheta
	ProcRegua(5)
	IncProc("Obtendo Dados...")

	cQuery:=  GetInfoGen(cVar, '001')  // Busca a query da tabela SZ6 ---Aqui ajustar a variável de sequencia

	If !vazio(cQuery)
		cQuery := GetNormQuery(cQuery) // Normaliza a query 

		GetSql(cQuery) // Faz a busca das informações
		GeraArq()

	EndIf

Return

/*
	3 - Cria tabela temporária 
*/
Static Function GetSql(cQuery)
	Local nOk	:= 0
	
	ChangeQuery("\sql\GerExcel_GetSql.sql",@cQuery)

	nOK := TCSQLExec(cQuery) 

	If nOK < 0
		Alert("Houve um erro no código SQL. Verifique tabela de Consultas!")
		Aviso("Query .",cQuery,{"Ok","Cancelar"},3,"GerExcel")
		Return (.T.)
	EndIf

	TcQuery cQuery New Alias (cAliasT:=GetNextAlias())

	If Select(cAliasT) = 0
		Alert("Não há dados.")
		Return .F.
	Endif

Return

/*
	4 -  Cria o arquivo/Tabela em .xml
*/
Static Function GeraArq()
	Local cWorkSheet 	:= ""
	Local cTable 		:= ""
	Local aStructure	:= (cAliasT)->(DBStruct())
	Local cCampo		:= ""
	Local cConteu		:=	""
	Local cContString	:= ""
	Local aRow 			:= {}
	Local i 			:= 0
	
	IncProc("Gerando arquivo Excel...")	
	
	oFwMsEx := FWMsExcel():New()

	cWorkSheet := cNomeArq
	cTable     := cNomeArq

	oFwMsEx:AddWorkSheet( cWorkSheet )
	oFwMsEx:AddTable( cWorkSheet, cTable )

	For i:= 1 to len(aStructure) // Lê a estrutura da tabela e pega os campos e cria as colunas
		oFwMsEx:AddColumn( cWorkSheet, cTable , ASTRUCTURE[i][1]  , 1,1) 		//1
	Next i

	While (cAliasT)->(!Eof()) 
		For i:= 1 to len(aStructure) // Lê a estrutura da tabela e pega os campos insere os dados
			cCampo	:= "(cAliasT)->"+ASTRUCTURE[i][1]
			cConteu	:= &cCampo
			aadd(aRow,cConteu)
		Next i

		oFwMsEx:AddRow( cWorkSheet, cTable,  aRow )		
		aRow := {}
		//Próximo Registro
		(cAliasT)->(dbSkip())
	EndDo	
		
	//Seleciona o arquivo Salvar como
	cDirTmp     := cGetFile( "Selecione o Diretorio | " , OemToAnsi( "Selecione Diretorio" ) , NIL , "" , .F. , _OPC_cGETFILE )
	cDirTmp 	:= UPPER(cDirTmp)

	oFwMsEx:Activate()
	cArq := CriaTrab( NIL, .F. ) + ".xml"
	LjMsgRun("Gerando o arquivo, aguarde...", cNomeArq, {|| oFwMsEx:GetXMLFile( cArq ) } )
	If __CopyFile( cArq, cDirTmp + cArq )
		oExcelApp := MsExcel():New()
		oExcelApp:WorkBooks:Open( cDirTmp + cArq )
		oExcelApp:SetVisible(.T.)
	Else
		MsgInfo("Arquivo não copiado para temporário do usuário." )
	Endif
	
	CursorArrow() //Libera Ampulheta	
	
Return

/*
	1 - Busca o conteudo do campo da Tabela Genéria SZ6
*/
Static Function GetInfoGen(cChave, cSequencia)
	Local cQuery1	:= ""
	
	If Select("SZ6") > 0
		SZ6->(DbCloseArea()) // Fecha a area
	Endif 
	
	DbselectArea("SZ6")
	DbSetOrder(1)
	If dbSeek(xFilial("SZ6")+cChave+cSequencia)
		cQuery1 	:= 	SZ6->Z6_CONTEU
		cNomeArq 	:=  SZ6->Z6_DESCRI
	Else
		Alert("Dados da tabela  SZ6 não localizada.")
	Endif

Return cQuery1

/*
	2 - Normatiza o SQL em comandos do Protheus
*/
Static Function GetNormQuery(cQuery)
	Local cChar			:= "#"
	Local cQueryNorm	:= ""
	
	IncProc("Normatizando Query...")
	
	//SELECT * FROM SE1070 WHERE E1_EMISSAO  <= '#DtoS(DDATABASE)#' AND E1_CLIENTE = '#CCLIENTE#'  AND E1_FILIAL = '001001'

	While At(cChar, allTrim(cQuery)) <> 0  

		nValIni	:= At(cChar, allTrim(cQuery)) 									// Onde encontrou o primeiro #
		nValFim := At(cChar, allTrim( substr(cQuery, nValIni+1, len(cQuery))))	// Onde encontrou o Fim do Primeiro #

		cQueryNorm	:= substr(cQuery, 1, nValIni -1  ) // Elimina o primeiro #
		cParam 		:= substr(cQuery,  nValIni + 1 , nValFim - 1 ) //Pega o primeiro parametro

		cQuery := cQueryNorm + &cParam + substr(cQuery, (nValIni + nValFim + 1) , len(cQuery))
	Enddo

	Aviso("SQL",cQuery,{"Ok","Cancelar"},3,"Regras")
    

Return cQuery

/*
	6 - Busca as descrições dos Relatórios(SQL) na Tabela SZ6
*/
Static Function GetItenSZ6()
	Local aDados	:= {}

	DbselectArea("SZ6")
	DbSetOrder(1)
	If dbSeek(xFilial("SZ6"))
		While ("SZ6")->(!Eof())
			aAdd(aDados,  ("SZ6")->Z6_CHAVE + " - " +("SZ6")->Z6_DESCRI )
			("SZ6")->(dbSkip())
		EndDo
	Else
		Alert("Dados da tabela  SZ6 não localizada.")
	Endif
	
Return aDados

/*
	Carrega Perguntas
*/
Static Function GetPerg(cParamVar)
	Local cVar 		:= Substr(cParamVar,1,3)
	Local cPergunt	:= ""

	DbselectArea("SZ6")
	DbSetOrder(1)
	If dbSeek(xFilial("SZ6")+cVar)
		cPergunt:= alltrim(("SZ6")->Z6_ROTINA)
	Else
		Alert("Dados da tabela  SZ6 não localizada.")
	Endif

	If cPergunt <> ''
		If Pergunte(cPergunt,.T.)
			GeraRel(cVar)
		Else
			Return
		Endif
	Endif

Return 
