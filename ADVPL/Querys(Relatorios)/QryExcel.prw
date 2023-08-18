#INCLUDE "TOTVS.CH"
#INCLUDE "PARMTYPE.CH"

/*/{Protheus.doc} QryExcel

Cadastro da Tabela SZ6 onde são armazenadas as Querys e Parametros das Funções Internas, sendo possível modificação em tempo de execução.

@type function
@author Cesar Lopes
@since 06/04/2022
@version P11,P12
@database MSSQL,Oracle

@table SZ6

@see GerExcel
/*/
User Function QryExcel()
	Local cAlias	:= "SZ6"
	Local cTitulo 	:= "Cadastro de Pesquisas e Parametros Genéricos "
	Local aRotAdic	:= {}
	Local aBotoes	:= {}

	//Seleciiona a Tabela e Posiciona
	DbSelectArea(cAlias)
	(cAlias)->(DbSetOrder(1))
	(cAlias)->(DbGotop())
	
	aAdd( aBotoes , {"SZ6",{|| NORMQRY() },"Norm Query","Normatizar Query"} )	  

	//AxCadastro - Tela padrão da mBrowse ( [ cAlias ] [ cTitle ] [ cDel ] [ cOk ] [ aRotAdic ] [ bPre ] [ bOK ] [ bTTS ] [ bNoTTS ] [ aAuto ] [ nOpcAuto ] [ aButtons ] [ aACS ] [ cTela ] )
	AxCadastro(cAlias, cTitulo, , , aRotAdic , , , , , , , aBotoes, , )

Return

/*
	normatiza a query
*/
Static Function NORMQRY()
	Local cQuery 		:= M->Z6_CONTEU
	Local cChar			:= "#"
	Local cQueryNorm 	:= ""
	Local cPerg			:= M->Z6_ROTINA

	If !Empty(cPerg)
		If !Pergunte(cPerg,.T.)
			Return
		Endif
	Endif

	If cQuery <> "" 
		While At(  cChar , alltrim(cQuery)) <> 0  

			nValIni:= At(  cChar , alltrim(cQuery)) 											        // Onde encontrou o primeiro #
			nValFim := At(  cChar , alltrim( substr(cQuery, nValIni+1, len(cQuery)   )      ))  // Onde encontrou o Fim do Primeiro #

			cQueryNorm := substr(cQuery, 1, nValIni -1  ) // Elimina o primeiro #
			cParam := substr(cQuery,  nValIni + 1 , nValFim - 1 ) //Pega o primeiro parametro

			cQuery := cQueryNorm + &cParam + substr(cQuery, (nValIni + nValFim + 1) , len(cQuery))

			//SELECT * FROM SE1070 WHERE E1_EMISSAO  <= '#DtoS(DDATABASE)#' AND E1_CLIENTE = '#CCLIENTE#'  AND E1_FILIAL = '001001'
		Enddo
		
		Aviso("Query Normatizada.",cQuery,{"Ok","Cancelar"},3,"Regras",,,.T.)
		
	Endif
	
Return
