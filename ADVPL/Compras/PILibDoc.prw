#include "PROTHEUS.CH"  
#include "TBICONN.CH"
#include "TBICODE.CH" 
STATIC __cPrgNom
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} PILibDoc
Rotina utilizada para realizar liberação de documento Multipla.
Usuário terá opçao para liberar varios documentos.

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		
@return    	
@obs        

Alterações Realizadas desde a Estruturação Inicial
/*/
//---------------------------------------------------------------------------------------
User Function PILibDoc()

	Local   cMsgProble	:= ''//Variavel de mensagem de problema
	Local   cMsgSoluca	:= ''//Variavel de mensagem de solução

	Private cCadastro 	:= "Liberação de documentos"//Titulo da tela                                                                
	Private aRotina		:= {}
	Private oMark		:= GetMarkBrow()//Objeto para marcar  
	Private aArraySCR	:= {}
	Private aCampos 	:= {}
	Private aCpos		:= {} 
	Private nTotReg	   	:= 0
	Private cArqTrab    := ""
	Private cIndOrdPag  := ""
	Private nRecnoSCR   := 0
	Private cSeek       := ""
	 
	/* Mensagem de problema*/
	cMsgProble	+= OemToAnsi('Usuário não está cadastrado como aprovador. O acesso desta rotina é destinada apenas aos ')
	cMsgProble	+= OemToAnsi('usuários envolvidos no processo de aprovação de pedido de compras definido')
	/* Mensagem de solução*/
	cMsgSoluca  += OemToAnsi('Verifique se o usuário deveria estar envolvido no processo de aprovação, através do grupo de aprovadores')

	/* Valida se usuario é um aprovador*/	
	SAK->(dbSetOrder(02))
	If ! SAK->(Dbseek(xFilial('SAK')+__CUSERID))
	   ShowHelpDlg("Liberação Documento",{cMsgProble},5,{cMsgSoluca},5)
		Return(.F.)
	EndIf

	FStrTela() // cria a estrutura da tabela para ser apresentado na tela da rotina.
	/*Chama rotina para criar menu da tela */
	FMenu()

	/*Função utilizada para buscar os dados*/
	FBusDados()
		
	If nTotReg == 0 
		MsgAlert(OemTOAnsi('Não existem dados a serem exibidos!!'))
		Return .T.
	Endif

	// 23/03/2023 - GUSTAVO BARCELOS - Implementada melhoria para verificar o saldo de liberações
	U_FSldAprov()

	/* Cria tela para usuario selecionar quais documento deseja liberar*/
	MarkBrow("TRB","OK",,aCpos,,,"U_PIMarkTud()",,,,"U_DuplClik()")
	
Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} FBusDados
Função utilizada para buscar os dados do aprovador

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum
/*/
//---------------------------------------------------------------------------------------
Static Function FBusDados()
	
	Local cQuery 		:= ''
	Local cTabDoc		:= GetNextAlias()
	Local nVlrLiq		:= 0
	Local cClassi		:= ''
	Local cVencimentos	:= ''
	Local aCondPag		:= {}

	DbSelectArea("TRB")
	TRB->(__DBZAP()) // Limpa a tabela temporaria.

	aArraySCR:={}     

	cQuery += Chr(13)+" SELECT CR_FILIAL AS FILIAL, CR_NUM AS NUM, CR_USER AS USUARIO, CR_TIPO AS TIPO, CR_TOTAL AS TOTAL FROM "+RetSqlName("SCR")+"  "	
	cQuery += Chr(13)+" WHERE D_E_L_E_T_ <> '*' "
	cQuery += Chr(13)+" AND SubString(CR_FILIAL,1,2) = '"+xFilial("SAK")+"' "
	cQuery += Chr(13)+" AND CR_USER = '"+__CUSERID+"' "
	cQuery += Chr(13)+" AND CR_STATUS IN ('02') "	
	cQuery += Chr(13)+" ORDER BY 1,2 "
		
	DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cTabDoc,.F.,.F.)  

	nRecnoSCR := SCR->(Reccount())

	(cTabDoc)->(DbGoTop())                                   
		
	Do While (cTabDoc)->(!Eof())
			
		//nVlrLiq:= FbusVlq((cTabDoc)->FILIAL,(cTabDoc)->NUM)
			
		SC7->(DbSetorder(01))
		SC7->(DbSeek((cTabDoc)->FILIAL+AvKey((cTabDoc)->NUM,'C7_NUM')))
		SC1->(DbSetOrder(01))
		SC1->(DbSeek((cTabDoc)->FILIAL+SC7->C7_NUMSC+SC7->C7_ITEMSC))
		SY1->(DbSetorder(03))
		SY1->(DbSeek(xFilial('SY1')+AvKey(SC7->C7_USER,'Y1_USER')))
		SE4->(dbSetOrder(1))
		SE4->(DbSeek(xFilial('SE4')+SC7->C7_COND))						
		SA2->(DbSetOrder(01))
		SA2->(DbSeek(xFilial('SA2')+SC7->C7_FORNECE+SC7->C7_LOJA))
		
		/*Retorna vencimentos conforme condição de pagamento*/
		aCondPag:=Condicao(1,SC7->C7_COND,,dDataBase,,,,,,)		
	    /*Carrega variavel vencimento com as datas de vencimento*/
	    AEval(aCondPag,{|x,y| cVencimentos+=Dtoc(x[1])+' | '})
	    nTotReg ++		
		  
		Reclock('TRB',.T.)
		Replace OK          With Space(02),;
			   FILIAL       With (cTabDoc)->FILIAL,;
			   NOMEFIL      With FWFilialName(,(cTabDoc)->FILIAL,2),;
			   NUMERO		With (cTabDoc)->NUM,;
			   VALOR		With (cTabDoc)->TOTAL,;
			   FORNECEDOR	With SA2->A2_COD,;
			   NOME_COMP    With SY1->Y1_NOME,;
			   DTINCLUSAO	With DTOC(SC7->C7_EMISSAO),;
   			   LOJA			With SA2->A2_LOJA,;
  			   NOME			With SA2->A2_NOME,;
			   OBS			With SC7->C7_OBS,;
			   TIPO			With (cTabDoc)->TIPO,;
			   TIPMAN		With SC7->C7_ZTMNT,;
			   PLACA 		With SC7->C7_ZPLACA,;
			   USUARIO 		With (cTabDoc)->USUARIO,;
			   REEMBOLSO	With SC7->C7_ZREEMB
		TRB->(MsUnLock())		                

		cClassi:= ''
		aCondPag:={}          
		cVencimentos:=''
		//CLASSIF	   With cClassi,;					
	   (cTabDoc)->(DbSkip())

	EndDo

	DbSelectArea("TRB")
	DbSetOrder(1)
	If ! Empty(cSeek)
	   DbSeek(Left(cSeek,6),.T.)
	   If Eof()
		  TRB->(DbGotop())
	   Endif
	Endif   

Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} DuplClik
Função valida duplo click no registro

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+---------------------------------------------------------
/*/
//---------------------------------------------------------------------------------------
User Function DuplClik() 
    
	Local aAreaAll	:= {SCR->(GetArea()),GetArea()}
	/* Caso usuario esteja selecionando um registro que está marcado, nesse caso ele está desmarcando*/
	If !Empty(TRB->OK)
		RecLock("TRB",.F.)
		Replace TRB->OK With "  "
		TRB->(MsUnLock())	
	Else		
		/* Caso nao tenha problema de saldo ou alçada deve-se preencher o campo */
		RecLock("TRB",.F.)
		Replace TRB->OK With ThisMark()
		TRB->(MsUnLock())					
	Endif
	
	SCR->(DbSetorder(02))
	If SCR->(DbSeek(AvKey(TRB->FILIAL,'CR_FILIAL')+AvKey(TRB->TIPO,'CR_TIPO')+Avkey(TRB->NUMERO,'CR_NUM')+TRB->USUARIO))
		/*Realizando a marcação do registro*/
		If !Empty(TRB->OK)
			/* Verifico se registro ja foi add no array*/
			nPosReg := aScan( aArraySCR , {|x| x[1] == SCR->(Recno())})
			If nPosReg == 0
				/* Add Recno SCR*/
				AADD(aArraySCR,{SCR->(Recno()),TRB->FILIAL, TRB->TIPO,SCR->CR_TOTAL})
			EndIf		
		Else
			nPosReg := aScan( aArraySCR , {|x| x[1] == SCR->(Recno())})
			If nPosReg <> 0
				/* Deleta a posiçao do array*/
				ADEL(aArraySCR, nPosReg)
				/* Redimensiona tamanho do array, sempre depois que deleta necessario redimensionar para nao ficar com posição Null no array*/
				ASIZE(aArraySCR,Len(aArraySCR)-1)
			EndIf	
		Endif
    Endif
    
    AEval(aAreaAll,{|nLem|RestArea(nLem)})
    
Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} FVisPCom
Função utilizada para visualizar pedido de compras ou NF

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+---------------------------------------------------------
/*/
//---------------------------------------------------------------------------------------
User Function FVisPCom()   

	Local aAreaOld		:= {SC7->(GetArea()),GetArea()}
	Local cPedSCR		:= TRB->NUMERO
	Local cTipPed		:= TRB->TIPO
	Local cFilDoc		:= TRB->FILIAL
	Local cFilAux		:= cFilAnt
	Private nTipoPed 	:= 1   //Define o tipo de pedido
	Private l120Auto 	:= .F. //Informa a rotina de pedidos que processo não é automatico
	Private aBackSC7  := {}
	
	cFilAnt:=cFilDoc 
	/* Verifica se documento é PC ou Nota Fiscal*/
	If cTipPed == 'PC'
		/* Posiciona no pedido de compras*/	
	    SC7->(DbSetOrder(01))
	    If SC7->(DbSeek(AvKey(cFilDoc,'C7_FILIAL')+Avkey(cPedSCR,'C7_NUM')))    	    
		    INCLUI := .F.
			ALTERA := .F.    
			A120Pedido( 'SC7', SC7->(Recno()), 2 )
		EndIf
		/*Sempre apos visualizar limpar o campo fleg*/
		FLimpTab()
	ElseIf cTipPed == 'NF'
		SF1->(DbSetorder(01))
		If SF1->(DbSeek(AvKey(cFilDoc,'F1_FILIAL')+Substr(cPedSCR,1,Len(SF1->F1_DOC+SF1->F1_SERIE+SF1->F1_FORNECE+SF1->F1_LOJA))))
		    INCLUI := .F.
			ALTERA := .F.    
			Pergunte("MTA103",.F.)
			A103NFISCAL('SF1',SF1->(Recno()),2)
		Endif            
		/*Sempre apos visualizar limpar o campo fleg*/
		FLimpTab()
	Endif 
	/* Restaura as areas*/                 
	AEval(aAreaOld,{|x|RestArea(x)}) 
	
	cFilAnt:= cFilAux

	If nRecnoSCR <> SCR->(Reccount())
	   FBusDados()
	Endif

Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} FLibDoc
Função utilizada para realizar liberação dos documentos

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+------'---------------------------------------------------
/*/
//---------------------------------------------------------------------------------------
User Function FLibDoc() 
	
	If Len(aArraySCR)>0
		If MsgYesNo("Deseja realizar liberação do(s) : "+cValToChar(Len(aArraySCR))+' documento(s) selecionado(s)?')
			Processa( {|| PILibDocs() , FBusDados()}, "Liberando documentos", "Aguarde...",.F.)
		EndIf
	Else
		MsgAlert('Favor selecionar no minimo 1 registro!!')
	Endif	

	aArraySCR := {} // limpa o array dos pedidos a serem liberados.
	SysRefresh()

Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} PILibDocs
Função utilizada para processar liberação dos documentos
                                  
@author		.iNi Sistemas
@since     	05/09/14
@version  	P.11
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+---------------------------------------------------------
/*/
//--------------------------------------------------------------------------------------- 
Static Function PILibDocs()
	
	Local nRegSCR	:= Len(aArraySCR)//Variavel recebe quantidade registros existentes no array
	Local nCont	   	:= 1
	Local nX	   	:= 1
	Local lRetAprov	:= .F. //Variavel recebe retorno caso documento tenha sido totalmente liberado
	Local cQuery	:= ''
	Local lErro		:= .T. 
	Local cCodUsrApr:= ''
	Local cFilAux	:= cFilAnt
	Local nSldLib	:= 0
	Local 	nTotSel	:= 0

	cSeek:= ""

    /* Definição de Regua*/
	ProcRegua(nRegSCR)

	// 23/03/2023 - GUSTAVO BARCELOS - Implementada melhoria para verificar o saldo de liberações
	nSldLib 	:= FGetSldLib()
	aEval(aArraySCR, { |x| nTotSel += x[4]})

	If (nSldLib + nTotSel) > SAK->AK_LIMITE
		FWAlertWarning(	"O valor limite de aprovação foi atingido para o aprovador " + SAK->AK_COD + " - " + AllTrim(SAK->AK_NOME) + " e, " +;
						"portanto, não será possível seguir com a operação de Aprovação." + CRLF + CRLF +;
						"- Valor Limite: R$ " + AllTrim(TRANSFORM(SAK->AK_LIMITE, GetSX3Cache("AK_LIMITE", "X3_PICTURE"))) + CRLF +;
						"- Saldo Restante: R$ " + AllTrim(TRANSFORM(SAK->AK_LIMITE - nSldLib, GetSX3Cache("AK_LIMITE", "X3_PICTURE"))) + CRLF +;
						"- Total Selecionado: R$ " + AllTrim(TRANSFORM(nTotSel, GetSX3Cache("AK_LIMITE", "X3_PICTURE"))) + CRLF + CRLF +;
						"Obs.: Quando é atingido o valor limite de aprovações do aprovador, não é realizada a operação de liberação " +;
						"de nenhum dos registros selecionados, ou seja, o processo de aprovação é completamente cancelado.",;
						"Foi atingido o valor limite de aprovações.")
	Else
	
		Begin Transaction 
		
			For nX:= 1 To Len(aArraySCR)
				/* Regua do processo*/
				IncProc("Processando Registro: "+ AllTrim(str(nCont)) +' de: '+Alltrim(Str(nRegSCR)))
				/*Atualiza a filial de acordo com o registro*/
				cFilAnt := aArraySCR[nX][2]
				/*Necessario posicionar via Recno, pois existe a possibilidade de pedidos antigos*/ 
				SCR->(dbGoTo(aArraySCR[nX][1]))
				If !SCR->(Eof())
					/* Seek do pedido de compras*/
					cSeek:=SCR->CR_FILIAL+AvKey(SCR->CR_NUM,'C7_NUM')
					/* Bloco 1 */
					If aArraySCR[nX][3] == 'PC' .Or. aArraySCR[nX][3] == 'AE'
						SC7->(DbSetOrder(01))
						SC7->(DbSeek(cSeek))
						cCodUsrApr:= SC7->C7_APROV
					ElseIf aArraySCR[nX][3] == 'NF'
						SF1->(DbSetOrder(01))
						SF1->(DbSeek(Substr(cSeek,1,Len(SF1->F1_DOC+SF1->F1_SERIE+SF1->F1_FORNECE+SF1->F1_LOJA))))
						cCodUsrApr:= SF1->F1_APROV
					Endif	
					/* Realiza liberação do documento*/
					lRetAprov:= MaAlcDoc({SCR->CR_NUM,SCR->CR_TIPO,SCR->CR_TOTAL,SCR->CR_APROV,,cCodUsrApr,,,,,},dDatabase,4)             									
				EndIf			 
				/* Se variavel lRetAprov = .T. documento foi totalmente liberado*/
				/* Realiza liberação do pedido de compras ou Pre-Nota*/
				If lRetAprov .And. SCR->CR_TIPO == 'PC'
					SC7->(dbSetorder(01))
					If SC7->(dbSeek(cSeek))
						Do While SC7->(!Eof()) .And. SC7->(C7_FILIAL+C7_NUM) == SCR->(CR_FILIAL+AllTrim(CR_NUM))
							If Reclock("SC7",.F.)
								Replace	SC7->C7_CONAPRO With "L"
								SC7->(MsUnlock())
							Endif
							SC7->(dbSkip())
						EndDo	
					Endif	
				ElseIf lRetAprov .And. SCR->CR_TIPO == 'NF'
					SF1->(DbSetOrder(01))
					If SF1->(DbSeek(xFilial("SF1")+Substr(SCR->CR_NUM,1,Len(SF1->F1_DOC+SF1->F1_SERIE+SF1->F1_FORNECE+SF1->F1_LOJA))))
						If RecLock('SF1',.F.)
							Replace SF1->F1_STATUS With If(SF1->F1_STATUS=="B"," ",SF1->F1_STATUS)
							SF1->(MsUnLock())
						EndIf
					EndIf				
				EndIf
			/*Atualiza variavel para nao carregar lixo*/
			lRetAprov:=.F.
			Next nX	
							
		End Transaction	

	EndIf
	/*Retorna a variavel da filial corrente*/
	cFilAnt:= cFilAux	
			
Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} PIMarkAll
Função acionada no momento em que usuario tenta usar objeto do markbrow para marcar
todos os registros.
Seu retorno será Nil, faz-se necessario para desabilitar essa opçao.

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+---------------------------------------------------------
/*/
//---------------------------------------------------------------------------------------	
User Function PIMarkTud()     
Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} FbusVlq
Função utilizada para buscar o valor liquido do pedido de compras

@author		.iNi Sistemas
@since     	03/03/15
@version  	P.11              
@param 		cFilPC - Filial do Pedido de compras
@param 		cNumPC - Numero do pedido de compras
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+---------------------------------------------------------
/*/
//---------------------------------------------------------------------------------------	
Static Function FbusVlq(cFilPC, cNumPC)

	Local nVlPedLiq	:= 0  
	Local cQuery	:= ''
	Local aPrefSom	:= GetNextAlias()
	
	cQuery+="SELECT SUM((C7_TOTAL+C7_VALFRE+C7_VALIPI ) - C7_VLDESC) AS TOTAL "
	cQuery+="FROM "+RetSqlName('SC7')+ " "
	cQuery+="WHERE C7_FILIAL = '"+cFilPC+"' "
	cQuery+="AND C7_NUM = '"+AllTrim(cNumPC)+"' "
	cQuery+="AND D_E_L_E_T_ <> '*'  "
	
	DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),aPrefSom,.F.,.F.) 
	
	(aPrefSom)->(DbGoTop())
	/*Recebe o valor total do PC*/
	nVlPedLiq:= (aPrefSom)->TOTAL
	/*Fecha tabela temporaria*/
	(aPrefSom)->(DbCloseArea())
		                         
Return(nVlPedLiq)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} FLimpTab
Função utilizada para desmarcar todos os itens da tabela e limpar array com os pedidos
selecionados para aprovação

@author		.iNi Sistemas
@since     	03/03/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum

Alterações Realizadas desde a Estruturação Inicial
------------+-----------------+---------------------------------------------------------
Data       	|Desenvolvedor    |Motivo                                                    
------------+-----------------+---------------------------------------------------------
/*/
//---------------------------------------------------------------------------------------	
Static Function FLimpTab()

	/*Sempre apos visualizar limpar o campo fleg*/
	TRB->(DbGoTop())
	Do While TRB->(!Eof())			
		RecLock("TRB",.F.)
		Replace TRB->OK With "  "
		TRB->(MsUnLock())
	TRB->(DbSkip())	
	EndDo
	
	/*Limpa array que contem os registros selecionados*/
	aArraySCR:={}			

Return(Nil)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} PILibDoc
Rotina utilizada para realizar liberação de documento Multipla.
Usuário terá opçao para liberar varios documentos.

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		
@return    	
@obs        

Alterações Realizadas desde a Estruturação Inicial
/*/
//---------------------------------------------------------------------------------------
Static Function FStrTela()

	aCpos := {}

	AADD(aCpos,{ "OK" 	   	 	 , "","" 					})
	AADD(aCpos,{ "FILIAL"  		 , "", "Filial"			,"@!"})
	AADD(aCpos,{ "NOMEFIL" 		 , "", "Nome Filial"	,"@!"})
	AADD(aCpos,{ "NUMERO"  		 , "", "Numero"			,"@!"})
	AADD(aCpos,{ "VALOR"   		 , "", "Valor"			,"@E 999,999,999.99"})
	AADD(aCpos,{ "FORNECEDOR" 	 , "", "Codigo"			,"@!"})
	AADD(aCpos,{ "LOJA" 	 	 , "", "Loja"			,"@!"})
	AADD(aCpos,{ "NOME" 	 	 , "", "Nome"			,"@!"})
    AADD(aCpos,{ "NOME_COMP" 	 , "", "Nome Comprador"	,"@!"})
	AADD(aCpos,{ "DTINCLUSAO"	 , "", "Data Insercao"	,"@D"})
	AADD(aCpos,{ "OBS" 			 , "", "Observacao"		,"@!"})
	AADD(aCpos,{ "TIPMAN"	 	 , "", "Tipo Manutencao","@!"})
	AADD(aCpos,{ "PLACA"	     , "", "Placa"			,"@!"})
	AADD(aCpos,{ "REEMBOLSO"	 , "", "Reembolso"		,"@!"}) 
	AADD(aCpos,{ "USUARIO"	     , "", "Usuario"		,"@!"})
	AADD(aCpos,{ "TIPO"	   	 	 , "", "Tipo Doc"		,"@!"})

	aCampos := {}
			  
	AADD(aCampos,{ "OK"    		 , "C", 2, 0 })                    
	AADD(aCampos,{ "FILIAL"   	 , "C", 5, 0})//TamSx3( "CR_FILIAL" )[1], 0 })
	AADD(aCampos,{ "NOMEFIL"     , "C", 20, 0 })                    
	AADD(aCampos,{ "NUMERO"   	 , "C", TamSx3( "C7_NUM"  )[1], 0 })
	AADD(aCampos,{ "VALOR"		 , "N", 10, 0})//( "C7_TOTAL")[1], 2 })
	AADD(aCampos,{ "FORNECEDOR"  , "C", TamSx3( "A1_COD"  )[1], 0 })
	AADD(aCampos,{ "LOJA"   	 , "C", TamSx3( "A2_LOJA" )[1], 0 })
	AADD(aCampos,{ "NOME"   	 , "C", 35, 0})//( "A2_NOME" )[1], 0 })
	AADD(aCampos,{ "NOME_COMP"	 , "C", TamSx3( "Y1_NOME" )[1], 0 })
	AADD(aCampos,{ "DTINCLUSAO"  , "C", 10, 0 })
	AADD(aCampos,{ "OBS"		 , "C", TamSx3( "C7_OBS"  )[1], 0 })
	AADD(aCampos,{ "TIPMAN"  	 , "C", TamSx3( "C7_ZTMNT")[1], 0 })
	AADD(aCampos,{ "PLACA" 	 	 , "C", TamSx3( "C7_ZPLACA" )[1], 0 })
	AADD(aCampos,{ "REEMBOLSO"	 , "C", TamSx3( "C7_ZREEMB")[1], 2 })
	AADD(aCampos,{ "USUARIO" 	 , "C", TamSx3( "CR_USER" )[1], 0 })
	AADD(aCampos,{ "TIPO" 		 , "C", TamSx3( "CR_TIPO" )[1], 0 })
	
	cArqTrab  := CriaTrab(aCampos,.T.) 
	cIndOrdPag:= CriaTrab(Nil,.F.)
		
	If Select('TRB')>0
	   TRB->(DbCloseArea())
	Endif

	dbUseArea(.T.,, cArqTrab,"TRB",.F.,.F.) // DbUseArea(lNovo, cDriver, cArquivo, cAlias, lComparilhado,lSoLeitura)    
	IndRegua("TRB",cIndOrdPag,"FILIAL+NUMERO",,,"Selecionando Registro") //"Selecionando Registros..."

Return(.T.)
//--------------------------------------------------------------------------------------
/*/
{Protheus.doc} FMenu
Função utilizada para definir Menu da rotina

@author		.iNi Sistemas
@since     	01/01/15
@version  	P.11              
@param 		Nenhum
@return    	Nenhum
@obs        Nenhum
/*/
//---------------------------------------------------------------------------------------
Static Function FMenu()   

	aRotina := {}
	/*Monta o Menu*/
	AAdd (aRotina,{"Pesquisar"     ,"U_FPesqDoc()"  ,0,1})
	AAdd (aRotina,{"Consulta Doc"  ,"U_FVisPCom()"  ,0,2})
	AAdd (aRotina,{"Liberar Doc"   ,"U_FLibDoc()"   ,0,3})
	// 23/03/2023 - GUSTAVO BARCELOS - Implementada melhoria para verificar o saldo de liberações
	AAdd (aRotina,{"Consulta Saldo - Aprovador"   ,"U_FSldAprov()"   ,0,3})

Return(aRotina)

/*/{Protheus.doc} GetX3CBox
Função responsável por recuperar o conteúdo da lista de opções (X3_CBOX).
          
@author 	Gustavo Barcelos
@since 		14/10/2019 
@version 	P12.1.25

/*/
User Function GetX3CBox(cCampo,xDado)

	Local aAreas    := {SX3->(GetArea()), GetArea()}
	Local aTmp 		:= {}
	Local aLegendas	:= {}

	Local cCBox		:= ""
	
	Local xRet

	// Carrega conteúdo do CBox do campo
	cCBox := GetSX3Cache(cCampo, "X3_CBOX")

	// Valida se o retorno do X3_CBOX é uma lista. Caso não seja, tenta executar.
	If At("=", cCBox) == 0
		cCBOx := &(cCBox)
	EndIf

	// Realiza a quebra e conversão do conteúdo para um array
	aTmp := STRTOKARR( cCBox , ";" )
	If Empty(xDado)
		xRet := aTmp
	Else
		// Busca legenda do conteúdo enviado
		aEval(aTmp, { |x| aAdd(aLegendas, STRTOKARR( x , "=" )) })
		xRet := Upper(aLegendas[aScan(aLegendas, { |x| x[1] == xDado })][2])
	EndIf

	// Restaura áreas anteriores
    AEval(aAreas, {|x| RestArea(x) })

Return xRet

/*/{Protheus.doc} FGetSldLib
Função responsável por buscar os dados de saldo liberado do Aprovador na Filial Logada.
@type function
@author gustavo.barcelos
@since 3/23/2023
/*/
Static Function FGetSldLib()

	Local cFilSld	:= SAK->AK_FILIAL
	Local cCodAprov	:= SAK->AK_COD
	Local cTipoLib	:= SAK->AK_TIPO
	Local cAliSCR	:= GetNextAlias()

	Local nRet		:= 0
	Local nDiaSema	:= 0

	Local dDtIni	:= StoD("")
	Local dDtFim	:= StoD("")

	If cTipoLib == 'D'
		dDtIni := Date()
		dDtFim := Date()
	ElseIf cTipoLib == 'S'
		nDiaSema := Dow( Date() )
		nSubInic := nDiaSema - 1
		nSumFina := 7 - nDiaSema

		dDtIni := DaySub(Date(), nSubInic)
		dDtFim := DaySum(Date(), nSumFina)		
	ElseIf cTipoLib == 'M'
		dDtIni := FirstDate(Date())
		dDtFim := LastDate(Date())
	EndIf

	cDtIni	:= DtoS(dDtIni)
	cDtFim	:= DtoS(dDtFim)

	BeginSql Alias cAliSCR

	SELECT SUM(CR_TOTAL) TOT_LIB
	FROM %Table:SCR% SCR
	WHERE SUBSTRING(CR_FILIAL,1,2) = %Exp:cFilSld%
	AND CR_APROV = %Exp:cCodAprov%
	AND CR_STATUS = '03'
	AND CR_DATALIB BETWEEN %Exp:cDtIni% AND %Exp:cDtFim%
	AND SCR.%NotDel%

	EndSql

	If !((cAliSCR)->(EOF()))
		nRet := (cAliSCR)->TOT_LIB
	EndIf

Return(nRet)

/*/{Protheus.doc} FSldAprov
Rotina responsável por apresentar em tela os saldos de limite do Aprovador.
@type function
@author gustavo.barcelos
@since 3/23/2023
/*/
User Function FSldAprov()

	Local nSldLib 	:= FGetSldLib()

	FWAlertInfo("Abaixo, segue informações de saldos do aprovador " + SAK->AK_COD + " - " + AllTrim(SAK->AK_NOME) + ":" + CRLF + CRLF +;
				"- Tipo de Limite: " + SAK->AK_TIPO + " - " + U_GetX3CBox("AK_TIPO",SAK->AK_TIPO) + CRLF +;
				"- Valor Limite: R$ " + AllTrim(TRANSFORM(SAK->AK_LIMITE, GetSX3Cache("AK_LIMITE", "X3_PICTURE"))) + CRLF +;
				"- Valores já Liberados: R$ " + AllTrim(TRANSFORM(nSldLib, GetSX3Cache("AK_LIMITE", "X3_PICTURE"))) + CRLF +;
				"- Saldo Restante: R$" + AllTrim(TRANSFORM(SAK->AK_LIMITE - nSldLib, GetSX3Cache("AK_LIMITE", "X3_PICTURE"))),;
				"Status atual de Liberações do Aprovador " + SAK->AK_COD + " - " + AllTrim(SAK->AK_NOME) + ".")

Return(Nil)
