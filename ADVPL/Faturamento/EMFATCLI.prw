#INCLUDE "PROTHEUS.CH"
#INCLUDE "MSOLE.CH"

User Function EMFATCLI()
	SetPrvt("CCADASTRO,ASAYS,ABUTTONS,NOPCA,CTYPE")
	SetPrvt("NVEZ,OWORD,CINICIO,CFIM,CFIL,CXINSTRU,CXLOCAL")
	SetPrvt("NPAG,CPATH,CARQLOC,NPOS")

	/*/
	ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
	±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
	±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
	±±³Fun‡„o    ³ EMCTRVND ³  Autor ³Alex T.Souza			³ Data ³ 05/08/15 ³±±
	±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
	±±³Descri‡„o ³ Relatorio                        - VIA WORD                ³±±
	±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
	±±³ Uso      ³ Especifico                                                 ³±±
	±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ´±±
	±±³ Revis„o  ³                                          ³ Data ³          ³±±
	±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÙ±±
	±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
	ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß/*/
	
	Processa({|| WORDIMP()})  // Chamada do Processamento// Substituido pelo assistente de conversao do AP5 IDE em 14/02/00 ==> 	Processa({|| Execute(WORDIMP)})  // Chamada do Processamento
	
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o    ³ WORDIMP  ³ Autor ³ Equipe Desenv. R.H.   ³ Data ³ 31.03.99 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Relatorio de Certificados dos cursos  - VIA WORD           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ Especifico                                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Revis„o  ³                                          ³ Data ³          ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß/*/
Static FUNCTION WORDIMP()
	Local cQuery 		:= ""
	Local cAliasTop 	:= GetNextAlias()
	Local cCNPJCli		:= ""
	Local cPrecFat		:= ""
	Local cArquivo    	:= GetSrvProfString("StartPath","")
	Local cPathSave		:= ""
	Local cDadCont		:= ""
	Local cStartPath	:= GetSrvProfString("Startpath","")
	Local cLogo			:= ""
	Local aItens		:= {}
	Local nCount        := 0
	Local nXy,nXi       := 0
	
 	cLogo :=  "LGRL" + SM0->M0_CODIGO + SM0->M0_CODFIL + ".BMP" // Empresa + Filial
  	If !File( cStartPath + cLogo )
   		cLogo := "LGRL" + SM0->M0_CODIGO + ".BMP" // Empresa
  	Endif	
     
	dbSelectArea("SA1")
	dbSetOrder(1)	
	
	dbSelectArea("SC5")
	dbSetOrder(1)		

	dbSelectArea("SC6")
	dbSetOrder(1)		
	
	
	cPathSave := cGetFile("\", "Selecione o diretorio para gerar o arquivo",,,,GETF_RETDIRECTORY+GETF_LOCALHARD+GETF_LOCALFLOPPY/*128+GETF_NETWORKDRIVE*/)//"Selecione o Diretorio p/ Gerar os Arquivos"
	cArquivo += "EMFATCLI" + ALLTRIM(SUBSTR(SM0->M0_CODFIL,1,2))+".dotm"

	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Copiar Arquivo .DOT do Server para Diretorio Local ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	nPos := Rat("\",cArquivo)
	If nPos > 0
		cArqLoc := AllTrim(Subst(cArquivo, nPos+1,20 ))
	Else
		nPos := Rat("/",cArquivo)
		If nPos > 0
			cArqLoc := AllTrim(Subst(cArquivo, nPos+1,20 ))
		Else'
			cArqLoc := cArquivo
		Endif
	EndIF
	
	cPath := GETTEMPPATH()
	If Right( AllTrim(cPath), 1 ) != "\"
		cPath += "\"
	Endif

	If !CpyS2T("\SYSTEM\"+cArqLoc, cPath, .T.)
   		Return
	Endif

	nPag 		:= 0

	cQuery := "SELECT SC5.C5_NUM, SC5.C5_EMISSAO,SC5.C5_DATA1, SC5.C5_NOTA, SC5.C5_SERIE, SC5.C5_MENNOT1, SA1.A1_NOME, SA1.A1_PESSOA, SA1.A1_CGC, SA1.A1_INSCR, SA1.A1_END, SA1.A1_MUN, SA1.A1_BAIRRO, SA1.A1_EST, SA1.A1_CEP, "
	cQuery += "SA1.R_E_C_N_O_ SA1REC, SC5.R_E_C_N_O_ SC5REC FROM "+RetSqlName("SC5")+" SC5 "
	cQuery += "INNER JOIN "+RetSqlName("SA1")+" SA1 ON SC5.C5_CLIENT = SA1.A1_COD AND SC5.C5_LOJACLI = SA1.A1_LOJA AND SA1.D_E_L_E_T_ <> '*' "	
	cQuery += "WHERE SC5.C5_FILIAL = '"+SC5->C5_FILIAL+"' AND SC5.C5_NUM = '"+Alltrim(SC5->C5_NUM)+"' AND SC5.D_E_L_E_T_ <> '*' "

	cQuery := ChangeQuery(cQuery)

	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTop,.T.,.T.)
	dbSelectArea(cAliasTop)

	if !(cAliasTop)->(Eof())
	
		SA1->(DBGoto( (cAliasTop)->SA1REC ))
		SC5->(DBGoto( (cAliasTop)->SC5REC ))

		cValExt := ""
		nTotVal := 0
		cDataVenc	:= DtoC(SC5->C5_DATA1)

		If SC6->(dbSeek(SC5->C5_FILIAL+SC5->C5_NUM))
			While SC6->(!Eof()) .AND. SC6->C6_FILIAL + SC6->C6_NUM == SC5->C5_FILIAL + SC5->C5_NUM
				nCount++

				//cValUnit    := AllTrim(Transform(SC6->C6_PRCVEN+(SC6->C6_VALDESC/SC6->C6_QTDVEN), "@E 999,999,999.99"))	
				cValUnit    := AllTrim(Transform((SC6->C6_VALOR+SC6->C6_VALDESC)/SC6->C6_QTDVEN, "@E 999,999,999.99"))	
				cValLiq		:= AllTrim(Transform(SC6->C6_PRCVEN, "@E 999,999,999.99")) 
				cValDesc	:= AllTrim(Transform(SC6->C6_VALDESC, "@E 999,999,999.99")) 
				cValLiq		:= AllTrim(Transform(SC6->C6_VALOR, "@E 999,999,999.99"))
				cQtdIt		:= AllTrim(Transform(SC6->C6_QTDVEN, "@E 999,999,999.99"))  
				cDescIt     := ALLTRIM(SC6->C6_DESCRI)

				nTotVal     += SC6->C6_VALOR

				aadd(aItens,{ 	{ "ITEM1"+Alltrim(Str(nCount)),cDescIt},;
								{ "ITEM2"+Alltrim(Str(nCount)),	cQtdIt},;
								{ "ITEM3"+Alltrim(Str(nCount)),	cValUnit},;
								{ "ITEM4"+Alltrim(Str(nCount)),	cValDesc},;
								{ "ITEM5"+Alltrim(Str(nCount)), cValLiq},;
								{ "ITEM6"+Alltrim(Str(nCount)), cDataVenc }})

				SC6->(dbSkip())
			EndDo

		Endif

		cValExt 	:= Extenso(nTotVal)
		cPrecFat	:= Transform(nTotVal,"@E 999,999,999.99")

		oWord := OLE_CreateLink('TMsOleWord97')
		OLE_NewFile(oWord,cPath+cArqLoc)	
	
		
		// Variaveis a serem usadas na Montagem do Documento no Word    
		//--Cadastro Funcionario
		OLE_SetDocumentVar(oWord,"TOTVS_NUMDOC",SC5->C5_NOTA)
		OLE_SetDocumentVar(oWord,"TOTVS_TIPODOC",SC5->C5_SERIE)
		

		OLE_SetDocumentVar(oWord,"TOTVS_VALFAT",cPrecFat)
		
		OLE_SetDocumentVar(oWord,"TOTVS_IMAGEM",cLogo)
				
		OLE_SetDocumentVar(oWord,"TOTVS_VALFAT2",cPrecFat)
		OLE_SetDocumentVar(oWord,"TOTVS_VALFAT3",cPrecFat)
		OLE_SetDocumentVar(oWord,"TOTVS_NOMEMP",Alltrim(SM0->M0_NOMECOM))
		OLE_SetDocumentVar(oWord,"TOTVS_ENDEMP",Alltrim(SM0->M0_ENDENT))
		OLE_SetDocumentVar(oWord,"TOTVS_CIDEMP",Alltrim(SM0->M0_CIDENT))		
		OLE_SetDocumentVar(oWord,"TOTVS_CEPEMP",Alltrim(SM0->M0_CEPENT))		
		OLE_SetDocumentVar(oWord,"TOTVS_TELEMP",Alltrim(SM0->M0_TEL))		
		OLE_SetDocumentVar(oWord,"TOTVS_FAXEMP",Alltrim(SM0->M0_FAX))				
		OLE_SetDocumentVar(oWord,"TOTVS_CNPJEMP",Transform(Alltrim(SM0->M0_CGC),"@R99.999.999/9999-99")) 
		OLE_SetDocumentVar(oWord,"TOTVS_INSCEST",Alltrim(SM0->M0_INSC))
		
		cMes := StrZero(Month(SC5->C5_EMISSAO),2)
		cDia := StrZero(Day(SC5->C5_EMISSAO),2)
		cDia := StrZero(Year(SC5->C5_EMISSAO),4)
		
		cData		:= DtoC(SC5->C5_EMISSAO)
		cDataExt	:= StrZero(Day(SC5->C5_EMISSAO),2)+" de "+MesExtenso(SC5->C5_EMISSAO)+" de "+StrZero(Year(SC5->C5_EMISSAO),4)     
		
		
		OLE_SetDocumentVar(oWord,"TOTVS_DATEXT",Alltrim(cDataExt))		
		OLE_SetDocumentVar(oWord,"TOTVS_VALEXT",Alltrim(cValExt))	
		OLE_SetDocumentVar(oWord,"TOTVS_DATFAT",Alltrim(cData))
		OLE_SetDocumentVar(oWord,"TOTVS_NOMCLI",	SA1->A1_NOME)
		OLE_SetDocumentVar(oWord,"TOTVS_ENDCLI",	SA1->A1_END)		
		OLE_SetDocumentVar(oWord,"TOTVS_MUNCLI",	SA1->A1_MUN)				
		OLE_SetDocumentVar(oWord,"TOTVS_UFCLI",	SA1->A1_EST)				
		OLE_SetDocumentVar(oWord,"TOTVS_CEPCLI",	SA1->A1_CEP)						

		If SA1->A1_PESSOA == "J"
			cCNPJCli := Transform(SA1->A1_CGC,"@R 99.999.999/9999-99")
		Else
			cCNPJCli := Transform(SA1->A1_CGC,"@R 999.999.999-99")
		Endif	


		OLE_SetDocumentVar(oWord,"TOTVS_CGCCLI",cCNPJCli)                                   
		OLE_SetDocumentVar(oWord,"TOTVS_IECLI",SA1->A1_INSCR)						
		

		cObs := SC5->C5_MENNOT1		
		            
		
		cObs := cObs + cDadCont

		OLE_SetDocumentVar(oWord,"TOTVS_NOTA1",cObs)

		OLE_SetDocumentVar( oWord, 'varQtdItens'  , LTrim( Str( Len( aItens ) ) ) )
		For nXi := 1 to Len(aItens)
			For nXy := 1 to len(aItens[nXi])
				OLE_SetDocumentVar(oWord,aItens[nXi][nXy][1],aItens[nXi][nXy][2])
			Next
		Next

		OLE_ExecuteMacro(oWord,"EMFATCLI01_ITENS")

		//Alterar nome do arquivo para Cada Pagina do arquivo para evitar sobreposicao.
		nPag ++

		OLE_UpdateFields(oWord) //Atualiza os campos dentro do word

		cFileSave := cPathSave
		OLE_SaveAsFile( oWord, cFileSave+"FAT_"+SC5->C5_NUM + "_" + Alltrim(SC5->C5_ZNNOTA)+".doc" ) 
		MsgInfo("Foi gerado o documento "+cFileSave+"FAT_"+SC5->C5_NUM+ "_" +Alltrim(SC5->C5_ZNNOTA)+".doc", "Atenção")

		OLE_CloseFile( oWord )
		OLE_CloseLink( oWord ) 			// Fecha o Documento
	Else
	
		Alert("Nao foi encontrado nenhum registro com os parametros informados")	

	Endif
	
	(cAliasTop)->(dbCloseArea())

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³  Apaga arquivo .DOT temporario da Estacao 		   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If File(cPath+cArqLoc)
	FErase(cPath+cArqLoc)
Endif

Return
      

