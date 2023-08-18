#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'

// Campos que serão impressos. Podem ser alterados desde que estejam nas tabelas consideradas na query da função FGetNFEnt.
// Para o tratamento do Centro de Custo e Rateio, estes campos não irão entrar nesta regra, ou seja, sempre serão exibidos os dados dos centros de custos.
#DEFINE aCampos    { { "F1_FILIAL"  , 1, .F., Nil        },;                  // Não remover!
                     { "F1_DOC"     , 1, .F., Nil        },;                  // Não remover!
                     { "F1_SERIE"   , 1, .F., Nil        },;                  // Não remover!
                     { "F1_FORNECE" , 1, .F., Nil        },;                  // Não remover!
                     { "F1_LOJA"    , 1, .F., Nil        },;                  // Não remover!
                     { "A2_NOME"    , 1, .F., Nil        },;
                     { "F1_EMISSAO" , 1, .F., Nil        },;
                     { "D1_PEDIDO"  , 1, .F., Nil        },;
                     { "C7_OBS"     , 1, .F., Nil        },;
                     { "F1_VALBRUT" , 3, .T., Nil        },;
                     { "F1_DTDIGIT" , 1, .F., Nil        },;
                     { "F1_ESPECIE" , 1, .F., Nil        },;
                     { "F1_USERLGI" , 1, .F., Nil        },;
                     { "F1_USERLGA" , 1, .F., Nil        }  }

/*/{Protheus.doc} FSCOMR01
Rotina responsável pela impressão de relatório de Notas Fiscais de Entrada da EMTEL.
@type function
@author gustavo.barcelos
@since 19/01/2023
/*/
User Function FSCOMR01()

    Local cCadastro	:= "Notas Fiscais de Entrada"
	Local cParam 	:= ""
	
    Local dDataDe   := StoD("")
    Local dDataAt   := StoD("")

	Local nOpca 	:= 0		// Flag de confirmacao para OK ou CANCELA

	Local aSays		:= {} 		// Array com as mensagens explicativas da rotina
	Local aButtons	:= {}		// Array com as botões da rotina
	Local aParams   := {}       // Array com parâmetros selecionados
	Local aDados    := {}       // Array com os dados encontrados para impressão

	Local lTemPar   := .F.      // Define se opção de parâmetros foi acionada

    //																							 |
	AADD(aSays, "Este programa tem como objetivo gerar um arquivo contendo o resumo das ")
	AADD(aSays, "Notas Fiscais de Entrada. ")

	aAdd( aButtons, { 5, .T., { || lTemPar := .T., aParams :=  FExibePerg() } } )
	aAdd( aButtons, { 6, .T., { || nOpca:=1, FechaBatch() } } )
	aAdd( aButtons, { 2, .T., { || FechaBatch() } } )

	FormBatch( cCadastro, aSays, aButtons )

	If ( nOpcA == 1 )

        If !lTemPar .AND. Empty(aParams)
			aParams := FExibePerg(.F.)

			aEval(aParams,{|x| cParam += "- " + IIF(ValType(x) == "D", DtoC(x),x) + CRLF })

			aOpcoes := {"Usar última seleção","Selecionar novos Parâmetros"}

			nOpcParam := Aviso( "Nenhum parâmetro informado",;
				"Parâmetros para processamento não informados." + CRLF +;
				"Últimos dados utilizados:" + CRLF + cParam + CRLF +;
				"Qual ação deseja executar?",aOpcoes,3)

			If nOpcParam == 2
				aParams := FExibePerg()
			EndIf

		EndIf

		// Caso não tenha lista de parâmetro preenchida, não executa o processo
		If !Empty(aParams)
			dDataDe     := aParams[1]
			dDataAt     := aParams[2]

			Processa({|| aDados := FGetNFEnt(dDataDe, dDataAt) }, "Carga de Dados...", "Carregando dados para Impressão...", .F.)

            If !Empty(aDados)
			    Processa({ || FGeraExcel(aDados) }, "Gerando arquivo Excel..." )
            Else
                FWAlertInfo("Nenhum registro foi encontrado para impressão. Informe novos parâmetros e tente novamente.",;
                            "Dados para impressão não encontrados!")
            EndIf

		EndIf

    EndIf

Return(Nil)

/*/{Protheus.doc} FExibePerg
Função responsável por exibir tela de perguntas ao usuário.
@type function
@author gustavo.barcelos
@since 19/01/2023
@param lExibePerg, logical, Define se tela de perguntas deve ser exibida (.T.) ou somente carregar parâmetros (.F.) salvos anteriormente
/*/
Static Function FExibePerg(lExibePerg)

	Local aPergs    := {}
	Local aRet      := {}

	Local nX        := 0
	Local nP        := 0

    Local dDataDe   := FirstDate(Date())
    Local dDataAt   := LastDate(Date())

	Local cNomPrg   := "FSCOMR01"

	Default lExibePerg := .T.

	// Lista de perguntas
	aAdd(aPergs, {1, "Data de Digitação De",  dDataDe, "", ".T.", "", ".T.", 80,  .T.})
	aAdd(aPergs, {1, "Data de Digitação Ate", dDataAt, "", ".T.", "", ".T.", 80,  .T.})

	//Carrega o array com os valores utilizados na última tela ou valores default de cada campo.
	For nX := 1 To Len(aPergs)
		aPergs[nX][3] := ParamLoad(cNomPrg,aPergs,nX,aPergs[nX][3])
	Next nX

	If lExibePerg
		If ParamBox(aPergs,"Parametros de Geracao...",@aRet)
			//Salva parametros
			ParamSave(cNomPrg,aPergs,"1")
		EndIf
	else
		For nP := 1 to Len(aPergs)

			aAdd(aRet,aPergs[nP][3])

		Next nP

	EndIf

Return(aRet)

/*/{Protheus.doc} FGetNFEnt
Função responsável por carregar os dados que serão impressos com base nos campos informados no aCampos.
@type function
@author gustavo.barcelos
@since 19/01/2023
/*/
Static Function FGetNFEnt(dDataDe, dDataAt)

    Local aRet      := {}
    Local aTmp      := {}
    Local aTotais   := Array(Len(aCampos) + 2)

    Local cAliNF    := GetNextAlias()
    Local cCampos   := ""
    Local cConsulta := ""
    Local cGroup    := ""

    Local nTotReg   := 0
    Local nC        := 0

    aEval(aCampos, { |x| cCampos += x[1] + "," })
    cConsulta   := "%" + cCampos + " D1_CC, CTT.CTT_DESC01 DESCC, DE_CC, CTT2.CTT_DESC01 DESCC2 %"
    cGroup      := "%" + cCampos + " D1_CC, CTT.CTT_DESC01, DE_CC, CTT2.CTT_DESC01 %"

    BeginSql Alias cAliNF
        
        SELECT %Exp:cConsulta%
        FROM %Table:SF1% SF1
        INNER JOIN %Table:SD1% SD1
            ON D1_FILIAL = F1_FILIAL
            AND D1_DOC = F1_DOC
            AND D1_FORNECE = F1_FORNECE
            AND D1_LOJA = F1_LOJA
            AND SD1.%NotDel%
        INNER JOIN %Table:SA2% SA2
            ON A2_FILIAL = %xFilial:SA2%
            AND A2_COD = F1_FORNECE
            AND A2_LOJA = F1_LOJA
            AND SA2.%NotDel%
        LEFT JOIN %Table:CTT% CTT
            ON CTT.CTT_FILIAL = %xFilial:CTT%
            AND CTT.CTT_CUSTO = D1_CC
            AND CTT.%NotDel%
        LEFT JOIN %Table:SDE% SDE
            ON DE_FILIAL = D1_FILIAL
            AND DE_DOC = D1_DOC
            AND DE_SERIE = D1_SERIE
            AND DE_FORNECE = D1_FORNECE
            AND DE_LOJA = D1_LOJA
            AND SDE.%NotDel%
        LEFT JOIN %Table:CTT% CTT2
            ON CTT2.CTT_FILIAL = %xFilial:CTT%
            AND CTT2.CTT_CUSTO = DE_CC
            AND CTT2.%NotDel%
        LEFT JOIN %Table:SC7% SC7
            ON C7_FILIAL = D1_FILIAL
            AND C7_NUM = D1_PEDIDO
            AND C7_ITEM = D1_ITEMPC
            AND SC7.%NotDel%
        WHERE F1_DTDIGIT BETWEEN %Exp:dDataDe% AND %exp:dDataAt%
        AND SF1.%NotDel%
        GROUP BY    %Exp:cGroup%
        ORDER BY    F1_FILIAL,
                    F1_DOC

    EndSql

    Count To nTotReg

    ProcRegua(nTotReg)

    (cAliNF)->(DbGoTop())

    While !((cAliNF)->(EOF()))

        IncProc()

        aTmp := {}

        For nC := 1 to Len(aCampos)
            If "USERLGI" $ aCampos[nC][1] .OR. "USERLGA" $ aCampos[nC][1]
                aAdd(aTmp, USRFULLNAME(SUBSTR(EMBARALHA(&(cAliNF + "->" + aCampos[nC][1]),1),3,6)))
            ElseIf GetSX3Cache(aCampos[nC][1], "X3_TIPO") == "D"
                aAdd(aTmp, DtoC(StoD(&(cAliNF + "->" + aCampos[nC][1]))))
            Else
                aAdd(aTmp, &(cAliNF + "->" + aCampos[nC][1]))
            EndIf

            If aCampos[nC][3]
                If Empty(aTotais[nC])
                    aTotais[nC] := 0
                EndIf

                aTotais[nC] += &(cAliNF + "->" + aCampos[nC][1])
            EndIf

        Next nC
        
        If !Empty((cAliNF)->D1_CC)
            aAdd(aTmp, (cAliNF)->D1_CC)
            aAdd(aTmp, (cAliNF)->DESCC)
        Else
            aAdd(aTmp, (cAliNF)->DE_CC)
            aAdd(aTmp, (cAliNF)->DESCC2)
        EndIf

        aAdd(aRet, aTmp)

        (cAliNF)->(DbSkip())
    EndDo

    aAdd(aRet, aTotais)

    (cAliNF)->(DbCloseArea())

Return(aRet)

/*/{Protheus.doc} FGeraExcel
Função responsável por gerar o arquivo excel com os dados carregados.
@type function
@author gustavo.barcelos
@since 19/01/2023
/*/
Static Function FGeraExcel(aDados)

    Local nC            := 0
    Local nAux          := 0
    Local nCol          := 0

    Local oFWMsExcel    
    Local oExcel

    Local cArquivo      := GetTempPath()+'NF_Ent_' + DtoS(Date()) + '_' + StrTran(Time(),":") + '.xls'
    Local cTable        := "Notas Fiscais de Entrada"

    Local aCabec        := {}
    Local aStyleTot     := {}

    ProcRegua(Len(aDados))

    // Monta cabeçalho de Campos que serão usados
    aCabec := FGetCabec()

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcelEx():New()
    oFWMsExcel:SetLineBgColor("#FFFFFF")
    oFWMsExcel:SetTitleSizeFont(14)
     
     // Compondo a aba do relatório
    oFWMsExcel:AddworkSheet(cTable) //Adiciona uma Worksheet ( Planilha ) Nome da planilha que será adicionada

    //Criando a Tabela
    oFWMsExcel:AddTable(cTable, cTable)

    //Criando Colunas
    For nC := 1 To Len(aCabec)
        oFWMsExcel:AddColumn(cTable, cTable, aCabec[nC][1], 1, aCabec[nC][2])
    Next

    oFWMsExcel:SetCelBold(.T.)
    oFWMsExcel:SetCelFrColor("#FFFFFF")
    oFWMsExcel:SetCelBgColor("#4682B4")

    //Criando Linhas
    For nAux := 1 To Len(aDados)-1
        IncProc(AllTrim(Str(nAux)) + " de " + AllTrim(Str(Len(aDados))))
        
        oFWMsExcel:AddRow(cTable, cTable, aDados[nAux])
    Next 

    // Inclui linha de totais
    aEval(aCampos, { |x| nCol++, aAdd(aStyleTot, nCol) })
    nCol++
    aAdd(aStyleTot, nCol)
    nCol++
    aAdd(aStyleTot, nCol)
    oFWMsExcel:AddRow(cTable, cTable, aDados[Len(aDados)], aStyleTot)

    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
            
    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)       //Abre uma planilha
    oExcel:SetVisible(.T.)                //Visualiza a planilha
    oExcel:Destroy()                      //Encerra o processo do gerenciador de tarefas

Return(Nil)

/*/{Protheus.doc} FGetCabec
Função responsável por montar o cabeçalho do Excel com base no campo do Dicionário.
@type function
@author gustavo.barcelos
@since 19/01/2023
/*/
Static Function FGetCabec()

    Local aRet      := {}

    Local nC        := 0

    Local cTitCol   := ""

    For nC := 1 to Len(aCampos)

        If Empty(aCampos[nC][4])
            cTitCol := GetSX3Cache(aCampos[nC][1], "X3_TITULO")
        Else
            cTitCol := aCampos[nC][4]
        EndIf

        aAdd(aRet, {cTitCol, aCampos[nC][2], aCampos[nC][3]})
       
    Next nC

    aAdd(aRet, {GetSX3Cache("D1_CC", "X3_TITULO"),1,.F.})
    aAdd(aRet, {"Desc. Centro de Custo",1,.F.})

Return(aRet)
