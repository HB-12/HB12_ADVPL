#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'

// Campos que serão impressos. Podem ser alterados desde que estejam nas tabelas consideradas na query da função FGetNFEnt.
/*
    Estrutura do array aCampos:
    [1] - Campo do banco de dados para exibição
    [2] - Tipo do Campo no arquivo Excel (1-General,2-Number,3-Monetário,4-DateTime)
    [3] - Indica se a coluna deve ser totalizada
    [4] - Indica o nome personalizado que a coluna deve assumir (quanto não informado, assume da SX3. Em casos de campos calculados, deve ser sempre preenchido)
*/
#DEFINE aCampos    { {"N1_FILIAL"    , 1, .F., Nil },;                      // Não remover!
                     {"N1_CBASE"     , 1, .F., Nil },;                      // Não remover!
                     {"N1_ITEM"      , 1, .F., Nil },;                      // Não remover!
                     {"N1_AQUISIC"   , 1, .F., Nil },;
                     {"N1_QUANTD"    , 2, .T., Nil },;
                     {"N1_DESCRIC"   , 1, .F., Nil },;
                     {"N1_CHAPA"     , 1, .F., Nil },;
                     {"N1_APOLICE"   , 1, .F., Nil },;
                     {"N1_CODSEG"    , 1, .F., Nil },;
                     {"N1_FORNEC"    , 1, .F., Nil },;                      // Não remover!
                     {"N1_LOJA"      , 1, .F., Nil },;                      // Não remover!
                     {"A2_NOME"      , 1, .F., Nil },;
                     {"N1_NFISCAL"   , 1, .F., Nil },;
                     {"N1_CHASSIS"   , 1, .F., Nil },;
                     {"N1_PLACA"     , 1, .F., Nil },;
                     {"T9_CIDEMPL"   , 1, .F., Nil },;
                     {"T9_UFEMPLA"   , 1, .F., Nil },;
                     {"T9_RENAVAM"   , 1, .F., Nil },;
                     {"N1_CODBEM"    , 1, .F., Nil },;                      // Não remover!
                     {"N1_DIACTB"    , 1, .F., Nil },;
                     {"N1_TAXAPAD"   , 1, .F., Nil },;
                     {"N1_NODIA"     , 1, .F., Nil },;
                     {"N1_DETPATR"   , 1, .F., Nil },;
                     {"N1_UTIPATR"   , 1, .F., Nil },;
                     {"N1_VLAQUIS"   , 3, .T., Nil },;
                     {"N1_NUMPRO"    , 1, .F., Nil },;
                     {"N1_ZBOSERV"   , 1, .F., Nil },;
                     {"N1_CONTR"     , 1, .F., Nil },;
                     {"N1_NFINANC"   , 1, .F., Nil },;
                     {"N1_DTINICC"   , 1, .F., Nil },;
                     {"N1_DTFCON"    , 1, .F., Nil },;
                     {"N3_CUSTBEM"   , 1, .F., Nil },;
                     {"N1_SLBMCON"   , 1, .F., Nil }    }

/*/{Protheus.doc} FSATFR01
Rotina responsável pela geração de um arquivo excel com a listagem dos bens do Ativo Fixo.
@type function
@author gustavo.barcelos
@since 3/2/2023
/*/
User Function FSATFR01()

    Local cCadastro	:= "Listagem de Ativos"

	Local nOpca 	:= 0		// Flag de confirmacao para OK ou CANCELA

	Local aSays		:= {} 		// Array com as mensagens explicativas da rotina
	Local aButtons	:= {}		// Array com as botões da rotina
	Local aDados    := {}       // Array com os dados encontrados para impressão

    //																							 |
	AADD(aSays, "Este programa tem como objetivo gerar um arquivo contendo o resumo dos ")
	AADD(aSays, "bens do Ativo Fixo que não foram baixados. ")

	aAdd( aButtons, { 6, .T., { || nOpca:=1, FechaBatch() } } )
	aAdd( aButtons, { 2, .T., { || FechaBatch() } } )

	FormBatch( cCadastro, aSays, aButtons )

	If ( nOpcA == 1 )

        Processa({|| aDados := FGetBens() }, "Carga de Dados...", "Carregando dados para Impressão...", .F.)

        If !Empty(aDados)
            Processa({ || FGeraExcel(aDados) }, "Gerando arquivo Excel..." )
        Else
            FWAlertInfo("Nenhum registro foi encontrado para impressão.",;
                        "Dados para impressão não encontrados!")
        EndIf

    EndIf

Return(Nil)

/*/{Protheus.doc} FGetBens
Função responsável por carregar os dados que serão impressos com base nos campos informados no aCampos.
@type function
@author gustavo.barcelos
@since 02/03/2023
/*/
Static Function FGetBens()

    Local aRet      := {}
    Local aTmp      := {}
    Local aTotais   := Array(Len(aCampos))

    Local cAliSN1   := GetNextAlias()
    Local cCampos   := ""
    Local cConsulta := ""

    Local nTotReg   := 0
    Local nC        := 0

    aEval(aCampos, { |x| cCampos += x[1] + "," })
    cConsulta   := "%" + SubStr(cCampos,1,Len(cCampos)-1) + " %"

    BeginSql Alias cAliSN1
        
        SELECT %Exp:cConsulta%
        FROM %Table:SN1% SN1
        LEFT JOIN %Table:SA2% SA2
            ON  A2_FILIAL = N1_FILIAL
            AND A2_COD = N1_FORNEC
            AND A2_LOJA = N1_LOJA
            AND SA2.%NotDel%
        LEFT JOIN %Table:SN3% SN3
            ON  N3_FILIAL = N1_FILIAL
            AND N3_CBASE = N1_CBASE
            AND N3_ITEM = N1_ITEM
            AND N3_TIPO = '01'
            AND N3_SEQ = '001'
            AND SN3.%NotDel%
        LEFT JOIN %Table:ST9% ST9
            ON  T9_FILIAL = %xFilial:ST9%
            AND T9_CODBEM = N1_CODBEM
            AND ST9.%NotDel%
        WHERE N1_CBASE <> ''
        AND N1_BAIXA = ''
        AND SN1.%NotDel%
        ORDER BY    N1_FILIAL,
                    N1_CBASE

    EndSql

    Count To nTotReg

    ProcRegua(nTotReg)

    (cAliSN1)->(DbGoTop())

    While !((cAliSN1)->(EOF()))

        IncProc()

        aTmp := {}

        For nC := 1 to Len(aCampos)
            If "USERLGI" $ aCampos[nC][1] .OR. "USERLGA" $ aCampos[nC][1]
                aAdd(aTmp, USRFULLNAME(SUBSTR(EMBARALHA(&(cAliSN1 + "->" + aCampos[nC][1]),1),3,6)))
            ElseIf GetSX3Cache(aCampos[nC][1], "X3_TIPO") == "D"
                aAdd(aTmp, DtoC(StoD(&(cAliSN1 + "->" + aCampos[nC][1]))))
            Else
                aAdd(aTmp, &(cAliSN1 + "->" + aCampos[nC][1]))
            EndIf

            If aCampos[nC][3]
                If Empty(aTotais[nC])
                    aTotais[nC] := 0
                EndIf

                aTotais[nC] += &(cAliSN1 + "->" + aCampos[nC][1])
            EndIf

        Next nC

        aAdd(aRet, aTmp)

        (cAliSN1)->(DbSkip())
    EndDo

    aAdd(aRet, aTotais)

    (cAliSN1)->(DbCloseArea())

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

    Local cArquivo      := GetTempPath()+'Ativos_' + DtoS(Date()) + '_' + StrTran(Time(),":") + '.xls'
    Local cTable        := "Listagem de Ativos"

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

Return(aRet)
