#include "protheus.ch"
#include "totvs.ch"
#INCLUDE "Topconn.ch"
#INCLUDE "PRTOPDEF.CH"

#define cEndLin Chr(13) + Chr(10)
//-------------------------------------------------------------------------------
/* Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
19/12/2018 Luiz Cruz: corrigidos alguns erros de digitação no SELECT que cria o arquvio temporário e alguns outros 
20/12/2018 Luiz Cruz: ajustei a linha do POSICIONE e acrescentei os campos D2_PEDIDO e D2_ITEM no SELECT de criação do arquivo temporário  
20/12/2018 Luiz Cruz: ajustei o "order by", acrescentando D2_CLIENTE,D2_LOJA,D2_COD 
03/01/2019 Luiz Cruz: passei a coluna Filial da posição 2 para a posição 1. Linhas 320 e 321
19/03/2023 GUSTAVO BARCELOS: corrigida a busca por centro de custos
*/
//-------------------------------------------------------------------------------
User Function IAPCPR06()  
	/* Variaveis criadas para classes TReport */
	Private oReport, oSection, oSection2
	Private cTitRel := "Faturamento por Cliente" /* Titulo do Relatório */
	Private cNomPro := "IAPCPR06" /* Nome do Programa do Relatório */
	Private cRelPrg := "IAPCPR06" /* Nome do Grupo de Perguntas do Relatório */

	/* Variáveis que receberão os valores do layout necessitado no relatório */
	Private cCodCli  := "" /* Cod. Produto */ 
	Private cLoja    := "" /*Loja Cliente*/
	Private cNomeCli := "" /* Nome Cliente */
	Private cCodPro  := "" /* Cod. Produto */
	Private cDscPro  := "" /* Descrição Produto */
	Private cTpNf    := "" /* Tipo NF */
	Private cNumDoc  := "" /* Numero Documento */
	Private cSerie   := "" /* Serie Documento */
	Private cCC      := "" /* Centre de Custo */
	Private cCCont   := "" /* Conta Contabil */
	PRIVATE CCF		 := "" /*                */
	Private cUm      := "" /* Unidade de Medida */
	Private nQtd     := 0  /* Quantidade*/
	Private nVlrUn   := 0  /* Vlr Unitario */
	Private nVlrTot  := 0  /* Vlr Total */
	Private nAlIcm   := 0  /* Aliquota de ICMS */
	Private nVlrIcm  := 0  /* Valor de ICMS  */
	Private nAlIss   := 0  /* Aliquota ICMS */
	Private nVlrIss  := 0  /* Valor ISS */
	Private nVlrIns  := 0   /* Valor INSS */
	Private nPis     := 0  /* Valor do PIS */  
	Private nCofins  := 0  /* Valor COFINS */
	Private nCsll    := 0  /* Valor CSLL */
	Private nIrrf    := 0  /* Valor IRFF */ 
	Private cCodFil := ""
    Private dEmis    := Ctod("")		

	/* Variáveis representantes de parâmetros */
	Private dDtaPar := Space(08) /* Ano + Mes (Competência) */
	Private cPrdIni := Space(TamSX3("D2_COD")[01]) /* Codigo do produto Inicial */
	Private cPrdFin := Space(TamSX3("D2_COD")[01]) /* Codigo do produto Final */
	Private cCliIni := Space(TamSX3("D2_CLIENTE")[01]) /* Codigo do Cliente Inicial */ 
	Private cCliFin := Space(TamSX3("D2_CLIENTE")[01]) /* Codigo do Cliente Inicial */  
	Private dDtEmisI := Space(TamSX3("D2_EMISSAO")[01]) /* Dt emissao Inicial */ 
	Private dDtEmisF := Space(TamSX3("D2_EMISSAO")[01]) /* Dt Emissao Final */  
	Private cCcIni := Space(TamSX3("D2_CCUSTO")[01]) /* Codigo CC Inicial */ 
	Private cCcFin := Space(TamSX3("D2_CCUSTO")[01]) /* Codigo CC Final*/   
	Private cFilIni := Space(TamSX3("D2_FILIAL")[01]) /* Codigo Filial Inicial */ 
	Private cFilFin := Space(TamSX3("D2_FILIAL")[01]) /* Codigo Filial Final   */
	  

	/* fGerPrg(). Chamada de Função para Criar as Perguntas. */
	Processa({||fGerPrg()},"Aguarde...","Gerando Parâmetros...")

	/* Chama a função CriaStru para Criar a estrutura XLS do relatór. */
	Processa({||fGerStr()},"Aguarde...","Gerando Estrutura...")

	/* Chama a função fTReport para Criar o relatório no Padrão TRep. */
	If Pergunte(cRelPrg,.T.)
		Processa({||fTReport()},"Aguarde...","Gerando Relatório...")
		oReport:PrintDialog()
	EndIf

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fGerPrg
Função - A funcao estática fGerPrg faz uma seleçao e cria automaticamente
         as perguntas no SX1.

Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
*/
//-------------------------------------------------------------------------------
Static Function fGerPrg()
	/* Ajusta SX1 segundo novos parametros */
	/* criacao de algumas variaveis para a inclusao das perguntas no sistema */
	Local nXi,nXj
	Local aSX1Reg := {}
	Local nTotReg := 0 /* Quantidade total de Registros */
	Local nRegLid := 0 /* Quantidade de Registros lidos */

	/* seleciona a tabela de perguntas SX1 */
	dbSelectArea("SX1")
	/* Determina ordenação 01 do índice de pesquisa */
	SX1->(dbSetOrder(01))

	//perguntas a serem criadas com seus respectivos parametros
	aAdd(aSX1Reg,{cRelPrg,"01","Data Referência:","","","mv_ch1","D",08,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","","","",".IAPCPR0601."})
	aAdd(aSX1Reg,{cRelPrg,"02","Produto De:     ","","","mv_ch2","C",TamSX3("D2_COD")[01],0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SB1","","",".IAPCPR0602."})
	aAdd(aSX1Reg,{cRelPrg,"03","Produto Até:    ","","","mv_ch3","C",TamSX3("D2_COD")[01],0,0,"G","!Empty(mv_par03) .And. (mv_par03>=mv_par02)","mv_par03","","","","","","","","","","","","","","","","","","","","","","","","","SB1","","",".IAPCPR0603."})
   
	aAdd(aSX1Reg,{cRelPrg,"04","Cliente De:       ","","","mv_ch4","C",TamSX3("D2_CLIENTE")[01],0,0,"G","","mv_par04","","","","","","","","","","","","","","","","","","","","","","","","","SA1","","",".IAPCPR0604."})
	aAdd(aSX1Reg,{cRelPrg,"05","Cliente Até:      ","","","mv_ch5","C",TamSX3("D2_CLIENTE")[01],0,0,"G","!Empty(mv_par05) .And. (mv_par05>=mv_par04)","mv_par05","","","","","","","","","","","","","","","","","","","","","","","","","SA1","","",".IAPCPR0605."})
	
	aAdd(aSX1Reg,{cRelPrg,"06","Filial De:     ","","","mv_ch6","C",TamSX3("D2_FILIAL")[01],0,0,"G","","mv_par06","","","","","","","","","","","","","","","","","","","","","","","","","SD2","","",".IAPCPR0606."})
	aAdd(aSX1Reg,{cRelPrg,"07","Filial Até:    ","","","mv_ch7","C",TamSX3("D2_FILIAL")[01],0,0,"G","!Empty(mv_par07) .And. (mv_par07>=mv_par06)","mv_par07","","","","","","","","","","","","","","","","","","","","","","","","","SD2","","",".IAPCPR0607."})
	
	aAdd(aSX1Reg,{cRelPrg,"08","Data Emissao De:     ","","","mv_ch8","D",TamSX3("D2_EMISSAO")[01],0,0,"G","","mv_par08","","","","","","","","","","","","","","","","","","","","","","","","","","","",".IAPCPR0608."})
	aAdd(aSX1Reg,{cRelPrg,"09","Data Emissao Até:    ","","","mv_ch9","D",TamSX3("D2_EMISSAO")[01],0,0,"G","!Empty(mv_par09) .And. (mv_par09>=mv_par08)","mv_par09","","","","","","","","","","","","","","","","","","","","","","","","","","","",".IAPCPR0609."})
    
    aAdd(aSX1Reg,{cRelPrg,"10","Centro de Custo De:     ","","","mv_ch10","C",TamSX3("D2_CCUSTO")[01],0,0,"G","","mv_par10","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","",".IAPCPR0610."})
	aAdd(aSX1Reg,{cRelPrg,"11","Centro de Custo Até:    ","","","mv_ch11","C",TamSX3("D2_CCUSTO")[01],0,0,"G","!Empty(mv_par11) .And. (mv_par11>=mv_par10)","mv_par11","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","",".IAPCPR0611."})
   
		/* Verifica a quantidade de perguntas */
	nTotReg := Len(aSX1Reg)

	/* Funcao para regua de processos */
	ProcRegua(nTotReg)

	/* laco for para fazer a inclusao das perguntas */
	For nXi := 01 to Len(aSX1Reg)
	   /* funcao da regua de processos incrementando */
		IncProc("Aguarde... Processando Registro " + Alltrim(Str(nRegLid)) + " de " + Alltrim(Str(nTotReg)))
		If !SX1->(dbSeek(cRelPrg+Space(Len(SX1->X1_GRUPO)-Len(cRelPrg))+aSX1Reg[nXi,2]))
			RecLock("SX1",.T.)
			For nXj := 01 to FCount()
				If nXj <= Len(aSX1Reg[nXi])
					FieldPut(nXj,aSX1Reg[nXi,nXj])
				Endif
			Next
			MsUnlock()
		EndIf
		nRegLid++
	Next nXi

	fSX1Hlp("P.IAPCPR0601.","","Código de data de referência para o filtro. A curba acontecerá referente a Mes/Ano desta data e seus próximos dois meses.","")
	fSX1Hlp("P.IAPCPR0602.","","Código inicial de intervalo de um produto para o filtro. Para um único produto insira o mesmo código no parâmetro abaixo.","")
	fSX1Hlp("P.IAPCPR0603.","","Código final de intervalo de um produto para o filtro. Para um único produto insira o mesmo código no parâmetro acima.","")   
	
	fSX1Hlp("P.IAPCPR0604.","","Código inicial de intervalo de Cliente para o filtro. Para um único Cliente insira o mesmo código no parâmetro abaixo.","")
	fSX1Hlp("P.IAPCPR0605.","","Código final de intervalo de um Cliente para o filtro. Para uma único Cliente insira o mesmo código no parâmetro acima.","")
	
	fSX1Hlp("P.IAPCPR0606.","","Código inicial de intervalo de uma Filial para o filtro. Para uma única Filial insira o mesmo código no parâmetro abaixo.","")
	fSX1Hlp("P.IAPCPR0607.","","Código final de intervalo de uma Filial para o filtro. Para uma única Filial insira o mesmo código no parâmetro acima.","")
	
	fSX1Hlp("P.IAPCPR0608.","","Código inicial de intervalo da Data para o filtro. Para uma única Data insira o mesmo código no parâmetro abaixo.","")
	fSX1Hlp("P.IAPCPR0609.","","Código final de intervalo de uma Data para o filtro. Para um única Data insira o mesmo código no parâmetro acima.","") 
	
	fSX1Hlp("P.IAPCPR06010.","","Código inicial de intervalo de um Centro de Custo para o filtro. Para um único Centro de Custo insira o mesmo código no parâmetro abaixo.","")
	fSX1Hlp("P.IAPCPR06011.","","Código final de intervalo de um Centro de Custo para o filtro. Para um único Centro de Custo insira o mesmo código no parâmetro acima.","")

Return

//-------------------------------------------------------------------------------
/*  {Protheus.doc} fSX1Hlp
F unção - A funcao estática fSX1Hlp gera o Help de Cada pergunta do relatorio de 
         acordo com os parametros.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12

@parms pKeyPrg - Chave da Pergunta a ser inserida
@parms cIn1Hlp - Descrição Help Inglês
@parms cPt1Hlp - Descrição Help Português
@parms cEs1Hlp - Descrição Help Espanhol
  */
//-------------------------------------------------------------------------------
Static Function fSX1Hlp(pKeyPrg,pIn1Hlp,pPt1Hlp,pEs1Hlp)
	Local nXi := 0 /* Contador */
	Local nNumLin := 0  /* Quantidade de caracteres por linha */
	Local aIngHlp := {} /* Help da Pergunta em inglês */
	Local aPorHlp := {} /* Help da Pergunta em português */
	Local aEspHlp := {} /* Help da Pergunta em espanhol */

	/* Montagem da estrutura de help de perguntas do parametro informado
	 * sendo neste momento feita para a lingua Português */
	If !Empty(Alltrim(pPt1Hlp))
		nNumLin := Mlcount(pPt1Hlp,40)
		For nXi := 01 To nNumLin
			If !Empty(Memoline(pPt1Hlp,40,nXi))
				aAdd(aPorHlp,OemToAnsi(Memoline(pPt1Hlp,40,nXi)))
			Endif
		Next nXi
	Else
		aAdd(aPorHlp,"") /* Grava vazio na primeira e segunda linha */
		aAdd(aPorHlp,"") /* Grava vazio na primeira e segunda linha */
    EndIf

	/* Montagem do estrutura de help de perguntas do parametro informado
	 * sendo neste momento feita para a lingua Inglês */
	If !Empty(Alltrim(pIn1Hlp))
		nNumLin := Mlcount(pIn1Hlp,40)
		For nXi := 01 To nNumLin
			If !Empty(Memoline(pIn1Hlp,40,nXi))
				aAdd(aIngHlp,OemToAnsi(Memoline(pIn1Hlp,40,nXi)))
			Endif
		Next nXi
	Else
		aIngHlp := aPorHlp /* Recebe mesmo conteudo da lingua português */
    EndIf

	/* Montagem do estrutura de help de perguntas do parametro informado
	 * sendo neste momento feita para a lingua Espanhol */
	If !Empty(Alltrim(pEs1Hlp))
		nNumLin := Mlcount(pEs1Hlp,40)
		For nXi := 01 To nNumLin
			If !Empty(Memoline(pEs1Hlp,40,nXi))
				aAdd(aEspHlp,OemToAnsi(Memoline(pEs1Hlp,40,nXi)))
			Endif
		Next nXi
	Else
		aEspHlp := aPorHlp /* Recebe mesmo conteudo da lingua português */
    EndIf

	/* Função utilizada para cadastro de helps no Protheus. */
	PutSX1Help(pKeyPrg,aPorHlp,aIngHlp,aEspHlp)

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fGerStr
Função - A funcao estática fGerStr faz todo o arquivo de trabalho verificando 
         sempre os parametros das perguntas.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
*/
//-------------------------------------------------------------------------------
Static Function fGerStr()
	Local cArqTrb := "" /* Arquivo de trabalho */
	Local aStrXls := {} /* Strutura de Arquivo de trabalho XLS */
	//Local cChkInd := "" /* Validação de Índice */

	/* Variáveis que receberão os valores do layout necessitado no relatório */
    
	aAdd(aStrXls,{"TMP_CODCLI","C",TamSX3("D2_CLIENTE")[01],00})
	aAdd(aStrXls,{"TMP_FIL","C",TamSX3("D2_FILIAL")[01],00}) 
	aAdd(aStrXls,{"TMP_LOJA","C",TamSX3("D2_LOJA")[01],00}) 
	aAdd(aStrXls,{"TMP_DSCCLI","C",TamSX3("A1_NOME")[01],00}) 
	aAdd(aStrXls,{"TMP_CODPRD","C",TamSX3("D2_COD")[01],00}) 
	aAdd(aStrXls,{"TMP_DSCPRD","C",TamSX3("B1_DESC")[01],00})
	aAdd(aStrXls,{"TMP_TPNF","C",TamSX3("D2_TIPO")[01],00})
	aAdd(aStrXls,{"TMP_NUMDOC","C",TamSX3("D2_DOC")[01],00})
    aAdd(aStrXls,{"TMP_SERIE","C",TamSX3("D2_SERIE")[01],00})
    aAdd(aStrXls,{"TMP_DATA","D",TamSX3("D2_EMISSAO")[01],00}) 
    aAdd(aStrXls,{"TMP_CCUSTO","C",TamSX3("D2_CCUSTO")[01],00}) 
    aAdd(aStrXls,{"TMP_CCONT","C",TamSX3("D2_CONTA")[01],00}) 
    aAdd(aStrXls,{"TMP_CF","C",TamSX3("D2_CF")[01],00})
    aAdd(aStrXls,{"TMP_GRUPO","C",15})
	aAdd(aStrXls,{"TMP_UNMPRD","C",TamSX3("D2_UM")[01],00})
    aAdd(aStrXls,{"TMP_QTD","N",14,02})
    aAdd(aStrXls,{"TMP_VLRUN","N",14,02})
    aAdd(aStrXls,{"TMP_VLRTOT","N",14,02})
    aAdd(aStrXls,{"TMP_ALQICM","N",14,02})
    aAdd(aStrXls,{"TMP_VLRICM","N",14,02})
    aAdd(aStrXls,{"TMP_ALQISS","N",14,02})
    aAdd(aStrXls,{"TMP_VLRISS","N",14,02})
    aAdd(aStrXls,{"TMP_VLRINS","N",14,02})
    aAdd(aStrXls,{"TMP_PIS","N",14,02})
    aAdd(aStrXls,{"TMP_CONFIN","N",14,02})
    aAdd(aStrXls,{"TMP_CSLL","N",14,02}) 
    aAdd(aStrXls,{"TMP_IRRF","N",14,02})
     
	//SELECT E FECHA O ARQ. DE TRAB.
	fCloseArea("XLS")

	cArqTrb := CriaTrab(aStrXls,.T.)
	dbUseArea(.T.,,cArqTrb,"XLS")

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fTReport
Função - A funcao estática fTReport faz algumas validacoes para o funcionamento
		 do Relatorio em Release 04.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
*/
//-------------------------------------------------------------------------------
Static Function fTReport()

	/* A classe TReport permite que o usuário personalize as informações que serão apresentadas 
	   no relatório, alterando fonte (tipo, tamanho, etc), cor, tipo de linhas, cabeçalho, rodapé, etc.*/
   
   // ALTEAÇÃO MARIO 20/12 oReport := TReport():New(cNomPro,cTitRel,/*cRelPrg*/,{|oReport| fGerTmp()},"Este relatório irá imprimir o Faturamento de Cliente")
	oReport := TReport():New(cNomPro,cTitRel,cRelPrg,{|oReport| fGerTmp()},"Este relatório irá imprimir o Faturamento de Cliente")
	Pergunte(cRelPrg,.F.)
	oReport:SetPortrait()
	oReport:SetTotalInLine(.F.)

	/* A classe TRSection pode ser entendida como um layout do relatório, por conter 
	   células, quebras e totalizadores que darão um formato para sua impressão. */
	//TRSection():New ( < oParent > , [ cTitle ] , [ uTable ] , [ aOrder ] , [ lLoadCells ] , [ lLoadOrder ] ) --> TRSection
	oSection := TRSection():New(oReport,"Faturamento de Clientes ",{"XLS"},{"Filial","Cod. Cliente"})
	//oSection2 := TRSection():New(oSection,"Contratos",{"XL2","Contratos"})

	/*Se o nome da célula informada pelo parametro for encontrado no Dicionário de
	Campos (SX3), as informações do campo serão carregadas para a célula, respeitando
	os parametros de título, picture e tamanho. Dessa forma o relatório sempre estará
	atualizado com as informações do Dicionário de Campos (SX3).*/
	//TRCell():New(/*OBJETO*/,/*CAMPO*/,/*ARQ.TRAB.*/,/*TITULO*/,/*PICTURE*/,/*TAMANHO*/,/*lPixel*/,/*{|| codblock de impressao }*/
	
	TRCell():New(oSection,"TMP_FIL","XLS","Filial","@!",TamSX3("D2_FILIAL")[01]) 
	TRCell():New(oSection,"TMP_CODCLI","XLS","Cod. Cliente",,TamSX3("D2_CLIENTE")[01]) 
	TRCell():New(oSection,"TMP_LOJA","XLS","Loja",,TamSX3("D2_LOJA")[01]) 
	TRCell():New(oSection,"TMP_DSCCLI","XLS","Cliente",,TamSX3("A1_NOME")[01]) 
	TRCell():New(oSection,"TMP_CODPRD","XLS","Cód. Produto",,TamSX3("D2_COD")[01]) 
    TRCell():New(oSection,"TMP_DSCPRD","XLS","Produto",,TamSX3("B1_DESC")[01])
	TRCell():New(oSection,"TMP_TPNF","XLS","Tipo Documento",,TamSX3("D2_TIPO")[01])
    TRCell():New(oSection,"TMP_NUMDOC","XLS","Numero Documento",,TamSX3("D2_DOC")[01])
    TRCell():New(oSection,"TMP_SERIE","XLS","Serie",,TamSX3("D2_SERIE")[01])
    TRCell():New(oSection,"TMP_DATA","XLS","Data de Emissao",,TamSX3("D2_EMISSAO")[01]) 
    TRCell():New(oSection,"TMP_CCUSTO","XLS","Centro de Custo",,TamSX3("D2_CCUSTO")[01]) 
    TRCell():New(oSection,"TMP_CCONT","XLS","Conta Contabil",,TamSX3("D2_CONTA")[01]) 
    TRCell():New(oSection,"TMP_CF","XLS","Classificacao Fiscal",,TamSX3("D2_CF")[01]) 
	TRCell():New(oSection,"TMP_GRUPO","XLS","GRUPO",,15)
   	TRCell():New(oSection,"TMP_UNMPRD","XLS","U.M",,TamSX3("D2_UM")[01])
    TRCell():New(oSection,"TMP_QTD","XLS","Quantidade","@E 999,999,999.9999",14)
    TRCell():New(oSection,"TMP_VLRUN","XLS","Vlr Unitario","@E 999,999,999.9999",14)
    TRCell():New(oSection,"TMP_VLRTOT","XLS","Vlr Total","@E 999,999,999.9999",14)
    TRCell():New(oSection,"TMP_ALQICM","XLS","Aliq. ICMS","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_VLRICM","XLS","Vlr. ICMS","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_ALQISS","XLS","Aliq. ISS","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_VLRISS","XLS","Vlr. ISS","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_VLRINS","XLS","Vlr. INSS","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_PIS","XLS","Pis","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_CONFIN","XLS","Cofins","@E 99,999,999,999.99",14)
    TRCell():New(oSection,"TMP_CSLL","XLS","Vlr. Csll","@E 99,999,999,999.99",14) 
    TRCell():New(oSection,"TMP_IRRF","XLS","Vlr. Irrf","@E 99,999,999,999.99",14)
	
	
//	TRCell():New(oSection2,"TMP_CODCO2","XL2","CONTRATO","",TamSX3("Z1_NUM")[01]) //Código do Pedido de Venda
//	TRCell():New(oSection2,"TMP_CLIENT","XL2","CLIENTE","",50) //Cliente

	/* neste caso verifica o arquivo de trabalho gerado e faz a exclusão
	dos dados e da tabela temporária */
	fCloseArea("IATRB")

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fGerTmp
Função - A funcao estática fGerTmp faz todo o arquivo de trabalho verificando 
         sempre os parametros das perguntas.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
*/
//-------------------------------------------------------------------------------
Static Function fGerTmp()
   Local cQrySql := "" //Variavel na qual é armazenado a query de consulta ao banco

	/* Variáveis representantes de parâmetros */
	 cPrdIni  :=MV_PAR02
	 cPrdFin  :=MV_PAR03
	 cCliIni  :=MV_PAR04 
	 cCliFin  :=MV_PAR05
	 dDtEmisI :=MV_PAR08 
	 dDtEmisF :=MV_PAR09  
	 cCcIni   :=MV_PAR10  
	 cCcFin   :=MV_PAR11  
	 cFilIni  :=MV_PAR06  
	 cFilFin  :=MV_PAR07 

	/* montagem da clausula SELECT */
     cQrySql += " SELECT D2_FILIAL, D2_LOJA, D2_CLIENTE, A1_NOME, D2_COD, B1_DESC, D2_TIPO,"
     cQrySql += " D2_DOC,D2_SERIE,CAST(D2_EMISSAO AS DATE) AS DT_EMISS, D2_CONTA,D2_CF, "
 	 cQrySql += " ISNULL(CT1.CT1_ZGRUPO,'')  GRUPO,	
	 cQrySql += " D2_UM,D2_QUANT,D2_PRCVEN, D2_TOTAL, D2_PICM,D2_VALICM,D2_ALIQISS,D2_VALISS," 
	 cQrySql += " D2_VALINS,D2_VALIMP5,D2_VALIMP6,D2_VALCSL, D2_VALIRRF, D2_PEDIDO,D2_ITEMPV,D2_ITEM,D2_VALPIS,D2_VALCOF"

	/* montagem da clausula FROM */
//     cQrySql += " FROM "+ RetSqlName("SD2")
     cQrySql += " FROM "+ RetSqlName("SD2") + " AS D2 "
     cQrySql += " INNER JOIN "+ RetSqlName("SA1") + " AS A1 ON D2_CLIENTE=A1_COD AND SUBSTRING(D2_FILIAL,1,2)=SUBSTRING(A1_FILIAL,1,2)AND A1_LOJA=D2_LOJA "
     cQrySql += " INNER JOIN "+ RetSqlName("SB1") + " AS B1 ON D2_COD=B1_COD   AND D2_FILIAL=B1_FILIAL "
     cQrySql += " INNER JOIN "+ RetSqlName("SF4") + " AS F4 ON D2_TES=F4_CODIGO  AND SUBSTRING(D2_FILIAL,1,2)=SUBSTRING(F4_FILIAL,1,2) "
     cQrySql += " LEFT OUTER JOIN "+ RetSqlName("CT1") + " AS CT1 ON CT1_FILIAL = '"+xFilial("CT1")+"' AND  D2_CONTA = CT1.CT1_CONTA AND CT1.D_E_L_E_T_ <> '*' "

	//inserir inner do sf4 filtrando gera dpl
	
	cQrySql += " WHERE A1.D_E_L_E_T_='' AND D2.D_E_L_E_T_=''  AND  B1.D_E_L_E_T_=''  AND  F4.D_E_L_E_T_='' "
     cQrySql += " AND D2_FILIAL>= '"+cFilIni+ "' AND D2_FILIAL<= '"+cFilFin+"' "  
//     cQrySql += " AND D2_EMISSAO>= '"+dDtEmisI+ "' AND D2_EMISSAO<= '"+dDtEmisF+"' "  
     cQrySql += " AND D2_EMISSAO>= '"+dtos(dDtEmisI)+ "' AND D2_EMISSAO<= '"+dtos(dDtEmisF)+"' "
     cQrySql += " AND D2_CLIENTE>= '"+cCliIni+ "' AND D2_CLIENTE<= '"+cCliFin+"' "  
     cQrySql += " AND D2_COD>= '"+cPrdIni+ "' AND D2_COD<= '"+cPrdFin+"'  AND F4_DUPLIC='S'  "  

	/* implementação da classe para imprimir totalizadores s
	TRFunction():New(oSection:Cell("TMP_MES001"),Nil,"SUM",,,,,.F.,,)
	TRFunction():New(oSection:Cell("TMP_MES002"),Nil,"SUM",,,,,.F.,,)
	TRFunction():New(oSection:Cell("TMP_MES003"),Nil,"SUM",,,,,.F.,,)
	TRFunction():New(oSection:Cell("TMP_TOTGER"),Nil,"SUM",,,,,.F.,,)
	TRFunction():New(oSection:Cell("TMP_VMEDIA"),Nil,"SUM",,,,,.F.,,)
	*/
	
	fCloseArea("IATRB")
	dbUseArea(.T., "TOPCONN",TCGenQry(,,cQrySql),"IATRB",.T., .T.)

	/* Chama a função fProTmp para processar temporario e gravar. */
	Processa({||fProTmp()},"Aguarde...","Gravando Informações...")

	/* Chama a função fImpRel para Imprimir informações do Relatorio */
	fImpRel()

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fProTmp
Função - A funcao estática fProTmp faz todo o processode trabalho verificando 
         sempre os parametros das perguntas.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
//-------------------------------------------------------------------------------   
*/
Static Function fProTmp()
	Local nTotReg := 0 /* Quantidade total de Registros */
	Local nRegLid := 0 /* Quantidade de Registros lidos */
	Private nTotMed := 0 /* Total das Médias */

	/* seleciono o arquivo de trabalho gerado pela query e coloco no inicio */
	dbSelectArea("IATRB")
	/* Totaliza os registros da tabela */
	IATRB->(dbEval({||nTotReg++}))
	/* Posiciona no primeiro registro */
	IATRB->(dbGoTop())

	/* Funcao para regua de processos */
	ProcRegua(nTotReg)

	Do While !IATRB->(Eof())
		/* funcao da regua de processos incrementando */
		IncProc("Aguarde... Processando Registro " + Alltrim(Str(nRegLid)) + " de " + Alltrim(Str(nTotReg)))

		/* Chama a função fGrvTmp para Gravar o Tmp. no Arq. Trab. XLS. */
		fGrvTmp()

		/* seleciono o arquivo de trabalho gerado pela query e coloco no inicio */
		dbSelectArea("IATRB")
		/* avança para proximo registro */
		IATRB->(dbSkip())
		/* atualiza contador da régua */
		nRegLid++
	EndDo

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fGrvTmp
Função - A funcao estática fGrvTmp armazena os valores gravados nas variaveis de 
         memória no arquivo de trabalho XLS.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
*/
//-------------------------------------------------------------------------------
Static Function fGrvTmp()
     
//     cCusto	:= POSICIONE("SC6",1,XFILIAL("SC6")+(cSD2)->D2_PEDIDO+(cSD2)->D2_ITEM+(cSD2)->D2_COD,"C6_ZCCUSTO") 
	// GUSTAVO BARCELOS - 19/03/2023 - Alterado o campo de busca do item do Pedido de Vendas, sendo que agora a busca usará o campo D2_ITEMPV.
    //  cCusto	:= POSICIONE("SC6",1,IATRB->D2_FILIAL+IATRB->D2_PEDIDO+IATRB->D2_ITEM+IATRB->D2_COD,"C6_ZCCUSTO") 
     cCusto	:= POSICIONE("SC6",1,IATRB->D2_FILIAL+IATRB->D2_PEDIDO+IATRB->D2_ITEMPV+IATRB->D2_COD,"C6_ZCCUSTO") 
     If cCusto = " "
//		cCusto:= POSICIONE("DTC",17,xFilial("DTC")+(cSD2)->D2_DOC+(cSD2)->D2_SERIE,"DTC_ZCUSTO")				  
		// GUSTAVO BARCELOS - 19/03/2023 - Alterado o índice da busca para carregar o centro de custo corretamente.
		// cCusto:= POSICIONE("DTC",18,D2_FILIAL+IATRB->D2_DOC+IATRB->D2_SERIE,"DTC_ZCUSTO")				  
		cCusto:= POSICIONE("DTC",20,IATRB->D2_FILIAL+IATRB->D2_DOC+IATRB->D2_SERIE,"DTC_ZCUSTO")				  
    EndIF
                                                                  
     If cCusto >= cCcIni .And. cCusto <= cCcFin 
//   If cCcIni >= cCusto .And. cCcFin <= cCusto
	     cCodCli  := IATRB->D2_CLIENTE
	     cCodFil  := IATRB->D2_FILIAL
		 cLoja    := IATRB->D2_LOJA
		 cNomeCli := IATRB->A1_NOME
		 cCodPro  := IATRB->D2_COD
		 cDscPro  := IATRB->B1_DESC
		 cTpNf    := IATRB->D2_TIPO
//		 cNumDoc  := IATRB->D2_NUM 
  		 dEmis    := IATRB->DT_EMISS
		 cNumDoc  := IATRB->D2_DOC
		 cSerie   := IATRB->D2_SERIE
		 cCC      := cCusto 
		 cCCont   := IATRB->D2_CONTA
		 CCF	  := IATRB->D2_CF
		 CGRUPO	  := IATRB->GRUPO
		 cUm      := IATRB->D2_UM
		 nQtd     := IATRB->D2_QUANT
		 nVlrUn   := IATRB->D2_PRCVEN
		 nVlrTot  := IATRB->D2_TOTAL
		 nAlIcm   := IATRB->D2_PICM
		 nVlrIcm  := IATRB->D2_VALICM
		 nAlIss   := IATRB->D2_ALIQISS
		 nVlrIss  := IATRB->D2_VALISS
		 nVlrIns  := IATRB->D2_VALINS
//		 nCofins  := IATRB->D2_VALIMP5 
//		 nPis     := IATRB->D2_VALIMP6
		 nCofins  := IATRB->D2_VALCOF 
		 nPis     := IATRB->D2_VALPIS
		 nCsll    := IATRB->D2_VALCSL  
		 nIrrf    := IATRB->D2_VALIRRF
		 		
		dbSelectArea("XLS")
		If XLS->(RecLock("XLS",.T.)) 
		     
			XLS->TMP_CODCLI := cCodCli
			XLS->TMP_FIL := cCodFil 
			XLS->TMP_LOJA := cLoja 
			XLS->TMP_DSCCLI := cNomeCli
			XLS->TMP_CODPRD := cCodPro 
		    XLS->TMP_DSCPRD := cDscPro
			XLS->TMP_TPNF:= cTpNf
		    XLS->TMP_NUMDOC := cNumDoc
		    XLS-> TMP_SERIE := cSerie
//		    XLS->TMP_DATA := dData 
		    XLS->TMP_DATA := dEmis 
		    XLS->TMP_CCUSTO := cCC
		    XLS->TMP_CCONT := cCCont 
		    XLS->TMP_CF    := CCF
		    XLS->TMP_GRUPO :=CGRUPO
			XLS->TMP_UNMPRD := cUm
		    XLS->TMP_QTD := nQtd
		    XLS->TMP_VLRUN := nVlrUn
		    XLS->TMP_VLRTOT := nVlrTot
		    XLS->TMP_ALQICM := nAlIcm
		    XLS->TMP_VLRICM := nVlrIcm
		    XLS->TMP_ALQISS := nAlIss
		    XLS->TMP_VLRISS := nVlrIss
		    XLS->TMP_VLRINS := nVlrIns
		    XLS->TMP_PIS := nPis
		    XLS->TMP_CONFIN := nCofins
		    XLS->TMP_CSLL := nCsll
		    XLS->TMP_IRRF := nIrrf        
	    
			XLS->(MsUnlock())
		EndIf
	EndIf
  
Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fImpRel
Função - A funcao estática fImpRel faz a impressão dos valores capturados e 
         gravados na XLS gerando o Relatório no Microsiga.
            
Relatório - FATURAMENTO CLIENTE
Copyright I AGE©
@author MARIO RONCA
@since 19/12/2018
@version P12
*/
//-------------------------------------------------------------------------------
Static Function fImpRel()
	//seleciono o arquivo de trabalho gerado pela query e coloco no inicio
	dbSelectArea("XLS")
	XLS->(dbGoTop())

	//Seta o contador da regua
	oReport:SetMeter(XLS->(RecCount()))	

	//Posiciona no primeiro registro
	dbSelectArea("XLS")
	XLS->(dbGoTop())
	
	//Inicializa a Seção
	oSection:Init()

	Do While !XLS->(Eof())
		//Verifica se Cancelou
		If oReport:Cancel()
			Exit
		EndIf

		/*Processa as informações da tabela principal ou 
		da query definida pelo Embedded SQL com os métodos BeginQuery e EndQuery*/
		oSection:PrintLine()

		/*Incrementa a régua da tela de processamento do relatório*/
		oReport:IncMeter()

		dbSelectArea("XLS")
		XLS->(dbSkip())
	EndDo

	oSection:Finish()

Return

//-------------------------------------------------------------------------------
/* {Protheus.doc} fCloseArea
Função - A funcao estática fCloseArea recebe como parametro um nome de uma 
         tabela, seja existente ou temporaria e fecha arquivo
*/
//-------------------------------------------------------------------------------
Static Function fCloseArea(pCodTabe)

	If (Select(pCodTabe)!= 0)
		dbSelectArea(pCodTabe)
		dbCloseArea()
		If File(pCodTabe+GetDBExtension())
			FErase(pCodTabe+GetDBExtension())
		EndIf
	EndIf

Return


// pedido da nayrane , o que nao gera financeiro nao sair no relatorio solicitado em 14/09/2021 
