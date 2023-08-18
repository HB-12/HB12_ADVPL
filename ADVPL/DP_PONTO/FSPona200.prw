#INCLUDE "FILEIO.CH"
#INCLUDE "PROTHEUS.CH"

//Posi��es do Array
Static nPosFilial 	:= 1 
Static nPosMatr	    := 2
Static nPosData		:= 3 
Static nPosEvento  	:= 4
Static nPosLanc  	:= 5

/*/{Protheus.doc} FSPona200
Fun��o copia do browser PONA200 para incluir importacao de arquivo
@author Alex Teixeira
@since 08/02/2023
@version 1.0
@type function
/*/

User Function FSPona200

Local cFiltraSRA	:= ""
Local aIndexSRA		:= {}
Local aArea			:= GetArea()
Local aAreaSPI		:= SPI->( GetArea() )

Private bFiltraBrw	:= {|| NIL}
Private cCadastro	:= OemToAnsi("Manuten��o Banco de Horas" ) 

/*
��������������������������������������������������������������Ŀ
� Define array contendo as Rotinas a executar do programa      �
� ----------- Elementos contidos por dimensao ------------     �
� 1. Nome a aparecer no cabecalho                              �
� 2. Nome da Rotina associada                                  �
� 3. Usado pela rotina                                         �
� 4. Tipo de Transa��o a ser efetuada                          �
�    1 - Pesquisa e Posiciona em um Banco de Dados             �
�    2 - Simplesmente Mostra os Campos                         �
�    3 - Inclui registros no Bancos de Dados                   �
�    4 - Altera o registro corrente                            �
�    5 - Remove o registro corrente do Banco de Dados          �
����������������������������������������������������������������*/
Private aRotina	:= MenuDef()

// - Valida se o usu�rio tem acesso
If BloqPer()

	Return (Nil)

EndIf

/*
������������������������������������������������������������������������Ŀ
�So executa se o Modo de Acesso do SPI e SRA foram iguais e se este  ulti�
�mo nao estiver vazio.                                                   �
��������������������������������������������������������������������������*/
IF ( ValidArqPon() .and. ChkVazio("SRA") )

	/*
	������������������������������������������������������������������������Ŀ
	� Inicializa o filtro utilizando a funcao FilBrowse                      �
	��������������������������������������������������������������������������*/
	cFiltraRh := CHKRH("PONA200","SRA","1")
	bFiltraBrw 	:= {|| FilBrowse("SRA",@aIndexSRA,@cFiltraRH) }
	Eval(bFiltraBrw)

	/*
	��������������������������������������������������������������Ŀ
	� Endereca a funcao de BROWSE                                  �
	����������������������������������������������������������������*/
	dbSelectArea("SRA")
	mBrowse( 6, 1,22,75,"SRA",,,,,,fCriaCor() )

	/*
	������������������������������������������������������������������������Ŀ
	� Deleta o filtro utilizando a funcao FilBrowse                     	 �
	��������������������������������������������������������������������������*/
	EndFilBrw("SRA",aIndexSra)

EndIF

RestArea( aAreaSPI )
RestArea( aArea )

Return( NIL )


Static Function MenuDef()

Local aRotina		:= {	{ "Pesquisar" 	,"PesqBrw"	 	, 0 , 1, ,.F.},; 	
		 		            { "Visualizar" 	,"Pona200Atu" 	, 0 , 2},; 		
                     		{ "Incluir" 	,"Pona200Atu" 	, 0 , 3,,,.T.},; 	
                     		{ "Alterar" 	,"Pona200Atu" 	, 0 , 4},; 		
                     		{ "Excluir" 	,"Pona200Atu" 	, 0 , 5},; 	
							{ "Importar" 	,"U_FSSPIArq" 	, 0 , 6},; 	 	
                     		{ 'Legenda' 	,"gpLegend"	 	, 0 , 7, ,.F.}} 	

Return aRotina



User Function FSSPIArq()
    Local aArea     := GetArea()
    Private cArqOri := ""
 
    //Mostra o Prompt para selecionar arquivos
    cArqOri := tFileDialog( "CSV files (*.csv) ", 'Sele��o de Arquivos', , , .F., )
     
    //Se tiver o arquivo de origem
    If ! Empty(cArqOri)
         
        //Somente se existir o arquivo e for com a extens�o CSV
        If File(cArqOri) .And. Upper(SubStr(cArqOri, RAt('.', cArqOri) + 1, 3)) == 'CSV'
            Processa({|| fImporta() }, "Importando...")
        Else
            MsgStop("Arquivo e/ou extens�o inv�lida!", "Aten��o")
        EndIf
    EndIf
     
    RestArea(aArea)
Return


/*-------------------------------------------------------------------------------*
 | Func:  fImporta                                                               |
 | Desc:  Fun��o que importa os dados                                            |
 *-------------------------------------------------------------------------------*/
 Static Function fImporta()
    Local aArea      := GetArea()
    Local cArqLog    := "zImpCSV_" + dToS(Date()) + "_" + StrTran(Time(), ':', '-') + ".log"
    Local nTotLinhas := 0
    Local cLinAtu    := ""
    Local nLinhaAtu  := 0
    Local aLinha     := {}
    Local oArquivo
    Local aLinhas
	Local cFilBKP    := cFilAnt
    Private cDirLog    := GetTempPath() + "x_importacao\"
    Private cLog       := ""
     
    //Se a pasta de log n�o existir, cria ela
    If ! ExistDir(cDirLog)
        MakeDir(cDirLog)
    EndIf
 
    //Definindo o arquivo a ser lido
    oArquivo := FWFileReader():New(cArqOri)
     
    //Se o arquivo pode ser aberto
    If (oArquivo:Open())
 
        //Se n�o for fim do arquivo
        If ! (oArquivo:EoF())
 
            //Definindo o tamanho da r�gua
            aLinhas := oArquivo:GetAllLines()
            nTotLinhas := Len(aLinhas)
            ProcRegua(nTotLinhas)
             
            //M�todo GoTop n�o funciona (dependendo da vers�o da LIB), deve fechar e abrir novamente o arquivo
            oArquivo:Close()
            oArquivo := FWFileReader():New(cArqOri)
            oArquivo:Open()

            //Enquanto tiver linhas
            While (oArquivo:HasLine())
 
                //Incrementa na tela a mensagem
                nLinhaAtu++
                IncProc("Analisando linha " + cValToChar(nLinhaAtu) + " de " + cValToChar(nTotLinhas) + "...")
                 
                //Pegando a linha atual e transformando em array
                cLinAtu := oArquivo:GetLine()
                aLinha  := StrTokArr(cLinAtu, ";")

				if len(aLinha) <> 5
					 cLog += "Layout do arquivo invalido;" + CRLF
					Exit
				Endif

                //Se n�o for o cabe�alho (encontrar o texto "C�digo" na linha atual)
                If !("filial" $ Lower(cLinAtu) .or. "pi_filial" $ Lower(cLinAtu))

                    //Zera as variaveis
					cCodFil	    := AVKey(Alltrim(aLinha[nPosFilial]),"PI_FILIAL")
					cCodMatr    := AVKey(Alltrim(aLinha[nPosMatr]),"PI_MAT")
                    dData	   	:= Ctod(aLinha[nPosData])
					cCodEvento  := AVKey(Alltrim(aLinha[nPosEvento]),"PI_PD")
					nQtdLanc    := aLinha[nPosLanc]
					nQtdLanc	:= Val(STRTRAN(nQtdLanc,',','.'))

					cFilAnt	   := cCodFil

                    DbSelectArea('SPI')
                    SPI->(DbSetOrder(1)) // Filial + C�digo + Loja
 
                     DbSelectArea('SRA')
                    SRA->(DbSetOrder(1)) // Filial + C�digo + Loja
 
					If !SRA->(DBSeek(xFilial("SRA")+cCodMatr))
                        cLog += "+ Lin" + cValToChar(nLinhaAtu) + " MAtricula [" + cCodMatr + "] nao foi localizado;" + CRLF
						Loop
					Endif	
					
					If ExistSP9(cCodEvento) <= 0
                        cLog += "+ Lin" + cValToChar(nLinhaAtu) + " Evento [" + cCodEvento + "] nao foi localizado;" + CRLF
						Loop
					Endif	

                    If SPI->(DbSeek(FWxFilial('SPI') + cCodMatr + cCodEvento+DtoS(dData)))
                        cLog += "+ Lin" + cValToChar(nLinhaAtu) + " Filial " + xFilial('SPI') + " Matricula "+SRA->RA_MAT+" Cod Verba "+cCodEvento+" ja informada ;" + CRLF
                    Else
						RecLock("SPI", .T.)
						
						SPI->PI_FILIAL  	:= xFilial('SPI')
						SPI->PI_MAT	  		:= SRA->RA_MAT
						SPI->PI_DATA		:= dData
						SPI->PI_PD			:= cCodEvento
						SPI->PI_CC	  		:= SRA->RA_CC
						SPI->PI_QUANTV  	:= nQtdLanc
						SPI->PI_QUANT   	:= nQtdLanc
						SPI->PI_DEPTO   	:= SRA->RA_DEPTO
						SPI->PI_CODFUNC 	:= SRA->RA_CODFUNC
						SPI->PI_FLAG    	:= "I"
						SPI->PI_STATUS  	:= ""
						SPI->PI_DTBAIX  	:= CtoD("")						
						SPI->(MSUnLock())

                        cLog += "- Lin" + cValToChar(nLinhaAtu) + ", Cadastrado Evento [" + cCodEvento + "]  para Matricula [" + SRA->RA_MAT + "] ;" + CRLF
                    EndIf
                     
                Else
                    cLog += "- Lin" + cValToChar(nLinhaAtu) + ", linha n�o processada - cabe�alho;" + CRLF
                EndIf
                 
            EndDo
 
            //Se tiver log, mostra ele
            If ! Empty(cLog)
                cLog := "Processamento finalizado, abaixo as mensagens de log: " + CRLF + cLog
                MemoWrite(cDirLog + cArqLog, cLog)
                ShellExecute("OPEN", cArqLog, "", cDirLog, 1)
            EndIf
 
        Else
            MsgStop("Arquivo n�o tem conte�do!", "Aten��o")
        EndIf
 
        //Fecha o arquivo
        oArquivo:Close()
    Else
        MsgStop("Arquivo n�o pode ser aberto!", "Aten��o")
    EndIf
  	
	  cFilAnt := cFilBKP 

    RestArea(aArea)
Return

/*/{Protheus.doc} ExistSP9
Fun��o copia do browser PONA200 para incluir importacao de arquivo
@author Alex Teixeira
@since 08/02/2023
@version 1.0
@type function
/*/
Static Function ExistSP9(cCodEveAux)
	Local aArea		:= GetArea()
	Local cAliasQry := GetNextAlias()
	Local nInd		:= 0
	LOCAL cQuery    := ""

	cQuery := " SELECT P9_CODIGO "
	cQuery += " FROM " + RetSqlName("SP9")
	cQuery += " WHERE P9_CODIGO = " +"'"+cCodEveAux+"'"
	cQuery += " AND D_E_L_E_T_<>'*' AND P9_FILIAL = '"+ xFilial('SP9') +"'"
	cQuery += " ORDER BY  P9_FILIAL,P9_CODFOL"	

	cQuery := ChangeQuery(cQuery)
	dbUseArea( .T., "TOPCONN", TCGENQRY(,,cQuery),cAliasQry, .F., .F.)
	dbSelectArea(cAliasQry)
	(cAliasQry)->( DbGoTop() )
	
	While !(cAliasQry)->( Eof() )
		nInd := nInd +1
		(cAliasQry)->( DbSkip() )
	EndDo

	(cAliasQry)->( DbCloseArea() )
	RestArea(aArea)
Return (nInd)
