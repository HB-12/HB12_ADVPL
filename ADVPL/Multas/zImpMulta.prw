//Bibliotecas
#Include "TOTVS.ch"
#Include "TopConn.ch"
 
//Posições do Array
Static nPosMulta 	:= 1 
Static nPosDtEnv 	:= 30 
Static nPosTpFat   	:= 32 
Static nPosNumFat 	:= 33
Static nPosDtRec 	:= 35
Static nPosFilial 	:= 42



/*/{Protheus.doc} zImpMulta
Função para importar informações da planilha de multa
@author Alex Teixeira
@since 11/04/2023
@version 1.0
@type function
/*/
 
User Function zImpMulta()
    Local aArea     := GetArea()
    Private cArqOri := ""
 
    //Mostra o Prompt para selecionar arquivos
    cArqOri := tFileDialog( "CSV files (*.csv) ", 'Seleção de Arquivos', , , .F., )
     
    //Se tiver o arquivo de origem
    If ! Empty(cArqOri)
         
        //Somente se existir o arquivo e for com a extensão CSV
        If File(cArqOri) .And. Upper(SubStr(cArqOri, RAt('.', cArqOri) + 1, 3)) == 'CSV'
            Processa({|| fImporta() }, "Importando...")
        Else
            MsgStop("Arquivo e/ou extensão inválida!", "Atenção")
        EndIf
    EndIf
     
    RestArea(aArea)
Return
 
/*-------------------------------------------------------------------------------*
 | Func:  fImporta                                                               |
 | Desc:  Função que importa os dados                                            |
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
    Local aTpFat     := Separa("FATURA;FOLHA;EMPRESA;RECUSADO;OUTROS;RESCISAO;DIARIAS;SERVICOS;RESTITUICAO;DINHEIRO;DEPOSITO",";",.t.)           

    Private cDirLog    := GetTempPath() + "x_importacao\"
    Private cLog       := ""
     
    //Se a pasta de log não existir, cria ela
    If ! ExistDir(cDirLog)
        MakeDir(cDirLog)
    EndIf
 
    //Definindo o arquivo a ser lido
    oArquivo := FWFileReader():New(cArqOri)
     
    //Se o arquivo pode ser aberto
    If (oArquivo:Open())
 
        //Se não for fim do arquivo
        If ! (oArquivo:EoF())
 
            //Definindo o tamanho da régua
            aLinhas := oArquivo:GetAllLines()
            nTotLinhas := Len(aLinhas)
            ProcRegua(nTotLinhas)
             
            //Método GoTop não funciona (dependendo da versão da LIB), deve fechar e abrir novamente o arquivo
            oArquivo:Close()
            oArquivo := FWFileReader():New(cArqOri)
            oArquivo:Open()

			cCodProd := ""
 
            //Enquanto tiver linhas
            While (oArquivo:HasLine())
 
                //Incrementa na tela a mensagem
                nLinhaAtu++
                IncProc("Analisando linha " + cValToChar(nLinhaAtu) + " de " + cValToChar(nTotLinhas) + "...")
                 
                //Pegando a linha atual e transformando em array
                cLinAtu := oArquivo:GetLine()
                aLinha  := Separa(cLinAtu, ";",.t.)
                 
                //Se não for o cabeçalho (encontrar o texto "Código" na linha atual)
                    If !("filial" $ Lower(cLinAtu) .or. "b1_filial" $ Lower(cLinAtu))

                    //Zera as variaveis
					cCodFil	    := Alltrim(aLinha[nPosFilial])
                    If len(cCodFil) <= 3
                        cCodFil := "0"+cCodFil
                    Endif
                    cCodFil     := AVKey(Alltrim(cCodFil),"TRX_FILIAL")    
                    dDtEnv      := Ctod(AVKey(Alltrim(aLinha[nPosDtEnv]),"TRX_ZDTREC"))
                    cMulta      := AVKey(Alltrim(aLinha[nPosMulta]),"TRX_MULTA")
                    cTPFat      := Alltrim(aLinha[nPosTpFat])

                    If  (nPos := ascan(aTPFat,{|x| upper(x) == upper(cTPFat)})) > 0
                        cTPFat := Alltrim(str(nPos))
                    Else    
                        cTPFat := ""
                    Endif    

                    cNumFat     := AVKey(Alltrim(aLinha[nPosNumFat]),"TRX_ZFLFIN")
                    dDtRec      := Ctod(AVKey(Alltrim(aLinha[nPosDtRec]),"TRX_ZDTREC"))		
					
					cFilAnt	   := cCodFil

                    DbSelectArea('TRX')
                    TRX->(DbSetOrder(1)) // Filial + Código + Loja

                    //Se conseguir posicionar no fornecedor
                    If TRX->(DbSeek(FWxFilial('TRX') + cMulta))
                        cLog += "+ Lin" + cValToChar(nLinhaAtu) + ", Multa [" + cMulta + "] Localizada ;" + CRLF
 
                        RecLock('TRX', .F.)
                        TRX->TRX_ZFATUR  := cTPFat
                        TRX->TRX_ZFLFIN  := cCodFil
                        TRX->TRX_ZDTREC  := dDtRec
                        TRX->TRX_ZDFATU  := dDtEnv
                        TRX->TRX_ZFTFIN  := cNumFat 
                        TRX->(MsUnlock())
 
                    Else
                       cLog += "- Lin" + cValToChar(nLinhaAtu) + ", Multa [" + cMulta + "] Não foi localizada ;" + CRLF
                    EndIf
                     
                Else
                    cLog += "- Lin" + cValToChar(nLinhaAtu) + ", linha não processada - cabeçalho;" + CRLF
                EndIf
                 
            EndDo
 
            //Se tiver log, mostra ele
            If ! Empty(cLog)
                cLog := "Processamento finalizado, abaixo as mensagens de log: " + CRLF + cLog
                MemoWrite(cDirLog + cArqLog, cLog)
                ShellExecute("OPEN", cArqLog, "", cDirLog, 1)
            EndIf
 
        Else
            MsgStop("Arquivo não tem conteúdo!", "Atenção")
        EndIf
 
        //Fecha o arquivo
        oArquivo:Close()
    Else
        MsgStop("Arquivo não pode ser aberto!", "Atenção")
    EndIf
  	
	  cFilAnt := cFilBKP 

    RestArea(aArea)
Return
