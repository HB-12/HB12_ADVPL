#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "TOTVS.CH"

/*/{Protheus.doc} FSGPEP01
Rotina respons�vel por buscar dados do Funcion�rio e enviar e-mail de Comunicado de Vencimento de Aviso Pr�vio.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
User Function FSGPEP01()

    Local aMail         := {}

    Local cMailDest     := ""
    Local cMailRH       := SuperGetMV("FS_AVIEMRH",,"")
    Local cHTMLMsg      := ""
    Local cCmpAtu       := ""

    Local dVencExp45    := SRA->RA_VCTOEXP
    Local dVencExp90    := SRA->RA_VCTEXP2
    Local dDatAvis      := StoD("")
    Local dToday        := Date()

    If  (dVencExp45 >= dToday)
        dDatAvis := dVencExp45
        cCmpAtu  := "RA_ZCOMAV1"
    ElseIf (dVencExp90 >= dToday)
        dDatAvis := dVencExp90
        cCmpAtu  := "RA_ZCOMAV2"
    EndIf

    If !Empty(dDatAvis)

        cMailDest := FGetMail()
        cHTMLMsg  := FGetHMsg(SRA->RA_MAT, SRA->RA_NOME, dDatAvis)

        If FWAlertYesNo("Deseja enviar o Comunicado de Vencimento de Aviso Pr�vio para " + AllTrim(cMailDest) + "?","Envia Comunicado de Vencimento de Aviso Pr�vio?")
            MsgRun( "Enviando e-mail...", , { || aMail := U_FSSendMail(cMailDest, "Comunicado de Vencimento de Aviso Pr�vio", cHTMLMsg, cMailRH)} )

            If aMail[1]
                FAtuComAvi(cCmpAtu)
                FWAlertSuccess(aMail[2], "E-mail enviado para " + AllTrim(cMailDest) + "!")
            Else
                FWAlertError(aMail[2], "Houveram inconsist�ncias e o e-mail n�o foi enviado.")
            EndIf

        EndIf

    Else
        
        FWAlertWarning( "O e-mail de Comunicado de Vencimento de Aviso Pr�vio s� pode ser enviado para funcion�rios que est�o em fase de experi�ncia.",;
                        "Este funcion�rio n�o est� em per�odo de experi�ncia.")

    EndIf

Return(Nil)

/*/{Protheus.doc} FGetMail
Fun��o respons�vel por apresentar tela para usu�rio informar o endere�o de e-mail que ir� receber o comunicado de Aviso Pr�vio.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FGetMail()

    Local nTop          := 180
    Local nLeft         := 180
    Local nBottom       := 310
    Local nRight        := 670

    Local oDialog	    := Nil
    Local oGetMail      := Nil

    Local cEndMail      := Space(TamSX3("RA_EMAIL")[1])

    oDialog     := TDialog():New(nTop, nLeft, nBottom, nRight,'Comunicado de Vencimento de Aviso Pr�vio',,,,DS_MODALFRAME,,,,,.T.)

    oGetMail    := TGet():New( 1 /* nRow */, 1 /* nCol */, {|u| If(PCount()>0,cEndMail:=u,cEndMail)}/* bSetGet */, oDialog /* oWnd */, 175 /* nWidth */, 13 /* nHeight */, /* cPict */, /* bValid */, 0/* nClrFore */,;
                        /* nClrBack */, /* oFont */, /* uParam12 */, /* uParam13 */, /* lPixel */, /* uParam15 */, /* uParam16 */, /* bWhen */,;
                        /* uParam18 */, /* uParam19 */, /* bChange */, .F. /* lReadOnly */, /* lPassword */, /* uParam23 */, "cEndMail"/* cReadVar */, /* uParam25 */,;
                        /* uParam26 */, /* uParam27 */, .T./* lHasButton */, .F. /* lNoButton */, /* uParam30 */, "E-mail do Destinat�rio "/* cLabelText */, 2 /* nLabelPos */,;
                        /* oLabelFont */, /* nLabelColor */, /* cPlaceHold */, /* lPicturePriority */, /* lFocSel */ )

    TButton():New( 040, 025, "Enviar E-mail",oDialog,{|| oDialog:End() }, 200,12,,,.F.,.T.,.F.,,.F.,,,.F. )

    oDialog:lEscClose := .F.

    oDialog:Activate( ,,,.T.,,, )

Return(cEndMail)

/*/{Protheus.doc} FGetHMsg
Fun��o respons�vel pela montagem do c�digo HTML com a mensagem que ser� enviada.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FGetHMsg(cMatFun, cNomFun, dDatAvis)

    Local cRet      := ""

    cRet += " <p>Prezado(a),</br> "
	cRet += " </br> "
	cRet += " <p>O aviso pr�vio do funcion�rio " + cMatFun + " - " + AllTrim(cNomFun) + " vence no dia " + DtoC(dDatAvis) + " e necessita de sua avalia��o. "
	cRet += " </br> "
	cRet += " <p>Entre em contato com o setor de RH para maiores informa��es. "
	cRet += " </br> "
	cRet += " <p>Obrigado! "
	cRet += " <p> "

Return(cRet)

/*/{Protheus.doc} FAtuComAvi
Fun��o respons�vel por atualizar o campo de Data de Comunicado de Aviso.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FAtuComAvi(cCmpAtu)

    RecLock("SRA", .F.)
    &("SRA->" + cCmpAtu + " := Date()")
    SRA->(MsUnlock())

Return(Nil)
