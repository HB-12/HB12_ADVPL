#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "TOTVS.CH"

/*/{Protheus.doc} FSGPEP01
Rotina responsável por buscar dados do Funcionário e enviar e-mail de Comunicado de Vencimento de Aviso Prévio.
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
    Local lVenc45       := .f.
    Local lVenc90       := .f.

    If  (dVencExp45 >= dToday)
        dDatAvis := dVencExp45
        cCmpAtu  := "RA_ZCOMAV1"
        lVenc45  := .t.
    ElseIf (dVencExp90 >= dToday)
        dDatAvis := dVencExp90
        cCmpAtu  := "RA_ZCOMAV2"
        lVenc90  := .t.
    EndIf

    If !Empty(dDatAvis)

        cMailDest := FGetMail()
        cHTMLMsg  := FGetHMsg(lVenc45,lVenc90,SRA->RA_MAT,dDatAvis)

        If !Empty(cHTMLMsg)

            If FWAlertYesNo("Deseja enviar o Comunicado de Vencimento de Aviso Prévio para " + AllTrim(cMailDest) + "?","Envia Comunicado de Vencimento de Aviso Prévio?")
                MsgRun( "Enviando e-mail...", , { || aMail := U_FSSendMail(cMailDest, "Comunicado de Vencimento de Aviso Prévio", cHTMLMsg, cMailRH)} )

                If aMail[1]
                    FAtuComAvi(cCmpAtu)
                    FWAlertSuccess(aMail[2], "E-mail enviado para " + AllTrim(cMailDest) + "!")
                Else
                    FWAlertError(aMail[2], "Houveram inconsistências e o e-mail não foi enviado.")
                EndIf

            EndIf
        Else
            FWAlertWarning( "É necessário localizar o fomrmulario de e-mail configurado no sistema.",;
                            "Entre em contato com o Depto. de TI.")
        Endif                        

    Else
        
        FWAlertWarning( "O e-mail de Comunicado de Vencimento de Aviso Prévio só pode ser enviado para funcionários que estão em fase de experiência.",;
                        "Este funcionário não está em período de experiência.")

    EndIf

Return(Nil)

/*/{Protheus.doc} FGetMail
Função responsável por apresentar tela para usuário informar o endereço de e-mail que irá receber o comunicado de Aviso Prévio.
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

    oDialog     := TDialog():New(nTop, nLeft, nBottom, nRight,'Comunicado de Vencimento de Aviso Prévio',,,,DS_MODALFRAME,,,,,.T.)

    oGetMail    := TGet():New( 1 /* nRow */, 1 /* nCol */, {|u| If(PCount()>0,cEndMail:=u,cEndMail)}/* bSetGet */, oDialog /* oWnd */, 175 /* nWidth */, 13 /* nHeight */, /* cPict */, /* bValid */, 0/* nClrFore */,;
                        /* nClrBack */, /* oFont */, /* uParam12 */, /* uParam13 */, /* lPixel */, /* uParam15 */, /* uParam16 */, /* bWhen */,;
                        /* uParam18 */, /* uParam19 */, /* bChange */, .F. /* lReadOnly */, /* lPassword */, /* uParam23 */, "cEndMail"/* cReadVar */, /* uParam25 */,;
                        /* uParam26 */, /* uParam27 */, .T./* lHasButton */, .F. /* lNoButton */, /* uParam30 */, "E-mail do Destinatário "/* cLabelText */, 2 /* nLabelPos */,;
                        /* oLabelFont */, /* nLabelColor */, /* cPlaceHold */, /* lPicturePriority */, /* lFocSel */ )

    TButton():New( 040, 025, "Enviar E-mail",oDialog,{|| oDialog:End() }, 200,12,,,.F.,.T.,.F.,,.F.,,,.F. )

    oDialog:lEscClose := .F.

    oDialog:Activate( ,,,.T.,,, )

Return(cEndMail)

/*/{Protheus.doc} FGetHMsg
Função responsável pela montagem do código HTML com a mensagem que será enviada.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FGetHMsg(lVenc45,lVenc90,cMatFun,dDatAvis)
Local aArea         := GetArea()
Local cArqHTML	    := ""
Local cHTML         := ""
Local nQtdDias      := SuperGetMV("FS_PRZRETRH", .F., 3)
Local dDtPrazo      := dDataBase+nQtdDias
Local cNomeFun      := ""
Local cNomGestor    := ""

    If lVenc45
        cArqHTML := GetSrvProfString( "StartPath", "" )+"\email_45dias.html"
    Endif

    If lVenc90
        cArqHTML := GetSrvProfString( "StartPath", "" )+"\email_90dias.html"
    Endif

    If File(cArqHTML)

        Do While DateWorkDay(dDataBase,dDtPrazo,.t.,.t. ,.t. ) < nQtdDias
            dDtPrazo++
        EndDo
        

        dbSelectArea("SRA")
        dbSetOrder(1)

        dbSelectArea("CTT")
        dbSetOrder(1)

        If SRA->(DBSeek(xFilial("SRA")+cMatFun))

            cNomeFun := SRA->RA_NOME

            If CTT->(DBSeek(xFilial("CTT")+SRA->RA_CC))
                cNomGestor := CTT->CTT_ZGESTO
            Endif
            
            cHtml := MemoRead( cArqHTML )
            //cHtml := StrTRan(cHtml,"%cGestor%",Alltrim(cNomGestor))   
            cHtml := StrTRan(cHtml,"%cFuncionario%",Alltrim(cNomeFun))   
            cHtml := StrTRan(cHtml,"%cPrazo%",DtoC(dDtPrazo))   
            cHtml := StrTRan(cHtml,"%cData%",DtoC(dDatAvis))   

        Endif

    Endif   

    RestArea(aArea) 

Return(cHtml)

/*/{Protheus.doc} FAtuComAvi
Função responsável por atualizar o campo de Data de Comunicado de Aviso.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FAtuComAvi(cCmpAtu)

    RecLock("SRA", .F.)
    &("SRA->" + cCmpAtu + " := Date()")
    SRA->(MsUnlock())

Return(Nil)
