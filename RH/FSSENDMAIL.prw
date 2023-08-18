#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "TOTVS.CH"

/*/{Protheus.doc} FSSendMail
Rotina respons�vel por enviar e-mails.
@type function
@author gustavo.barcelos
@since 3/6/2023
@param cMailDest, character, Endere�o de e-mail do destinat�rio
@param cTitMail, character, T�tulo do e-mail
@param cMailMsg, character, Mensagem a ser enviada
@param cMailCc, character, Endere�o de e-mail que receber� a c�pia
@param cMailCCo, character, Endere�o de e-mail que receber� a c�pia (oculta)
/*/
User Function FSSendMail(cMailDest, cTitMail, cMailMsg, cMailCc, cMailCCo)

    Local aRet          := {}

    Local lRet          := .T.

    Local cRet          := ""
    Local cSendSrv      := ""
    Local cFromEmail    := ""
    Local cPassFron     := ""

    Local nServer       := 0
    Local nSendPort     := 0
    Local nTimeout      := 0

    Local oServer       := Nil

    Default cMailCc     := ""
    Default cMailCCo    := ""

    oServer := TMailManager():New()

    cSendSrv   := Left(SuperGetMV('MV_RELSERV'), RAt( ':', SuperGetMV('MV_RELSERV'))-1)
    cFromEmail := SuperGetMV('MV_RELACNT')
    cPassFron  := SuperGetMV('MV_RELPSW')
    nSendPort  := Val(Right(SuperGetMV('MV_RELSERV', .F.),;
                        Len(SuperGetMV('MV_RELSERV'))-RAt( ':', SuperGetMV('MV_RELSERV'))))
    nTimeout  := SuperGetMV('MV_RELTIME')

    oServer := TMailManager():New()

    oServer:SetUseSSL(SuperGetMV('MV_RELSSL'))
    oServer:SetUseTLS(SuperGetMV('MV_RELTLS'))

    //Cria a conex�o com o server STMP ( Envio de e-mail )
    nServer := oServer:Init( "", cSendSrv, cFromEmail, cPassFron, , nSendPort )
    If nServer <> 0
        cRet := "Nao foi possivel inicializar o servidor SMTP: " + oServer:GetErrorString( nServer )
        lRet := .F.
    EndIf

    // O m�todo define o tempo limite para o servidor SMTP
    nServer := oServer:SetSMTPTimeout( nTimeout )
    If nServer <> 0
        cRet := "Nao foi possivel definir o tempo limite de " + cValToChar( nTimeout ) + " segundos."
        lRet := .F.
    EndIf
   
    // Estabelecer a conex�o com o servidor SMTP
    nServer := oServer:SMTPConnect()
    If nServer <> 0
        cRet := "Nao foi possivel inicializar o servidor SMTP: " + oServer:GetErrorString( nServer )
        lRet := .F.
    EndIf
   
    // Autenticar no servidor SMTP (se necess�rio)
    nServer := oServer:SmtpAuth( cFromEmail, cPassFron )
    If nServer <> 0
        cRet := "Nao foi possivel inicializar o servidor SMTP: " + oServer:GetErrorString( nServer )
        lRet := .F.
        oServer:SMTPDisconnect()
    EndIf

    If lRet
        
        If Empty(cMailDest)
            cRet := "Endere�o de e-mail de destino inv�lido ou n�o preenchido."
            lRet := .F.
        Else
            
            oMessage := TMailMessage():New()
            oMessage:Clear()

            oMessage:cDate     := cValToChar( Date() )
            oMessage:cFrom     := cFromEmail
            oMessage:cTo       := cMailDest
            oMessage:cCC       := cMailCc
            oMessage:cBCC      := cMailCCo
            oMessage:cSubject  := cTitMail
            oMessage:cBody     := cMailMsg

            nServer := oMessage:Send( oServer )

            If nServer <> 0
                cRet := "Nao foi possivel enviar mensagem: " + oServer:GetErrorString( nServer )
                lRet := .F.
            Else
                cRet := "Mensagem enviada para " + AllTrim(cMailDest) + " com sucesso!"
                lRet := .T.
            EndIf

        EndIf

        oServer:SMTPDisconnect()

    EndIf

    aRet := {lRet, cRet}

Return(aRet)
