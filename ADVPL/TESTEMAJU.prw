USER FUNCTION EnviarEmailNotificacao()

   LOCAL cDestinatario := "henrique.bruno@grupoemtel.com.br"  // Endereço de email do destinatário
   LOCAL cAssunto := "Notificação de data atingida" // Assunto do email
   LOCAL cMensagem := "A data limite foi atingida!" // Corpo da mensagem

   // Obtém a data atual
   LOCAL dHoje := Cronos()

   // Obtém a data limite do campo
   LOCAL dDataLimite := SRA020->RA_VCTOEXP

   // Compara as datas
   IF dHoje >= dDataLimite

      // Envia o email de notificação
      MSendMail(cDestinatario, cAssunto, cMensagem)

   ENDIF

RETURN
