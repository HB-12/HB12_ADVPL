#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "TOTVS.CH"

/*/{Protheus.doc} FSGPEJ01
Rotina responsável por enviar e-mail com as matrículas com vencimento de aviso prévio próximas.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
User Function FSGPEJ01()

	Local aMat      := {}
    Local aMail     := {}

	Local cHTMLMsg  := ""
    Local cMailRH   := ""

	// Prepara o Ambiente
	RpcSetType(3)
	RpcSetEnv("02", "0101")

	// Carrega matrículas que deverão aparecer na lista de Avisos que irão vencer
	aMat        := FGetMatAvi()

	// Monta HTML da mensagem que será enviada
	cHTMLMsg    := FMontaHTML(aMat)

    // Envia e-mail com as matrículas carregadas
    cMailRH     := SuperGetMV("FS_AVIEMRH",,"henrique.bruno@grupoemtel.com.br")
    aMail       := U_FSSendMail(cMailRH, "Vencimento de Aviso Prévio", cHTMLMsg)

    If !(aMail[1])
        FWLogMsg("WARN", "LAST", "FSGPEJ01", , , , aMail[2])
    EndIf

	// Encerra o Ambiente
	RpcClearEnv()

Return(Nil)

/*/{Protheus.doc} FGetMatAvi
Função responsável por carregar as matrículas que estão com vencimento de aviso prévio.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FGetMatAvi()

	Local nDiaLimNot    := SuperGetMV("FS_AVIDLIM",,7)

	Local aRet          := {}

	Local cToday        := DtoS(Date())
	Local cAliMat       := GetNextAlias()

	BeginSql Alias cAliMat

        Column RA_VCTOEXP as Date
        Column RA_VCTEXP2 as Date

        SELECT  RA_FILIAL,
                RA_MAT,
                RA_NOME,
                RA_VCTOEXP,
                RA_VCTEXP2,
                RA_ZCOMAV1,
                RA_ZCOMAV2
        FROM %Table:SRA% SRA
        WHERE ((RA_VCTOEXP >= %Exp:cToday% AND RA_ZCOMAV1 = '') OR (RA_VCTEXP2 >= %Exp:cToday% AND RA_ZCOMAV2 = ''))
        AND RA_DEMISSA = ''
        AND SRA.%NotDel%
        ORDER BY RA_FILIAL, RA_MAT, RA_NOME

	EndSql

	(cAliMat)->(DbGoTop())

	While !((cAliMat)->(EOF()))

		If Empty((cAliMat)->RA_ZCOMAV1) .AND. (DateDiffDay(Date(), (cAliMat)->RA_VCTOEXP) <= nDiaLimNot)
			aAdd(aRet, {(cAliMat)->RA_FILIAL + " - " + AllTrim(FwFilialName( "02", (cAliMat)->RA_FILIAL, 1 )),;
                        (cAliMat)->RA_MAT + " - " + AllTrim((cAliMat)->RA_NOME),;                        
                        DtoC((cAliMat)->RA_VCTOEXP),;
                        cValToChar(DateDiffDay(Date(), (cAliMat)->RA_VCTOEXP))})
		ElseIf Empty((cAliMat)->RA_ZCOMAV2) .AND. (DateDiffDay(Date(), (cAliMat)->RA_VCTEXP2) <= nDiaLimNot)
			aAdd(aRet, {(cAliMat)->RA_FILIAL + " - " + AllTrim(FwFilialName( "02", (cAliMat)->RA_FILIAL, 1 )),;
                        (cAliMat)->RA_MAT + " - " + AllTrim((cAliMat)->RA_NOME),;
                        DtoC((cAliMat)->RA_VCTEXP2),;
                        cValToChar(DateDiffDay(Date(), (cAliMat)->RA_VCTEXP2))})
		EndIf

		(cAliMat)->(DbSkip())
	EndDo

Return(aRet)

/*/{Protheus.doc} FMontaHTML
Função responsável por montar o HTML com os dados coletados.
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
Static Function FMontaHTML(aMat)

    Local cRet      := ""

    Local nM        := 0

    cRet += "<html>"
    cRet += "   <head>"
    cRet += "      <style>     body {"
    cRet += "							 font-family: Arial;"
    cRet += "							 font-size:11pt;"
    cRet += "						}"
    cRet += "						 .cabec{"
    cRet += "							 text-align: center;"
    cRet += "							 color: white;"
    cRet += "							 background-color: DarkBlue;"
    cRet += "							 font-size: 30px;"
    cRet += "							 margin: 0px;"
    cRet += "						}"
    cRet += "						 td{"
    cRet += "							 color: rgb(0, 80, 159);"
    cRet += "							 text-align: left;"
    cRet += "							 background-color: rgb(239, 239, 239);"
    cRet += "							 margin: 0px;"
    cRet += "						}"
    cRet += "						 .conteudoag{"
    cRet += "							 color: #000;"
    cRet += "							 font-size: 14px;"
    cRet += "						}"
    cRet += "						 .conteudose{"
    cRet += "							 color: Blue;"
    cRet += "						}"
    cRet += "						 .conteudosc{"
    cRet += "							 color: Green;"
    cRet += "						}"
    cRet += "						 .conteudonf{"
    cRet += "							 color: Red;"
    cRet += "						}"
    cRet += "						 .tdcab{"
    cRet += "							 font-weight: bold;"
    cRet += "							 text-align: center;"
    cRet += "							 height:20px;"
    cRet += "							 background-color: Silver;"
    cRet += "							 font-size: 16px;"
    cRet += "							 margin: 0px;"
    cRet += "						}"
    cRet += "						 .salto{"
    cRet += "							 font-size: 0px;"
    cRet += "							 margin: 0px;"
    cRet += "							 color:rgb(239, 239, 239);"
    cRet += "						}"
    cRet += "    </style>"
    cRet += "   </head>"
    cRet += "   <body>"
    cRet += "      <p>Prezado(a),"
    cRet += "	  </br>"
    cRet += "	  </br>"
    cRet += "	  <p>Segue abaixo as matrículas que estão próximas do vencimento do aviso prévio."
    cRet += "	  </br>"
    cRet += "	  </br>"
    cRet += "      <p>    "
    cRet += "      <table style='width:100%' align='center'>"
    cRet += "         <tr>"
    cRet += "            <td class='tdcab'>       		FILIAL      	</td>"
    cRet += "            <td class='tdcab'>  	 		FUNCIONÁRIO      	</td>"
    cRet += "            <td class='tdcab'>       		VENC. AVISO      	</td>"
    cRet += "            <td class='tdcab'>       		DIAS P/ VENC.      	</td>"
    cRet += "         </tr>"

    For nM  := 1 to Len(aMat)

        cRet += "<tr>"
        cRet += "   <td>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "      <div class='conteudoag'>" + aMat[nM][1] + "</div>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "   </td>"
        cRet += "   <td>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "      <div class='conteudoag'>" + aMat[nM][2] + "</div>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "   </td>"
        cRet += "   <td>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "      <div class='conteudoag' align='center'>" + aMat[nM][3] + "</div>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "   </td>"
        cRet += "   <td>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "      <div class='conteudoag' align='center'>" + aMat[nM][4] + "</div>"
        cRet += "      <p class='salto'>-</p>"
        cRet += "   </td>"
        cRet += "</tr>"

    Next nM

    cRet += "</table>"
    cRet += "      <p>--</p>"
    cRet += "   </body>"
    cRet += "</html>"

Return(cRet)
