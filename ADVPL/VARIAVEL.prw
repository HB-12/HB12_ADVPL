#INCLUDE 'protheus.ch'
#INCLUDE 'parmtype.ch'

/**
    Tipo de Dados

NUMERICO: 3 / 21.000 / 0.4 / 20000
LOGICO: .T. / .F.
CARACTERE: "D" / 'C'
DATA: DATE()
ARRAY: {"VALOR1", "VALOR2", "valor3"}
BLOCO DE CODIGO: {|| VALOR := 1, MsgAlert("VALOR É IGUAL A "+cVALtOcHAR(VALOR))}

**/

user function VARIAVEL()
    Local nNum := "12"
    Local lLogic := .T.
    Local cCarac := "String"
    Local dData := DATE()
    Local aArray := {"Henrique", "Gabriel", "Gian"}
    Local bBloco := {|| nValor := 2, MsgAlert("O numero é: "+ CValToChar(nValor))}

    Alert(nNum)
    Alert(lLOgic)
    Alert(CValToChar(cCarac))
    Alert(dData)
    Alert(aArray[1])
    Eval(bBloco)



RETURN
