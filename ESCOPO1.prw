#INCLUDE 'protheus.ch'
#INCLUDE 'parmtype.ch'

Static cStat :=''

user function ESCOPO1 ()
// Variaveis Locais
Local nVar0 := 1
Local nVar1 := 20


//variaveis private
Private cPri := 'private1'


//Variavel public
Public _cPublic := 'HB'

TestEscop (nVar0, @nVar1)





RETURN
//----------- Função Static -----

Static function TestEscop (nValor1, nValor2)
    Local _cPublic := 'Alterei'
    Default nValor1 := 0
    
    // Alterando conteudo da Variavel
    nValor2 := 10

    //mostrar conteudo da variavel private
    Alert("Private: "+ cPri)

    //Alterar valor da variavel public
    Alert("Publica: "+ _cPublic)

    MsgAlert(nValor2)
    

RETURN
