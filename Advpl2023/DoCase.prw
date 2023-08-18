#INCLUDE 'protheus.ch'
#INCLUDE 'parmtype.ch'


user function Docase()

local cData := '25/12/2023'

Do CASE
case cData == '20/12/2017'
alert("Não é Natal" + cData)

Case cData == '25/12/2023'
Alert("É Natal")

OTHERWISE
MSGALERT("Não se qual o dia de hoje")

ENDCASE


RETURN
