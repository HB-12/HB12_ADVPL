#INCLUDE 'protheus.ch'
#INCLUDE 'parmtype.ch'


user function Docase()

local cData := '25/12/2023'

Do CASE
case cData == '20/12/2017'
alert("N�o � Natal" + cData)

Case cData == '25/12/2023'
Alert("� Natal")

OTHERWISE
MSGALERT("N�o se qual o dia de hoje")

ENDCASE


RETURN
