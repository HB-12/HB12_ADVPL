User function GetCampoSize()

Local oDict:= GetDD()
Local nFieldSize := oDict:FieldGet("SX6020", "X6_CONTEUD", "size")
Return nFieldSize

