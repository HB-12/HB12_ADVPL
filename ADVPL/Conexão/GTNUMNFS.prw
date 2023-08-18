User Function GTNUMNFS()

Local cNumNFS := Paramixb

Local cNumAux := ""

Local cNumAux2 := ""

    If Len(cNumNFS) > 9

        cNumAux := SubStr(cNumNFS,1,4)

   

        // Pega as 9 posições da direita para esquerda

        cNumAux2 := Right(Alltrim(cNumNFS),9)

        // Remove zeros a esquerda

        Do While Len(cNumAux2) > 0 .And. SubStr(cNumAux2,1,1) == "0"

            cNumAux2 := SubStr(cNumAux2,2)

        EndDo

        // Incrementa 2021 + 503 = 2021503

        cNumAux += PadL(AllTrim(cNumAux2),5,"0")

    Else

        cNumAux := cNumNFS

    EndIf

   

Return cNumAux
