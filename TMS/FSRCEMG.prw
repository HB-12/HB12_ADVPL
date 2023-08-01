#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TBICONN.CH" 
#INCLUDE "REPORT.CH"
 
/*/{Protheus.doc} FsFinR01
    
    RelatÃ³rio de Vendas.
    
    @type  Function
    @author user
    @since 02/03/2020
  
  Executa a consulta no banco de dados e retorna area com dados para impressÃ£o.

    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
    /*/
User Function FsRCemg()

Local aAreOld       := {GetArea()} 
Private dDataIni  	:= ''
Private dDataFim  	:= ''

If FPergunte('FSRCEMG')
    dDataIni	:= MV_PAR01
    dDataFim	:= MV_PAR02

    Processa({|| reportDef() })

  EndIf 

aEval(aAreOld, {|xAux| RestArea(xAux)})
    
Return Nil


//-------------------------------------------------------------------
/*/{Protheus.doc} reportDef
Funcao para chamada das definicoes do relatorio TReport

@protected
@author    Alex Teixeira de Souza
@since     23/07/2020
@obs      
@param	    	
            			
Alteracoes Realizadas desde a Estruturacao Inicial
Data       Programador     Motivo
/*/
//------------------------------------------------------------------- 
static function reportDef()
	Local oReport
	Local oSection1
	Local cTitulo	:= "Relatorio CEMIG - Período "+Dtoc(dDataIni)+" a "+Dtoc(dDataFim)
    Local cPictVal  := "@E 999,999,999.99"

    Private cAliasQry   		:= GetNextAlias()

	oReport := TReport():New('FSRCEMIG', cTitulo, , {|oReport| PrintReport(oReport)},"Impressão do Relatório")
	oreport:nfontbody	:= 7
    oReport:DisableOrientation(.T.)
	oReport:SetLandscape()
	oReport:SetTotalInLine(.F.)
	oReport:ShowHeader()	

    oSection1 := TRSection():New(oReport,"CEMIG",cAliasQry) 
    oSection1:LAUTOSIZE := .T.

    TRCell():New(oSection1,"NOTA",,"NOTA",,10,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"NVIAGEM",,"N Viagem",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"CTE",,"CTE",,10,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"SERIE","DTC","SERIE",,20,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"RECEBIDO","DTC","Recebido",,20,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"DATAENTREGA",,"Data Entrega",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
   	TRCell():New(oSection1,"HORAENTREGA",,"Hora Entrega",,10,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
   	TRCell():New(oSection1,"PESOTRANSP",,"Peso Transp",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"VALOR",,"Valor Mercad",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"FRETE",,"Valor Frete",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"IMPOSTO",,"Valor Imposto",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"ORIGEM",,"Origem",,25,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"UFORIGEM",,"UF Origem",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
	TRCell():New(oSection1,"DESTINO",,"Destino",,25,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"UFDESTINO",,"UF Destino",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"PLACA",,"Placa",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"PESONOMINAL",,"Peso Nominal",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"PEROCUP",,"% Ocupacao",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"MODAL",,"Modal",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"KM",,"KM",cPictVal,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"MOTORISTA",,"Motorista",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"E5MOTBX",,"Mot Baixa",,5,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"E5DOCUMEN",,"Documento",,5,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )
    TRCell():New(oSection1,"CCUSTO",,"CC",,15,/*lPixel*/,/*{|| code-block de impressao }*/,/*nALign*/ "LEFT",/*lLineBreak*/,/*cHeaderAlign*/"LEFT",/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/.T. )

 	oReport:PrintDialog()
	
	If (Select(cAliasQry)!= 0)
		dbSelectArea(cAliasQry)
		dbCloseArea()
	
		If File(cAliasQry+GetDBExtension())
			FErase(cAliasQry+GetDBExtension())
		EndIf
	EndIf


return (oReport)

Static Function PrintReport(oReport)
	Local oSection1   
    Local lProc			:= .f.
	Local nQtdRec		:= 0

    oSection1   	:= 	oReport:Section(1)

	DBSelectArea("DT6")
	DT6->(DBSetOrder(1))

    Processa({|| lProc := FSeleDados()})//,"Aguarde...","Coletando Dados..."

	If lProc

		//Define a mensagem apresentada durante a geração do relatório.
		oReport:SetMsgPrint(Space(30)+"Lendo Registros")

		//Seta o contador da regua   
		(cAliasQry)->(dbGoTop())          
		(cAliasQry)->(DBEval({||  nQtdRec++})) 

		oReport:SetMeter(nQtdRec)

		//seleciono o arquivo de trabalho gerado pela query e coloco no inicio
		dbSelectArea(cAliasQry)
		(cAliasQry)->(dbGoTop())          

		oSection1:Init()	

		Do While !(cAliasQry)->(Eof()) .and. !oReport:Cancel()

			If oReport:Cancel()
				Exit
			EndIf

			oReport:IncMeter()

			oSection1:Cell("NOTA"):Setvalue( (cAliasQry)->DTC_NUMNFC )
            oSection1:Cell("CTE"):Setvalue( (cAliasQry)->DT6_DOC )
            oSection1:Cell("SERIE"):Setvalue( (cAliasQry)->DT6_SERIE )
            oSection1:Cell("RECEBIDO"):Setvalue( (cAliasQry)->DUA_RECEBE )
           	oSection1:Cell("DATAENTREGA"):Setvalue( Dtoc(Stod((cAliasQry)->DUA_DATCHG)) )
			oSection1:Cell("NVIAGEM"):Setvalue( (cAliasQry)->DTR_VIAGEM )
			cHora := ""
			If !Empty((cAliasQry)->DUA_HORCHG)
				cHora :=  Substr((cAliasQry)->DUA_HORCHG,1,2)+":"+Substr((cAliasQry)->DUA_HORCHG,3,2)  
          	Endif

          	oSection1:Cell("HORAENTREGA"):Setvalue( cHora )
            oSection1:Cell("PESOTRANSP"):Setvalue( (cAliasQry)->DT6_PESO )
    	    oSection1:Cell("VALOR"):Setvalue( (cAliasQry)->DT6_VALMER )
    	    oSection1:Cell("FRETE"):Setvalue( (cAliasQry)->DT6_VALFRE )
    	    oSection1:Cell("IMPOSTO"):Setvalue( (cAliasQry)->DT6_VALIMP )
           	oSection1:Cell("ORIGEM"):Setvalue( (cAliasQry)->ORIGEM )
			oSection1:Cell("UFORIGEM"):Setvalue( (cAliasQry)->UFORIGEM )
           	oSection1:Cell("DESTINO"):Setvalue( (cAliasQry)->DESTINO )
			oSection1:Cell("UFDESTINO"):Setvalue( (cAliasQry)->UFDESTINO )
           	oSection1:Cell("PLACA"):Setvalue( (cAliasQry)->DA3_PLACA )
          	oSection1:Cell("PESONOMINAL"):Setvalue( (cAliasQry)->DA3_CAPACN )

			nPercOcup := 0
			IF (cAliasQry)->DT6_PESO <> 0 .and. (cAliasQry)->DA3_CAPACN <> 0
				nPercOcup := ((cAliasQry)->DT6_PESO*100)/((cAliasQry)->DA3_CAPACN*1000)  
			Endif
          	oSection1:Cell("PEROCUP"):Setvalue(nPercOcup )
          	oSection1:Cell("KM"):Setvalue( (cAliasQry)->DTC_KM )
			oSection1:Cell("MODAL"):Setvalue( (cAliasQry)->DUT_DESCRI )
			oSection1:Cell("MOTORISTA"):Setvalue( (cAliasQry)->DA4_NOME )
   
          	oSection1:Cell("E5MOTBX"):Setvalue( (cAliasQry)->E5_MOTBX )	
          	oSection1:Cell("E5DOCUMEN"):Setvalue( (cAliasQry)->E5_DOCUMEN )
			oSection1:Cell("CCUSTO"):Setvalue((cAliasQry)->DTC_ZCUSTO )  

            oSection1:PrintLine()

			(cAliasQry)->(DBSkip())	
	
		EndDo    

	Endif

Return


Static Function FSeleDados()
	/***********************************************
	*  Seleciona os titulos marcados pelo usuario. *
	***********************************************/
	Local cQrySE5		:= ""
	Local lRet  		:= .t.

	If Select(cAliasQry) <> 0
		dbSelectArea(cAliasQry)
		dbCloseArea()
	EndIf


	cQrySE5 := "SELECT DISTINCT DTC_NUMNFC, DT6_DOC, DT6_SERIE, ISNULL(DUA_RECEBE,'') DUA_RECEBE, DT6_PESO, DT6_VALMER,DT6_VALFRE, DTC.DTC_ZCUSTO,"+CRLF
	cQrySE5 += "DT6_VALIMP, ISNULL(ORI.DUY_DESCRI,'') ORIGEM, ISNULL(ORI.DUY_EST,'') UFORIGEM, "+CRLF
	cQrySE5 += "ISNULL(DES.DUY_DESCRI,'') DESTINO,  ISNULL(DES.DUY_EST,'') UFDESTINO,  "+CRLF
	cQrySE5 += "ISNULL(DA3_PLACA,'') DA3_PLACA, "+CRLF
	cQrySE5 += "DT6_PESO,ISNULL(DA3_CAPACN,0) DA3_CAPACN, ISNULL(DUT_DESCRI,'') DUT_DESCRI, DTC_KM, ISNULL(DA4_NOME,'') DA4_NOME, "+CRLF
	cQrySE5 += "ISNULL(DUA_DATCHG,'') DUA_DATCHG,ISNULL(DUA_HORCHG,'') DUA_HORCHG, "+CRLF
	cQrySE5 += "ISNULL(SD2.D2_CONTA,'') D2_CONTA, "+CRLF   
	cQrySE5 += "ISNULL(SE5.E5_MOTBX,'') E5_MOTBX, "+CRLF   
	cQrySE5 += "ISNULL(SE5.E5_DOCUMEN,'') E5_DOCUMEN, "+CRLF   
	cQrySE5 += "ISNULL(DTR.DTR_VIAGEM,'') DTR_VIAGEM "+CRLF
	cQrySE5 += "FROM " + RetSqlName("DT6") + " DT6 "+CRLF
	cQrySE5 += "INNER JOIN " + RetSqlName("DTC") + " DTC ON DTC.DTC_FILDOC = DT6.DT6_FILDOC AND DTC_DOC = DT6.DT6_DOC AND DTC.DTC_SERIE = DT6.DT6_SERIE AND DT6.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "INNER JOIN " + RetSqlName("SD2") + " SD2 ON DT6.DT6_FILDOC = SD2.D2_FILIAL AND DTC.DTC_DOC = SD2.D2_DOC AND DT6.DT6_SERIE = SD2.D2_SERIE AND SD2.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("SE5") + " SE5 ON SD2.D2_FILIAL = SE5.E5_FILIAL AND SD2.D2_DOC = SE5.E5_NUMERO AND SD2.D2_CLIENTE = SE5.E5_CLIFOR AND SD2.D2_LOJA = SE5.E5_LOJA AND SE5.E5_TIPOLAN = '' AND SE5.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER  JOIN " + RetSqlName("DUA") + " DUA ON DTC.DTC_FILIAL  = DUA.DUA_FILIAL AND DTC.DTC_FILORI = DUA.DUA_FILDOC AND DTC.DTC_DOC = DUA.DUA_DOC AND DTC.DTC_SERIE = DUA.DUA_SERIE AND DUA.D_E_L_E_T_ <> '*' "+CRLF
	//cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DTR") + " DTR ON DUA.DUA_FILIAL = DTR.DTR_FILIAL AND DUA.DUA_FILDOC = DTR.DTR_FILORI AND DUA.DUA_VIAGEM = DTR.DTR_VIAGEM AND DTR.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DTR") + " DTR ON DT6.DT6_FILIAL = DTR.DTR_FILIAL AND DT6.DT6_FILVGA = DTR.DTR_FILORI AND DT6.DT6_NUMVGA = DTR.DTR_VIAGEM AND DTR.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DA3") + " DA3 ON DTR.DTR_FILIAL = DA3.DA3_FILIAL AND DTR.DTR_CODVEI = DA3.DA3_COD AND DA3.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DUT") + " DUT ON DA3.DA3_TIPVEI = DUT.DUT_TIPVEI AND DUT.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DUY") + " ORI ON DTC.DTC_FILIAL = ORI.DUY_FILIAL AND DTC.DTC_CDRORI = ORI.DUY_GRPVEN AND ORI.D_E_L_E_T_ <> '*' "+CRLF
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DUY") + " DES ON DTC.DTC_FILIAL = DES.DUY_FILIAL AND DTC.DTC_CDRDES = DES.DUY_GRPVEN AND DES.D_E_L_E_T_ <> '*' "+CRLF	
	cQrySE5 += "LEFT OUTER JOIN " + RetSqlName("DA4") + " DA4 ON DA3.DA3_FILIAL = DA4.DA4_FILIAL AND DA3.DA3_MOTORI = DA4.DA4_COD AND DA4.D_E_L_E_T_ <> '*' "+CRLF	

	cQrySE5 += "WHERE  DT6.DT6_FILIAL = '"+FwxFilial("DT6")+"' AND DT6.DT6_DATEMI  >= '"+DTOS(dDataIni)+"' AND  DT6.DT6_DATEMI  <= '"+DTOS(dDataFim)+"' "+CRLF	
	cQrySE5 += "ORDER BY DTC_NUMNFC"+ CRLF

	cQrySE5 := ChangeQuery(cQrySE5)

	dbUseArea(.T.,"TOPCONN",TCGENQRY(,,cQrySE5),cAliasQry,.T.,.T.)     

	If (cAliasQry)->(Eof())
		lRet := .f.
	Endif		

Return lRet


/*/{Protheus.doc} FPergunte
Perguntas da rotina
@author rafaelalmeida
@since 10/06/2019
@version 1.0
@return ${return}, ${return_description}

@param lView, logical, descricao
@param lEdit, logical, descricao
@type function
/*/
Static Function FPergunte(cNomRot,lView,lEdit)

Local aParambox	:= {}
Local aRet 		:= {}

Local lRet		:= .T.
Local nX		:= 0

Default cNomRot	:= "FSRCEMIG"
Default lView	:= .T.
Default lEdit	:= .T.

Private lWhen	:= lEdit

aAdd( aParambox ,{1,"Data Inicial : "      , Ctod(Space(8)),,".T.",,"lWhen",50,.T.})	
aAdd( aParambox ,{1,"Data Final   : "      , Ctod(Space(8)),,".T.",,"lWhen",50,.T.})	

//Carrega o array com os valores utilizados na Ãºltima tela ou valores default de cada campo.
For nX := 1 To Len(aParamBox)
	aParamBox[nX][3] := ParamLoad(cNomRot,aParamBox,nX,aParamBox[nX][3])
Next nX

//Define se ira apresentar tela de perguntas
If lView
	lRet := ParamBox(aParamBox,"Parametros",aRet,{|| .T.},{},.T.,Nil,Nil,Nil,cNomRot,.F.,.F.)
Else
	For nX := 1 To Len(aParamBox)
		Aadd(aRet, aParamBox[nX][3])
	Next nX
EndIf

If lRet
	//Carrega perguntas em variaveis usadas no programa
	If ValType(aRet) == "A" .And. Len(aRet) == Len(aParamBox)
		For nX := 1 to Len(aParamBox)
			If aParamBox[nX][1] == 2 .And. ValType(aRet[nX]) == "C"
				&("Mv_Par"+StrZero(nX,2)) := aScan(aParamBox[nX][4],{|x| Alltrim(x) == aRet[nX]})
			ElseIf aParamBox[nX][1] == 2 .And. ValType(aRet[nX]) == "N"
				&("Mv_Par"+StrZero(nX,2)) := aRet[nX]
			Else
				&("Mv_Par"+StrZero(nX,2)) := aRet[nX]
			Endif
		Next nX
	EndIf

	If lEdit
		//Salva parametros
		ParamSave(cNomRot,aParamBox,"1")
	EndIf
EndIf

Return(lRet)
