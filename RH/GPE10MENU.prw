#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "TOTVS.CH"

/*/{Protheus.doc} GPE10MENU
Ponto de Entrara utilizado para acrescentar op��es no menu da tela de Cadastro de Funcion�rios.

Link https://tdn.totvs.com/pages/releaseview.action?pageId=6079250
@type function
@author gustavo.barcelos
@since 3/6/2023
/*/
User Function GPE10MENU()

    //��������������������������������������������������������������Ŀ
	//� Define Array contendo as Rotinas a executar do programa      �
	//� ----------- Elementos contidos por dimensao ------------     �
	//� 1. Nome a aparecer no cabecalho                              �
	//� 2. Nome da Rotina associada                                  �
	//� 3. Usado pela rotina                                         �
	//� 4. Tipo de Transacao a ser efetuada                          �
	//�    1 - Pesquisa e Posiciona em um Banco de Dados             �
	//�    2 - Simplesmente Mostra os Campos                         �
	//�    3 - Inclui registros no Bancos de Dados                   �
	//�    4 - Altera o registro corrente                            �
	//�    5 - Remove o registro corrente do Banco de Dados          �
	//����������������������������������������������������������������
    aAdd(aRotina, { "Comunica Vencimento de Aviso - EMTEL", "U_FSGPEP01", 0, 2 })

Return(Nil)
