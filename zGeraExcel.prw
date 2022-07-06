//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
 
//Constantes
#Define STR_PULA    Chr(13)+Chr(10)
 
/*/{Protheus.doc} zGeraExcel
Listagem de NFs em Excel via AdvPL
@author Matheus Costa
@since 05/07/2022
@version 1.0
    U_zGeraExcel()
/*/


///// Criação de Tela de Parametros
USER FUNCTION xTela1()

///Variaveis Locais
Local _aArea := GetArea()

Local _oDlgExcel /// Tela MSDIALOG
Local _nOpc := 0 //Validação
 
/// Separador e especifico
Local nTamBtn      := 40
Local oBtnConf
Local oBtnCanc

Local	_oGetFil01 //oGet Filial 01
Local	_oGetFil02 //oGet Filial 02

Local _oGetDt01 //oGet Emissão de
Local _oGetDt02 //oGet Emissão Até


//// Variaveis Privadas 
Private	_cGetFil01 := Space(2) // Filial de ?
Private _cGetFil02 := Space(2) // Filial Até ?

Private _cGetDt01 := SToD("") //Emissão de ?
Private _cGetDt02 := SToD("") //Emissão Até ?


    Define Font _oFont Name "Arial" Size 0,-12 Bold

	/// Criação da Tela
	Define MsDialog _oDlgExcel From 333, 227 To 545, 600  Title OemToAnsi( "-Informe os parametros" ) Pixel Of oMainWnd

    //// Campos Filial
	@ 007, 010 Say "Filial De    ?" Pixel
	@ 007, 090 MsGet _oGetFil01 Var _cGetFil01  F3 "SM0" Size 070, 009  Pixel HASBUTTON

	@ 025, 010 Say "Filial Até   ? *" Font _oFont Pixel
	@ 025, 090 MsGet _oGetFil02 Var _cGetFil02  F3 "SM0" Size 070, 009 Pixel HASBUTTON


	////////////////////////////////////// - Datas de Emissão

	////// - Campo Data De?
	@ 045, 010 Say "Data De    ? " Pixel
	@ 045, 090 MsGet _oGetDt01  Var _cGetDt01  Picture "@D" Size 070, 009 Pixel HASBUTTON  Of _oDlgExcel

          
	////// - Campos Data Até?
	@ 060, 010 Say "Data Até    ? *" Font _oFont Pixel
	@ 060, 090 MsGet _oGetDt02  Var _cGetDt02  Picture "@D" Size 070, 009 Pixel HASBUTTON Of _oDlgExcel
   

	////////////////////////////////////// - BUTTONS
	   
    @ 080, 030 BUTTON oBtnConf PROMPT "Confirmar" SIZE nTamBtn, 013 OF _oDlgExcel Action ( _nOpc := 01, _oDlgExcel:End() ) Of _oDlgExcel Pixel
    @ 080, 100 BUTTON oBtnCanc PROMPT "Cancelar"  SIZE nTamBtn, 013 OF _oDlgExcel Action ( _nOpc := 02, _oDlgExcel:End() ) Of _oDlgExcel Pixel
             

	Activate Dialog _oDlgExcel Centered


    //// Validação dos Parâmetros
    if _nOpc == 1
    Processa({|| zGeraExcel()}, "Exportando...")  
    endif

RestArea(_aArea)
RETURN


///// - Geração do Excel
Static Function zGeraExcel()

    Local aArea        := GetArea()
    Local cQuery        := "" // SF1 - Entradas
    Local cQuery1 := "" /// SF2 - Saídas
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'zGeraExcel.xml'
    
    //Pegando os dados
    cQuery := " SELECT "                                                    + STR_PULA
    cQuery += "     SF1.F1_EMISSAO, "                                            + STR_PULA
    cQuery += "     SA2.A2_NOME, "                                            + STR_PULA
    cQuery += "     SF1.F1_DTDIGIT, "                                            + STR_PULA
    cQuery += "     SF1.F1_DOC, "                                        + STR_PULA
    cQuery += "     SF1.F1_SERIE, "                                        + STR_PULA
    cQuery += "     SF1.F1_FORNECE, "                                        + STR_PULA
    cQuery += "     SF1.F1_LOJA, "                                        + STR_PULA
    cQuery += "     SF1.F1_VALBRUT "                                        + STR_PULA
    cQuery += " FROM "                                                    + STR_PULA
    cQuery += " "+RetSQLName('SF1')+" SF1 "                            + STR_PULA
    cQuery += " INNER JOIN  "+RetSQLName('SA2')+" SA2 "                                                    + STR_PULA
    cQuery += " ON SA2.A2_COD = SF1.F1_FORNECE "                                                    + STR_PULA
    cQuery += " WHERE "                                                    + STR_PULA
    cQuery += " SF1.F1_FILIAL = '"+_cGetFil01+"' OR SF1.F1_FILIAL = '"+_cGetFil02+"' " + STR_PULA
    cQuery += " AND  SF1.F1_DTDIGIT >= '"+DTOS(_cGetDt01)+"' AND SF1.F1_DTDIGIT <= '"+DToS(_cGetDt02)+"' "            + STR_PULA
    cQuery += " AND SF1.D_E_L_E_T_ = ' ' "            + STR_PULA
    
    MemoWrite('C:\Temp\REL001.txt',cQuery) 

    TCQuery cQuery New Alias "QRYPRO"  

    ///// SF2 - CABEÇALHO 

    cQuery1 := " SELECT "                                                    + STR_PULA
    cQuery1 += "     SF2.F2_EMISSAO, "                                            + STR_PULA
    cQuery1 += "     SA2.A2_NOME, "                                            + STR_PULA
    cQuery1 += "     SF2.F2_DTDIGIT, "                                            + STR_PULA
    cQuery1 += "     SF2.F2_DOC, "                                        + STR_PULA
    cQuery1 += "     SF2.F2_SERIE, "                                        + STR_PULA
    cQuery1 += "     SF2.F2_CLIENTE, "                                        + STR_PULA
    cQuery1 += "     SF2.F2_LOJA, "                                        + STR_PULA
    cQuery1 += "     SF2.F2_VALBRUT "                                        + STR_PULA
    cQuery1 += " FROM "                                                    + STR_PULA
    cQuery1 += " "+RetSQLName('SF2')+" SF2 "                            + STR_PULA
    cQuery1 += " INNER JOIN  "+RetSQLName('SA2')+" SA2 "                                                    + STR_PULA
    cQuery1 += " ON SA2.A2_COD = SF2.F2_CLIENTE "                                                    + STR_PULA
    cQuery1 += " WHERE " 
    cQuery1 += " SF2.F2_FILIAL = '"+_cGetFil01+"' OR SF2.F2_FILIAL = '"+_cGetFil02+"' " + STR_PULA                                                   + STR_PULA
    cQuery1 += " AND  SF2.F2_EMISSAO >= '"+DToS(_cGetDt01)+"' AND SF2.F2_EMISSAO <= '"+DToS(_cGetDt02)+"' "            + STR_PULA
    cQuery1 += " AND SF2.D_E_L_E_T_ = ' ' "            + STR_PULA

    MemoWrite('C:\Temp\REL002.txt',cQuery1) 

    TCQuery cQuery1 New Alias "QRYPRO1"  

    ProcRegua(0)
    IncProc("Adicionando registros no Excel")
    
    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New() 

    //Aba 01 - Teste
    oFWMsExcel:AddworkSheet("Saidas") //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable("Saidas","Lista de NFs de Saida")
        //Criando Colunas
        
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Data de Emissão",1) //1 = Modo Texto
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Nota Fiscal",1) 
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Série",1) 
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Cliente",1)
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Loja",1)
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Nome",1)
        oFWMsExcel:AddColumn("Saidas","Lista de NFs de Saida","Valor",3,3)

        While !(QRYPRO1->(EoF()))       
            
        nEmissao := SToD(QRYPRO1->F2_EMISSAO)   //Conversão de Data 
        oFWMsExcel:AddRow("Saidas","Lista de NFs de Saida",{;
                                                                    CValToChar(nEmissao),;
                                                                    QRYPRO1->F2_DOC,;
                                                                    QRYPRO1->F2_SERIE,;
                                                                    QRYPRO1->F2_FORNECE,;
                                                                    QRYPRO1->F2_LOJA,;
                                                                    QRYPRO1->A2_NOME,;
                                                                    QRYPRO1->F2_VALBRUT;
                                                                    })     
            //Pulando Registro
            QRYPRO->(DbSkip())
        EndDo
     

        //Criando as Linhas
      /*  oFWMsExcel:AddRow("Saidas","Lista de NFs de Saida",{11,12,13,sToD('20140317')})
        oFWMsExcel:AddRow("Saidas","Lista de NFs de Saida",{21,22,23,sToD('20140217')})
        oFWMsExcel:AddRow("Saidas","Lista de NFs de Saida",{31,32,33,sToD('20140117')})
        oFWMsExcel:AddRow("Saidas","Lista de NFs de Saida",{41,42,43,sToD('20131217')}) */

        //Aba 01 - Teste
         oFWMsExcel:AddworkSheet("Entrada") //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable("Entrada","Lista de NFs de Entrada")
        //Criando Colunas
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Data de Emissão",1) //1 = Modo Texto
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Nota Fiscal",1) 
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Série",1) 
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Cliente",1)
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Loja",1)
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Nome",1)
        oFWMsExcel:AddColumn("Entrada","Lista de NFs de Entrada","Valor",3,3) // Valor com R$ 3 / sem 2
         //Criando as Linhas... Enquanto não for fim da query   
                 
        While !(QRYPRO->(EoF()))       
            
        nEmissao := SToD(QRYPRO->F1_EMISSAO)   //Conversão de Data 
        oFWMsExcel:AddRow("Entrada","Lista de NFs de Entrada",{;
                                                                    CValToChar(nEmissao),;
                                                                    QRYPRO->F1_DOC,;
                                                                    QRYPRO->F1_SERIE,;
                                                                    QRYPRO->F1_FORNECE,;
                                                                    QRYPRO->F1_LOJA,;
                                                                    QRYPRO->A2_NOME,;
                                                                    QRYPRO->F1_VALBRUT;
                                                                    })     
            //Pulando Registro
            QRYPRO->(DbSkip())
        EndDo
     

    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
         
    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
    oExcel:SetVisible(.T.)                 //Visualiza a planilha
    oExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas
     
    QRYPRO->(DbCloseArea())
    QRYPRO1->(DbCloseArea())
    RestArea(aArea)
Return
