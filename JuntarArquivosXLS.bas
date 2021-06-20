Attribute VB_Name = "Módulo1"
Sub Bradesco()

Dim wbArquivo       As Workbook
Dim wbCliente       As Workbook
Dim Cliente         As Worksheet
Dim Cliente_        As Worksheet
Dim Endereco        As Worksheet
Dim Endereco_       As Worksheet
Dim Comp            As Worksheet
Dim Comp_           As Worksheet
Dim Cotista         As Worksheet
Dim Cotista_        As Worksheet
Dim Conta           As Worksheet
Dim Conta_          As Worksheet
Dim Termos          As Worksheet
Dim Termos_         As Worksheet
Dim Perfil          As Worksheet
Dim Perfil_         As Worksheet
Dim Nomes           As Worksheet
Dim varNomes        As String
Dim VarEnd          As String
Dim VarConta        As Long
Dim Codigo          As String
Dim CodErrado       As String
Dim ContaCorrente   As String
Dim NomeDoCliente   As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False

'ALTERA TEMPORARIAMENTE OS NOMES DAS ABAS
Sheets("Cliente").Name = "Cliente_"
Sheets("Endereço").Name = "Endereço_"
Sheets("Cliente Complemento").Name = "Cliente Complemento_"
Sheets("Cotista").Name = "Cotista_"
Sheets("Conta Externa").Name = "Conta Externa_"
Sheets("Termo de Adesao").Name = "Termo de Adesao_"
Sheets("Cotista Perfil Investimento").Name = "Cotista Perfil Investimento_"

Set Cliente_ = Sheets("Cliente_")
Set Endereco_ = Sheets("Endereço_")
Set Comp_ = Sheets("Cliente Complemento_")
Set Cotista_ = Sheets("Cotista_")
Set Conta_ = Sheets("Conta Externa_")
Set Termos_ = Sheets("Termo de Adesao_")
Set Perfil_ = Sheets("Cotista Perfil Investimento_")
Set Nomes = Sheets("NOMES")

'LIMPA O CONTEÚDO DA PLANILHA
Cliente_.Activate
Range("A2:BF50").ClearContents
Endereco_.Activate
Range("A2:W100").ClearContents
Comp_.Activate
Range("A2:AB50").ClearContents
Cotista_.Activate
Range("A2:CU50").ClearContents
Conta_.Activate
Range("A2:T100").ClearContents
Termos_.Activate
Range("A2:F150").ClearContents
Perfil_.Activate
Range("A2:E50").ClearContents

'SELECIONA O NOME, CÓDIGO E CONTA CORRENTE A SER UTILIZADO
Nomes.Select
Nomes.Cells(Nomes.Rows.Count, 1).End(xlUp).Select
varNomes = ActiveCell
Codigo = ActiveCell.Offset(0, 2)
ContaCorrente = ActiveCell.Offset(0, 3)

Set wbArquivo = Workbooks("ARQUIVO PARA CADASTROS.xlsm")
wbArquivo.Activate
Nomes.Activate

Do While ActiveCell.Row >= 2

Workbooks.Open Filename:="C:\Users\adria\Desktop\Validações de Cadastro\" & Trim$(varNomes) & ".xls"


    Set wbCliente = ActiveWorkbook
    wbCliente.Activate
    Set Cliente = Sheets("Cliente")
    Set Endereco = Sheets("Endereço")
    Set Comp = Sheets("Cliente Complemento")
    Set Cotista = Sheets("Cotista")
    Set Conta = Sheets("Conta Externa")
    Set Termos = Sheets("Termo de Adesao")
    Set Perfil = Sheets("Cotista Perfil Investimento")


    'ABA CLIENTE
    
    wbCliente.Activate
    Cliente.Activate
    'Range("U2") = Range("T2").Value
    'Range("T2").ClearContents
    'Range("AB2").ClearContents
    'NomeDoCliente = Range("AD2").Value
    'Range("H2").Value = NomeDoCliente
    'Range("I2").Value = NomeDoCliente
    If Range("U2").Value = "CAPITALISTA" Or Range("U2").Value = "ESTUDANTE" Or Range("U2").Value = "APOSENTADO" Then
    
        Range("av2").Value = ""
    
    End If
    
    Cliente.Cells(Cliente.Rows.Count, 1).End(xlUp).Select
    
    ActiveCell.EntireRow.Copy
    
    wbArquivo.Activate
    Cliente_.Activate
    Cliente_.Cells(Cliente_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveCell.PasteSpecial
    CodErrado = ActiveCell
    Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    
    'ABA ENDEREÇO
    
    wbCliente.Activate
    Endereco.Activate
    VarEnd = Range("C2").Value
    'Range("V2").Value = VarEnd
    'Range("V3").Value = VarEnd
    'Range("U2").Value = VarEnd
    'Range("U3").Value = VarEnd
    Range("A2:W3").Select
    Selection.Copy
    
    wbArquivo.Activate
    Endereco_.Activate
    Endereco_.Cells(Endereco_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveCell.PasteSpecial
    Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'ABA CLIENTE COMPLEMENTO
    
    wbCliente.Activate
    Comp.Activate
    If Range("A2") <> "" Then
        'Comp.Cells(Comp.Rows.Count, 3).End(xlUp).Select
        'Prim_e_Ult_Nome = Split(ActiveCell, " ")
        'ActiveCell.Offset(0, 1) = Trim$(Prim_e_Ult_Nome(LBound(Prim_e_Ult_Nome))) & " " & Trim$(Prim_e_Ult_Nome(UBound(Prim_e_Ult_Nome)))
        Range("A2").EntireRow.Select
        Selection.Copy
        
        wbArquivo.Activate
        Comp_.Activate
        Comp_.Cells(Comp_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
        ActiveCell.PasteSpecial
        Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
    
    
    'ABA COTISTA
    
    wbCliente.Activate
    Cotista.Activate
    'Range("AV2").ClearContents
    'Range("BL2").ClearContents
    'Range("CG2").Value = "N"
    
    Range("A2").EntireRow.Select
    Selection.Copy
    
    wbArquivo.Activate
    Cotista_.Activate
    Cotista_.Cells(Cotista_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveCell.PasteSpecial
    ActiveCell.Offset(0, 5).Value = "'" & ContaCorrente
    ActiveCell.Offset(0, 28).Value = "'" & ContaCorrente
    Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'ABA CONTA
    
    wbCliente.Activate
    Conta.Activate
    Range("A2:T3").Select
    Selection.Copy
    
    wbArquivo.Activate
    Conta_.Activate
    Conta_.Cells(Conta_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveCell.PasteSpecial
    Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'ABA TERMO DE ADESAO
    
    wbCliente.Activate
    Termos.Activate
    Range("A2:F4").Select
    Selection.Copy
    
    wbArquivo.Activate
    Termos_.Activate
    Termos_.Cells(Termos_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveCell.PasteSpecial
    Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'ABA COTISTA PERFIL INVESTIMENTO
    
    wbCliente.Activate
    Perfil.Activate
    Range("A2").EntireRow.Select
    Selection.Copy
    
    wbArquivo.Activate
    Perfil_.Activate
    Perfil_.Cells(Perfil_.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveCell.PasteSpecial
    Cells.Replace What:=CodErrado, Replacement:=Codigo, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Application.CutCopyMode = False
        
    'FECHA A PLANILHA DO CLIENTE
    wbCliente.Close
       
    'MUDA O CLIENTE NA PLANILHA NOMES
    wbArquivo.Activate
    Nomes.Select
    ActiveCell.Offset(-1, 0).Select
    varNomes = ActiveCell
    Codigo = ActiveCell.Offset(0, 2)
    ContaCorrente = ActiveCell.Offset(0, 3)

Loop

'EXCLUINDO O ZERO ANTES DA CONTA CORRENTE
Conta_.Activate
Conta_.Cells(Conta_.Rows.Count, 4).End(xlUp).Select
Do While ActiveCell.Row >= 2
    VarConta = ActiveCell.Value
    ActiveCell = VarConta
    ActiveCell.Offset(-1, 0).Select
Loop

wbArquivo.Activate

'RETORNA O NOME CORRETO DAS ABAS NO ARQUIVO PRINCIPAL
wbArquivo.Activate
Cliente_.Name = "Cliente"
Endereco_.Name = "Endereço"
Comp_.Name = "Cliente Complemento"
Cotista_.Name = "Cotista"
Conta_.Name = "Conta Externa"
Termos_.Name = "Termo de Adesao"
Perfil_.Name = "Cotista Perfil Investimento"


Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Lembrou de lavar as mãos hoje?"

End Sub
