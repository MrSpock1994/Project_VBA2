Attribute VB_Name = "Módulo1"
Sub Atualizar_1()

Dim wkOrigem As Worksheet
Dim wkDestino5 As Worksheet
Dim wkDestino As Worksheet
Dim x As Double
Dim y As Double
Dim z As Double
Dim Linha As Double



Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.DisplayAlerts = False


Workbooks.Open Filename:="X:\CO\BI\Cargo Drop\No_Show_Project.xlsm"

Set wkOrigem = Workbooks("No_Show_Project.xlsm").Worksheets("No_Show")
Set wkDestino = Workbooks("Controle.xlsm").Worksheets("Controle")
Set wkDestino1 = Workbooks("Controle.xlsm").Worksheets("Navios")
Set wkDestino2 = Workbooks("Controle.xlsm").Worksheets("Portos")
Set wkDestino3 = Workbooks("Controle.xlsm").Worksheets("Disponibilizados")
Set wkDestino5 = Workbooks("Controle.xlsm").Worksheets("Atualizados")



'Copiando todos os dados da planilha de origem
x = 9
With wkOrigem
verificaCel = Sheets("No_Show").Cells(x, 6).Value
Do While verificaCel <> ""
    x = x + 1
    verificaCel = Sheets("No_Show").Cells(x, 6).Value
Loop
y = x

ultima_linha = "AA" & y
Sheets("No_Show").Range("F9:" + ultima_linha).Copy
wkDestino5.Range("A2").PasteSpecial xlPasteValues

End With
wkDestino5.Activate
With wkDestino5
'INSERIR AQUI A EXCLUSAO DOS FATURADOS, PENDENTES E SEM CNPJ  (Considerar os vazios)
wkDestino5.Range("U1").AutoFilter Field:=21, Criteria1:=""
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="Pendente"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="Faturado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="Cancelado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="Substituído"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="TI/Outros"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="FaturadoCanceladoCancelado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="TI/OutrosTI/OutrosTI/Outros"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="CanceladoCanceladoFaturado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="CanceladoCanceladoCanceladoCancelado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="CanceladoCanceladoCanceladoFaturado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If
wkDestino5.Range("V1").AutoFilter Field:=22, Criteria1:="CanceladoCanceladoCancelado"
wkDestino5.Range("A2:V1000").ClearContents
If wkDestino5.AutoFilterMode Or wkDestino5.FilterMode Then
    wkDestino5.ShowAllData
End If



Columns("A:V").Sort Key1:=Range("U2"), Order1:=xlAscending, Header:=xlYes

    Do
    Linha = Linha + 1
    If .Cells(Linha, 21).Value <> "" Then
    Valor = .Cells(Linha, 21).Value
    .Cells(Linha, 21).Value = Valor
    End If
    Loop Until .Cells(Linha, 21).Value = ""
    
        
End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.DisplayAlerts = True
wkDestino.Activate

End Sub
