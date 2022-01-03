Attribute VB_Name = "Módulo4"
Sub FATURA()

Dim wkDestino As Worksheet
Dim wkFatura As Worksheet
Dim tabela_fatura As Range
Dim Valor As Double
Dim Linha As Double
Dim data As String
Linha = 1
Application.ScreenUpdating = False

Set wkDestino = Workbooks("Controle.xlsm").Worksheets("Controle")
Set wkFatura = Workbooks("Controle.xlsm").Worksheets("Planilha_fatura")
With wkFatura
    Do
    Linha = Linha + 1
    If .Cells(Linha, 2).Value <> "" Then
    Valor = .Cells(Linha, 2).Value
    .Cells(Linha, 2).Value = Valor
    End If
    
    Loop Until .Cells(Linha, 2).Value = ""
    
        
End With



On Error Resume Next
Set tabela_fatura = ThisWorkbook.Sheets("Planilha_fatura").Range("B2:L10000")
k = 2
With wkDestino
Do While wkDestino.Cells(k, 6).Value <> ""
    wkDestino.Cells(k, 3) = WorksheetFunction.VLookup(CDbl(ThisWorkbook.Sheets("Controle").Range("U" & k)), tabela_fatura, 10, 0)
    wkDestino.Cells(k, 4) = WorksheetFunction.VLookup(CDbl(ThisWorkbook.Sheets("Controle").Range("U" & k)), tabela_fatura, 9, 0)
    wkDestino.Cells(k, 5) = WorksheetFunction.VLookup(CDbl(ThisWorkbook.Sheets("Controle").Range("U" & k)), tabela_fatura, 5, 0)
    k = k + 1
   
Loop




End With

wkFatura.Range("A2:M10000").ClearContents
Application.ScreenUpdating = True

wkDestino.Activate
End Sub

