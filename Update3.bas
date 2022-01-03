Attribute VB_Name = "Módulo3"
Sub POS_ABSORCAO()

Dim wkDestino As Worksheet
Dim wkAbsorcao As Worksheet
Dim tabela_absorcao As Range

Application.ScreenUpdating = False

Set wkDestino = Workbooks("Controle.xlsm").Worksheets("Controle")
Set wkAbsorcao = Workbooks("Controle.xlsm").Worksheets("Planilha_absorcao")

wkAbsorcao.Range("A:A").ClearContents
ThisWorkbook.Sheets("Planilha_absorcao").Range("C:C").Copy
wkAbsorcao.Range("A1").PasteSpecial xlPasteValues

On Error Resume Next

Set tabela_absorcao = ThisWorkbook.Sheets("Planilha_absorcao").Range("A2:B1000")
k = 2
With wkDestino
Do While wkDestino.Cells(k, 6).Value <> ""
    wkDestino.Cells(k, 21) = WorksheetFunction.VLookup(CDbl(ThisWorkbook.Sheets("Controle").Range("F" & k)), tabela_absorcao, 2, 0)
    k = k + 1
   
Loop

End With

wkAbsorcao.Range("A2:R20000").ClearContents
Application.ScreenUpdating = True
wkDestino.Activate
End Sub

