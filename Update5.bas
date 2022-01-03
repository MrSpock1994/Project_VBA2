Attribute VB_Name = "Módulo5"
Sub Criar_Arquivo()
Dim nome As String
Dim wkDestino As Worksheet
Dim wkDestino10 As Worksheet
Dim x As Double
nome = "Faturar"
Application.ScreenUpdating = False
Set wkDestino10 = Workbooks("Controle.xlsm").Worksheets("Faturar")
Set wkDestino = Workbooks("Controle.xlsm").Worksheets("Controle")
wkDestino.Range("B1").AutoFilter Field:=2, Criteria1:="Pendente"

x = 1
With wkDestino
verificaCel = Sheets("Controle").Cells(x, 6).Value
Do While verificaCel <> ""
    x = x + 1
    verificaCel = Sheets("Controle").Cells(x, 6).Value
Loop

ultima_linha = "R" & x
Sheets("Controle").Range("F1:" + ultima_linha).Copy
wkDestino10.Range("A1").PasteSpecial xlPasteValues

End With

wkDestino10.Activate
ActiveSheet.Copy
With ActiveWorkbook
.SaveAs ThisWorkbook.Path & "\" & nome & ".xlsx"
ActiveWorkbook.Close

End With
wkDestino10.Range("A1:M1000").ClearContents
Application.ScreenUpdating = True
wkDestino.Activate
End Sub

