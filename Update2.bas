Attribute VB_Name = "Módulo2"
Sub Atualizar_2()
Dim wkOrigem As Worksheet
Dim wkDestino As Worksheet
Dim wkDestino1 As Worksheet
Dim wkDestino2 As Worksheet
Dim wkDestino3 As Worksheet
Dim wkDestino5 As Worksheet
Dim x As Double
Dim y As Double
Dim z As Double
Dim w As Double
Dim k As Double
Dim i As Double
Dim a As Double
Dim n As Double
Dim c As Double
Dim p As Double
Dim b As Double
Dim max As Double
Dim tabela As Range
Dim tabela_dois As Range


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.DisplayAlerts = False



Set wkOrigem = Workbooks("No_Show_Project.xlsm").Worksheets("No_Show")
Set wkDestino = Workbooks("Controle.xlsm").Worksheets("Controle")
Set wkDestino1 = Workbooks("Controle.xlsm").Worksheets("Navios")
Set wkDestino2 = Workbooks("Controle.xlsm").Worksheets("Portos")
Set wkDestino3 = Workbooks("Controle.xlsm").Worksheets("Disponibilizados")
Set wkDestino5 = Workbooks("Controle.xlsm").Worksheets("Atualizados")



'Verificar ultima linha aba controle
With wkDestino
i = 2
verificalcel = ThisWorkbook.Sheets("Controle").Cells(i, 1).Value
Do While verificalcel <> ""
    i = i + 1
    verificalcel = ThisWorkbook.Sheets("Controle").Cells(i, 1).Value
Loop

End With
'Verificar ultima linha da aba atualizados
With wkDestino5
z = 1
verificacelfinal = ThisWorkbook.Sheets("Atualizados").Cells(z, 1).Value
Do While verificacelfinal <> ""
    z = z + 1
    verificacelfinal = ThisWorkbook.Sheets("Atualizados").Cells(z, 1).Value
Loop

End With

'Copia o booking
ultimalinha = "B" & z
ThisWorkbook.Sheets("Atualizados").Range("B2:" + ultimalinha).Copy
wkDestino.Range("A" & i).PasteSpecial xlPasteValues
'Copia o CNPJ - AJUSTAR PARA NUMERO

ultimalinha = "U" & z
ThisWorkbook.Sheets("Atualizados").Range("U2:" + ultimalinha).Copy
wkDestino.Range("H" & i).PasteSpecial xlPasteValues
'Copia cliente
ultimalinha = "D" & z
ThisWorkbook.Sheets("Atualizados").Range("D2:" + ultimalinha).Copy
wkDestino.Range("G" & i).PasteSpecial xlPasteValues
'Copia Reduzidos
ultimalinha = "S" & z
ThisWorkbook.Sheets("Atualizados").Range("S2:" + ultimalinha).Copy
wkDestino.Range("P" & i).PasteSpecial xlPasteValues
'Copia Valor
ultimalinha = "T" & z
ThisWorkbook.Sheets("Atualizados").Range("T2:" + ultimalinha).Copy
wkDestino.Range("Q" & i).PasteSpecial xlPasteValues
'Copia Key
ultimalinha = "A" & z
ThisWorkbook.Sheets("Atualizados").Range("A2:" + ultimalinha).Copy
wkDestino.Range("T" & i).PasteSpecial xlPasteValues
'Copia ShortName
ultimalinha = "C" & z
ThisWorkbook.Sheets("Atualizados").Range("C2:" + ultimalinha).Copy
wkDestino.Range("S" & i).PasteSpecial xlPasteValues
'Copia disponibilizados
ultimalinha = "J" & z
ThisWorkbook.Sheets("Atualizados").Range("J2:" + ultimalinha).Copy
wkDestino3.Range("A2").PasteSpecial xlPasteValues
ultimalinha = "K" & z
ThisWorkbook.Sheets("Atualizados").Range("K2:" + ultimalinha).Copy
wkDestino3.Range("B2").PasteSpecial xlPasteValues
ultimalinha = "L" & z
ThisWorkbook.Sheets("Atualizados").Range("L2:" + ultimalinha).Copy
wkDestino3.Range("C2").PasteSpecial xlPasteValues
'Copiar portos
ultimalinha = "E" & z
ThisWorkbook.Sheets("Atualizados").Range("F2:" + ultimalinha).Copy
wkDestino2.Range("E2").PasteSpecial xlPasteValues
ultimalinha = "F" & z
ThisWorkbook.Sheets("Atualizados").Range("G2:" + ultimalinha).Copy
wkDestino2.Range("J2").PasteSpecial xlPasteValues
'Copia Navio
ultimalinha = "C" & z
ThisWorkbook.Sheets("Atualizados").Range("C2:" + ultimalinha).Copy
wkDestino1.Range("H2").PasteSpecial xlPasteValues

'Copia Observação concatenando
p = 2
b = i
Do While ThisWorkbook.Sheets("Controle").Cells(b, 1).Value <> "" And ThisWorkbook.Sheets("Atualizados").Cells(p, 17).Value <> ""
    concatenados = "Booking" & ":" & ThisWorkbook.Sheets("Controle").Cells(b, 1) & "-" & ThisWorkbook.Sheets("Atualizados").Cells(p, 17)
    ThisWorkbook.Sheets("Controle").Cells(b, 18) = concatenados
    p = p + 1
    b = b + 1
Loop


With wkDestino3
w = 2

'Ajusta disponibilizados
For w = 2 To z
    max = WorksheetFunction.max(CDbl(ThisWorkbook.Sheets("Disponibilizados").Cells(w, 1)), CDbl(ThisWorkbook.Sheets("Disponibilizados").Cells(w, 2)), CDbl(ThisWorkbook.Sheets("Disponibilizados").Cells(w, 3)))
    wkDestino3.Cells(w, 4).Value = max
    
Next

ultima_linha = "D" & z - 1
ThisWorkbook.Sheets("Disponibilizados").Range("D2:" + ultima_linha).Copy
wkDestino.Range("O" & i).PasteSpecial xlPasteValues


End With

'Ajusta portos
With wkDestino2
k = 2
Set tabela_dois = ThisWorkbook.Sheets("Portos").Range("A1:C23")
Do While wkDestino2.Cells(k, 5).Value <> ""
    wkDestino2.Cells(k, 6) = WorksheetFunction.VLookup(CStr(ThisWorkbook.Sheets("Portos").Range("E" & k)), tabela_dois, 3, 0)
    wkDestino2.Cells(k, 7) = WorksheetFunction.VLookup(CStr(ThisWorkbook.Sheets("Portos").Range("E" & k)), tabela_dois, 2, 0)
    wkDestino2.Cells(k, 11) = WorksheetFunction.VLookup(CStr(ThisWorkbook.Sheets("Portos").Range("J" & k)), tabela_dois, 3, 0)
    wkDestino2.Cells(k, 12) = WorksheetFunction.VLookup(CStr(ThisWorkbook.Sheets("Portos").Range("J" & k)), tabela_dois, 2, 0)
    k = k + 1
   
Loop

ultima_linha = "F" & k
ThisWorkbook.Sheets("Portos").Range("F2:" + ultima_linha).Copy
wkDestino.Range("K" & i).PasteSpecial xlPasteValues
ultima_linha = "G" & k
ThisWorkbook.Sheets("Portos").Range("G2:" + ultima_linha).Copy
wkDestino.Range("J" & i).PasteSpecial xlPasteValues
ultima_linha = "K" & k
ThisWorkbook.Sheets("Portos").Range("K2:" + ultima_linha).Copy
wkDestino.Range("M" & i).PasteSpecial xlPasteValues
ultima_linha = "L" & k
ThisWorkbook.Sheets("Portos").Range("L2:" + ultima_linha).Copy
wkDestino.Range("L" & i).PasteSpecial xlPasteValues

End With

'Ajustar navio
With wkDestino1
n = 2
Set tabela = ThisWorkbook.Sheets("Navios").Range("A1:B1000")
wkDestino1.Columns("H").Replace What:=" ", Replacement:=""
verificaCel = ThisWorkbook.Sheets("Navios").Cells(n, 8)
Do While verificaCel <> ""
    novovalor = Left(wkDestino1.Cells(n, 8).Value, 5)
    wkDestino1.Cells(n, 9) = novovalor
    novovalor2 = Right(wkDestino1.Cells(n, 8).Value, 4)
    wkDestino1.Cells(n, 10) = novovalor2
    wkDestino1.Cells(n, 11) = WorksheetFunction.VLookup(CStr(ThisWorkbook.Sheets("Navios").Range("I" & n)), tabela, 2, 0)
    wkDestino1.Cells(n, 12) = ThisWorkbook.Sheets("Navios").Range("K" & n) & "/" & ThisWorkbook.Sheets("Navios").Range("J" & n)
    n = n + 1
    verificaCel = ThisWorkbook.Sheets("Navios").Cells(n, 8)
Loop

ultima_linha = "L" & n
ThisWorkbook.Sheets("Navios").Range("L2:" + ultima_linha).Copy
wkDestino.Range("I" & i).PasteSpecial xlPasteValues

End With

'Ajusta referencia
With wkDestino
a = 2
verificaCel = ThisWorkbook.Sheets("Controle").Cells(a, 1).Value
Do While verificaCel <> ""
    ThisWorkbook.Sheets("Controle").Cells(a, 6).Value = (a - 1)
    ThisWorkbook.Sheets("Controle").Cells(a, 14).Value = "No Show"
    a = a + 1
    verificaCel = ThisWorkbook.Sheets("Controle").Cells(a, 1).Value
Loop


verificaCel_status = ThisWorkbook.Sheets("Controle").Cells(i, 1).Value
verificaCel_status2 = ThisWorkbook.Sheets("Controle").Cells(i, 2).Value
Do While verificaCel_status <> "" And verificaCel_status2 = ""
    ThisWorkbook.Sheets("Controle").Cells(i, 2).Value = "Pendente"
        i = i + 1
     verificaCel_status = ThisWorkbook.Sheets("Controle").Cells(i, 1).Value
     verificaCel_status2 = ThisWorkbook.Sheets("Controle").Cells(i, 2).Value
Loop

End With



wkDestino1.Range("H2:L10000").ClearContents
wkDestino2.Range("E2:L10000").ClearContents
wkDestino3.Range("A2:D10000").ClearContents
wkDestino5.Range("A2:V10000").ClearContents


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.DisplayAlerts = True
wkDestino.Activate
End Sub
