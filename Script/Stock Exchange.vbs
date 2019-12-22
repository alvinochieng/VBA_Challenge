Attribute VB_Name = "Module1"
Sub stock()
    Dim ws As Worksheet
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim vol As Double
    Dim first_value As Long
    Dim lastrow As Long
    Dim f As Long
    Dim i As Long
    Dim reng As Range
    Dim tl As Range
    
    For Each ws In ThisWorkbook.Worksheets
    
    f = 0
    vol = 0
    percentchange = 0
    yearlychange = 0
    first_value = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Value"
    ws.Cells(1, 16).Value = "Ticker"
    
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
     For i = 2 To lastrow
     
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        vol = ws.Cells(i, 7).Value + vol
        
        If vol = 0 Then
            ws.Range("i" & 2 + f).Value = ws.Cells(i, 1).Value
            ws.Range("j" & 2 + f).Value = 0
            ws.Range("k" & 2 + f).Value = "%" & 0
            ws.Range("l" & 2 + f).Value = 0
        Else
            If ws.Cells(first_value, 3) = 0 Then
                For y = first_value To i
                        If ws.Cells(y, 3).Value <> 0 Then
                    first_value = y
                Exit For
            End If
        Next
    End If
        
            yearlychange = (ws.Cells(i, 6) - ws.Cells(first_value, 3))
            percentchange = (yearlychange / ws.Cells(first_value, 3)) * 100
            
            first_value = i + 1
            
            ws.Range("i" & 2 + f).Value = ws.Cells(i, 1).Value
            ws.Range("j" & 2 + f).Value = yearlychange
            ws.Range("k" & 2 + f).Value = Round(percentchange, 2)
            ws.Range("l" & 2 + f).Value = vol
     
    Select Case yearlychange
        Case Is < 0
            ws.Range("j" & 2 + f).Interior.ColorIndex = 3
        Case Is >= 0
            ws.Range("j" & 2 + f).Interior.ColorIndex = 4
         Case Else
            ws.Range("j" & 2 + f).Interior.ColorIndex = 0
    End Select
    End If
        
    Set reng = ws.Range("k:k")
    Set tl = ws.Range("l:l")
    
    dblmax = Application.WorksheetFunction.Max(reng)
    dblmin = Application.WorksheetFunction.Min(reng)
    maxtotal = Application.WorksheetFunction.Max(tl)
    
    If ws.Cells(i, 11).Value = 0 Then
        ws.Cells(2, 15).Value = Round(dblmax, 2)
    ElseIf ws.Cells(i, 9).Value = 0 Then
         ws.Cells(i, 9).Value = ws.Cells(2, 16).Value
    End If
    If ws.Cells(i, 11).Value = 0 Then
        ws.Cells(3, 15).Value = Round(dblmin, 2)
    End If
    If ws.Cells(i, 12).Value = 0 Then
        ws.Cells(4, 15).Value = maxtotal
    End If
     
    vol = 0
    yearlychange = 0
    f = f + 1
    
    Else
      vol = ws.Cells(i, 7).Value + vol
    End If
      Next
    Next
End Sub


