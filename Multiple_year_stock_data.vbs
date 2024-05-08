Attribute VB_Name = "Module1"
Sub Stock()
'Declaring Worksheet
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
'Declaring Variables
Dim opening As Double
Dim quatPercentage As Double
Dim total As LongLong
Dim closing As Double
Dim quatchange As Double
Dim Min As Double
Dim Max As Double
Dim Stock As Double
Dim rng1 As Range
Dim rng2 As Range

'Initializing Variables
total = 0
opening = 0
closing = 0
quatPercentage = 0
quatchange = 0
l = 2
k = 1
j = 2
Min = 0
Max = 0
Stock = 0
'Conditional formatting
Set rng1 = ws.Range("K2:K" & ws.Cells(ws.Rows.count, 11).End(xlUp).Row)
rng1.NumberFormat = "0.00%"
Set rng2 = ws.Range("L2:L" & ws.Cells(ws.Rows.count, 12).End(xlUp).Row)
rng2.NumberFormat = ""
Range("Q2:Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = ""

'To calculate Quaterly change, Percentage Change and Total stock volume
For i = l To ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        total = total + ws.Cells(i, 7).Value
    Else
        total = total + ws.Cells(i, 7).Value
        opening = ws.Cells(j, 3).Value
        closing = ws.Cells(i, 6).Value
        quatchange = closing - opening
        quatPercentage = quatchange / opening
        ws.Cells(k + 1, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(k + 1, 10).Value = quatchange
            If ws.Cells(k + 1, 10).Value < 0 Then
            ws.Cells(k + 1, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(k + 1, 10).Value > 0 Then
            ws.Cells(k + 1, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(k + 1, 10).Interior.ColorIndex = 0
            End If
         ws.Cells(k + 1, 11).Value = quatPercentage
         ws.Cells(k + 1, 12).Value = total
         quatchange = 0
         quatPercentage = 0
         k = k + 1
         j = i + 1
        total = 0
    End If
        l = i
Next i


'Calculating Greatest % increase, Greatest % Decrease and Greatest Volume
For col1 = 2 To ws.Cells(Rows.count, 11).End(xlUp).Row
   
        If (ws.Cells(col1, 11).Value > ws.Cells(col1 + 1, 11).Value And Max < ws.Cells(col1 + 1, 11).Value) Then

        Max = ws.Cells(col1, 11).Value
         ws.Cells(2, 16).Value = ws.Cells(col1, 9).Value
        ElseIf (ws.Cells(col1 + 1, 11).Value > ws.Cells(col1, 11).Value And Max < ws.Cells(col1 + 1, 11).Value) Then
         ws.Cells(2, 16).Value = ws.Cells(col1 + 1, 9).Value
        Max = ws.Cells(col1 + 1, 11).Value
        Else
         ws.Cells(2, 16).Value = ws.Cells(2, 16).Value
        Max = Max
        End If

        If (ws.Cells(col1, 11).Value < ws.Cells(col1 + 1, 11).Value And Min > ws.Cells(col1, 11).Value) Then

        Min = ws.Cells(col1, 11).Value
           ws.Cells(3, 16).Value = ws.Cells(col1, 9).Value
        ElseIf (ws.Cells(col1 + 1, 11).Value < ws.Cells(col1, 11).Value And Min > ws.Cells(col1 + 1, 11).Value) Then

        Min = ws.Cells(col1 + 1, 11).Value
           ws.Cells(3, 16).Value = ws.Cells(col1 + 1, 9).Value
        Else
        Min = Min
           ws.Cells(3, 16).Value = ws.Cells(3, 16).Value
        End If
        
        If (ws.Cells(col1, 12).Value > ws.Cells(col1 + 1, 12).Value And Stock < ws.Cells(col1, 12).Value) Then

        Stock = ws.Cells(col1, 12).Value
          ws.Cells(4, 16).Value = ws.Cells(col1, 9).Value
        ElseIf (ws.Cells(col1 + 1, 12).Value > ws.Cells(col1, 12).Value And Stock < ws.Cells(col1 + 1, 12).Value) Then
          ws.Cells(4, 16).Value = ws.Cells(col1 + 1, 9).Value
        Stock = ws.Cells(col1 + 1, 12).Value
        Else
        Stock = Stock
          ws.Cells(4, 16).Value = ws.Cells(4, 16).Value
        End If


 Next col1
    ws.Cells(2, 17).Value = Max
    ws.Cells(3, 17).Value = Min
    ws.Cells(4, 17).Value = Stock
                   
Next ws
End Sub

