Sub stockCounter()
    
    'Dim wb As Workbook
    Dim ws As Worksheet
    'Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Dim i As Long
        Dim j As Integer
        Dim NextRow As Long
        Dim x, y, z As Long 'placeholder for open/close date
        Dim ticker As String 'Contains ticker symbol
        Dim strOpenDate As String 'String value for Open Date
        'dim openDate as Double 'Open date for symbol
        Dim strCloseDate As String 'String value for Close Date
        'dim lastDate as Double 'Last date for symbol
        Dim openPrice As Double 'Open price for symbol
        Dim closePrice As Double 'Close price for symbol
        Dim yearlyChng As Double 'Yearly Change of Price in stock symbol
        Dim pctChng As Double 'Percentage Change of price in stock symbol
        Dim TotVolume As Double 'Stores total volume
        Dim finalRow As Long 'find the last entry on worksheet
        'Dim HaltNow As Boolean 'Testing Loop
        finalRow = Cells(Rows.Count, 1).End(xlUp).Row

        ws.Range("I1").Resize(1, 1).Value = "Ticker"
        ws.Range("J1").Resize(1, 1).Value = "Yearly Change ($)"
        ws.Range("K1").Resize(1, 1).Value = "Percent Change (%)"
        ws.Range("L1").Resize(1, 1).Value = "Total Stock Volume"
        x = 0
        y = 0
        z = 2
        'Compile all of the columns indicated in part 1
        For i = 2 To finalRow
        'HaltNow = False
            'MsgBox (finalRow)
            ticker = Cells(i, "A").Value
            If (Cells(i, "A").Offset(-1, 0).Value <> ticker And Cells(i, "C").Value >= 0) Then
                x = i
                'MsgBox (x)
                ticker = Cells(x, "A").Value
                'MsgBox (ticker)
                strOpenDate = CStr(Cells(x, "B").Value)
                'MsgBox (strOpenDate)
                'openDate = dateserial(CInt(Left(strOpenDate,4)),CInt(Mid(strOpenDate,5,2)),CInt(Right(strOpenDate,2)))
                If (Cells(i, "C").Value = 0) Then
                    NextRow = i + 1
                    Do While (Cells(NextRow, "C").Value >= 0 And Cells(i, "A").Offset(-1, 0).Value = ticker)
                        If (Cells(NextRow, "C").Value > 0 And Cells(i, "A").Offset(-1, 0).Value = ticker) Then
                            openPrice = Cells(NextRow, "C").Value
                        ElseIf Cells(i, "A").Offset(1, 0).Value <> ticker Then
                            openPrice = Null
                            Exit Do
                        End If
                    Loop
                Else: openPrice = CDbl(Cells(x, "C").Value)
                End If
                'MsgBox (openPrice)
                TotVolume = CLng(Cells(x, "G").Value)
                'MsgBox (TotVolume)
            ElseIf (Cells(i, "A").Offset(0, 1).Value = ticker And Cells(i, "C").Value >= 0) Then
                TotVolume = TotVolume + CLng(Cells(i, "G").Value)
                'MsgBox (TotVolume)
            ElseIf (Cells(i, "A").Offset(1, 0).Value <> ticker) Then 'And (Right(Cells(i, "B").Value, 4) = 1231 Or Right(Cells(i, "B").Value, 4) = 1230)
                y = i
                'MsgBox (y)
                strCloseDate = CStr(Cells(y, "B").Value)
                'lastDate = dateserial(CInt(Left(strCloseDate,4)),CInt(Mid(strCloseDate,5,2)),CInt(Right(strCloseDate,2)))
                closePrice = CDbl(Cells(y, "F").Value)
                TotVolume = TotVolume + CLng(Cells(i, "G").Value)
                yearlyChng = closePrice - openPrice
                If (openPrice = Null Or openPrice = 0) Then
                    pctChng = FormatNumber(CDbl(0), 3)
                Else: pctChng = FormatNumber(CDbl(closePrice / openPrice), 3) 'CDbl
                End If
                'MsgBox (z)
                Cells(z, "I").Value = ticker
                Cells(z, "J").Value = yearlyChng
                Cells(z, "K").Value = pctChng
                Cells(z, "L").Value = TotVolume
                If Cells(z, "J").Value > 0 Then
                    Cells(z, "J").Resize(1, 1).Interior.ColorIndex = 4
                ElseIf Cells(z, "J").Value < 0 Then
                    Cells(z, "J").Resize(1, 1).Interior.ColorIndex = 3
                End If
                z = z + 1
                'Reset variables to null or '0'
                ticker = ""
                strOpenDate = ""
                strCloseDate = ""
                TotVolume = 0
                yearlyChng = 0
                pctChng = 0
                x = 0
                y = 0
            
            End If
            'MsgBox (i)
            
            'If i = 263 Then
                'HaltNow = True
             '   Exit For
            'End If
        Next i
        
        ws.Columns("I:L").AutoFit
    Next ws
End Sub
