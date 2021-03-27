Attribute VB_Name = "Module11"
Sub Ticker():

    ' set variable
    Dim Ticker As String
    Dim Total As Double
    Dim PercentChange As Double
    Dim RowCount As Long
    Dim i As Long
    Dim j As Long
    Dim change As Double
    
    'setting the title row
    Range("I1").Value = "ticker"
    Range("J1").Value = "yearly change"
    Range("K1").Value = "percent change"
    Range("L1").Value = "total stock volume"
    
    'create a variable to find number of last row with data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' setting initial values
    j = 1
    Total = 0
    Ticker = Cells(2, 1).Value
    Range("I2").Value = Ticker
    Start = 2
    
    'loop through all the the stocks for one year
    For i = 2 To RowCount
       
        'Check if the ticker changes and get results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Total = Total + Cells(i, 7).Value
            Ticker = Cells(i, 1).Value
            
            ' calculate change
            If Total = 0 Then
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = 0
            Range("L" & 2 + j).Value = 0
            
        Else
            If Cells(Start, 3) = 0 Then
                For findValue = Start To i
                    If Cells(findValue, 3).Value <> 0 Then
                        Start = findValue
                        Exit For
                    End If
                Next findValue
            End If
            
            'calculate change
            change = (Cells(i, 6) - Cells(Start, 3))
            PercentChange = Round((change / Cells(Start, 3)) * 100, 2)
            Start = i + 1
            j = j + 1
            Range("I" & j).Value = Ticker
            Range("J" & j).Value = Round(change, 2)
            Range("K" & j).Value = PercentChange & "%"
            Range("L" & j).Value = Total
            
            ' Highlight the positive changes in green and negative changes in red
                Select Case change
                    Case Is > 0
                        Range("J" & j + 1).Interior.ColorIndex = 4
                Case Is < 0
                        Range("J" & j + 1).Interior.ColorIndex = 3
                Case Else
                        Range("J" & j + 1).Interior.ColorIndex = 0
                End Select
        
            Total = 0
            
            End If
                
        Else
            Total = Total + Cells(i, 7).Value
    
            End If
       
    Next i
        ' We want to take the max and the min and seperate them in another place in the worksheet
        Range("Q2") = PercentChange & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
        Range("Q3") = PercentChange & WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & RowCount))
        
        ' Need to return one less due to the header row not being a factor
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
        volumeNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
End Sub
