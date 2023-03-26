Attribute VB_Name = "Module1"
Sub Stock()
Attribute Stock.VB_ProcData.VB_Invoke_Func = " \n14"

Dim LastRow As Long         ' last row of worksheet
Dim Ticker As String        ' ticker name
Dim OpenPrice As Double     ' first open price of year
Dim ClosePrice As Double    ' last close price of year
Dim Vol As LongLong         ' total volume of year Source: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary
Dim Counter As Long         ' counter to determine how many times each ticker appears (for loop)
Dim TckrCount As Long       ' counter for number of tickers

Dim GInc As Double          ' greatest % increase
Dim GDec As Double          ' greatest % decrease
Dim GTotVol As LongLong     ' greatest total volume

Dim GIncTckr As String      ' greatest % increase ticker
Dim GDecTckr As String      ' greatest % decrease ticker
Dim GTotVolTckr As String   ' greatest total volume ticker

Dim i, j, z As Long
    

For Each ws In Worksheets
    
    TckrCount = 2
    z = 2
    GInc = 0
    GDec = 0
    GTotVol = 0
        
    ' Determine total rows per worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
    ' Add header for data
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Columns("J:Q").AutoFit 'Source: https://learn.microsoft.com/en-us/office/vba/api/Excel.Range.AutoFit
    
    ' Sort by ticker and date to ensure data is organized (Source: macro recorded with excel and code generated was adapted for challenge)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range( _
        "B2:B" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ws.Sort
        .SetRange Range("A1:G" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Find Open Price, Close Price for Each Ticker
    For i = 2 To LastRow

        ' Get Open Price
        OpenPrice = ws.Cells(i, 3).Value
        
        ' Ticker
        Ticker = ws.Cells(i, 1).Value
                
        ' Count number of times ticker appears
        Counter = WorksheetFunction.CountIf(ws.Range("A" & i & ":A" & LastRow), Ticker)  ' Source: https://www.wallstreetmojo.com/vba-countif/
    
        ' Find Close Price for Ticker and total volume
        For j = 2 To Counter + 1
                
            If ws.Cells(z + 1, 1) <> ws.Cells(z, 1) Then
                ClosePrice = ws.Cells(z, 6).Value
                Vol = Vol + ws.Cells(z, 7).Value
                z = z + 1
            Else
                Vol = Vol + ws.Cells(z, 7).Value
                z = z + 1
            End If
                
        Next j
            
        ' Skip ahead and retain row place for close price and total volume values
        i = ((TckrCount - 1) * Counter) + 1

    ' Write data for current ticker
    ws.Cells(TckrCount, 9).Value = Ticker
    ws.Cells(TckrCount, 10).Value = ClosePrice - OpenPrice
    If ws.Cells(TckrCount, 10).Value < 0 Then
        ws.Cells(TckrCount, 10).Interior.ColorIndex = 3
    Else
        ws.Cells(TckrCount, 10).Interior.ColorIndex = 4
    End If
    ws.Cells(TckrCount, 11).Value = (ClosePrice - OpenPrice) / OpenPrice
    If ws.Cells(TckrCount, 11).Value < 0 Then
        ws.Cells(TckrCount, 11).Interior.ColorIndex = 3
    Else
        ws.Cells(TckrCount, 11).Interior.ColorIndex = 4
    End If
    ws.Cells(TckrCount, 12).Value = Vol
    
    ' Calcualte Greatest Increase, Greatest Decrease, and Greatest Total Volume
    If ws.Cells(TckrCount, 11).Value >= GInc Then
        GInc = ws.Cells(TckrCount, 11).Value
        GIncTckr = Ticker
    ElseIf ws.Cells(TckrCount, 11).Value <= GDec Then
        GDec = ws.Cells(TckrCount, 11).Value
        GDecTckr = Ticker
    End If
    
    If Vol > GTotVol Then
        GTotVol = Vol
        GTotVolTckr = Ticker
    End If

    'Reste Variables for next ticker
    ClosePrice = 0
    OpenPrice = 0
    Vol = 0
    Counter = 0
    Ticker = ""
    TckrCount = TckrCount + 1

    Next i
    
'Formatting
ws.Range("J2:J" & TckrCount - 1).NumberFormat = "$#,##0.00"
ws.Range("K2:K" & TckrCount - 1).NumberFormat = "0.00%"
ws.Range("L2:L" & TckrCount - 1).NumberFormat = "#,###"

'Write Greatest Increase, Greatest Decrease, and Greatest Total Volume
ws.Range("P2").Value = GIncTckr
ws.Range("Q2").Value = GInc

ws.Range("P3").Value = GDecTckr
ws.Range("Q3").Value = GDec

ws.Range("P4").Value = GTotVolTckr
ws.Range("Q4").Value = GTotVol

ws.Range("P1:Q1").Columns.AutoFit
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "#,###"


LastRow = 0

Next

End Sub
