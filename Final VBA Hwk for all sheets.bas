Sub stock()

'Declare Variables
Dim I As Long
Dim ticker As String
Dim volume As Double
Dim lastrow As Double
Dim Summary_Table_Row As Long
Dim change As Double
Dim percent As Double
Dim open_ As Double
Dim open2 As Double
Dim close_ As Double

Dim J As Long
Dim g As Long
Dim inc As Double
Dim dec As Double
Dim tot As Double
Dim lastsummaryrow As Long
Dim ws As Worksheet


For Each ws In Worksheets
       

    ' Set an initial variable for holding the total volume per ticker
    volume = 0

    ' Keep track of the location for each ticker in the summary table
    Summary_Table_Row = 2

    'find last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'loop through each ticker
    For I = 2 To lastrow
        'Get first open
        If I = 2 Then
            open_ = ws.Cells(I, 3).Value
            For J = 1 To lastrow
                If open_ = 0 Then
                    open_ = ws.Cells(I + J, 3).Value
                Else
                Exit For
                End If
            Next J
                
        ' Check if we are still within the same ticker, if it is not..
        ElseIf ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

         'get next open
        open2 = ws.Cells(I + 1, 3).Value
            For J = 1 To lastrow
                If I = lastrow Then
                Exit For
                ElseIf open2 = 0 And J <> lastrow Then
                open2 = ws.Cells(I + J, 3).Value
                Else
                Exit For
                End If
            Next J
        
        
        ' Set the ticker name
        ticker = ws.Cells(I, 1).Value

        ' Add to the volume total
        volume = volume + ws.Cells(I, 7).Value

        ' Print the ticker name in the Summary Table and title at top
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("I1").Value = "Ticker Name"
      

        ' Print the volume Amount to the Summary Table and volume total at top
        ws.Range("L" & Summary_Table_Row).Value = volume
        ws.Range("L1").Value = "Volume Total"
        'get close
        close_ = ws.Cells(I, 6).Value
        'get change
        If close_ > 0 Then
        change = close_ - open_
        Else
        End If
        
        'get percent change and add conditional for zero
            If open_ > 0 Then
            percent = (change / (open_))
            Else
            End If
        'put change into change column
            If change <> 0 Then
            ws.Range("J" & Summary_Table_Row).Value = change
            ws.Range("J1").Value = "Change"
            Else
            End If
        'change decimal type
        ws.Range("J" & Summary_Table_Row).NumberFormat = "0.000000000000000"
        'color code change
            If change < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
        
      
        'put percentage change into percent column and change column to percent style
        ws.Range("K" & Summary_Table_Row).Value = percent
       ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        ws.Range("K1").Value = "Percent Change"
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
      
        ' Reset the Volume
        volume = 0

        'get next open
        open_ = open2

            
        Else

        ' Add to the Volume
        volume = volume + ws.Cells(I, 7).Value

        End If

    Next I
    'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

    'Row Titles and changing numberformat in column 16
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    'wrap text
    ws.Cells(2, 14).WrapText = True
    ws.Cells(3, 14).WrapText = True
    ws.Cells(4, 14).WrapText = True

    'last row
    lastsummaryrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'declare variables
    inc = 0
    ticker = 0
    dec = 0
    tot = 0
    'For loop to check through greatest inc
    For g = 2 To lastsummaryrow
    'Greatest percent increase
       
        If ws.Cells(g, 11).Value > inc Then
            inc = ws.Cells(g, 11).Value
            ticker = ws.Cells(g, 9).Value
        Else
        End If
    Next g
        ws.Cells(2, 15).Value = ticker
        ws.Cells(2, 16).Value = inc
        ticker = 0
    'for loop for greatest decrease
    For g = 2 To lastsummaryrow
    'Greatest percent increase
       
        If ws.Cells(g, 11).Value < dec Then
            dec = ws.Cells(g, 11).Value
            ticker = ws.Cells(g, 9).Value
        Else
        End If
    Next g
        ws.Cells(3, 15).Value = ticker
        ws.Cells(3, 16).Value = dec
        ticker = 0
    'for loop for greatest volume
    For g = 2 To lastsummaryrow
    'Greatest percent increase
       
        If ws.Cells(g, 12).Value > tot Then
            tot = ws.Cells(g, 12).Value
            ticker = ws.Cells(g, 9).Value
        Else
        End If
    Next g
        ws.Cells(4, 15).Value = ticker
        ws.Cells(4, 16).Value = tot
        ticker = 0



Next ws

End Sub









