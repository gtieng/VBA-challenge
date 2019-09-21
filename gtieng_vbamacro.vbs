'BEGIN SUBROUTINE
Sub analyze()

'APPLY TO ALL WORKSHEETS
Dim ws as Worksheet
For Each ws In Worksheets
With ws

 'SET DECLARATIONS
    Dim symbol As String
    Dim ticker as String
    Dim yrChange As Double
    Dim yrChangeStart as Double
    Dim yrChangeEnd as Double
    Dim prChange As Double
    Dim totalVol As Double

    'SET STARTING VALUES
    summaryRow = 2 'sets pointer to print at second row
    totalVol = 0 'initial counter value
    lastRowA = .Cells(Rows.Count, 1).End(xlUp).Row 'calculates the total number of data rows in variable sheets
    

    'TABLE FORMATTING: ADD CELL LABELS FOR "INSTRUCTIONS"
    .Range("I1") = "Ticker"
    .Range("L1") = "Total Vol."
    .Range("K1") = "% Change"
    .Range("J1") = "Yr Change"

    'TABLE FORMATTING: ADD CELL LABELS FOR "CHALLENGE"
    .Range("P1") = "Ticker"
    .Range("Q1") = "Value"
    .Range("O2") = "Greatest % Increase"
    .Range("O3") = "Greatest % Decrease"
    .Range("O4") = "Greatest Total Volume"
        .Columns("O").ColumnWidth = 25

    'BEGIN "INSTRUCTIONS" LOOP
    For i = 2 to lastRowA

        'CONDITIONAL: BEFORE TICKER SYMBOLS ARE MATCHING
        If .Cells(i, 1).Value <> .Cells(i-1, 1).Value Then
            'SET THE YRCHANGE START VARIABLE
            yrChangeStart = .Cells(i, 3).Value

        'CONDITIONAL: AS TICKER SYMBOLS ARE MATCHING
        ElseIf .Cells(i, 1).Value = .Cells(i+1, 1).Value Then
            'ADD TO THE TICKERTOTAL
            totalVol = totalVol + .Cells(i, 7).Value

        'CONDITIONAL: WHEN TICKER SYMBOLS NO LONGER MATCHING
        ElseIF .Cells(i, 1).Value <> .Cells(i+1, 1).Value Then
            'SET THE YRCHANGEEND VARIABLE
            yrChangeEnd = .Cells(i,6).Value

            'SET THE STOCK SYMBOL
            symbol = .Cells(i,1).Value

            'PRINT THE SUMMARIES
            .Range("I" & summaryRow) = symbol
            .Range("J" & summaryRow) = yrChangeEnd-yrChangeStart
            
            '*CONDITIONAL: TO PROVENT DIVIDE BY ZERO*
                If yrChangeStart <> 0 Then
                .Range("K" & summaryRow) = .Range("J" & summaryRow)/yrChangeStart
                Else
                .Range("K" & summaryRow) = .Range("J" & summaryRow)/ 1
                End If

            .Range("L" & summaryRow) = totalVol

            'CONDITIONAL: HIGHLIGHT THE % COLUMN RED OR GREEN
            If .Range("K" & summaryRow) < 0 Then
                .Range("K" & summaryRow).Interior.ColorIndex = 3
            Else
                .Range("K" & summaryRow).Interior.ColorIndex = 4

            End If
        
            'MOVE THE SUMMARY ROW
            summaryRow = summaryRow + 1

            'RESET THE TICKER TOTAL
            totalVol=0

        End If

    Next i

    

    'BEGIN "CHALLENGE" LOOP
    lastRowB = .Cells(Rows.Count, 9).End(xlUp).Row 'calculates the total number of summary rows in variable sheets

    For j = 2 to lastRowB

        'SET THE TICKER SYMBOL
        ticker = .Cells(j, 9).Value

        'SEARCH FOR MAX & MIN %
        If .Cells(j, 11).Value = WorksheetFunction.Max(.Columns("K")) Then
            'PRINT ITS SYMBOL
            .Range("P2") = ticker
            '& PRINT ITS VALUE
            .Range("Q2") = WorksheetFunction.Max(.Columns("K"))

        ElseIF .Cells(j, 11).Value = WorksheetFunction.Min(.Columns("K")) Then
            'PRINT ITS SYMBOL
            .Range("P3") = ticker
            '& PRINT ITS VALUE
            .Range("Q3") = WorksheetFunction.Min(.Columns("K"))
        End If

        'SEARCH FOR MAX VOL
        If .Cells(j, 12).Value = WorksheetFunction.Max(.Columns("L")) Then
                'PRINT ITS SYMBOL
                .Range("P4") = ticker
                '& PRINT ITS VALUE
                .Range("Q4") = WorksheetFunction.Max(.Columns("L"))
        End If

    Next j

End With

Next ws

End Sub