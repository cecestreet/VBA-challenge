Attribute VB_Name = "Module1"
'All of my code has stopped working saying "external
Sub FunnyBusiness()

'declare to the computer the variables you will be referencing. ws means the Worksheet you are on
'cell is the Range, aka where you want the function to start and end
'apparently you don't have access to the scripting library function thingy so the alternative is making a collection object to keep unique values
'this will be shortened to safe when referencing this memory
'the key for a variant datatype will take on the value of what it is being iterated through as it goes, like a waterfall. If the key is stored as
'Variant it allows the value to be pretty much anything and it will still work, it can "vary". The key value will chameleon with each loop.
'Since we are working with just letters here though, I'm just gonna use Str - oh jk it won't let me. I guess you have to use Variant for collections.

    Dim ws As Worksheet
    Dim cell As Range
    Dim safe As New Collection
    Dim key As Variant
    Dim tickerheader As Range
    Dim tickerRow As String
    
    'I want the unique tickers to populate under column J of this current sheet, and I want to make the heading top set as tickerheader.
     Set tickerheader = ActiveSheet.Columns("J")

    'I want it to draw data from each sheet available in the workbook, so For Each ws In the workbook,
        'And then for every cell under the first column, A, or for each cell in the worksheet range of A1, A aka "row" A1, column A, and then every other
        'row in the ws, ws.Cells(ws.Rows.Count,"A") because also column A. Guess you could also do row and col but we already defined cell by range, not cell
        'this For Each loop will iterate through the cells in column A from A2 to the last non-empty cell in column A,
        '
        
     For Each ws In ThisWorkbook.Sheets
     'so we are using cell lingo, telling it to go through each cell range A2:A___ but we have to combine the row answer into here,
     'so we use an & to merge it, counting all of the rows for A, by counting up from the bottom and stopping once it hits a value.
     
        For Each cell In ws.Range("A2, A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        'if the cell is not empty of content
        
            If Not IsEmpty(cell.Value) Then

            'then, add the cell's value into your safe.
           safe.Add cell.Value
    
        End If
    Next cell
Next ws
'Okay, now I want a specific row, my ticker row to change increasing by 1 each row number to fill with the unique key output from the safe
'and have to use range not "to" cause doing cells not row, col, w/range we are doing Column,Row format like J1, w range, but switch back when
'when we start with Cells, then we need to do row, col
'this will make sure tickerRow knows the Range it is allowed to populate
tickerRow = tickerheader.Range("J" & tickerheader.Rows.Count).End(xlUp).Row + 1 'a 1 is added to the last row number that a value was found in

For Each key In safe
    tickerheader.Cells(tickerRow, 10).Value = key
Next key

End Sub

        

'Yearly change from the opening price at the beginning of a
'Okie dokie, trying to do the next three all in one go kept freezing my excel and it was not working for me, so I am breaking it up into smaller steps, the most useful
'Thing I have learned from this challenge.
'Beginning with iterating through the data site to find the sum of open prices for each unique ticker listed in column J.

'Defining/Declaring time. Still need to dim ws as worksheet so it knows range, then declare the unique tickers as range, as well as the opening price
'Column cells as range, and then finally, where you want to display the totals as a range. We are working with three main ranges.
'Sum’s should be as double because we don’t want to lose any decimal places in calculations
Sub OpenPriceSums()
    Dim ws As Worksheet
    Dim ticker As Range
    Dim openPrice As Range
    Dim OpenTotal As Range
    Dim OpenSum As Double

    ' Set OpenTotal to column K range
    Set OpenTotal = ActiveSheet.Columns("K")

'Go through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
'For each unique marker the corresponding sum starts at 0
        For Each ticker In ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
            ' need the sum to start out at 0, it won’t do it automatically unless you say it

            totalSum = 0

'For each opening price in the range of column C, if it is not empty and if the value of that row’s A column equals the string letter value from our ticker column, the cells for J, then add the value to our opening sum for that ticker. Aka, add it to the column next to the proper ticker in Cell J.

            For Each openPrice In ws.Range("C2", ws.Cells(ws.Rows.Count, "C").End(xlUp))
                If Not IsEmpty(openPrice.Value) And ws.Cells(openPrice.Row, "A").Value = ticker.Value Then
                    OpenSum = OpenSum + openPrice.Value
                End If
            Next openPrice
            OpenTotal.Cells(ticker.Row, 1).Value = OpenSum
        Next ticker
    Next ws
End Sub

'Now same thing but with closing prices! Woo hoo!
Sub ClosePriceSums()
    Dim ws As Worksheet
    Dim ticker As Range
    Dim closePrice As Range
    Dim CloseTotal As Range
    Dim CloseSum As Double

    Set CloseTotal = ActiveSheet.Columns("L")

    For Each ws In ThisWorkbook.Sheets
        For Each ticker In ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
            totalSum = 0

For Each closePrice In ws.Range("F2", ws.Cells(ws.Rows.Count, "F").End(xlUp))
    If Not IsEmpty(ticker.Value) And ws.Cells(closePrice.Row, 1).Value = ticker.Value Then
        CloseSum = CloseSum + closePrice.Value
    End If
Next closePrice

            CloseTotal.Cells(ticker.Row, 1).Value = CloseSum
        Next ticker
    Next ws
End Sub


Sub YearlyChange()
    Dim ws As Worksheet
    Dim ticker As Range
    Dim openPrice As Range
    Dim closePrice As Range
    Dim YearlyChangeColumn As Range
    Dim YearlyChange As Double

'Going to be populating column M of this sheet with the yearly change values, this is one of our ranges
    Set YearlyChangeColumn = ActiveSheet.Columns("M")

'Need to do this for every sheet in the workbook, for each ticker in column J, we will use the open price drawn from the open price totals for that ticker ‘row from column K
'same for closing price
'Then yearly change is close - open.
'Then store under our yearly change column, same row, WHY IS IT TICKER.ROW, 1, NOT TICKER.ROW, M
    For Each ws In ThisWorkbook.Sheets
        For Each ticker In ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
            Set openPrice = ws.Cells(ticker.Row, "K")

            Set closePrice = ws.Cells(ticker.Row, "L")

            YearlyChange = closePrice.Value - openPrice.Value

'It could be ticker.Row, M instead of ticker.Row, 1, but apparently ticker.Row, 1 means same row as the ticker, and “first column” of the ‘YearlyChangeColumn, there is only one Yearly Change Column so it’s kinda silly, but it still works and would provide more flexibility if we needed multiple columns of deciphering in the yearly change chart

            YearlyChangeColumn.Cells(ticker.Row, 1).Value = YearlyChange

        Next ticker
    Next ws
End Sub


'Okie dokie so when I tried this for the whole workbook instead of just sheet A it froze my excel, and I don’t really need to run it for the whole
'Workbook because my data for this particular step is all on sheet A, and it could have been getting confused by all of the empty cells

Sub yearlyChangePercentageForSheetA()
    Dim ws As Worksheet
    Dim ticker As Range
    Dim openPrice As Range
    Dim closePrice As Range
    Dim yearlyChangePercentageColumn As Range
'Percentage column definitely needs to be stored as a double
    Dim yearlyChangePercentage As Double

    Set ws = ThisWorkbook.Sheets("A")

    Set yearlyChangePercentageColumn = ws.Columns("N")

    For Each ticker In ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
        Set openPrice = ws.Cells(ticker.Row, "K")

        Set closePrice = ws.Cells(ticker.Row, "L")

'If the opening price does not equal 0, because we do not want to divide by a zero and get funky calculations, then do the yearly change percentage
‘calculation
'If it does equal zero, then put zero

        If openPrice.Value <> 0 Then
            yearlyChangePercentage = (closePrice.Value - openPrice.Value) / openPrice.Value * 100
        Else
            yearlyChangePercentage = 0
        End If

'Populate all of these in column N, the first, and only column for the yearly change percentage, the same row as the corresponding ticker value
        yearlyChangePercentageColumn.Cells(ticker.Row, 1).Value = yearlyChangePercentage
    Next ticker
End Sub


'Now on to the total stock volume
Sub TSVolume()
    Dim ws As Worksheet
    Dim ticker As Range
    Dim volume As Range
     Dim TSVolumeColumn As Range
    Dim totalSum As Double

    Set TSVolumeColumn = ActiveSheet.Columns("O")

    For Each ws In ThisWorkbook.Sheets
        For Each ticker In ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
‘Counter
            totalSum = 0

            For Each volume In ws.Range("G2", ws.Cells(ws.Rows.Count, "G").End(xlUp))
                If Not IsEmpty(volume.Value) And ws.Cells(volume.Row, "A").Value = ticker.Value Then
                    totalSum = totalSum + volume.Value
                End If
            Next volume

            TSVolumeColumn.Cells(ticker.Row, 1).Value = totalSum
        Next ticker
    Next ws
End Sub


Sub GreatestPercentIncrease()
    Dim ws As Worksheet
    Dim YC As Range
    Dim maxPercentIncrease As Double
    Dim maxTicker As String
    Dim percentChange As Range
   Dim tickers As Range

    Set percentChange = ActiveSheet.Columns("N")

    Set tickers = ActiveSheet.Columns("J")

'Counter the variables that are iterating, unsure if I need to counter the ticker too since it is not adding
    maxPercentIncrease = 0
    maxTicker = ""

'Go through workbook
    For Each ws In ThisWorkbook.Sheets
'Just made it absolute to save time
        For Each YC In ws.Range("N2:N1321")
            Dim percentChangeValue As Double
            percentChangeValue = YC.Value

'Update as you go
            If percentChangeValue > maxPercentIncrease Then
                maxPercentIncrease = percentChangeValue

                maxTicker = tickers.Cells(YC.Row, 1).Value
            End If
        Next YC
    Next ws

    Range("R2").Value = maxTicker
    Range("S2").Value = maxPercentIncrease
End Sub

Sub GreatestPercentDecrease()
    Dim ws As Worksheet
    Dim YC As Range
    Dim GPercentDecrease As Double
     Dim MinTicker As String
    Dim percentChange As Range
    Dim tickers As Range

    Set percentChange = ActiveSheet.Columns("N")
    Set tickers = ActiveSheet.Columns("J")

    GPercentDecrease = 0
    MinTicker = ""

    For Each ws In ThisWorkbook.Sheets
        For Each YC In ws.Range("N2:N1321")
            Dim percentChangeValue As Double
            percentChangeValue = YC.Value

            If percentChangeValue < GPercentDecrease Then
                GPercentDecrease = percentChangeValue
                MinTicker = tickers.Cells(YC.Row, 1).Value
            End If
        Next YC
    Next ws

    Range("R3").Value = MinTicker
    Range("S3").Value = GPercentDecrease
    
End Sub


Sub PopulateGreatestVolumeValues()
    Dim ws As Worksheet
    Dim stockvol As Range
    Dim tick As Range
    Dim maxVolume As Double
    Dim maxTicker As String
    Dim TSVolume As Range
     Dim tickers As Range

    Set TSVolume = ActiveSheet.Columns("O")

    Set tickers = ActiveSheet.Columns("J")

    maxVolume = 0
    maxTicker = ""

    For Each ws In ThisWorkbook.Sheets
        For Each stockvol In ws.Range("O2", ws.Cells(ws.Rows.Count, "O").End(xlUp))
            Dim volumeValue As Double
            volumeValue = stockvol.Value

'if the stock value, which then becomes the volumeValue, is bigger than the maxVolume(which starts as ‘zero), then the maxVolume becomes the volumeValue and continues from there. The next value will have ‘to be bigger in order to override it.
            If volumeValue > maxVolume Then
                maxVolume = volumeValue

'Get the ticker associated with it, the value in the tickers column, which we have defined as column J, the ‘first and only column of that set, row=same row as stockvol. It’s within the loop though so it will keep ‘changing according to the previous.
                Set tick = tickers.Cells(stockvol.Row, 1)
                maxTicker = tick.Value
            End If
        Next stockvol
    Next ws

    ' Populate the results
    Range("R4").Value = maxTicker
    Range("S4").Value = maxVolume
End Sub


Sub Test()
MsgBox ("some of my code has stopped working referencing an external environment as I have changed the placement of routines and made a complete mess")
End Sub




