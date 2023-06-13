Sub Stock_Data()
    'Create a worksheet loop that cycles through each sheet in the worksheet
    For Each ws In Worksheets
        'Defining Variables and creating a function to find the lastrow of stocks in each sheet
        Dim Total As Double
        Dim Summary_Row As Integer
        Dim Greatest_Volume As Double
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'Creating column headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        'Setting summary row to the beginning row of data
        'Setting Greatest_Volume equal to 0 in order to find the largest stock volume
        Summary_Row = 2
        Greatest_Volume = 0
        'This loops goes through every row with stock data in it
        For i = 2 To LastRow
            'This if statement looks at if the stock ticker is different than the one before it
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Stores stock ticker name and adds the last value of the stock to the total
                Ticker_name = ws.Cells(i, 1).Value
                Total = Total + ws.Cells(i, 7).Value
                'Stores the stock ticker name and stock volume in the I and L columns
                ws.Range("I" & Summary_Row).Value = Ticker_name
                ws.Range("L" & Summary_Row).Value = Total
                'Resets total to 0 so the next stock volume doesnt include the previous stocks
                Total = 0
                'Creating another if statement to see if the stock volume if > Greatest_VOlume
                If ws.Range("L" & Summary_Row).Value > Greatest_Volume Then
                    'If it is then it will become the new Greatest_Volume and stored in Cell(4, 17)
                    Greatest_Volume = ws.Range("L" & Summary_Row).Value
                    ws.Cells(4, 17) = Greatest_Volume
                    'The name that matches the stock volume will be displayed in the row before it
                    ws.Cells(4, 16) = ws.Range("I" & Summary_Row).Value
                End If
                'Adding a row the summary table
                Summary_Row = Summary_Row + 1
            Else
                'If the stock ticker names in the rows match add the stock volume of the 1st to the Total
                Total = Total + ws.Cells(i, 7).Value
            End If
        Next i
        'Defining variables for the beginning of the year and end of the year stock prices
        Dim open_price As Double
        Dim close_price As Double
        Dim stock_year_change As Double
        'Creating a 2nd summary row variable in order to not override the 1st summary row variable that is now at the lastrow of the previous Range
        Dim Summary_Row_2 As Integer
        'Defining two variables that will hold the Greatest Percentage Increase and Decrease of stock values
        Dim Greatest As Double
        Dim Least As Double
        Summary_Row_2 = 2
        Greatest = 0
        Least = 0
    
        'This for loop works the same as the 1st
        For i = 2 To LastRow
            'This statement looks at the last 4 numbers of the date column and if they match takes the opening price
            If Right(ws.Cells(i, 2), 4) = "0102" Then
                open_price = Cells(i, 3)
            'This one looks at the last 4 numbers of the date column and if they match returns the ticker and closing price
            'Then it calculates the percentage change and yearly change before storing the yearly change in Row J
            ElseIf Right(ws.Cells(i, 2), 4) = "1231" Then
                Name = ws.Cells(i, 1)
                close_price = ws.Cells(i, 6)
                stock_year_change = close_price - open_price
                percentage_change = stock_year_change / open_price
                ws.Range("J" & Summary_Row_2).Value = stock_year_change
                'This statement changes the interior to green if yearly change is > 0 and red if it's < 0
                If ws.Range("J" & Summary_Row_2).Value >= 0 Then
                    ws.Range("J" & Summary_Row_2).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Row_2).Interior.ColorIndex = 3
                End If
                'Store percentage change in Column K
                ws.Range("K" & Summary_Row_2).Value = percentage_change
                'Evaluates each value in column to see if it's > Greatest
                'If it is it becomes Greatest and the value is then stored in Cell(2,17)
                If ws.Range("K" & Summary_Row_2).Value > Greatest Then
                    Greatest = ws.Range("K" & Summary_Row_2).Value
                    ws.Cells(2, 17) = Greatest
                    ws.Cells(2, 16) = Name
                'Evaluates each value in column K to see if it's < Least
                'If it is it becomes Least and the value is then stored in Cell(3,17)
                ElseIf ws.Range("K" & Summary_Row_2).Value < Least Then
                    Least = ws.Range("K" & Summary_Row_2).Value
                    ws.Cells(3, 17) = Least
                    ws.Cells(3, 16) = Name
                End If
                'Add another row to summary row so the if statements above evaluate the next row down
                Summary_Row_2 = Summary_Row_2 + 1
            End If
        Next i
        'Change column K and Greatest and Least values to percentages
        'Autofit the columns so the headers and stock volume can be fully seen
        ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit
    Next ws

End Sub