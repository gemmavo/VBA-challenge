
Sub RunMultipleSheets()

        
        Dim WS As Worksheet
        Application.ScreenUpdating = False
        For Each WS In Worksheets
            WS.Select
            Call WallStreetStock_data
       Next
       Application.ScreenUpdating = True
End Sub

  Sub WallStreetStock_data()
  
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"

        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim Volume As Double
        
'Set initial volume
        Volume = 0
        
        Dim R As Long
        R = 2

'Set first ticker's open price
        OpenPrice = Cells(2, "C").Value
        
        For I = 2 To LastRow
            If Cells(I + 1, "A").Value <> Cells(I, "A").Value Then
                Ticker = Cells(I, "A").Value
                Cells(R, "I").Value = Ticker

                ClosePrice = Cells(I, "F").Value
                YearlyChange = ClosePrice - OpenPrice
                Cells(R, "J").Value = YearlyChange
                
                If YearlyChange >= 0 Then
                    Cells(R, "J").Interior.ColorIndex = 4
                Else
                    Cells(R, "J").Interior.ColorIndex = 3
                End If

                If OpenPrice = 0 And ClosePrice = 0 Then
                    PercentChange = 0
                ElseIf OpenPrice = 0 And ClosePrice <> 0 Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If

'Remember to format PercentChange as Percentage
                Cells(R, "K").Value = PercentChange
                Cells(R, "K").NumberFormat = "0.00%"

'Add on volume
                Volume = Volume + Cells(I, "G").Value
                Cells(R, "L").Value = Volume

'Add on to next row of summary table, set new open price, and reset volume
                R = R + 1
                OpenPrice = Cells(I + 1, "C").Value
                Volume = 0
            Else
                Volume = Volume + Cells(I, "G").Value

            End If

        Next I
        
        Cells(2, "O").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Cells(4, "O").Value = "Greatest Total Volume"
        Cells(1, "P").Value = "Ticker"
        Cells(1, "Q").Value = "Value"
        
        Dim LastSumRow As Long
        LastSumRow = Cells(Rows.Count, "I").End(xlUp).Row

        Dim GreatestPercentInc As Double
        Dim GreatestPercentDec As Double
        Dim GreatestVolume As Currency
        
        GreatestPercentInc = Application.WorksheetFunction.Max(Columns("K"))
        GreatestPercentDec = Application.WorksheetFunction.Min(Columns("K"))
        GreatestVolume = Application.WorksheetFunction.Max(Columns("L"))

        For m = 2 To LastSumRow
            If Cells(m, "K").Value = GreatestPercentInc Then
                Cells(2, "Q").Value = Cells(m, "K").Value
                Cells(2, "Q").NumberFormat = "0.00%"
                Cells(2, "P").Value = Cells(m, "I").Value
            ElseIf Cells(m, "K").Value = GreatestPercentDec Then
                Cells(3, "Q").Value = Cells(m, "K").Value
                Cells(3, "Q").NumberFormat = "0.00%"
                Cells(3, "P").Value = Cells(m, "I").Value
            ElseIf Cells(m, "L").Value = GreatestVolume Then
                Cells(4, "Q").Value = Cells(m, "L").Value
                Cells(4, "P").Value = Cells(m, "I").Value

            End If
        Next m
        
        
End Sub
