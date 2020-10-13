Attribute VB_Name = "stockAnalysis_All_Sheets"
Sub stocksAnalysis():

sheetCount = ActiveWorkbook.Worksheets.Count

For a = 1 To sheetCount
    ActiveWorkbook.Worksheets(a).Activate
  
  'Headers
    Range("i1").Value = "Ticker Symbol"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Variables for Ticker Symbol
    Dim tickerSymbol As String
    Dim lastRow As Long
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim stockCounter As Long
        stockCounter = 2
    
    'Formating issues, change Column B to Date, Change Columns C to G to Currency
    Range("C1:F" & lastRow).Style = "Currency"
    Range("b1:B" & lastRow).NumberFormat = "####-##-##"
    
    Dim columnHeader As String
    Dim remove1 As String
    Dim remove2 As String
    remove1 = "<"
    remove2 = ">"
    For i = 1 To 7
        columnHeader = Cells(1, i).Value
        columnHeader = Replace(columnHeader, remove1, "")
        columnHeader = Replace(columnHeader, remove2, "")
        Cells(1, i).Value = columnHeader
        Cells(1, i).Value = StrConv(Cells(1, i).Value, vbProperCase)
        Next i
    
    'Variables for Open and Close
    Dim yearOpen As Double
    Dim yearClose As Double
    
    yearOpen = Range("C2").Value
    
    'variables for total stock volume
    Dim totalStockVolume As Double
    totalStockVolume = 0
    
    
            
     'For loop and if statement for ticker Symbol, Yearly Change, and Percent Change ca
    For i = 2 To lastRow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
          
          'insert Stock symbol in column i
           tickerSymbol = Cells(i, 1).Value
            Range("i" & stockCounter).Value = tickerSymbol
            
           'take open price from and close price from row i value, subtract, insert in column J, update open price to i+1
           yearClose = Cells(i, 6).Value
           Range("J" & stockCounter).Value = yearClose - yearOpen
           
                If Range("J" & stockCounter) < 0 Then
                    Range("J" & stockCounter).Interior.ColorIndex = 3
                    
                    Else: Range("J" & stockCounter).Interior.ColorIndex = 4
                    End If
                
            If yearOpen = 0 Then
                Range("K" & stockCounter).Value = 0
                Else: Range("K" & stockCounter).Value = ((yearClose - yearOpen) / yearOpen)
                End If
                
          
                      
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
           Range("L" & stockCounter).Value = totalStockVolume
                      
           yearOpen = Cells(i + 1, 3).Value
            stockCounter = stockCounter + 1
            totalStockVolume = 0
        
        Else
            totalStockVolume = totalStockVolume + Range("G" & i).Value

        End If
            
        
    Next i

'Challenges

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Greatest % Increase
Range("O2").Value = "Greatest % Increase"
Dim greatestIncrease As Double
greatestIncrease = WorksheetFunction.Max(Range("K2:K" & lastRow))
Range("Q2").Value = greatestIncrease

    For i = 2 To lastRow
        If Range("K" & i).Value = greatestIncrease Then
            Range("P2").Value = Range("i" & i).Value
            Exit For
        End If
        Next i

'Greatest % decrease
Range("O3").Value = "Greatest % Decrease"
Dim greatestDecrease As Double
greatestDecrease = WorksheetFunction.Min(Range("K2:K" & lastRow))
Range("Q3").Value = greatestDecrease

    For i = 2 To lastRow
        If Range("K" & i).Value = greatestDecrease Then
            Range("P3").Value = Range("i" & i).Value
            Exit For
        End If
        Next i


'Greatest Total Volume
Range("O4").Value = "Greatest Total Volume"
Dim greatestTotalVolume As Single
greatestTotalVolume = WorksheetFunction.Max(Range("L2:L" & lastRow))
Range("Q4").Value = greatestTotalVolume
    For i = 2 To lastRow
        If Range("L" & i).Value = greatestTotalVolume Then
            Range("P4").Value = Range("i" & i).Value
            Exit For
        End If
        Next i


'adjut columns and formatting
Columns("A:Q").AutoFit
 Range("K2:K" & lastRow).Style = "Percent"
 Range("Q2").Style = "Percent"
 Range("Q3").Style = "Percent"
Columns("L").NumberFormat = "###,###,###,###,###"
Range("Q4").NumberFormat = "###,###,###,###,###"
Next a

MsgBox ("Stock Analysis Complete")

End Sub
