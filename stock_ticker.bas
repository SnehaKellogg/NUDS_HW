Attribute VB_Name = "Module1"
Option Explicit

Sub WorksheetLoop()
'Define all variables used in this sub
Dim lastrow As Long
Dim i As Long
Dim j As Long
Dim y As Long
Dim years As Integer
Dim stockvolume As Double
Dim uniquetickers As Long
Dim n As Long 'to mark the worksheet
         
'Loop through all of the worksheets in the active workbook.
    For n = 1 To ThisWorkbook.Worksheets.Count
        MsgBox ActiveWorkbook.Worksheets(n).Name

        ' Calculate the end of the worksheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
                     
        'Advance filter to list unique tickers
        Range("A1:A" & lastrow).AdvancedFilter xlFilterCopy, CopyToRange:=Range("I1"), unique:=True
        
         'Define number of unique tickers in the worksheet
        uniquetickers = Cells(Rows.Count, 9).End(xlUp).Row
         'Give headers to summary table
        Cells(1, 9).Value = "Unique Tickers"
        Cells(1, 10).Value = "Total Volume"
        
        'Extract year from dates- excluding loop since this already is segregated by year
            'For i = 1 To lastrow
                'Cells(1, 10).Value = Left(Cells(2, 2), 4)
           ' Next i
                           
            'Set stockvolume for initial
                stockvolume = 0
                
            'Find price total for ticker in column I using for loop - runs for j (tickers) , y (years) and all the rows of initial data as i for loop
            For j = 2 To uniquetickers
                    For i = 2 To lastrow
                        If Cells(i, 1).Value = Cells(j, 9).Value Then
                            stockvolume = stockvolume + Cells(i, 7).Value
                         End If
                    Next i
                    Cells(j, 10).Value = stockvolume
                    stockvolume = 0
            Next j
    Next n

End Sub
