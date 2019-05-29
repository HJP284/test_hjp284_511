Attribute VB_Name = "Module1"
Option Explicit
Sub stocks()

' Set up variables


Dim ws As Worksheet

Dim Ticker As String

Dim Total_Stock_Volume As Double

Dim Summary_Table_Row As Integer

Dim LastRow As Double

Dim i As Double

' Create for loop

For Each ws In Worksheets
ws.Activate
    
    Total_Stock_Volume = 0

    Summary_Table_Row = 2


    ws.Range("i1") = "Ticker"
    ws.Range("j1") = "Total Stock Volume"

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For i = 2 To LastRow
    
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                Ticker = Cells(i, 1).Value
            
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
                ws.Range("i" & Summary_Table_Row).Value = Ticker
            
                ws.Range("j" & Summary_Table_Row).Value = Total_Stock_Volume
            
                Summary_Table_Row = Summary_Table_Row + 1
            
                Total_Stock_Volume = 0
    
            Else
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
            End If

        Next i
                  
    Next ws
    
End Sub

