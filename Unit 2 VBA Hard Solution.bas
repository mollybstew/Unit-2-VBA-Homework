Attribute VB_Name = "Module1"
Sub Hard_Stock()
    
' Loop through all the worksheet
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
        
'Create Variable to hold Value
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim Ticker As String
Dim PercentChange As Double
Dim Volume As Double
Volume = 0
Dim SummaryTableRow As Double
SummaryTableRow = 2
Dim i As Long
        
' Determine the Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

' Add Heading for summary
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Set Greatest % Increase, % Decrease, and Total Volume
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
       
'Set Initial Open Price
OpenPrice = Cells(2, 3).Value
         
' Loop through all ticker symbol
For i = 2 To LastRow
    ' Check if we are still within the same ticker symbol.
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Set Ticker name and Close Price
    Ticker = Cells(i, 1).Value
    Cells(SummaryTableRow, 9).Value = Ticker
    ClosePrice = Cells(i, 6).Value
                
    ' Add Yearly Change
    YearlyChange = ClosePrice - OpenPrice
    Cells(SummaryTableRow, 10).Value = YearlyChange
    Cells(SummaryTableRow, 10).NumberFormat = "0.000000000"
    
    ' Add Percent Change
    If (OpenPrice = 0 And ClosePrice = 0) Then
        PercentChange = 0
    ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
        PercentChange = 1
    Else
        PercentChange = YearlyChange / OpenPrice
        Cells(SummaryTableRow, 11).Value = PercentChange
        Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
    End If
    
    ' Add Total Volumn
    Volume = Volume + Cells(i, 7).Value
    Cells(SummaryTableRow, 12).Value = Volume
                
    ' Add one to the summary table row
    SummaryTableRow = SummaryTableRow + 1
    
    ' reset the Open Price and the Volumn Total
    OpenPrice = Cells(i + 1, 3)
    Volume = 0
    
    'if cells are the same ticker
    Else
    Volume = Volume + Cells(i, 7).Value
    End If
        
    Next i
        
    ' Determine the Last Row of Yearly Change per WS
    YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row

    ' Set the Cell Colors
        
    For j = 2 To YCLastRow
        If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
        Cells(j, 10).Interior.Color = vbGreen
            
        ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.Color = vbRed
            
        End If
        
    Next j
    
' Look through each rows to find the greatest value and its associate ticker
For Z = 2 To YCLastRow
    If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
    Cells(2, 16).Value = Cells(Z, 9).Value
    Cells(2, 17).Value = Cells(Z, 11).Value
    Cells(2, 17).NumberFormat = "0.00%"
        
    ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
    Cells(3, 16).Value = Cells(Z, 9).Value
    Cells(3, 17).Value = Cells(Z, 11).Value
    Cells(3, 17).NumberFormat = "0.00%"
            
    ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
    Cells(4, 16).Value = Cells(Z, 9).Value
    Cells(4, 17).Value = Cells(Z, 12).Value
        
    End If
        
Next Z
    
Next WS
        
End Sub




