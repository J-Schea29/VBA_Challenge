VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock()

'Declare Variables
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim Ticker As String
Dim Chart_row As Integer
Dim Start As Long
Dim i As Long
Dim ws As Worksheet


    'In each sheet, add headers for charts
    For Each ws In ActiveWorkbook.Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'BONUS section chart headers
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest total volume"
     
     
        'Find last row for each sheet.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Set Chart_row equal to 2 upon starting each sheet.
        Chart_row = 2
        
        'Set initial total_volume to zero and Start to 2.
        total_volume = 0
        Start = 2
    
        'Loop through data till last row.
        For i = 2 To lastrow
                
                'If current Ticker is different then one below, perform actions.
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'When cell below is different take Ticker value and imput to Collum I
                    Ticker = ws.Cells(i, 1).Value
                    ws.Cells(Chart_row, 9).Value = Ticker
                    
                    'add final cell volume, then place in Collum L
                    total_volume = total_volume + ws.Cells(i, 7).Value
                    ws.Cells(Chart_row, 12).Value = total_volume
                    
                    'yearly_change equals the final closing value minus the intial opening value.
                    yearly_change = ws.Cells(i, 6).Value - ws.Cells(Start, 3)
                    'Place in Collum J.
                    ws.Cells(Chart_row, 10).Value = yearly_change
                    
                        'If the change is greater than zero make interior green.
                        If yearly_change > 0 Then
                            ws.Cells(Chart_row, 10).Interior.ColorIndex = 4
                        'If the change is less than zero make interior red.
                        ElseIf yearly_change < 0 Then
                            ws.Cells(Chart_row, 10).Interior.ColorIndex = 3
                        End If
                        
                    'Find percent_change by dividing the yearly_change by the intial opening value. Round to second decimal place.
                    If ws.Cells(Start, 3) <> 0 Then
                    percent_change = Round(yearly_change / ws.Cells(Start, 3), 2)
                    Else
                    End If
            
                    'Place percent_change in Collum K.
                    ws.Cells(Chart_row, 11).Value = percent_change
                    
                    'New Start value
                    Start = i + 1
                    'Next chart row
                    Chart_row = Chart_row + 1
                    'Reset total_volume
                    total_volume = 0
                    
                'If current Ticker is not different then one below, perform actions.
                Else
                    'Add total_volume
                    total_volume = total_volume + ws.Cells(i, 7).Value
                End If
            Next i
            
            'Change formating for Colomn K.
            ws.Range("K:K").NumberFormat = "0.00%"
            
            'BONUS SECTION
            'Set Variables for next chart
            Dim greatest_inc As Double
            Dim least_inc As Double
            Dim greatest_total As Double
            Dim Row_number As Long
            Dim Ticker_2 As String
            Dim rng_1 As Range
            Dim rng_2 As Range
            
            'Define rng_1 and rng_2
            Set rng_1 = ws.Range("K:K")
            Set rng_2 = ws.Range("L:L")
            
            'Find greatest percentage increase
            greatest_inc = WorksheetFunction.Max(rng_1)
            
            'Find greatest percentage decrease
            least_inc = WorksheetFunction.Min(rng_1)
            
            'Find greatest total
            greatest_total = WorksheetFunction.Max(rng_2)
         
            'Find what row the greatest percentage increase is in and use that value to find its Ticker
            Row_number = WorksheetFunction.Match(greatest_inc, ws.Range("K:K"), 0)
            Ticker_2 = ws.Cells(Row_number, 1).Value
            
            'Insert into chart the Ticker and greatest percentage increase
            ws.Cells(2, 15).Value = Ticker_2
            ws.Cells(2, 16).Value = greatest_inc
    
            'Find what row the greatest percentage decrease is in and use that value to find its Ticker
            Row_number = WorksheetFunction.Match(least_inc, ws.Range("K:K"), 0)
            Ticker_2 = ws.Cells(Row_number, 1).Value
            
            'Insert into chart the Ticker and greatest percentage decrease
            ws.Cells(3, 15).Value = Ticker_2
            ws.Cells(3, 16).Value = least_inc
    
            'Find what row the greatest total is in and use that value to find its Ticker
            Row_number = WorksheetFunction.Match(greatest_total, ws.Range("L:L"), 0)
            Ticker_2 = ws.Cells(Row_number, 1).Value
            
            'Insert into chart the ticker and greatest total
            ws.Cells(4, 15).Value = Ticker_2
            ws.Cells(4, 16).Value = greatest_total
    
    Next ws
            

End Sub

