Attribute VB_Name = "Module2"
Sub Stockmarkettotals()

    'select all worksheets
    For Each ws In Worksheets
    
    'set initial variables
    Dim Stock_Name As String
    Dim Total_stockvolume As Double
    Dim Summary_Table_Row As Long
    Dim LastRow As Long
    
    'keep track of location for each stock name in summary table
    Total_stockvolume = 0
    Summary_Table_Row = 2
        
    'Loop through all stock volume
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
    
    
    'Check if still within same stock name, if not
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
        'set stock name
        Stock_Name = Cells(i, 1).Value
        
        'Add to total per stock
        Total_stockvolume = Total_stockvolume + Cells(i, 7).Value
        
        'print stock name in summary table
        Range("I" & Summary_Table_Row).Value = Stock_Name
        
       'print total volume per stock in summary table
        Range("J" & Summary_Table_Row).Value = Total_stockvolume
        
        'add 1 to summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset stock total
        Total_stockvolume = 0
        
    'if following cell is same stock name
    Else
    
        'add to stock total
        Total_stockvolume = Total_stockvolume + Cells(i, 7).Value
    
    End If

Next i

Next ws

End Sub


