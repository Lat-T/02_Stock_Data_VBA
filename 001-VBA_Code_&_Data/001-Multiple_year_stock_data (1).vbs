Sub WallStreetSummary1()
    
'------------------------
'Do it for each worksheet
    For Each ws In Worksheets
'-------------------------------
    
  
    '----------------------------
    '1st summary
    '----------------------------
        'Find the last row for the Ticker column (source data)
        Dim Source_Table_LastRow As LongLong
        Source_Table_LastRowTicker = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Find the the 1 first row to be poulated (summary data)
        Dim Summary_Table_Row As LongLong
        Summary_Table_Row = 2

        'Create headers for the summary columns
        Dim Header_Ticker As String
        Dim Header_YearlyChange As String
        Dim Header_PercentChange As String
        Dim Header_TotalStockVolume As String

        'Assign header names for the summary columns
        ws.Range("K" & Summary_Table_Row - 1) = "Ticker"
        ws.Range("L" & Summary_Table_Row - 1) = "Yearly Change"
        ws.Range("M" & Summary_Table_Row - 1) = "Percent Change"
        ws.Range("N" & Summary_Table_Row - 1) = " Total Stock Volume"
 
        'Define summary Columns to be retrieved
        Dim Summary_Ticker As String
        Dim Summary_YearlyChange As Double
        Dim Summary_PercentChange As Double
        Dim Summary_TotalStockVolume As LongLong
        Dim Summary_Greatest_Increase_Percent As Double
        Dim Summary_Greatest_Decrease_Percent As Double
        Dim Summary_Greatest_Total_Volume As LongLong
        Dim TickerRowTracker As LongLong
        
        'Assign default values
        Summary_Ticker = ""
        Summary_YearlyChange = 0
        Summary_PercentChange = 0
        Summary_TotalStockVolume = 0
        Summary_Greatest_Increase_Percent = 0
        Summary_Greatest_Decrease_Percent = 0
        Summary_Greatest_Total_Volume = 0
        
        'First row with data
        TickerRowTracker = 2
        
    
        'Create a loop to very the conditions
        For i = Summary_Table_Row To Source_Table_LastRowTicker
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                'Set Ticker
                Summary_Ticker = ws.Cells(i, 1).Value
            
                'Calculate volume
                Summary_TotalStockVolume = Summary_TotalStockVolume + ws.Cells(i, 7).Value
            
                'Calculate Yearly Change
                Summary_YearlyChange = ((ws.Cells(i, 6).Value) - (ws.Cells(TickerRowTracker, 3).Value))
            
                'Calculate Percent Change
                'If ws.Cells(i, 3) = 0 Then
                If Summary_YearlyChange = 0 Or ws.Cells(TickerRowTracker, 3).Value = 0 Then
                    Summary_PercentChange = 0
                    
                Else
                    Summary_PercentChange = (Summary_YearlyChange / ws.Cells(TickerRowTracker, 3).Value)
                
                End If
                    
                   
                'Summary_Ticker
                'Print
                ws.Range("K" & Summary_Table_Row) = Summary_Ticker
                ws.Range("N" & Summary_Table_Row) = Summary_TotalStockVolume
                ws.Range("L" & Summary_Table_Row) = Summary_YearlyChange
                ws.Range("M" & Summary_Table_Row) = Summary_PercentChange
             
                'Move to next row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'Reset Values
                Summary_TotalStockVolume = 0
                Summary_YearlyChange = 0
                Summary_PercentChange = 0
                TickerRowTracker = i + 1
            
            
            Else
                'Add Total Stock Volume to Summary
                Summary_TotalStockVolume = Summary_TotalStockVolume + ws.Cells(i, 7).Value
            
               
            End If
    
        
        Next i

         

    '-------------------------
    '2nd summary
    '-------------------------
        'Find the last row for the 1st summary that is used a data source for the 2nd summary
        Dim Source_Table2_LastRow As LongLong
        Source_Table2_LastRowTicker = ws.Cells(Rows.Count, "K").End(xlUp).Row

        'Find the the 1 first row to be poulated (summary data)
        Dim Summary_Table2_Row As LongLong
        Summary_Table2_Row = 2

        'Define output values
        Dim Summary_Max_Increase_Percent As Double
        Dim Summary_Max_Decrease_Percent As Double
        Dim Summary_Max_Total_Volume As LongLong
    

        ' Define headers
        ws.Range("R" & Summary_Table2_Row - 1) = "Ticker"
        ws.Range("Q" & Summary_Table2_Row) = " Greatest Increase Percent"
        ws.Range("Q" & Summary_Table2_Row + 1) = " Greatest Decrease Percent  "
        ws.Range("Q" & Summary_Table2_Row + 2) = " Greatest Total Volume"
        ws.Range("S" & Summary_Table2_Row - 1) = " Value"

        'Create a loop
        For i = Summary_Table2_Row To Source_Table2_LastRowTicker
        
            'Calculate the Minimum and Maximum for the summary
            Summary_Max_Increase_Percent = Application.WorksheetFunction.Max(ws.Range("M" & Summary_Table2_Row & ":M" & (Source_Table2_LastRowTicker)))
        
            Summary_Max_Decrease_Percent = Application.WorksheetFunction.Min(ws.Range("M" & Summary_Table2_Row & ":M" & (Source_Table2_LastRowTicker)))
            Summary_Max_Total_Volume = Application.WorksheetFunction.Max(ws.Range("N" & Summary_Table2_Row & ":N" & (Source_Table2_LastRowTicker)))
            '==> ws.Range("M" & Summary_Table2_Row & ":M" & (Source_Table2_LastRowTicker ))
            
            'Print values and %
            ws.Range("S2") = Summary_Max_Increase_Percent
            ws.Range("S3") = Summary_Max_Decrease_Percent
            ws.Range("S4") = Summary_Max_Total_Volume

        
        Next i
  
  
    '------------------------
    'Step 3
    '-------------------------

        'Find the last row for the 1st summary that is used a data source for the 2nd summary
        Dim Source_Table3_LastRow As LongLong
        Source_Table3_LastRowTicker = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
        'Find the the 1 first row to be poulated (summary data)
        Dim Summary_Table3_Row As LongLong
        Summary_Table3_Row = 2
    
        Dim Greatest_Increase_Percent_Value3 As Double
        Greatest_Increase_Percent_Value3 = ws.Range("S2")
        'MsgBox (Greatest_Increase_Percent_Value3 * 100)
    
        Dim Greatest_Decrease_Percent_Value3 As Double
        Greatest_Decrease_Percent_Value3 = ws.Range("S3")
        'MsgBox (Greatest_Decrease_Percent_Value3 * 100)
    
        Dim Greatest_Volume_Value3 As LongLong 'do not accept integer or long
        Greatest_Volume_Value3 = ws.Range("S4")
        'MsgBox (Greatest_Volume_Value)
    
        Dim SourcePercent3 As Double
        Dim SourceVolume3 As LongLong
        Dim Ticker3 As String
        
        'Ticker = ws.Range("K" & Summary_Table3_Row & ":K" & (Source_Table3_LastRowTicker))
    
        For i = Summary_Table3_Row To Source_Table3_LastRowTicker
            
                'Compare values to the range
                SourcePercent3 = ws.Range("M" & i).Value
                SourceVolume3 = ws.Range("N" & i).Value
                Ticker3 = ws.Range("K" & i).Value
              
              
                'Verify the condition #1
                If Greatest_Increase_Percent_Value3 = SourcePercent3 Then
                ws.Range("R2") = Ticker3
                End If
            
                'Verify the condition #2
                If Greatest_Decrease_Percent_Value3 = SourcePercent3 Then
                ws.Range("R3") = Ticker3
                End If
            
                'Verify the condition #3
                If Greatest_Volume_Value3 = SourceVolume3 Then
                ws.Range("R4") = Ticker3
                End If
            
            
                If ws.Range("M" & i) > 0 Then
                    ws.Range("M" & i).Interior.ColorIndex = 4
            
                ElseIf ws.Range("M" & i) < 0 Then
                    ws.Range("M" & i).Interior.ColorIndex = 3
                    
                ElseIf ws.Range("M" & i) = 0 Then
                    ws.Range("M" & i).Interior.ColorIndex = 7
                
                Else
                    ws.Range("M" & i).Interior.ColorIndex = 2
                
            End If
            
             
        Next i
        
        
        MsgBox ("Task complete !")
        
'------------------------
'Closing it for each worksheet
    Next ws
'------------------------

End Sub







