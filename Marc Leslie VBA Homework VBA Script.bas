Attribute VB_Name = "Module1"
Sub Alpha_Testing():
'YEARLY CHANGE IS BASED ON TAKE THE DIFFERENCE BETWEEN OPENING PRICE DAY 1 AND CLOSING PRICE THE LAST DAY

'Set varaibles
    Dim TheRows As Long
    Dim SummaryRow As Integer
    Dim TotalVolume
    Dim ClosingPrice As Double
    Dim YearlyChange
    Dim HoldOpen As Double
    Dim Ticker As String
    Dim PerctChange As Double
    Dim Max As Double
            
'Go through each WS and active each
    For Each ws In Worksheets
    ws.Activate

'Initialize variables
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row 'Cant do same formula for columns due to non data cells
    SummaryRow = 2
    TotalVolume = 0
    YearlyChange = 0
    PerctChange = 0
    HoldOpen = 0
    Ticker = ""
    
     
'Headers for summary table and making them stand out
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").Interior.ColorIndex = 16
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("I1:L1").EntireColumn.AutoFit
    ws.Range("I1:L1").HorizontalAlignment = xlCenter
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1, Q1, O2, O3, O4, P2, P3, P4, Q2, Q3, Q4").Font.Bold = True
    ws.Range("P1, Q1, O2, O3, O4").EntireColumn.AutoFit
'-----------------------------------------------------------------------
For TheRows = 2 To LastRow
        
'Step 1: Right now it assumes Ticker is blank.  When ticker is not blank, grab that opening day value
If Cells(TheRows, 1) <> Ticker Then
HoldOpen = Cells(TheRows, 3).Value
Ticker = Cells(TheRows, 1).Value
End If

    'Step 2: Look through each row and see if the stock on row x does not match the stock on row x+1.
    If Cells(TheRows + 1, 1) <> Cells(TheRows, 1) Then
        
            'Step 7: You found <>! Add up total volume in col 7, grab the closing day price, calc yearly change...
            TotalVolume = TotalVolume + Cells(TheRows, 7).Value
            ClosingPrice = Cells(TheRows, 6).Value
            YearlyChange = HoldOpen - ClosingPrice
            
            If HoldOpen = 0 Then
            PerctChange = 0
            Else
            PerctChange = (YearlyChange / HoldOpen) * 100
            End If
                     
            'Step 8:...and put stock name, total volume, and yearly change in the summary table, make it look nice...
            ws.Range("L" & SummaryRow).Value = TotalVolume
            ws.Range("I" & SummaryRow).Value = Cells(TheRows, 1).Value
            ws.Range("J" & SummaryRow).Value = YearlyChange
            ws.Range("K" & SummaryRow).Value = PerctChange
            ws.Range("K" & SummaryRow).NumberFormat = "0.00"
            ws.Range("L" & SummaryRow).HorizontalAlignment = xlLeft
            ws.Range("I" & SummaryRow).HorizontalAlignment = xlLeft
            ws.Range("J" & SummaryRow).HorizontalAlignment = xlLeft
            ws.Range("K" & SummaryRow).HorizontalAlignment = xlLeft
            ws.Range("K" & SummaryRow).HorizontalAlignment = xlLeft
           
            'Step 9: ...and re-set the volume to 0 so you can start adding the next stock up, make the cells highlight...
            TotalVolume = 0
                        
            If ws.Range("J" & SummaryRow).Value < 0 Then
            ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
            Else
            ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
            End If
                                     
            'Step 10...and move down to the next line of the summary table so you don't overwrite
            SummaryRow = SummaryRow + 1
                       
  'Step 3: You did not find < > which means still trying to find the end, so do something "Else" which is...
    Else

  'Step 4: ...start adding up those total stock volume values...and
    TotalVolume = TotalVolume + Cells(TheRows, 7).Value
     
  'Step 5: you would "end if" you found < >, but you did not, so...
    
    End If

  'Step 6:...go to the next row down and keep looking for <>
    
 Next TheRows
 
 
 
'Loop for max, min, greatest volume, which is basically the same type of thing as above
'It must go here because it has to go AFTER the results are all in

For TheRows = 2 To LastRow
 
'This finds the max
'WHY DOES < 0 or <0 not work???

If Cells(TheRows, 11).Value > ws.Range("Q2").Value Then
ws.Range("Q2").Value = Cells(TheRows, 11).Value
ws.Range("Q2").NumberFormat = "0.00"

'This puts the max ticker in the cell
ws.Range("P2").Value = Cells(TheRows, 9).Value

End If

'This finds the min
If Cells(TheRows, 11).Value < ws.Range("Q3").Value Then
ws.Range("Q3").Value = Cells(TheRows, 11).Value
ws.Range("Q3").NumberFormat = "0.00"

'This puts the min ticker in the cell
ws.Range("P3").Value = Cells(TheRows, 9).Value

End If

'This finds largest total vol
If Cells(TheRows, 12).Value > ws.Range("Q4").Value Then
ws.Range("Q4").Value = Cells(TheRows, 12).Value

'This puts the largest vol ticker in cel
ws.Range("P4").Value = Cells(TheRows, 9).Value

ws.Range("P1:Q1").EntireColumn.AutoFit
End If



Next TheRows
 
 
 Next ws
  
   
 End Sub


             
      
        
 


 
 
 

    
    
        

        





