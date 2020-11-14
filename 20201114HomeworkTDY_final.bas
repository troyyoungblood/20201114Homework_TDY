Attribute VB_Name = "Module1"
Sub VBAHomework()

' ---- Code Layout -----
'
'  The overall For statement that runs across all worksheets
'
'  Dim Yes As Integer
'   For Yes = 1 to Workheets.Count
'    Worksheets(Yes).Select
' - - - - - - - - - - - - -
'     Code that worked on single sheet
' - - - - - - - - - - - - -
'    Next Yes
' End Sub


' - - - - Beginning of Mother Loop

Dim Yes As Integer

For Yes = 1 To Worksheets.Count  'For Mother Loop

    Worksheets(Yes).Select
    
    Range("K:U").Clear

' - - - - Code that worked on single sheet

' Set Variables
    Dim Ticker As String  'Set an initial variable for holding the stock name
    Dim Total_Vol As Double ' Set an initial variable for holding the total volume traded per ticker for year
    Dim Summary_Table_Row As Integer  ' Keep track of the location for each ticker in the summary table
    Dim Opening As Currency   'Opening stock price.  Initially set as Long.  Try to set as currency later
    Dim Last_Close As Currency  'Last closing stock for the year.  Initially set as Long.  Try to set as currency later.
    Dim Percent_Chg As Double  'Percent change of stock from first opening to last close of the year
    Dim Last_Row As Double  'Last row of a given ticker symbol
    Dim I As Long ' Counter in Ticker loop
    Dim j As Long ' Counter in Percent Change loop
    Dim LastRowPercent As Double 'Last row in summary table
    Dim MaxPerInc As Double  'Maximum percent increase value
    Dim MinPerInc As Double  'Minimum percent increase value
    Dim GreatTot As Double  'Maximum total volume value
    Dim r As Long ' Counter in Bonus Loop
    
  
' Set Variable Initial Values as Needed
    Total_Vol = 0
    Summary_Table_Row = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Opening_Counter = 1
  
' Loop through all ticker symbols to capture opening price for each specific ticker - tickers must in repctive date order
    For I = 2 To LastRow   'For Loop F1
    
    ' Store opening price of ticker dedicated IF loop - Chunky but it works - Smarter person will be cooler :)
        If Cells(I + 1, 1).Value = Cells(I, 1).Value Then  'IF Statement 1
        
        ' Check if Opening Counter = 1.  If yes, this captures first opening price of Ticker
            If Opening_Counter = 1 Then   'IF 1a
            
            ' Set opening price of ticker to variable
                Opening = Cells(I, 3).Value
                
            ' Place Ticker symbol in the Summary Table under Ticker Header
                Range("L" & Summary_Table_Row).Value = Opening
                
            ' Advances Opening_Counter value.  Once >1, value will not be entered into summary table
                Opening_Counter = Opening_Counter + 1
            
            End If   ' End IF 1a
            
        End If  ' End IF 1
            

    ' Check if next ticker equals current - if no, set designated variables - if yes - the Else statement sums designated values
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then   'IF 2

        ' Set the Ticker symbol
            Ticker = Cells(I, 1).Value

        ' Add the final value to the Total Volume traded sum from the last row of ticker
            Total_Vol = Total_Vol + Cells(I, 7).Value

        ' Place Ticker symbol in the Summary Table under Ticker Header
            Range("K" & Summary_Table_Row).Value = Ticker

        ' Place Total volume traded in the Summary Table under Total Volume Header
            Range("P" & Summary_Table_Row).Value = Total_Vol

        ' Add one to the summary table row to advance data entry to cell below current
            Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the Total volume for the ticker
            Total_Vol = 0
            
         ' Reset the Opening Counter for the ticker
            Opening_Counter = 1

     ' If the cell immediately following a row is the same ticker.
        Else    'Else for IF 2
        
            ' Add to the Ticker's total volume traded
            Total_Vol = Total_Vol + Cells(I, 7).Value
            
            ' Captures last closing price of ticker in last row of ticker - found this by accident when trying to determine opening price - got lucky
            Last_Close = Cells(I, 6).Value
            
            ' Place last close price in the Summary Table under Last Closing Header
            Range("M" & Summary_Table_Row).Value = Last_Close
        
            
        End If    'End IF 2

  Next I    'This loop counter is to collect all the raw data and summarize / tabulate designated vaules


'Calculate Yearly Change, Percent Change loop - this loop works off the tabulated data collected from above
    
' Set Variable Initial Value as Needed
' LastRowPercent looks for the last row in the tabulated data
' Note - Several Tickers had a final closing price of $0.  Some even had volume traded
'        Since direction was not provided on methodolgy to manage those situations -
'        Those Tickers were assigned a percent change value of 0

    LastRowPercent = Cells(Rows.Count, 11).End(xlUp).Row
    MaxPerInc = 0
    MinPerInc = 0
    
    For j = 2 To LastRowPercent   'For loop F2
      
        Cells(j, 14).Value = Cells(j, 13).Value - Cells(j, 12).Value
              
            If Cells(j, 12).Value <> 0 Then   'IF 3 - Calculates percent change for Ticker that did not end year at 0
        
                Cells(j, 15).Value = Cells(j, 14).Value / Cells(j, 12).Value
            
            Else  ' Else for IF 3 - Assigns 0 to Ticker that closed year at 0
                
                Cells(j, 15).Value = 0
                
            End If ' End If IF3
    
    Next j  'Next for loop F2
    
    
'Calculate Bonus loop using tabulated data

' Using previous variable LastRowPercent from last loop
' Set Variable Initial Value as Needed
    MaxPerInc = 0
    MinPerInc = 0
    GreatTot = 0
    
    For r = 2 To LastRowPercent  'For loop F3
      
        If Cells(r, 15).Value > MaxPerInc Then   'IF 4 - loop finds highest percent change
        
            MaxPerInc = Cells(r, 15).Value
            Cells(3, 21).Value = Cells(r, 15).Value
            Cells(3, 20).Value = Cells(r, 11).Value
            
        End If  ' End IF 4
        
        If Cells(r, 15).Value < MinPerInc Then  'IF 5 - loop finds minimum percent change
        
            MinPerInc = Cells(r, 15).Value
            Cells(4, 21).Value = Cells(r, 15).Value
            Cells(4, 20).Value = Cells(r, 11).Value
            
        End If  'End IF 5
        
        If Cells(r, 16).Value > GreatTot Then  'IF 6 - loop finds highest total volume traded
        
            GreatTot = Cells(r, 16).Value
            Cells(5, 21).Value = Cells(r, 16).Value
            Cells(5, 20).Value = Cells(r, 11).Value
            
        End If  'End If 6
            
             
    Next r  'End loop F3
    
    
' Formatting Activities
     
     Range("K1").Value = "Ticker"
     Range("L1").Value = "Opening Price"   'Column added to help with trouble shooting and data management
     Range("M1").Value = "Last Closing"    'Column added to help with trouble shooting and data management
     Range("N1").Value = "Yearly Change"
     Range("O1").Value = "Percent Change"
     Range("P1").Value = "Total Volume"
     Range("S3").Value = "Greatest % Increase"
     Range("S4").Value = "Greatest % Decrease"
     Range("S5").Value = "Greatest Total Volume"
     Range("T2").Value = "Ticker"
     Range("U2").Value = "Value"
     [s2].Value = ActiveSheet.Name
     
     Range("A:U").HorizontalAlignment = xlCenter
     Range("C:F").NumberFormat = "$#,##0.00"
     Range("L:N").NumberFormat = "$#,##0.00"    '"$#,##0.00"
     Range("O:O").NumberFormat = "0.00%"
     Range("P:P").NumberFormat = "0,000"
     Range("U3").NumberFormat = "0,000%"
     Range("U4").NumberFormat = "0%"
     Range("U5").NumberFormat = "#,###0"
     Columns("A:U").AutoFit
     
     
     'Color loop for Percent Change
     
     Dim ColorFormat As Integer
     
     For ColorFormat = 2 To LastRowPercent   'For loop F4
      
        If Cells(ColorFormat, 15).Value > 0 Then  'IF 6
        
            Cells(ColorFormat, 15).Interior.ColorIndex = 4
            
        Else  'IF 6
        
            Cells(ColorFormat, 15).Interior.ColorIndex = 3
            
        End If  'If 6
        
      Next ColorFormat
      
      'Color Bonus Table
      
        Cells(3, 21).Interior.ColorIndex = 4
            
        Cells(4, 21).Interior.ColorIndex = 3
     
     
' - - - End of code that ran on single worksheet sheet

Next Yes  'End Mother Loop
    
'  - - - - - End of Mother Loop

End Sub


