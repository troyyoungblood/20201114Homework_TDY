# 20201114Homework_TDY
VBA homework due Nov 14 2020

The homework was completed by first getting it to run on one sheet.  After it was operational on one sheet, it was then transferred to the actual homework excel file.  The information below shows how the single tab file was updated to cycle through all the tabs.  The first 3 lines and the next to last line, Next Yes, were the only additional lines of code needed to make the code work on all tabs in the spreadsheet.  

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


An addtional two columns were added to the summary table.  Opening Price and Last Closing.  I added these intermediate data points as an aid to allow me to:

	1) Ensure the first data point was captured
        2) Ensure the last data point was captured
        3) Ensure my yearly change was being calculated correctly, and
        4) Visually verify percentage was being calculated correctly.

These columns can always be removed, but they really give real time feedback that everthing is working correctly.
