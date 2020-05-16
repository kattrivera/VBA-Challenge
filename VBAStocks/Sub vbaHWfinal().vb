Sub vbaHWfinal()

'loop through all the worksheets
For Each ws In Worksheets

'set up summary table headings as per "hard solution" picture
ws.Cells(1, 9).Value = "Ticker Name"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
    
'set up additional table to the right
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
    
' declare variables
Dim Summary_Table_Row As Long
Dim vTicker As String
Dim vPercentageChange As Variant
Dim vStartofStock As Long
Dim vYearlyChange As Double
Dim fdOpeningPrice As Double
Dim ldClosingPrice As Double
Dim vTotalStockVolume As LongLong
Dim gpIncrease As Variant
Dim gpDecrease As Variant
Dim gtVolume As LongLong

'assign data to variables
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2
vStartofStock = 2
gpIncrease = 0
gpDecrease = 0
gtVolume = 0

'Start for loop:
     For r = 2 To lastrow
'Start conditional statement if the next cell doesnt match the current cell
     If ws.Cells(r + 1, 1) <> ws.Cells(r, 1).Value Then

'grab the ticker name
         vTicker = ws.Cells(r, 1).Value
         ws.Range("I" & Summary_Table_Row).Value = vTicker

'assign value to determine where a new stock starts, the location of last closing price, calculate yearly change and insert value in summary table         
         vStartofStock = r + 1
         ldClosingPrice = ws.Cells(r, 6).Value
         vYearlyChange = (ldClosingPrice - fdOpeningPrice)
         ws.Range("J" & Summary_Table_Row).Value = vYearlyChange

'color code positive change to green and negative change to red        
             If vYearlyChange > 0 Then
             ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
             ElseIf vYearlyChange <= 0 Then
             ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
             End If
        
'calculate percentage change and insert value in summary table
'originally stock PLNT has 0 value in all categories. Deemed an error and removed.        
         vPercentageChange = (vYearlyChange / fdOpeningPrice) * 100
         ws.Range("K" & Summary_Table_Row).Value = vPercentageChange

'conditional statement to determine greatest % increase and place value and ticker in additional table        
             If vPercentageChange > gpIncrease Then
             gpIncrease = vPercentageChange
             ws.Range("P2").Value = gpIncrease
             ws.Range("O2").Value = vTicker
             End If

'conditional statement to determine greatest % decrease and place value and ticker in additional table                  
             If vPercentageChange < gpDecrease Then
             gpDecrease = vPercentageChange
             ws.Range("P3").Value = gpDecrease
             ws.Range("O3").Value = vTicker
             End If
  
'calculate total stock volume and insert into summary table        
        vTotalStockVolume = vTotalStockVolume + ws.Cells(r, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = vTotalStockVolume

'conditional statement to determine greatest total volume and place value and ticker in additional table                 
        If vTotalStockVolume > gtVolume Then
        gtVolume = vTotalStockVolume
        ws.Range("P4").Value = gtVolume
        ws.Range("O4").Value = vTicker
        End If           

'increase summary table row by 1                
        Summary_Table_Row = Summary_Table_Row + 1

'reset total stock volume        
        vTotalStockVolume = 0

'exception to pick up first opening price of stock before it loops through        
        ElseIf r = vStartofStock Then
        fdOpeningPrice = ws.Cells(r, 3).Value

'recalculate total stock volume         
        Else: vTotalStockVolume = vTotalStockVolume + ws.Cells(r, 7).Value
            
        End If
             
    Next r

'autofit columns
ws.Columns.AutoFit

Next ws

End Sub