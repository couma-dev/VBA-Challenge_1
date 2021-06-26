Sub Wall_Street1()

'Create a Script that will loop throught all the stocks for one year and output
'--- Ticker Symbol
'--- Yearly change from opening price to closing price
'--- percentage from opening price to closing price
'--- totatl stock Volume of the stock.



' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
        
        ' Declare variable for last row
        Dim LastRow As Double
        'determine last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Add headers "Ticker", "yearly Change", "percentage Change", "Total Stock Volume". begging with column I and moving to the right
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
' Declare Variables to hold Ticker, year change, percentage change, stock open and stock close values

        Dim Ticker As String
        Dim Yearly_change As Double
        Dim Percentage_change As Double
        Dim Volume As Double
        Dim Stock_Open As Double
        Dim Stock_Close As Double
    
' set Volume value to starts at 0

    Volume = 0
' Declare Summary Table Row And starting row number

   Dim Summary_Table_Row As Double
   Summary_Table_Row = 2
    
' iterate from the 2nd row to the last row

 For i = 2 To LastRow
 
 'checking through column A row by row for ticket value that is different from the previous to put in summary column I
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    
    Ticker = ws.Cells(i, 1).Value
    
    
    Volume = Volume + ws.Cells(i, 7)
    
'assigning Ticker and Volume numbers to column I, row 2 and column L row 2 respectively
    
  ws.Range("I" & Summary_Table_Row).Value = Ticker
  ws.Range("L" & Summary_Table_Row).Value = Volume
  

  Volume = 0
  
  Stock_Close = ws.Cells(i, 6)
  
    If Stock_Open = 0 Then
  
  Yearly_change = 0
  
  Percentage_change = 0
  
    Else
  
 Yearly_change = Stock_Close - Stock_Open
 
 Percentage_change = (Stock_Close - Stock_Open) / Stock_Open
    End If
' Assigning Yearly change and percentage change to columns J and K

 ws.Range("J" & Summary_Table_Row).Value = Yearly_change
 ws.Range("K" & Summary_Table_Row).Value = Percentage_change
 ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00"
 ws.Range("K" & Summary_Table_Row).Style = "Percent"
 
 ' moving to the next row
 Summary_Table_Row = Summary_Table_Row + 1
    
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    
    Stock_Open = ws.Cells(i, 3)
    
    
    Else: Volume = Volume + ws.Cells(i, 7).Value
 
 End If
 
 
 Next i
 
 
'conditional formatting to highlight positive change in green and negative change in red

For j = 2 To LastRow
    If ws.Range("j" & j).Value > 0 Then
        ws.Range("j" & j).Interior.ColorIndex = 4
        
ElseIf ws.Range("j" & j).Value < 0 Then
        ws.Range("j" & j).Interior.ColorIndex = 3
End If
    
    Next j
 
    Next ws
    
End Sub



