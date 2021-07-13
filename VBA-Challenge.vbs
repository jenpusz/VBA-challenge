
Sub Homework()

  ' Set initial variable
  Dim Ticker As String
  
  Dim Yearly_Change As Double
  Yearly_Change = 0
  
  Dim Summary_Table_Row As Integer
  
  Dim Open_Price As Double
  Dim Close_Price As Double
  Dim Volume As Double
  Dim Percent_Change As Double
  Dim Last As Long
  
  Dim ws As Worksheet
  
  For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
  
  Summary_Table_Row = 2
  
  
    'Set Title Rows
     Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Volume"
 
 
 
 
    'Define Last
    
    
    
    'Last = ActiveWorkbook.Worksheets.Count
     'Last = ActiveWorkbook.Worksheets("A" & Rows.Count).End(xlUp).Row
     
     Last = Cells(Rows.Count, "A").End(xlUp).Row
     
     
     
    
      ' Loop through all Ticker symbols
      
   For i = 2 To Last
   
   
   'Set first Open_Price
   
   If i = 2 Then
   
   Open_Price = Cells(i, 3)
   
   Range("N" & Summary_Table_Row).Value = Open_Price
   
   
    ' Check if we are still within the same Ticker symbol, if it is not update Summary_Table_Rows
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    Ticker = Cells(i, 1).Value
    
    'Add to the Volume
    
    Volume = Volume + Cells(i, 7).Value
    
    'Update Open Price
    
    Open_Price = Cells(i + 1, 3)
  
    
    'Calculate Yearly_Change
    Yearly_Change = Close_Price - Open_Price
    
    'Calculate Percent_Change
    Percent_Change = Yearly_Change / (Open_Price + 0.00000001) * 100
    
    
    

      

      ' Print the Ticker Symbol in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker
      
      'Print the Yearly Change in the Summary Table
       Range("K" & Summary_Table_Row).Value = Yearly_Change
       
      'Print the Percent_Change in the Summary Table
       Range("L" & Summary_Table_Row).Value = Percent_Change

      ' Print the Volume to the Summary Table
      Range("M" & Summary_Table_Row).Value = Volume
      
     'Print Open and Close Price for testing code
     
      'Print Open Price
      Range("N" & Summary_Table_Row + 1).Value = Open_Price
      'Print Closing Price
      Range("O" & Summary_Table_Row).Value = Close_Price
      
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume
      Volume = 0
      

    ' If the cell immediately following a row is the same ticker
    Else

        ' Add to the Volume
         Volume = Volume + Cells(i, 7).Value
        'Determine Close_Price
         Close_Price = Cells(i + 1, 6)
    
    
    
    End If
     
  Next i


For j = 2 To Last
'Cells(Rows.Count, "A").End(xlUp).Row
        
        'Color the Summary Tables
        If Cells(j, 11).Value > 0 Then
        
        Cells(j, 11).Interior.Color = vbGreen
        
        ElseIf Cells(j, 11).Value < 0 Then
        Cells(j, 11).Interior.Color = vbRed
        End If

Next j

Next ws

End Sub
