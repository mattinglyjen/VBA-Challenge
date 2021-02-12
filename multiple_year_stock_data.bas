Attribute VB_Name = "Module1"
Sub multipleyrstock()

For Each sht In ThisWorkbook.Sheets

sht.Select

Dim ticker As String

Dim perchange As Variant

Dim yrchange As Variant
Dim yrOpen As Variant
Dim yrClose As Variant
Dim volume As Double
  ' Set an initial variable for volume per ticker
 
 volume = 0
  
  ' Keep track of the location for each ticker brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all tickers
  
  Dim RowCount As Double
  RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  yrOpen = Cells(2, 3).Value
  
     For I = 2 To RowCount

    ' Check if we are still within the same ticker, if it is not
    
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            
      ' Set the ticker
            ticker = Cells(I, 1).Value
      
       'set the year close value

            yrClose = Cells(I, 6).Value
       
      ' Add to the Volume
            volume = volume + Cells(I, 7).Value
            
            
      'establish the percent change difference
            yrchange = Round((yrClose - yrOpen), 2)
       
     'establish the year change difference to be printed
      If yrOpen <> 0 Then
        perchange = yrchange / yrOpen
         
         
      End If
       'Print the ticker in the Summary Table
       
            Range("I" & Summary_Table_Row).Value = ticker
      
     'print year change
             Range("J" & Summary_Table_Row).Value = yrchange
      
    'print the year change
    
            Range("K" & Summary_Table_Row).Value = perchange
            Range("K" & Summary_Table_Row).NumberFormat = "00.0%"
            
      ' Print the Volume to the Summary Table
      
             Range("L" & Summary_Table_Row).Value = volume

      ' Add one to the summary table row
             Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume and year open values
            yrOpen = Cells(I + 1, 3).Value
     
            volume = 0
     
     Else
 
        yrchange = yrClose - yrOpen
   
        volume = volume + Cells(I, 7).Value
     
    End If
      
  Next I
  
      For I = 2 To RowCount

         If Cells(I, 10).Value > 0 Then
        
        Cells(I, 10).Interior.ColorIndex = 4
    Else
        Cells(I, 10).Interior.ColorIndex = 3
    
    End If
    
    Next I
      Cells(1, 9).Value = "Ticker"
      Cells(1, 10).Value = "Yearly Change"
      Cells(1, 11).Value = "Percent Change"
      Cells(1, 12).Value = "Volume"


Next sht

     
End Sub
