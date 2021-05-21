Attribute VB_Name = "Module2"
'start onclick sub
Sub oneclick()
'select each worksheet
For Each ws In Worksheets
ws.Select
'call the ticker sub
Call ticker
'analyze next worksheet    
Next ws
'end oneclick sub
End Sub
'start ticker sub
Sub ticker()
'set cells values
Cells(1, 9).Value = "Ticker"
Cells(2, 9).Value = Cells(2, 1).Value
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
'set variables as integers
Dim year_open_value As Double
Dim year_end_value As Double
Dim sum_of_stock As Double

'set lastrow
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'set variable j
j = 2
'set variable sum_of_stock to 0
sum_of_stock = 0
'loop through rows
For i = 2 To lastrow

   If Cells(i, 1).Value = Cells(j, 9).Value Then
        sum_of_stock = sum_of_stock + Cells(i, 7).Value
      Cells(j, 12).Value = sum_of_stock
      
   Else
   
       Cells(j + 1, 9).Value = Cells(i, 1).Value
       j = j + 1
       sum_of_stock = 0
    End If
Next i


b = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For a = 2 To lastrow

     If Cells(a, 1).Value = Cells(b, 9).Value And Cells(a, 2).Value < Cells(a - 1, 2).Value Then
   
         year_open_value = Cells(a, 3).Value
   
    
      ElseIf Cells(a, 1).Value = Cells(b, 9).Value And Cells(a, 2).Value > Cells(a + 1, 2).Value Then
      
         year_end_value = Cells(a, 6).Value
      
      
         Cells(b, 10).Value = year_end_value - year_open_value
             If Cells(b, 10).Value > 0 Then
                Cells(b, 10).Interior.ColorIndex = 4
             Else
                Cells(b, 10).Interior.ColorIndex = 3
             End If
            
             If year_open_value = 0 Then
                Cells(b, 11).Value = "not valid"
             
    
              Else
                Cells(b, 11).Value = (year_end_value - year_open_value) / year_open_value
                Cells(b, 11).NumberFormat = "0.00%"
              End If
              b = b + 1
   
       End If
Next a
Call bonus

End Sub
 Sub bonus()
 
  
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 15).Value = "Greatest Total Volume"
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Volume = 0
lastrow = Cells(Rows.Count, 11).End(xlUp).Row
For d = 2 To lastrow
 If Cells(d, 11).Value < Greatest_Decrease Then
 
       Greatest_Decrease = Cells(d, 11).Value

       Cells(3, 17).Value = Cells(d, 11).Value
 
       Cells(3, 16).Value = Cells(d, 9).Value
 End If

 If Cells(d, 11).Value > Greatest_Increase And Cells(d, 11).Value <> "not valid" Then

        Greatest_Increase = Cells(d, 11).Value

        Cells(2, 17).Value = Cells(d, 11).Value
        Cells(2, 16).Value = Cells(d, 9).Value
 End If
 
 If Cells(d, 12).Value > Greatest_Volume Then
        Greatest_Volume = Cells(d, 12).Value

        Cells(4, 17).Value = Cells(d, 12).Value
        Cells(4, 16).Value = Cells(d, 9).Value
  End If

Next d

End Sub
    
    

    







