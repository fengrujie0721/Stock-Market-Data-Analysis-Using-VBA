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
'check if cell(i,1)'s value equals to cell(j,9)'s value
   If Cells(i, 1).Value = Cells(j, 9).Value Then
'add cell(i,7)'s value to sum_of_stock
        sum_of_stock = sum_of_stock + Cells(i, 7).Value
'set cell(j,12)'s value to sum_of_stock
      Cells(j, 12).Value = sum_of_stock
      
   Else
'set cell(j+1,9)'s value to cell(i,1)'s value   
       Cells(j + 1, 9).Value = Cells(i, 1).Value
'add 1 to j
       j = j + 1
'set sum_of_stock to 0
       sum_of_stock = 0
    End If
'end loop
Next i

'set variable b to 2
b = 2
'set lastrow
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'loop through rows
For a = 2 To lastrow
'check if cell(a,1)'s value equals to cell(b,9)'s value and cell(a,2)'s value is less than cell(a-1,2)'s value
     If Cells(a, 1).Value = Cells(b, 9).Value And Cells(a, 2).Value < Cells(a - 1, 2).Value Then
'set year_open_value equals to cell(a,3)'s value  
         year_open_value = Cells(a, 3).Value
   
'check if cell(a,1)'s value equals to cell(b,9)'s value and cell(a,2)'s value is larger than cell(a+1,2)'s value  
      ElseIf Cells(a, 1).Value = Cells(b, 9).Value And Cells(a, 2).Value > Cells(a + 1, 2).Value Then
'set year_end_value to cell(a,6)'s value      
         year_end_value = Cells(a, 6).Value
      
'set cell(b,10)'s value to year_end_value subtracted by year_open_value      
         Cells(b, 10).Value = year_end_value - year_open_value
'check if cell(b,10)'s value is larger than 0
             If Cells(b, 10).Value > 0 Then
'set cell(b,10)'s color to 4
                Cells(b, 10).Interior.ColorIndex = 4
             Else
'otherwise set cell(b,10)'s color to 3
                Cells(b, 10).Interior.ColorIndex = 3
             End If
'check if year_open_value equals to 0            
             If year_open_value = 0 Then
'set cell(b,11)'s value to 'not valid'
                Cells(b, 11).Value = "not valid"
             
    
              Else
'set cell(b,11)'s value to year_end_value subtracted by year_open_value than divided by year_open_value
                Cells(b, 11).Value = (year_end_value - year_open_value) / year_open_value
'set cell(b,11)'s format to percentage and 2 digits after decimal point
                Cells(b, 11).NumberFormat = "0.00%"
              End If
'add 1 to b
              b = b + 1
   
       End If
'end loop
Next a
'execute bonus sub
Call bonus
'end ticker sub
End Sub
'start bonus sub
 Sub bonus()
 
'set cells values  
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
'set cells formats
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
'set variables as integer
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
'set variables values
Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Volume = 0
'set lastrow
lastrow = Cells(Rows.Count, 11).End(xlUp).Row
'loop through rows
For d = 2 To lastrow
'check if cell(d,11)'s value is less than Greatest_Decrease
 If Cells(d, 11).Value < Greatest_Decrease Then
'set Greastest_Decrease to cell(d,11)'s value
       Greatest_Decrease = Cells(d, 11).Value
'set cell(3,17)'s value to cell(d,11)'s value
       Cells(3, 17).Value = Cells(d, 11).Value
'set cell(3,16)'s value to cell(d,9)'s value
       Cells(3, 16).Value = Cells(d, 9).Value
 End If
'check if cell(d,11)'s value is greater than Greatest_Increase and cell(d,11)'s value is not 'not valid'
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
    
    

    







