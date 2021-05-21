Attribute VB_Name = "Module3"
'start sub oneclick
Sub oneclick()

Application.ScreenUpdating = False
'select each worksheet
For Each ws In Worksheets
ws.Select
'execute sub tickername
Call tickername

'analyze next worksheet

Next
Application.ScreenUpdating = True
'end oneclick sub
End Sub
'start tickername sub
Sub tickername()
'set cells values
Cells(1, 9).Value = "Ticker"
Cells(2, 9).Value = Cells(2, 1).Value
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"

Cells(1, 12).Value = "Total Stock Volume"
'set variable as integer
Dim sum_of_stock As Double
'set lastrow
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'set variables
j = 2
sum_of_stock = 0
'loop through all rows
For i = 2 To lastrow
'check if cell value equals cells(j,9)'s value
If Cells(i, 1).Value = Cells(j, 9).Value Then
'add cell(i,7)'s value to sum_of_stock
     sum_of_stock = sum_of_stock + Cells(i, 7).Value
'set cell(j,12)'s value to sum_of_stock
      Cells(j, 12).Value = sum_of_stock
'check if cell value not equals cell(j,9)'s value      
Else
'set cell(j+1,9)'s value to cell(i,1)'s value   
          Cells(j + 1, 9).Value = Cells(i, 1).Value
'set j to j plus 1
          j = j + 1
'set sum_of_stock to 0
          sum_of_stock = 0

          End If
'end loop
Next i
'execute yearchange sub
Call yearchange
'end tickername sub
End Sub
'start yearchange sub
Sub yearchange()
'set variables as integers
Dim year_open_value As Double
Dim year_end_value As Double
Dim sum_of_stock As Double
'set variable b
b = 2
'set lastrow
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'loop through all rows
For a = 2 To lastrow
'check if cell(a,1)'s value equals cell(b,9)'s value and cell(a,2)'s value less than cell(a-1,2)'s value
     If Cells(a, 1).Value = Cells(b, 9).Value And Cells(a, 2).Value < Cells(a - 1, 2).Value Then
'set year_open_value to cell(a,3)'s value   
   year_open_value = Cells(a, 3).Value
   
'check if cell(a,1)'s value equals cell(b,9)'s value and cell(a,2)'s value less than cell(a+1,2)'s value    
    ElseIf Cells(a, 1).Value = Cells(b, 9).Value And Cells(a, 2).Value > Cells(a + 1, 2).Value Then
'set year_end_value to cell(a,6)'s value      
       year_end_value = Cells(a, 6).Value
'set cell(b,10)'s value to year_end_value subtracted by year_open_value
       Cells(b, 10).Value = year_end_value - year_open_value
'check if cells(b,10)'s value is more than 0
           If Cells(b, 10).Value > 0 Then
'set cell(b,10)'s color to 4
             Cells(b, 10).Interior.ColorIndex = 4
            Else
'set cell(b,10)'s color to 3
             Cells(b, 10).Interior.ColorIndex = 3
             End If
'check if cell(b,10)'s value equals 0
            If Cells(b, 10).Value = 0 Then
'set cell(b,11)'s value to "not valid"
              Cells(b, 11).Value = "not valid"
              
    
   Else
'set cell(b,11)'s value to year_end_value subtracted by year_open_value then divided by year_open_value
    Cells(b, 11).Value = (year_end_value - year_open_value) / 
'set cell(b,11)'s format to percentage and two digits after decimal point
    Cells(b, 11).NumberFormat = "0.00%"
    End If
'set b to b plus 1
    b = b + 1
   
   End If
'end loop
  Next a
'execute bonus sub
  Call bonus
'end yearchange sub
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

'set variables as integers
Dim minimal_change As Double
Dim maximal_change As Double
Dim maximal_volume As Double


'set lastrow
 lastrow = Cells(Rows.Count, 11).End(xlUp).Row
'set variables
 minimal_change = 0
 maximal_change = 0
 maximal_volume = 0
'loop through rows
For d = 2 To lastrow
'check if cell(d,11)'s value is less than minimal_change
 If Cells(d, 11).Value < minimal_change Then
'set minimal_change to cell(d,11)'s value 
 minimal_change = Cells(d, 11).Value
'set cell(3,17)'s value to cell(d,11)'s value
Cells(3, 17).Value = Cells(d, 11).Value
'set cell(3,16)'s value to cell(d,9)'s value 
 Cells(3, 16).Value = Cells(d, 9).Value
  
 End If
'check if cell(d,11)'s value is more than maximal change 
If Cells(d, 11).Value > maximal_change and Cells(d,11).Value<>"not valid" Then
'set maximal_change to cell(d,11)'s value
maximal_change = Cells(d, 11).Value
'set cell(2,17)'s value to cell(d,11)'s value
Cells(2, 17).Value = Cells(d, 11).Value
'set cell(2,16)'s value to cell(d,9)'s value
 Cells(2, 16).Value = Cells(d, 9).Value
  
End If
'chekc if cell(d,12)'s value is larger than maximal_volume
If Cells(d, 12).Value > maximal_volume Then
'set maximal_volume to cell(d,12)'s value
 maximal_volume = Cells(d, 12).Value
'set cell(4,17)'s value to cell(d,12)'s value
Cells(4, 17).Value = Cells(d, 12).Value
'set cell(4,16)'s value to cell(d,9)'s value
Cells(4, 16).Value = Cells(d, 9).Value

End If
'end loop
Next d



'end bonus sub

End Sub
    
    
   
    






