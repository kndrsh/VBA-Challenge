Attribute VB_Name = "Module1"
Sub Data()
   
 
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
       
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        Dim Open_Price As Double
        Open_Price = 0
        
        Dim Close_Price As Double
        Close_Price = 0
        
        Dim Yearly_Change As Double
        Dim Ticker_ As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim r As Long
        
      
        Open_Price = Cells(2, Column + 2).Value
        
        
        For r = 2 To LastRow
         
         If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                
            If Cells(r + 2, Column).Value <> Cells(r, Column).Value Then
                
                Ticker = Cells(r, Column).Value
                Cells(Row, Column + 8).Value = Ticker
               
                Close_Price = Cells(r, Column + 5).Value
                
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
               
                
               
                Volume = Volume + Cells(r, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                Row = Row + 1

                Open_Price = Cells(r + 1, Column + 2)
               
                Volume = 0
            
            Else
                Volume = Volume + Cells(r, Column + 6).Value
            End If
        Next r
      
        
        
    Next WS
        
End Sub

