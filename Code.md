# Visual-Basic-Homework
##Easy 

Sub Stock_Volume()

  Dim LastRow As Long
  
  Dim Tick As String
    
  Dim Volume_Total As Double
  Volume_Total = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Tick = Cells(i, 1).Value
 
      Volume_Total = Volume_Total + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = Tick
     
      Range("J" & Summary_Table_Row).Value = Volume_Total

      Summary_Table_Row = Summary_Table_Row + 1

      Volume_Total = 0

    Else

      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i
  
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Total Stock Volume"
  
End Sub

