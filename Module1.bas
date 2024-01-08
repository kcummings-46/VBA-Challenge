Attribute VB_Name = "Module1"
Sub dataStockChangeVolume():
    Dim WSheet As Worksheet
    
     For Each WSheet In Worksheets
        
        Dim Ticker As String
        Dim Table_Row As Integer
        
        Dim TVol As Double
        Dim Price_Row As Long
        Dim YrChange As Double
        Dim YrOpen As Double
        Dim YrClose As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim G_Increase As Double
        Dim G_Decreas As Double
        Dim G_Total As Double
        Dim G_Decrease_Ticker As String
        
        TVol = 0
      
       
        Table_Row = 2
        Price_Row = 2
        
      
       WSheet.Range("I1").Value = "Ticker"
       WSheet.Range("J1").Value = "Yearly Change"
       WSheet.Range("K1").Value = "Percent Change"
       WSheet.Range("L1").Value = "Total Stock Volume"
       WSheet.Range("O2").Value = "Greatest % Increase"
       WSheet.Range("O3").Value = "Greatest % Decrease"
       WSheet.Range("O4").Value = "Greatest Total Volume"
       WSheet.Range("P1").Value = "Ticker"
       WSheet.Range("Q1").Value = "Value"
       
       
       Lastrow = WSheet.Cells(Rows.Count, 1).End(xlUp).Row
      
       
       For i = 2 To Lastrow:
       YrOpen = WSheet.Cells(i, 3).Value
       
            
            If WSheet.Cells(i + 1, 1).Value <> WSheet.Cells(i, 1).Value Then
            
               Ticker = WSheet.Cells(i, 1).Value
               TVol = TVol + WSheet.Range("G" & i).Value
               WSheet.Range("I" & Table_Row).Value = Ticker
               WSheet.Range("L" & Table_Row).Value = TVol
               
               
               YrOpen = WSheet.Range("C" & Price_Row).Value
               YrClose = WSheet.Range("F" & i).Value
               YrChange = YrClose - YrOpen
               
               If YrOpen = 0 Then
                      Percent_Change = 0
                Else
                 Percent_Change = YrChange / YrOpen
            End If
               
                  WSheet.Range("J" & Table_Row).Value = YrChange
                  WSheet.Range("K" & Table_Row).Value = Percent_Change
                  WSheet.Range("K" & Table_Row).NumberFormat = "0.00%"
               
                        
                If WSheet.Range("J" & Table_Row).Value > 0 Then
                    WSheet.Range("J" & Table_Row).Interior.ColorIndex = 4
                Else
                    WSheet.Range("J" & Table_Row).Interior.ColorIndex = 3
            End If
               Table_Row = Table_Row + 1
               Price_Row = i + 1
               TVol = 0
                Else
                    TVol = TVol + WSheet.Range("G" & i).Value

           End If

        Next i
        
        G_Increase = WSheet.Range("K2").Value
        G_Decrease = WSheet.Range("K2").Value
        G_Total = WSheet.Range("L2").Value
        Lastrow_Ticker = WSheet.Cells(Rows.Count, "I").End(xlUp).Row
                
        For r = 2 To Lastrow_Ticker:
               If WSheet.Range("K" & r + 1).Value > G_Increase Then
                  G_Increase = WSheet.Range("K" & r + 1).Value
                  G_Increase_Ticker = WSheet.Range("I" & r + 1).Value

               ElseIf WSheet.Range("K" & r + 1).Value < G_Decrease Then
                  G_Decrease = WSheet.Range("K" & r + 1).Value
                  G_Decrease_Ticker = WSheet.Range("I" & r + 1).Value

                ElseIf WSheet.Range("L" & r + 1).Value > G_Total Then
                  G_Total = WSheet.Range("L" & r + 1).Value
                  G_Total_Ticker = WSheet.Range("I" & r + 1).Value

                End If
            Next r
            
            WSheet.Range("P2").Value = G_Increase_Ticker
            WSheet.Range("P3").Value = G_Decrease_Ticker
            WSheet.Range("P4").Value = G_Total_Ticker
            WSheet.Range("Q2").Value = G_Increase
            WSheet.Range("Q3").Value = G_Decrease
            WSheet.Range("Q4").Value = G_Total
            WSheet.Range("Q2:Q3").NumberFormat = "0.00%"
    Next WSheet
      
End Sub

