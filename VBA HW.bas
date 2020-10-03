Attribute VB_Name = "Module1"
Sub Ticker()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Volume Traded"

Dim LastRow As Double
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim Tick As Integer
Tick = 2

Dim beginning As Double
beginning = Cells(2, 3).Value
Dim final As Double
Dim difference As Double

Dim volumecount As LongLong
volumecount = 0

For i = 2 To LastRow

    volumecount = volumecount + Cells(i, 7).Value
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(Tick, 9).Value = Cells(i, 1).Value
            final = Cells(i, 6).Value
            difference = final - beginning
            Cells(Tick, 10).Value = difference
            Cells(Tick, 12).Value = volumecount
            
                If beginning = 0 Then
                    Cells(Tick, 11).Value = "n/a"
                    
                Else: Cells(Tick, 11).Value = difference / beginning
                    Cells(Tick, 11).Value = FormatPercent(Cells(Tick, 11).Value, 2)
                End If
                
            Tick = Tick + 1
            
            volumecount = 0
            beginning = Cells(i + 1, 3).Value
            
        End If
        
        If Cells(Tick, 11).Value < 0 Then
            Cells(Tick, 11).Interior.ColorIndex = 3
        Else: Cells(Tick, 11).Interior.ColorIndex = 4
        End If
        
        
Next i



End Sub
