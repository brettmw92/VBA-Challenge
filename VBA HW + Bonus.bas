Attribute VB_Name = "Module1"
Sub Ticker()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Volume Traded"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

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

Dim LastRow2 As Double
LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row

Dim highestp As Double
Dim lowestp As Double
Dim greatestv As LongLong
Dim highestpt As String
Dim lowestpt As String
Dim greatestvt As String

greatestv = Cells(2, 12).Value
greatestvt = Cells(2, 9).Value
highestp = Cells(2, 11).Value
highestpt = Cells(2, 9).Value
lowestp = Cells(2, 11).Value
lowestpt = Cells(2, 9).Value

For j = 2 To LastRow2

    If Cells(j, 12).Value > greatestv Then
        greatestv = Cells(j, 12).Value
        greatestvt = Cells(j, 9).Value
    End If
    
    If Cells(j, 11).Value <> "n/a" Then
        If Cells(j, 11).Value > highestp Then
            highestp = Cells(j, 11).Value
            highestpt = Cells(j, 9).Value
        End If
    End If
    
    If Cells(j, 11).Value < 0 Then
        If Abs(Cells(j, 11).Value) > Abs(lowestp) Then
         lowestp = Cells(j, 11).Value
         lowestpt = Cells(j, 9).Value
        End If
    End If

Next j

Cells(4, 16).Value = greatestv
Cells(4, 15).Value = greatestvt
Cells(2, 16).Value = highestp
Cells(2, 15).Value = highestpt
Cells(3, 16).Value = lowestp
Cells(3, 15).Value = lowestpt

Cells(2, 16).Value = FormatPercent(Cells(2, 16).Value, 2)
Cells(3, 16).Value = FormatPercent(Cells(3, 16).Value, 2)

End Sub
