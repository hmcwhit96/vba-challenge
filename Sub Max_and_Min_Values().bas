Sub Max_and_Min_Values()
' copied over from the alphabetical testing file,
' the actual code is in the "Summary Table Final Product" module

Dim lastrow As Double
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("Q1") = "Ticker"
Range("R1") = "Value"
Range("P2") = "Greatest Increase"
Range("P3") = "Greatest Decrease"
Range("P4") = "Greatest Total Volume"

Range("R2").Value = Application.WorksheetFunction.Max(Range("K2:K290"))
Range("R3").Value = Application.WorksheetFunction.Min(Range("K2:K290"))
Range("R4").Value = Application.WorksheetFunction.Max(Range("L2:L290"))

Range("R2").NumberFormat = "0.00%"
Range("R3").NumberFormat = "0.00%"

    For i = 2 To lastrow
    
        'max increase
    
        If Cells(i, 11).Value = Range("R2").Value Then
    
            Range("Q2").Value = Cells(i, 9).Value
    
        Else
        End If
    
        'max("minimum")decrease
    
        If Cells(i, 11).Value = Range("R3").Value Then
     
            Range("Q3").Value = Cells(i, 9).Value
        Else
        End If
    
        'max volume
    
        If Cells(i, 12).Value = Range("R4").Value Then
    
            Range("Q4").Value = Cells(i, 9).Value
        Else
        End If
    
Next i

End Sub
