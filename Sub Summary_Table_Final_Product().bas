Sub Summary_Table_Final_Product()

'Repeat the loop for each sheet
For Each ws In ActiveWorkbook.Worksheets
    'activate the next sheet
    
    ws.Activate
    'setting up variables
Dim ticker As String
Dim totvol As Double
    totvol = 0
Dim begopenvalue As Double
Dim endcloseval As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim lastrow As Double
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim sumtablerow As Double
    sumtablerow = 2
'Dim StartRow As Double
    'StartRow = 2
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Volume"
Range("M1") = "Open Price"
Range("N1") = "Close Price"
Range("I1:N1").Interior.Color = vbCyan

'begin for loop
    For i = 2 To lastrow
            'Dim NextTicker As Integer
                'NextTicker = 262
        If Cells(i - 1, 1).Value <> Cells(i, 1) Then
            begopenvalue = Cells(i, 3).Value
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            totvol = totvol + Cells(i, 7).Value
                Range("I" & sumtablerow).Value = ticker
                Range("L" & sumtablerow).Value = totvol
            totalvol = 0
            'begopenvalue = Cells(StartRow, 3).Value
            endcloseval = Cells(i, 6).Value
                Range("M" & sumtablerow).Value = begopenvalue
                Range("N" & sumtablerow).Value = endcloseval
            If begopenvalue = 0 Then
                yearlychange = 0
                percentchange = 0
            Else
                yearlychange = endcloseval - begopenvalue
                percentchange = (endcloseval - begopenvalue) / begopenvalue
                Range("J" & sumtablerow).Value = yearlychange
                
                    If Range("J" & sumtablerow).Value < 0 Then
                        Range("J" & sumtablerow).Interior.Color = vbRed
                    Else
                        Range("J" & sumtablerow).Interior.Color = vbGreen
                    End If
                    
                
                Range("K" & sumtablerow).Value = percentchange
                Range("K" & sumtablerow).NumberFormat = "0.00%"
                    If Range("K" & sumtablerow).Value < 0 Then
                        Range("K" & sumtablerow).Interior.Color = vbRed
                    Else
                        Range("K" & sumtablerow).Interior.Color = vbGreen
                    End If
                
                
                
            End If
                sumtablerow = sumtablerow + 1
                totvol = 0
                
        Else: totvol = totvol + Cells(i, 7).Value
        End If
        'StartRow = StartRow + i
    Next i

    Range("Q1") = "Ticker"
    Range("R1") = "Value"
    Range("P2") = "Greatest % Increase"
    Range("P3") = "Greatest % Decrease"
    Range("P4") = "Greatest Total Volume"

    Range("R2").Value = Application.WorksheetFunction.Max(Range("K2:K3169"))
    Range("R3").Value = Application.WorksheetFunction.Min(Range("K2:K3169"))
    Range("R4").Value = Application.WorksheetFunction.Max(Range("L2:L3169"))
        
        Range("P2:P4").EntireColumn.AutoFit
        Range("R2:R4").EntireColumn.AutoFit
        
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

Next ws


End Sub

