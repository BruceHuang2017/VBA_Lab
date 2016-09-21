
Sub pollutantsTable()
    
    Dim polType As Integer, polAmt As Double, miles As Double

    polType = Range("B3").Value
    miles = Range("C3").Value
    polAmt = Range("D3").Value

    Dim FML As Double, SML As Double
    FML = Range("D9").Value
    SML = Range("E9").Value

    Dim fgl As Double, sgl As Double '  two gram limits

    Select Case polType '  gram limits for each pollutant
        Case 1 ' Carbon Monoxide
            fgl = Range("D10").Value
            sgl = Range("E10").Value
            Range("E3").Value = Range("C10").Value
        Case 2 ' Hydrocarbons
            fgl = Range("D11").Value
            sgl = Range("E11").Value
            Range("E3").Value = Range("C11").Value
        Case 3 ' Nitrogen
            fgl = Range("D12").Value
            sgl = Range("E12").Value
            Range("E3").Value = Range("C12").Value
        Case 4 ' Non-methane Hydrocarbons
            fgl = Range("D13").Value
            sgl = Range("E13").Value
            Range("E3").Value = Range("C13").Value
    End Select
    
' If statement. <= 50,000, between 50,000 and 100,000. < 100,000 miles.
' “Within Limit” or “Exceeds Limit”
' Google colorindex and get result: http://www.excel-pratique.com/en/vba/colors.php

    If miles <= FML Then
    
        If polAmt <= fgl Then
            Range("F3").Interior.ColorIndex = 4
            Range("F3").Value = "Within Limit"
        Else
            Range("F3").Interior.ColorIndex = 3
            Range("F3").Value = "Exceeds Limit"
        End If
        
    ElseIf miles > SML Then ' exceeds top limit.

            Range("F3").Interior.ColorIndex = 3
            Range("F3").Value = "Exceeds Limit"
            
    Else
    
        If polAmt <= sgl Then
            Range("F3").Interior.ColorIndex = 4
            Range("F3").Value = "Within Limit"
        Else
            Range("F3").Interior.ColorIndex = 3
            Range("F3").Value = "Exceeds Limit"
        End If
        
    End If

End Sub
