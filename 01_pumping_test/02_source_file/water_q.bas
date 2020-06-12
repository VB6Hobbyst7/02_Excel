Attribute VB_Name = "water_q"

Function ss_water(qhp As Integer, strPurpose As String) As Double

    Dim mypos As Integer


    mypos = InStr(1, strPurpose, "일")
    If (mypos <> 0) Then
        ss_water = Round(3.154 + qhp * 0.023, 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "가")
    If (mypos <> 0) Then
        ss_water = Round(0.173 + 2.63 * 0.21, 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "기")
    If (mypos <> 0) Then
        ss_water = Round(0.173 + 2.63 * 0.21, 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "간")
    If (mypos <> 0) Then
        ss_water = Round(7.13 + 30 * 0.001, 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "농")
    If (mypos <> 0) Then
        ss_water = Round(0.173 + 2.63 * 0.21, 2)
        Exit Function
    End If
    
   ss_water = 900
      
End Function


Function aa_water(qhp As Integer, strPurpose As String) As Double

    Dim mypos As Integer


    mypos = InStr(1, strPurpose, "전")
    If (mypos <> 0) Then
        aa_water = Round(5.66 + qhp * 0.014, 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "답")
    If (mypos <> 0) Then
        aa_water = Round(1.98 + qhp * 0.044, 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "원")
    If (mypos <> 0) Then
        aa_water = Round(2.789 + qhp * 0.011, 2)
        Exit Function
    End If
    
    
   aa_water = 900
      
End Function









