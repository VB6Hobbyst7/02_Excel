Attribute VB_Name = "water_q"
Public SS(1 To 5, 1 To 2) As Double
Public AA(1 To 6, 1 To 2) As Double

Public SS_CITY As Double

Public Enum SS_VALUE
    
    ssv_gajung = 1
    ssv_ilban = 2
    ssv_school = 3
    ssv_apartment = 4
    ssv_town = 5

End Enum

Public Enum AA_VALUE
    
    aav_jeon = 1
    aav_dap = 2
    aav_wonye = 3
    aav_cow = 4
    aav_pig = 5
    aav_chicken = 6
    
End Enum


Sub init_nonsan()

   
    SS(ssv_gajung, 1) = 0.173
    SS(ssv_gajung, 2) = 0.21
    SS_CITY = 2.63
    
    SS(ssv_ilban, 1) = 3.154
    SS(ssv_ilban, 2) = 0.023
    
    SS(ssv_school, 1) = 7.986
    SS(ssv_school, 2) = 0.005
    
    SS(ssv_apartment, 1) = 0.173
    SS(ssv_apartment, 2) = 0.21
    
    SS(ssv_town, 1) = 7.13
    SS(ssv_town, 2) = 0.001
    
'----------------------------------------

    AA(aav_jeon, 1) = 5.66
    AA(aav_jeon, 2) = 0.014
    
    AA(aav_dap, 1) = 1.98
    AA(aav_dap, 2) = 0.044
    
    AA(aav_wonye, 1) = 2.789
    AA(aav_wonye, 2) = 0.011
    
    AA(aav_cow, 1) = 3.48
    AA(aav_cow, 2) = 0.009
    
    AA(aav_pig, 1) = 4.719
    AA(aav_pig, 2) = 0.001
    
    AA(aav_chicken, 1) = 5.492
    AA(aav_chicken, 2) = 0.041
    
End Sub

Sub init_daejeon()

   
    SS(ssv_gajung, 1) = 0.173
    SS(ssv_gajung, 2) = 0.21
    SS_CITY = 2.73

    SS(ssv_ilban, 1) = 3.154
    SS(ssv_ilban, 2) = 0.023
    
    SS(ssv_school, 1) = 7.986
    SS(ssv_school, 2) = 0.005
    
    SS(ssv_apartment, 1) = 0.173
    SS(ssv_apartment, 2) = 0.21
    
    SS(ssv_town, 1) = 7.13
    SS(ssv_town, 2) = 0.001
    
'----------------------------------------

    AA(aav_jeon, 1) = 5.66
    AA(aav_jeon, 2) = 0.014
    
    AA(aav_dap, 1) = 1.98
    AA(aav_dap, 2) = 0.044
    
    AA(aav_wonye, 1) = 2.789
    AA(aav_wonye, 2) = 0.011
    
    AA(aav_cow, 1) = 3.48
    AA(aav_cow, 2) = 0.009
    
    AA(aav_pig, 1) = 4.719
    AA(aav_pig, 2) = 0.001
    
    AA(aav_chicken, 1) = 5.492
    AA(aav_chicken, 2) = 0.041

End Sub

Sub init_boryoung()

   
    SS(ssv_gajung, 1) = 0.173
    SS(ssv_gajung, 2) = 0.21
    SS_CITY = 2.52
    
    SS(ssv_ilban, 1) = 3.154
    SS(ssv_ilban, 2) = 0.023
    
    SS(ssv_school, 1) = 7.986
    SS(ssv_school, 2) = 0.005
    
    SS(ssv_apartment, 1) = 0.173
    SS(ssv_apartment, 2) = 0.21
    
    SS(ssv_town, 1) = 7.13
    SS(ssv_town, 2) = 0.001
    
'----------------------------------------

    AA(aav_jeon, 1) = 6.964
    AA(aav_jeon, 2) = 0.013
    
    AA(aav_dap, 1) = 2.089
    AA(aav_dap, 2) = 0.043
    
    AA(aav_wonye, 1) = 2.789
    AA(aav_wonye, 2) = 0.011
    
    AA(aav_cow, 1) = 3.48
    AA(aav_cow, 2) = 0.009
    
    AA(aav_pig, 1) = 4.719
    AA(aav_pig, 2) = 0.001
    
    AA(aav_chicken, 1) = 5.492
    AA(aav_chicken, 2) = 0.041
    
End Sub


Sub initialize()
        
       Call init_nonsan
       'Call init_daejeon
       'Call init_boryoung
       
End Sub



Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double

    Dim mypos As Integer


    mypos = InStr(1, strPurpose, "일") '일반용
    If (mypos <> 0) Then
        ss_water = Round(SS(ssv_ilban, 1) + qhp * SS(ssv_ilban, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "가") '가정용
    If (mypos <> 0) Then
        ss_water = Round(SS(ssv_gajung, 1) + SS_CITY * SS(ssv_gajung, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "기") '기타
    If (mypos <> 0) Then
        ss_water = Round(SS(ssv_gajung, 1) + SS_CITY * SS(ssv_gajung, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "농") '농생활겸용
    If (mypos <> 0) Then
        ss_water = Round(SS(ssv_gajung, 1) + SS_CITY * SS(ssv_gajung, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "청") '청소용
    If (mypos <> 0) Then
        ss_water = Round(SS(ssv_gajung, 1) + SS_CITY * SS(ssv_gajung, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "상") '간이상수도
    If (mypos <> 0) Then
        ss_water = Round(SS(ssv_town, 1) + npopulation * SS(ssv_town, 2), 2)
        Exit Function
    End If
    
    
   ss_water = 900
      
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double

    'nhead - 축산업의 두수 ....


    Dim mypos As Integer


    mypos = InStr(1, strPurpose, "전") '전작용
    If (mypos <> 0) Then
        aa_water = Round(AA(aav_jeon, 1) + qhp * AA(aav_jeon, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "답") '답작용
    If (mypos <> 0) Then
        aa_water = Round(AA(aav_dap, 1) + qhp * AA(aav_dap, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "원") '원예용
    If (mypos <> 0) Then
        aa_water = Round(AA(aav_wonye, 1) + qhp * AA(aav_wonye, 2), 2)
        Exit Function
    End If
    
    '농생활겸용
    mypos = InStr(1, strPurpose, "농")
    If (mypos <> 0) Then
        aa_water = Round(AA(aav_jeon, 1) + qhp * AA(aav_jeon, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "축") '축산업
    If (mypos <> 0) Then
        aa_water = Round(AA(aav_cow, 1) + nhead * AA(aav_cow, 2), 2)
        Exit Function
    End If
    
   aa_water = 900
      
End Function









