Attribute VB_Name = "SetTime_LongTest"
Public myTime As Integer



'10-77 : 2880 (68)
'78-101: recover (24)

Sub set_daydifference()

    Dim nTime() As Integer
    Dim i As Integer
    Dim day1, day2 As Integer
    
    ReDim nTime(1 To 92)
     
    For i = 1 To 92
        nTime(i) = Cells(i + 9, "D").Value
        If (i > 68) Then
            nTime(i) = Cells(i + 9, "D").Value + 2880
        End If
    Next i
    
    For i = 1 To 92
        Cells(i + 9, "h").Value = Range("c10").Value + nTime(i) / 1440
    Next i
    
    Range("H10:H101").Select
    Selection.NumberFormatLocal = "yyyy""년"" m""월"" d""일"";@"
    Range("A1").Select
    Application.CutCopyMode = False

    
    Application.ScreenUpdating = False
    day1 = Day(Cells(10, "h").Value)

    For i = 2 To 92
        day2 = Day(Cells(i + 9, "h").Value)
        If (day2 = day1) Then
            Cells(i + 9, "h").Value = ""
        End If
        day1 = day2
    Next i
    
    Range("h77").Value = "양수종료"
    Range("h78").Value = "회복수위측정"
    Application.ScreenUpdating = True

End Sub

Function find_stable_time() As Integer

    Dim i As Integer
    
    
    For i = 30 To 50
            
        If Range("AB" & CStr(i)).Value = Range("AB" & CStr(i + 1)) Then
            'MsgBox "found " & "AB" & CStr(i) & " time : " & Range("Z" & CStr(i)).Value
            
            find_stable_time = i
            Exit For
        End If
        
    Next i
 
End Function

Function initialize_myTime() As Integer

    'Range("G17").Value = 840 + 60 * (i - 35)

    initialize_myTime = (Sheet9.Range("g17").Value - 840) / 60 + 35

   
End Function

Sub OptionButton_Setting(i As Integer)


    Select Case i
    Case 38:
        Sheet4.Frame1.Controls("OptionButton11").Value = True
        myTime = 38
    Case 39:
        Sheet4.Frame1.Controls("OptionButton12").Value = True
        myTime = 39
    Case 40:
        Sheet4.Frame1.Controls("OptionButton13").Value = True
        myTime = 40
    Case 41:
        Sheet4.Frame1.Controls("OptionButton14").Value = True
        myTime = 41
    Case 42:
        Sheet4.Frame1.Controls("OptionButton15").Value = True
        myTime = 42
    Case 43:
        Sheet4.Frame1.Controls("OptionButton16").Value = True
        myTime = 43
    Case 44:
        Sheet4.Frame1.Controls("OptionButton17").Value = True
        myTime = 44
    Case Else:
        Sheet4.Frame1.Controls("OptionButton14").Value = True
        myTime = 41
    End Select


End Sub

Sub TimeSetting()
    Dim stable, h1, h2, myRandom As Integer
    Dim myRange As String
           
    stable = find_stable_time()
    
    If myTime = 0 Then
    
        myTime = initialize_myTime
        myRandom = myTime
        OptionButton_Setting (myTime)
        'Frame1.Controls("OptionButton14").Value = True
    Else
        myRandom = myTime
    End If
    
    If stable < myRandom Then
        h1 = stable
        h2 = myRandom
        Range("ab" & CStr(h1)).Select
        myRange = "AB" & CStr(h1) & ":AB" & CStr(h2)
        
    ElseIf stable > myRandom Then
        h1 = myRandom
        h2 = stable
        Range("ab" & CStr(h2 + 1)).Select
        myRange = "AB" & CStr(h1 + 1) & ":AB" & CStr(h2 + 1)
    Else
        Exit Sub
    End If
              
    
    Selection.AutoFill Destination:=Range(myRange), Type:=xlFillDefault
    setSkinTime (myTime)
    
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    
End Sub

Sub setSkinTime(i As Integer)

    Application.ScreenUpdating = False
    
    Sheet9.Activate
    Range("G17").Value = 840 + 60 * (i - 35)
    Sheet4.Activate
    
    Application.ScreenUpdating = True

End Sub

Sub setForRandomTime(i As Integer)

    Select Case i
    Case 38:
        Sheet4.Frame1.Controls("OptionButton11").Value = True
        myTime = 38
    Case 39:
        Sheet4.Frame1.Controls("OptionButton12").Value = True
        myTime = 39
    Case 40:
        Sheet4.Frame1.Controls("OptionButton13").Value = True
        myTime = 40
    Case 41:
        Sheet4.Frame1.Controls("OptionButton14").Value = True
        myTime = 41
    Case 42:
        Sheet4.Frame1.Controls("OptionButton15").Value = True
        myTime = 42
    Case 43:
        Sheet4.Frame1.Controls("OptionButton16").Value = True
        myTime = 43
    Case 44:
        Sheet4.Frame1.Controls("OptionButton17").Value = True
        myTime = 44
    Case Else:
        Sheet4.Frame1.Controls("OptionButton14").Value = True
        myTime = 41
    End Select


    Call setSkinTime(i)
    

End Sub

Sub RandomTimeSetting()
    Dim myRandom As Integer
    Dim stable, h1, h2 As Integer
    Dim myRange As String
           
    Randomize                                    'Initialize the Rnd function
     
    myRandom = CInt(38 + Rnd * 6)                'Generate a random number between 5-100
    'MsgBox CStr(myRandom)
    
    stable = find_stable_time()
    
    If stable < myRandom Then
        h1 = stable
        h2 = myRandom
        Range("ab" & CStr(h1)).Select
        myRange = "AB" & CStr(h1) & ":AB" & CStr(h2)
        
    ElseIf stable > myRandom Then
        h1 = myRandom
        h2 = stable
        Range("ab" & CStr(h2 + 1)).Select
        myRange = "AB" & CStr(h1 + 1) & ":AB" & CStr(h2 + 1)
    Else
        Exit Sub
    End If
              
    Selection.AutoFill Destination:=Range(myRange), Type:=xlFillDefault
    
    Call setForRandomTime(myRandom)
    
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    
End Sub

Sub cellRED(ByVal strcell As String)

    Range(strcell).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13209
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub

Sub cellBLACK(ByVal strcell As String)

    Range(strcell).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub

Sub resetValue()
    
    Range("o3").ClearContents
    Range("s1").Value = 0.1
    Range("k6").Value = 0.2
        
    Range("n3:n14").ClearContents
   

End Sub

Sub findAnswer_LongTest()
    
    If (Range("O3").Value > 0) Then Exit Sub
    
    Range("K10").GoalSeek goal:=0, ChangingCell:=Range("S1")
    Range("o3").Value = Abs(Range("j10").Value)
    
    If Range("k8").Value < 0 Then
        cellRED ("k8")
    Else
        cellBLACK ("k8")
    End If
    
    Sheet9.Range("d5").Value = Round(Range("S1").Value, 4)
    
End Sub

Sub check_LongTest()

    Dim igoal, k0, k1 As Double
    
    k1 = Range("k8").Value
    k0 = Range("k6").Value
    
    If k0 = k1 Then Exit Sub
    If k1 > 0 Then Exit Sub
    
    If k0 <> "" Then
        igoal = k0
    Else
        igoal = 0.3
    End If
    
    Range("k8").GoalSeek goal:=igoal, ChangingCell:=Range("n3")
     
    If Range("k8").Value < 0 Then
        cellRED ("k8")
    Else
        cellBLACK ("k8")
    End If
    

End Sub

Sub findAnswer_StepTest()
   
    Range("Q4:Q13").ClearContents
    Range("T4").Value = 0.1
    Range("G12").GoalSeek goal:=1#, ChangingCell:=Range("T4")
    
    If Range("J11").Value < 0 Then
        Call cellRED("J11")
    Else
        Call cellBLACK("J11")
    End If
    
End Sub

Sub check_StepTest()

    Dim igoal, nj As Double
    
    igoal = 0.12
    
    Do While (Range("J11").Value < 0 Or Range("j11").Value >= 50)
        Range("J11").GoalSeek goal:=igoal, ChangingCell:=Range("Q4")
        igoal = igoal + 0.1
    Loop
    
    If Range("J11").Value < 0 Then
        cellRED ("J11")
    Else
        cellBLACK ("J11")
    End If
    
    

End Sub

