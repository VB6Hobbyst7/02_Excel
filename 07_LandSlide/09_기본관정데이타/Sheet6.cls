VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call find_average
End Sub

Private Sub CommandButton2_Click()
    Call main_drasticindex
    Call print_drastic_string
End Sub

Private Sub CommandButton3_Click()
    Call getWhpaData_AllWell
End Sub


Private Sub CommandButton4_Click()
    Call delete_allWhpaData
End Sub

Private Sub CommandButton5_Click()
    Call ToggleDirection
End Sub


Private Function get_rf_number() As String
    Dim rf_num As String

    '=(max*rf_1*E17/1000)
    get_rf_number = VBA.Mid(Range("F17").Formula, 10, 1)

End Function

Private Sub Set_RechargeFactor_1to2()

    Range("F17").Formula = "=(max*rf_2*E17/1000)"
    Range("F19").Formula = "=(max*rf_2*E19/1000)/365"
    
    Range("G17").Formula = "=F17*allow_ratio2"
    Range("G19").Formula = "=F19*allow_ratio2"
    
    
    Range("E13").Formula = "=Recharge!G25"
    Range("F13").Formula = "=rf_2"
    Range("G13").Formula = "=allow_ratio2"
    
    
    
    Range("E26").Formula = "=Recharge!D30"
End Sub

Private Sub Set_RechargeFactor_2to1()

    Range("F17").Formula = "=(max*rf_1*E17/1000)"
    Range("F19").Formula = "=(max*rf_1*E19/1000)/365"
    
    Range("G17").Formula = "=F17*allow_ratio"
    Range("G19").Formula = "=F19*allow_ratio"
    
    Range("E13").Formula = "=Recharge!G24"
    Range("F13").Formula = "=rf_1"
    Range("G13").Formula = "=allow_ratio"
    
    
    
    Range("E26").Formula = "=Recharge!C30"
End Sub


Private Sub CommandButton6_Click()
'Select Recharge Factor

    If Frame1.Controls("optionbutton1").value = True Then
        Call Set_RechargeFactor_2to1
    Else
        Call Set_RechargeFactor_1to2
    End If

End Sub

Private Sub CommandButton7_Click()
   Call getWhpaData_EachWell
End Sub

Private Sub Worksheet_Activate()

    ' MsgBox ("hello")
    If get_rf_number = "1" Then
         Frame1.Controls("optionbutton1").value = True
    Else
        Frame1.Controls("optionbutton2").value = True
    End If

End Sub


