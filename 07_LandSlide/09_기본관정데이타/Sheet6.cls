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


Private Sub Set_RechargeFactor_One()

    Range("F17").Formula = "=(max*rf_1*E17/1000)"
    Range("F19").Formula = "=(max*rf_1*E19/1000)/365"
    
    Range("G17").Formula = "=F17*allow_ratio"
    Range("G19").Formula = "=F19*allow_ratio"
    
    Range("E13").Formula = "=Recharge!I24"
    Range("F13").Formula = "=rf_1"
    Range("G13").Formula = "=allow_ratio"
    
    Range("E26").Formula = "=Recharge!C30"
    
End Sub

Private Sub Set_RechargeFactor_Two()

    Range("F17").Formula = "=(max*rf_2*E17/1000)"
    Range("F19").Formula = "=(max*rf_2*E19/1000)/365"
    
    Range("G17").Formula = "=F17*allow_ratio2"
    Range("G19").Formula = "=F19*allow_ratio2"
    
    
    Range("E13").Formula = "=Recharge!I25"
    Range("F13").Formula = "=rf_2"
    Range("G13").Formula = "=allow_ratio2"
    
    
    
    Range("E26").Formula = "=Recharge!D30"
End Sub


Private Sub Set_RechargeFactor_Three()

    Range("F17").Formula = "=(max*rf_3*E17/1000)"
    Range("F19").Formula = "=(max*rf_3*E19/1000)/365"
    
    Range("G17").Formula = "=F17*allow_ratio3"
    Range("G19").Formula = "=F19*allow_ratio3"
    
    Range("E13").Formula = "=Recharge!I26"
    Range("F13").Formula = "=rf_3"
    Range("G13").Formula = "=allow_ratio3"
    
    Range("E26").Formula = "=Recharge!E30"
    
End Sub



Private Sub CommandButton6_Click()
'Select Recharge Factor

    
   If Frame1.Controls("optionbutton1").value = True Then
        Call Set_RechargeFactor_One
   End If
    
   If Frame1.Controls("optionbutton2").value = True Then
        Call Set_RechargeFactor_Two
   End If
    
   If Frame1.Controls("optionbutton3").value = True Then
        Call Set_RechargeFactor_Three
   End If
    

End Sub

Private Sub CommandButton7_Click()
   Call getWhpaData_EachWell
End Sub

Private Sub Worksheet_Activate()

    Select Case get_rf_number
    
        Case "1"
             Frame1.Controls("optionbutton1").value = True
             
        Case "2"
             Frame1.Controls("optionbutton2").value = True
             
        Case "3"
        
        Case Else
            Frame1.Controls("optionbutton1").value = True
           
    End Select

End Sub


