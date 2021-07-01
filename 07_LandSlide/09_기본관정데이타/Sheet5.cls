VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub CommandButton1_Click()
    Call CopyOneSheet
End Sub

Private Sub CommandButton2_Click()
    Call JojungSheetData
    Call make_wellstyle
End Sub

Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub

