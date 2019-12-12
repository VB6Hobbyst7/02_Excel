Attribute VB_Name = "ShowHide"
Sub show_gachae()
    Sheets("가채수량").Visible = True
    Sheets("가채수량").Select
End Sub

Sub hide_gachae()
    Sheets("가채수량").Visible = False
    Sheets("스킨팩터").Select
End Sub

