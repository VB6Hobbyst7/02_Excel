Attribute VB_Name = "BaseData_MakePumpingWell"

Option Explicit


'쉬트를 생성할때에는 전체 관정데이타를 건들지 않고, 우선먼저 쉬트복제를 누르는것이 기본으로 정해져 있다.


Dim ColorValue(1 To 20) As Long


Public Sub InitialSetColorValue()

    ColorValue(1) = RGB(192, 0, 0)
    ColorValue(2) = RGB(255, 0, 0)
    ColorValue(3) = RGB(255, 192, 0)
    ColorValue(4) = RGB(255, 255, 0)
    ColorValue(5) = RGB(146, 208, 80)
    ColorValue(6) = RGB(0, 176, 80)
    ColorValue(7) = RGB(0, 176, 240)
    ColorValue(8) = RGB(0, 112, 192)
    ColorValue(9) = RGB(0, 32, 96)
    ColorValue(10) = RGB(112, 48, 160)
    
    ColorValue(11) = RGB(192 + 10, 10, 0)
    ColorValue(12) = RGB(255, 0 + 10, 0)
    ColorValue(13) = RGB(255, 192 + 10, 0)
    ColorValue(14) = RGB(255, 255, 10)
    ColorValue(15) = RGB(146 + 10, 208 + 10, 80 + 10)
    ColorValue(16) = RGB(0 + 10, 176 + 10, 80)
    ColorValue(17) = RGB(0 + 10, 176 + 10, 240 + 10)
    ColorValue(18) = RGB(0 + 10, 112 + 10, 192)
    ColorValue(19) = RGB(0 + 10, 32 + 10, 96)
    ColorValue(20) = RGB(112, 48 + 10, 160 + 10)

End Sub



Private Sub deleteCommandButton()

    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton7")).Select
    Selection.Delete
End Sub


Public Sub CopyOneSheet()

    Dim n_sheets As Integer

    n_sheets = sheets_count()
    
    '2020/5/30 관정리스트의 목록삽입해주는 부분 추가
    InsertOneRow (n_sheets)
    
    
    If (n_sheets = 1) Then
        Sheets("1").Select
        Sheets("1").Copy Before:=Sheets("Q1")
        Call deleteCommandButton
    Else
        Sheets("2").Select
        Sheets("2").Copy Before:=Sheets("Q1")
    End If
    
    ActiveSheet.name = CStr(n_sheets + 1)
    Range("b2").value = "W-" & (n_sheets + 1)
    Range("e15").value = CStr(n_sheets + 1)
      
    
    If n_sheets = 1 Then
        Call ChangeCellData(n_sheets + 1, 1)
    Else
        Call ChangeCellData(n_sheets + 1, 2)
    End If
    
    Sheets("Well").Select
End Sub

Private Sub InsertOneRow(ByVal n_sheets As Integer)

    n_sheets = n_sheets + 4
    Rows(n_sheets & ":" & n_sheets).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    Rows(CStr(n_sheets - 1) & ":" & CStr(n_sheets - 1)).Select
    Selection.Copy
    Rows(CStr(n_sheets) & ":" & CStr(n_sheets)).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False

End Sub

Private Sub ChangeCellData(ByVal nsheet As Integer, ByVal nselect As Integer)
'
' change sheet data direct to well sheet data value
' https://stackoverflow.com/questions/18744537/vba-setting-the-formula-for-a-cell

    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
        
    nsheet = nsheet + 3
    Selection.Replace What:=CStr(nselect + 3), Replacement:=CStr(nsheet), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 
    
    Range("E21").Select
    'Range("E21").Formula = "=Well!" & Cells(nsheet, 9).Address
    Range("E21").Formula = "=Well!" & Cells(nsheet, "I").Address
    
End Sub


Private Sub JojungData(ByVal nsheet As Integer)

    Dim nselect As String

    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate

    nsheet = nsheet + 3
    '=Well!D7
    nselect = Mid(Range("c2").Formula, 8)
    
    'Debug.Print Mid(Range("c2").Formula, 8) & ":" & nselect

    Selection.Replace What:=nselect, Replacement:=CStr(nsheet), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Range("E21").Select
    Range("E21").Formula = "=Well!" & Cells(nsheet, "I").Address
        
End Sub


Private Sub SetMyTabColor(ByVal index As Integer)

    With ActiveWorkbook.Sheets(CStr(index)).Tab
        .Color = ColorValue(index)
        .TintAndShade = 0
    End With

End Sub



'각각의 쉬트를 순회하면서, 셀의 참조값을 맟추어준다.
'
Public Sub JojungSheetData()

    Dim n_sheets As Integer
    Dim i As Integer

    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Sheets(CStr(i)).Activate
        Call JojungData(i)
        Call SetMyTabColor(i)
    Next i
    
End Sub
















