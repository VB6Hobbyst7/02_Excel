Attribute VB_Name = "CopyLongTerm"
Sub make_step_document()
    '
    ' 단계양수시험 복사
    '

    '
    Application.ScreenUpdating = False
    
    Sheets("단계양수시험").Select
    Sheets("단계양수시험").Copy Before:=Sheets(12)
    Application.GoTo Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("J:AU").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll Down:=-48
    Application.GoTo Reference:="Print_Area"
    With Selection.Font
        .name = "맑은 고딕"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("H2").Select
    
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    
    Application.ScreenUpdating = True
    
End Sub

Sub make_long_document()
    '
    ' 장기양수시험복사 매크로
    '

    Application.ScreenUpdating = False

    Sheets("장기양수시험").Select
    Sheets("장기양수시험").Copy Before:=Sheets(12)
    
    Application.GoTo Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("J:AP").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll Down:=96
    
    Rows("102:264").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-102
    
    Range("H5").Select
    Application.GoTo Reference:="Print_Area"
    With Selection.Font
        .name = "맑은 고딕"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    
    Application.GoTo Reference:="Print_Area"
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("J6").Select
    
    ActiveSheet.Shapes.Range(Array("Frame1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton6")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton7")).Select
    Selection.Delete
    
    
    Application.ScreenUpdating = True
    
    Call modify_cell_value
    
End Sub

'2019/11/24

Sub modify_cell_value()

    Dim i As Integer, j As Integer
    
    For i = 10 To 101
            
        Cells(i, "F").Value = Round(Cells(i, "F").Value, 2)
        Cells(i, "G").Value = Round(Cells(i, "G").Value, 2)
        
    Next i
    

End Sub


