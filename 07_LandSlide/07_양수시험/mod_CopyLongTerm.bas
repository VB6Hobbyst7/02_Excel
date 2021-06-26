Attribute VB_Name = "mod_CopyLongTerm"
Sub make_step_document()
    '
    ' StepTest ∫πªÁ
    ' select last sheet -- Sheets(Sheets.Count).Select

    Application.ScreenUpdating = False
    
    Sheets("StepTest").Select
    Sheets("StepTest").Copy Before:=Sheets(14)
  
    Application.Goto Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Columns("J:AO").Select
    Selection.Delete Shift:=xlToLeft
      
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
  
    Application.Goto Reference:="Print_Area"
    With Selection.Font
        .name = "∏º¿∫ ∞ÌµÒ"
    End With
    
    Range("J19").Select
  
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
End Sub


Sub make_long_document()
    '
    ' LongTest∫πªÁ ∏≈≈©∑Œ
   '

    Application.ScreenUpdating = False

    shLongTermTest.Select
    shLongTermTest.Copy Before:=Sheets(Sheets.Count)
    
    Application.Goto Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
        
    Columns("J:AP").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll Down:=96
    
    Rows("102:264").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-102
    
        
    Application.Goto Reference:="Print_Area"
    With Selection.Font
        .name = "∏º¿∫ ∞ÌµÒ"
        .ThemeFont = xlThemeFontNone
    End With
       
    Range("J6").Select
    
    ActiveSheet.Shapes.Range(Array("Frame1")).Select
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
    
    Application.CutCopyMode = False
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




