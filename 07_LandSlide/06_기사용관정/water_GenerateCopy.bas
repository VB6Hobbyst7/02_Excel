Attribute VB_Name = "water_GenerateCopy"
Option Explicit

Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function


Private Sub DoCopy(lastRow As Long)
Attribute DoCopy.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("F2:H" & lastRow).Select
    selection.Copy
    
    Range("M2").Select
    ActiveSheet.Paste
    
    
    Range("K2:K" & lastRow).Select
    selection.Copy
    
    Range("P2").Select
    selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("J2:J" & lastRow).Select
    selection.Copy
    
    Range("Q2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
    
End Sub


Private Sub CleanSection(lastRow As Long)

    Range("M2:Q" & lastRow).Select
    selection.ClearContents
    Range("P14").Select
    
End Sub

Sub MainMoudleGenerateCopy()

    Dim lastRow As Long
        
    lastRow = lastRowByKey("I1")
    Call DoCopy(lastRow)


End Sub


Sub SubModuleCleanCopySection()

    Dim lastRow As Long
        
    lastRow = lastRowByKey("I1")
    Call CleanSection(lastRow)
    
End Sub

Sub insertRow()
'

    Dim lastRow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    
    lastRow = lastRowByKey("A1")
    
    Rows(CStr(lastRow + 1) & ":" & CStr(lastRow + 2)).Select
    selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1"): j = i + 2
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "J" & i & ":L" & i
    selection_target = "J" & i & ":L" & j

    Range(selection_origin).Select
    selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("R" & i).Select
    selection.AutoFill Destination:=Range("R" & i & ":R" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False

End Sub



