Attribute VB_Name = "water_GenerateCopy"
Option Explicit

Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function


Private Sub DoCopy(lastRow As Long)
Attribute DoCopy.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("F2:H" & lastRow).Select
    Selection.Copy
    
    Range("M2").Select
    ActiveSheet.Paste
    
    
    Range("K2:K" & lastRow).Select
    Selection.Copy
    
    Range("P2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("J2:J" & lastRow).Select
    Selection.Copy
    
    Range("Q2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
    
End Sub


Private Sub CleanSection(lastRow As Long)

    Range("M2:Q" & lastRow).Select
    Selection.ClearContents
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





