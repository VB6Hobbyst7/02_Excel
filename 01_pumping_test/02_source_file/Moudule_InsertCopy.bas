Attribute VB_Name = "Moudule_InsertCopy"
Option Explicit

Public Enum CellCopyMode
    ccmValue = 0
    ccmFormula = 1
End Enum



Private Sub insert_in_a_sheet_by_input()
    Dim Sel As Range
    Dim strReturn As String
    Dim retVal As Variant
    
    Dim ns, ne As Long
    Dim rp, i As Long
    
    Set Sel = Application.InputBox("원하는 영역에 값적용", "범위선택", Type:=8)
       
    retVal = ExtractStartEnd_FromRange(Sel)
    
    ns = Val(retVal(0))
    ne = Val(retVal(1))
    
       
    rp = ns
    For i = ns To ne
        
        
        retVal = Application.Run("MainMoudule.insert_row2", rp)
        Debug.Print retVal
        
        rp = Selection.row + 1
        
    Next i
     
End Sub


Sub Relative2Absolute()
    For Each c In Selection
        If c.HasFormula = True Then
            c.Formula = Application.ConvertFormula(c.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next c
End Sub



Private Function ConvertAbsoluteAddress(rng As Range) As String
    Dim cell As Range
            
    For Each cell In rng
        If cell.HasFormula = True Then
                cell.Formula = Application.ConvertFormula(cell.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next rng
        
    ConvertAbsoluteAddress = cell.Formula
End Function


Private Sub test()
    Dim strText As String
        
    strText = ConvertAbsoluteAddress(Range("C5"))
    Debug.Print strText
End Sub


Private Sub NoSelect()
   With ActiveSheet
   .EnableSelection = xlNoSelection
   .Protect
   End With
End Sub


'nrow - copy down row
'ns - start column
'ne - end column

Sub MakeCopyCurrentRow(nrow As Long, Optional ByVal ns As String = "C", Optional ByVal ne As String = "I")
    
    Range(ns & CStr(nrow) & ":" & ne & CStr(nrow)).Select
    Selection.AutoFill Destination:=Range(ns & CStr(nrow) & ":" & ne & CStr(nrow + 1)), Type:=xlFillDefault
    Call NoSelect
    
End Sub

Private Sub callingPrivateSubTest()

    Dim i As Integer

    i = Application.Run("MainMoudule.getSheetIndex", "계산서")
    Debug.Print "hello", i
    
End Sub





