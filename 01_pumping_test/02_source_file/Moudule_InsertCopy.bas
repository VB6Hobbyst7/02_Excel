Attribute VB_Name = "Moudule_InsertCopy"
Option Explicit


Sub insert_in_a_sheet_by_input()
    Dim ReturnSel As Range
    Dim strReturn As String
    Dim retVal As Variant
    
    Dim ns, ne As Long
    Dim rp, i As Long
    
    Set ReturnSel = Application.InputBox("원하는 영역에 값적용", "범위선택", Type:=8)
       
    retVal = ExtractStartEnd_FromRange(ReturnSel)
    
    ns = Val(retVal(0))
    ne = Val(retVal(1))
    
       
    rp = ns
    For i = ns To ne
        
        Call insert_row2(rp)
        rp = Selection.row + 1
        
    Next i
     
End Sub






