Attribute VB_Name = "BaseData_ETC"
Option Explicit


Public Function GetLengthByColor(ByVal tabColor As Variant) As Integer
    Dim n_sheets, i, j, nTab As Integer
    n_sheets = sheets_count()
     
    nTab = 0
    
    For i = 1 To n_sheets
        If (Sheets(CStr(i)).Tab.Color = tabColor) Then
            nTab = nTab + 1
        End If
    Next i
    
    GetLengthByColor = nTab
End Function


Private Sub get_tabsize_by_well(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Variant, ByRef n_tabcolors As Variant)
    ' n_tabcolors : return value
    ' nof_unique_tab : return value

    Dim n_sheets, i, j As Integer
    Dim limit() As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
   
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim limit(0 To n_sheets)
    
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.Color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    
    For i = 0 To UBound(new_tabcolors)
        limit(i) = GetLengthByColor(new_tabcolors(i))
    Next i
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
    n_tabcolors = new_tabcolors
    
End Sub

