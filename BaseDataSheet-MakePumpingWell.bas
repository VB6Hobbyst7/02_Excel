Attribute VB_Name = "MakePumpingWell"

Option Explicit

Sub CopyOneSheet()

    Dim n_sheets As Integer

    n_sheets = sheets_count()
    
    If (n_sheets = 1) Then
        Sheets("1").Select
        Sheets("1").Copy Before:=Sheets("Q1")
        
        ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
    Else
        Sheets("2").Select
        Sheets("2").Copy Before:=Sheets("Q1")
    End If
    
    ActiveSheet.Name = CStr(n_sheets + 1)
    Range("b2").value = "W-" & (n_sheets + 1)
    Range("e15").value = CStr(n_sheets + 1)
      
    
    If n_sheets = 1 Then
        Call ChangeCellData(n_sheets + 1, 1)
    Else
        Call ChangeCellData(n_sheets + 1, 2)
    End If
    
End Sub

Sub ChangeCellData(ByVal nsheet As Integer, ByVal nselect As Integer)
'
' change sheet data direct to well sheet data value
'
'https://stackoverflow.com/questions/18744537/vba-setting-the-formula-for-a-cell

    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate
    
    nsheet = nsheet + 3
    Selection.Replace What:=CStr(nselect + 3), Replacement:=CStr(nsheet), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 
    
    Range("E21").Select
    Range("E21").Formula = "=Well!" & Cells(nsheet, 9).Address
    
End Sub










