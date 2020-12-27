Attribute VB_Name = "Module_ReplaceSumToRange"
Option Explicit

Sub Extract_Ranges_From_Formula(rng As Range)
    
    Dim rCell As Range
    Dim cellValue As String
    
    Dim openingParen As Integer
    Dim closingParen As Integer
    Dim colonParam As Integer
    
    Dim FirstValue As String
    Dim SecondValue As String
    
    
    'strRange = "C2:C3"
    
    For Each rCell In rng
    
        cellValue = rCell.Formula
    
        openingParen = InStr(cellValue, "(")
        colonParam = InStr(cellValue, ":")
        closingParen = InStr(cellValue, ")")
    
    
        FirstValue = Mid(cellValue, openingParen + 1, colonParam - openingParen - 1)
        SecondValue = Mid(cellValue, colonParam + 1, closingParen - colonParam - 1)
    
        Debug.Print FirstValue
        Debug.Print SecondValue
    
    Next rCell

End Sub


'Uses Range.Find to get a range of all find results within a worksheet
' Same as Find All from search dialog box
'
Function FindAll(rng As Range, What As Variant, Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False) As Range
    Dim SearchResult As Range
    Dim firstMatch As String
    With rng
        Set SearchResult = .Find(What, , LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
        If Not SearchResult Is Nothing Then
            firstMatch = SearchResult.Address
            Do
                If FindAll Is Nothing Then
                    Set FindAll = SearchResult
                Else
                    Set FindAll = Union(FindAll, SearchResult)
                End If
                Set SearchResult = .FindNext(SearchResult)
            Loop While Not SearchResult Is Nothing And SearchResult.Address <> firstMatch
        End If
    End With
End Function


Function RangeToStringArray(myRange As Range) As String()

    ReDim strArray(myRange.Cells.Count - 1) As String
    Dim idx As Long
    Dim c As Range
    
    For Each c In myRange
        strArray(idx) = c.Text
        idx = idx + 1
    Next c

    RangeToStringArray = strArray
End Function


Function SerializeRange(theRange As Excel.Range) As String()
    
    Dim cell As Range
    Dim values() As String
    Dim i As Integer
    
    i = 0
    
    ReDim values(theRange.Cells.Count)
    
    For Each cell In theRange
    
           values(i) = cell.Address
           i = i + 1
    Next cell
    
    SerializeRange = values
    
End Function

Sub serialize_test()
    Dim ar As Variant
    
    ar = SerializeRange(Range("C11:C14"))
    
    Debug.Print ar(0)
    
 End Sub



 Sub ReplaceSUMtoEachCell()

    Dim ws As Worksheet
    Dim iList As Range, iName, rCell  As Range
    Dim aLR As Long, cLR As Long
    
    Dim strRange As Variant
    Dim strResult As String
    Dim i As Integer
    
    Dim cellValue, Value As String
    Dim openingParen As Integer
    Dim closingParen As Integer
    Dim colonParam As Integer
    
    
    

    Set ws = ThisWorkbook.ActiveSheet

    Set iList = FindAll(ws.UsedRange, "SUM", xlFormulas, xlPart)
    
'    For Each iName In iList
'        Debug.Print iName.Address
'    Next
'
        
    
    For Each rCell In iList
    
        cellValue = rCell.Formula
    
        openingParen = InStr(cellValue, "(")
        closingParen = InStr(cellValue, ")")
    
    
        Value = Mid(cellValue, openingParen + 1, closingParen - openingParen - 1)
        'Debug.Print Value, " ", cellValue
    
        strResult = "="
        strRange = SerializeRange(Range(Value))
        
        For i = 0 To UBound(strRange)
            strResult = strResult & strRange(i) & "+"
        Next i
        
        strResult = Left(strResult, Len(strResult) - 2)
        'Debug.Print strResult
        
        strResult = Replace(strResult, "$", "")
        rCell.Formula = strResult

    Next rCell
    


End Sub
    
    
    
    
