Attribute VB_Name = "SavePrn"

Public Function MyDocsPath() As String
   
    MyDocsPath = Environ$("USERPROFILE") & "\\" & "My Documents"
   
End Function

Sub janggi_01()
    
    ActiveWorkbook.SaveAs Filename:= _
                          "janggi_01.prn", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False

End Sub

Sub janggi_02()
    
    ActiveWorkbook.SaveAs Filename:= _
                          "janggi_02.prn", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False

End Sub

Sub recover_01()
    
    ActiveWorkbook.SaveAs Filename:= _
                          "recover_01.prn", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False

End Sub

Sub step_01()
    
    Range("a1").Select
    
    ActiveWorkbook.SaveAs Filename:= _
                          "step_01.prn", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False

End Sub

Sub save_original()

    ActiveWorkbook.SaveAs Filename:="save_original.xlsm", FileFormat:= _
                          xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

End Sub


