Attribute VB_Name = "Module3"
Function GetUserDirectory() As String
    GetUserDirectory = Environ("USERPROFILE") & "\"
End Function

Sub GetCurrentUserDirectory()
    If IsEmpty(Range("A3").Value) Then
        MsgBox "Please enter folder name", vbExclamation
        Exit Sub
    End If
    
    Range("B3").Value = GetUserDirectory()
    
    Call Intro_End
End Sub

Sub Intro_End()
    ThisWorkbook.Sheets("Test_WorkingSheet").Activate
End Sub




