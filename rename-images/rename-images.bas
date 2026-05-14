Attribute VB_Name = "Module2"
Sub RenameImages()
    Dim ws As Worksheet
    Dim oldName As String
    Dim newName As String
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Test_WorkingSheet")

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Loop through each row in the table
    For i = 3 To lastRow
        ' Read the old and new names from the table
        oldName = ws.Cells(i, 4).Value
        newName = ws.Cells(i, 5).Value

        ' Rename file if both names are provided
        If oldName <> "" And newName <> "" Then
            RenameFile oldName, newName
        End If
    Next i
    
    Set ws = Nothing
    MsgBox "Image renaming process completed."
End Sub

Sub RenameFile(oldName As String, newName As String)
    Dim cmd As String

    ' Ensure the old name contains the full path
    If InStr(1, oldName, ":\", vbTextCompare) = 0 Then
        MsgBox "Invalid path: " & oldName
        Exit Sub
    End If

    ' Append the file extension to the new name if not provided
    If InStr(1, newName, ".jpg", vbTextCompare) = 0 Then
        newName = newName & ".jpg"
    End If

    ' Create the command string
    cmd = "cmd.exe /c ren """ & oldName & """ """ & newName & """"

    ' Run the command
    ExecCmd cmd
End Sub

