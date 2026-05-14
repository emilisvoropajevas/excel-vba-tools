Attribute VB_Name = "Module1"
Option Explicit

Sub PullData()
    GetFileList
End Sub

Sub GetFileList()

    Dim filePath As String
    Dim cmdOutput As String
    Dim filesArray() As String
    Dim i As Long
    Dim nextRow As Long
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Test_WorkingSheet")

    ' Get folder path
    filePath = Trim(ThisWorkbook.Sheets("Introduction").Range("D3").Value)

    ' Remove trailing slash if present
    If Right(filePath, 1) = "\" Then
        filePath = Left(filePath, Len(filePath) - 1)
    End If

    ' Validate path
    If Dir(filePath, vbDirectory) = "" Then
        MsgBox "Folder path not found:" & vbCrLf & filePath, vbCritical
        Exit Sub
    End If

    ' Clear old data
    ws.Range("A3:B" & ws.Rows.Count).ClearContents

    ' Get file list
    cmdOutput = ExecCmd("cmd /c dir /b """ & filePath & """")

    ' Split into array
    filesArray = Split(cmdOutput, vbCrLf)

    nextRow = 3

    ' Paste file names
    For i = LBound(filesArray) To UBound(filesArray)

        If Trim(filesArray(i)) <> "" Then
            ws.Cells(nextRow, 1).Value = filesArray(i)
            nextRow = nextRow + 1
        End If

    Next i

    ' Sort list
    SortFileList

    MsgBox "Data Imported", vbInformation

End Sub


Function ExecCmd(cmd As String) As String

    Dim objShell As Object
    Dim objExec As Object
    Dim cmdOutput As String

    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec(cmd)

    ' Wait until command finishes
    Do While objExec.Status = 0
        DoEvents
    Loop

    cmdOutput = objExec.StdOut.ReadAll

    ExecCmd = cmdOutput

    Set objExec = Nothing
    Set objShell = Nothing

End Function


Sub SortFileList()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fileName As String
    Dim nameWithoutExt As String

    Set ws = ThisWorkbook.Sheets("Test_WorkingSheet")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Create numeric helper column
    For i = 3 To lastRow

        fileName = ws.Cells(i, 1).Value

        If InStr(fileName, ".") > 0 Then
            nameWithoutExt = Left(fileName, InStrRev(fileName, ".") - 1)

            ' Extract numeric value
            ws.Cells(i, 2).Value = Val(nameWithoutExt)
        End If

    Next i

    ' Sort
    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            Key:=ws.Range("B3:B" & lastRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange ws.Range("A3:B" & lastRow)

        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom

        .Apply

    End With

    ' Remove helper column
    ws.Range("B3:B" & lastRow).ClearContents

End Sub

