Attribute VB_Name = "Module4"
Sub Transfer_Data()
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    Dim mappings(1 To 15, 1 To 2) As String
    mappings(1, 1) = "B" ' Source
    mappings(1, 2) = "C" ' Destination
    
    mappings(2, 1) = "E"
    mappings(2, 2) = "K"
    
    mappings(3, 1) = "F"
    mappings(3, 2) = "N"
    
    mappings(4, 1) = "G"
    mappings(4, 2) = "A"
    
    mappings(5, 1) = "H"
    mappings(5, 2) = "E"
    
    mappings(6, 1) = "I"
    mappings(6, 2) = "F"
    
    mappings(7, 1) = "J"
    mappings(7, 2) = "G"
    
    mappings(8, 1) = "K"
    mappings(8, 2) = "D"
    
    mappings(9, 1) = "L"
    mappings(9, 2) = "H"
    
    mappings(10, 1) = "M"
    mappings(10, 2) = "I"
    
    mappings(11, 1) = "N"
    mappings(11, 2) = "J"
    
    mappings(12, 1) = "O"
    mappings(12, 2) = "O"
    
    mappings(13, 1) = "P"
    mappings(13, 2) = "L"
    
    mappings(14, 1) = "Q"
    mappings(14, 2) = "B"
    
    mappings(15, 1) = "R"
    mappings(15, 2) = "M"
    
    Call Open_Working_Sheet
    
    ' Set the source and destination sheets
    Set srcSheet = ThisWorkbook.Sheets("Test_WorkingSheet")
    Set destSheet = ThisWorkbook.Sheets("Working Sheet")
    
    ' Find the last row with data in the source sheet
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "B").End(xlUp).Row
    
    For i = 3 To lastRow
        For j = 1 To 15
            destSheet.Cells(i - 1, mappings(j, 2)).Value = srcSheet.Cells(i, mappings(j, 1)).Value
        Next j
    Next i

    MsgBox "Data has been copied successfully!"
End Sub

Sub Open_Working_Sheet()
    ThisWorkbook.Sheets("Working Sheet").Activate
End Sub


