Option Explicit
Option Base 1
Public Sheet As String, Workbook As String
Public sht As Worksheet, wrkb As Workbook
Sub Populate_CS()
    
    Call delete_rows
    Dim i As Integer
    For i = 1 To Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
      Sheets("Copper Sulphate").Range("B" & i + 4) = Sheets("Sheet1").Range("A" & i + 2)
      Sheets("Copper Sulphate").Range("C" & i + 4) = Sheets("Sheet1").Range("D" & i + 2)
      Sheets("Copper Sulphate").Range("C" & i + 4).Offset(0, Sheets("Sheet1").Range("C" & i + 2).Value) = Sheets("Sheet1").Range("G" & i + 2).Value
    Next
    
End Sub
Sub Populate_R()
    
    Call delete_rows
    Dim i As Integer
    For i = 1 To Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
      Sheets("Roughness").Range("B" & i + 4) = Sheets("Sheet1").Range("A" & i + 2)
      Sheets("Roughness").Range("C" & i + 4) = Sheets("Sheet1").Range("D" & i + 2)
      Sheets("Roughness").Range("C" & i + 4).Offset(0, Sheets("Sheet1").Range("C" & i + 2).Value) = Sheets("Sheet1").Range("F" & i + 2).Value
    Next
    
    
End Sub
Sub Populate_TN()
    
    Call delete_rows
    Dim i As Integer
    For i = 1 To Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
      Sheets("Top Nodules").Range("B" & i + 4) = Sheets("Sheet1").Range("A" & i + 2)
      Sheets("Top Nodules").Range("C" & i + 4) = Sheets("Sheet1").Range("D" & i + 2)
      Sheets("Top Nodules").Range("C" & i + 4).Offset(0, Sheets("Sheet1").Range("C" & i + 2).Value) = Sheets("Sheet1").Range("E" & i + 2).Value
    Next
    
    
End Sub
Sub delete_rows()
    
On Error GoTo 10
    Sheets("Sheet1").Range("A3" & ":" & "H" & Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row).Select
    Selection.SpecialCells(xlCellTypeFormulas, 16).Select
    Selection.EntireRow.Delete
10:     MsgBox ("Unwanted rows deleted")
    
End Sub
