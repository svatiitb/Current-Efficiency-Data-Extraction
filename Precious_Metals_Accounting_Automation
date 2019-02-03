Option Explicit

Sub format()
    Application.ScreenUpdating = False
    Dim i As Integer, c As Integer
    Dim Au_1 As Double, Au_2 As Double
    Dim folderpath As String, filename As String, fil As String, cnsgt As String
    Dim awb As Workbook, owb As Workbook
    
    folderpath = Application.ActiveWorkbook.Path
    fil = "Assay Database updated.xls"
    filename = folderpath + "\" + fil
    Set awb = ThisWorkbook
    
    Workbooks.Open filename:=filename, UpdateLinks:=0
    Set owb = Workbooks("Assay Database updated.xls")
    'Windows(fil).Activate
    'Sheets("Latest assay").Select
    c = 4
    For i = 1 To owb.Sheets("Latest assay").Range("A" & Rows.Count).End(xlUp).Row
        'Windows(fil).Activate
        'Sheets("Latest assay").Select
        cnsgt = owb.Sheets("Latest assay").Range("B" & i)
        'awb.Activate
        If cnsgt = awb.Sheets("P to F Concentrate 17-18").Range("D" & c) Then
            'Windows(fil).Activate
            'Sheets("Latest assay").Select
            Au_1 = owb.Sheets("Latest assay").Range("D" & i).Value
            Au_2 = owb.Sheets("Latest assay").Range("E" & i).Value
            awb.Activate
            If Range("F" & c).Value = Au_1 Then
                Range("F" & c, "G" & c).Select
                With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                End With
            ElseIf Range("F" & c).Value = Au_2 Then
                Range("F" & c, "G" & c).Select
                With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                End With
                Selection.Font.Bold = True
            Else
                Range("F" & c, "G" & c).Select
                With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
                End With
            End If
            c = c + 1
        End If
    Next
    Workbooks(fil).Close Savechanges:=False
    
    For i = 1 To awb.Sheets("P to F Concentrate 17-18").Range("B" & Rows.Count).End(xlUp).Row
        If Range("B" & i) = "PROV" Then
            Range("A" & i & ":" & "T" & i).Select
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
            End With
        End If
    Next
    Application.ScreenUpdating = True
    MsgBox ("Formatting done!")
End Sub
