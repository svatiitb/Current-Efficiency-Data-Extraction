Option Explicit

Sub extract_data_1()
    Dim awb As Workbook
    Dim loc As Variant, fil As Variant, filename As Variant, dat As Variant, A As Variant, B As Variant
    Dim CuEf() As Double
    Dim location As Variant
    Dim month As String
    Dim n As Integer, m As Integer, i As Integer, c As Integer, j As Integer, k As Integer
    Dim vol As Double, CE As Double, sum As Double, avg As Double
    On Error Resume Next
    loc = InputBox("Please enter the location of the file")
    fil = InputBox("Please enter the name of the file")
    Err.Clear
    On Error GoTo 0
    filename = fil + ".xlsx"
    location = loc + "\" + fil + ".xlsx"
    Set awb = ThisWorkbook
    

    month = Mid(fil, 5, Len(fil) - 7)
    
    If month = "Jan" Or month = "Mar" Or month = "May" Or month = "July" Or month = "Aug" Or month = "Oct" Or month = "Dec" Then
        m = 33
    ElseIf month = "April" Or month = "Jun" Or month = "Sept" Or month = "Nov" Then
        m = 32
    Else: m = 30
    End If
    
    Application.ScreenUpdating = False
    Workbooks.Open filename:=location, Password:="1234", UpdateLinks:=0
    Windows(filename).Activate
    Sheets("Elec composition").Select
    Range("A3:A" & m).Select
    Selection.Copy
    awb.Activate
    Range("B3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("C3:C" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C3:C" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("C3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("F3:F" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("D3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("I3:I30").Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("E3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("BN3:BN30").Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("F3").Select
    ActiveSheet.Paste
    
    Windows(filename).Activate
    Sheets("Vol Cir 1").Select
    i = 1
    A = Format(Range("A" & i + 7), "mmm")
    B = Left(month, 3)
    Do While A = B
        If Range("A" & i + 7).Value <> Range("A" & i + 8).Value Then
        dat = Range("A" & i + 7).Value
        Range("N" & i + 7).Select
        Application.CutCopyMode = False
        Selection.Copy
        awb.Activate
        For n = 3 To Range("B" & Rows.Count).End(xlUp).Row
            If Range("B" & n).Value = dat Then
            Range("G" & n).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            End If
        Next
        End If
        Windows(filename).Activate
        Sheets("Vol Cir 1").Select
        i = i + 1
        A = Format(Range("A" & i + 7), "mmm")
        B = Left(month, 3)
    Loop
    
    awb.Activate
    Sheets("Cir 1").Select
    i = 3
    j = 8
    c = 1
    Do While i < Range("B" & Rows.Count).End(xlUp).Row + 1
        dat = Range("B" & i).Value
        Windows(filename).Activate
        Sheets("CE").Select
        For j = 8 To Range("B" & Rows.Count).End(xlUp).Row
            If dat = Range("B" & j).Value And Range("C" & j).Value < 9 Then
            ReDim Preserve CuEf(c) As Double
            CuEf(c) = Range("M" & j).Value
            c = c + 1
            End If
        Next
        sum = 0
        For k = 1 To c
            On Error Resume Next
            sum = sum + CuEf(k)
        Next
        CE = sum / (c - 1)
        If CE > 1000 Then
        CE = 0
        End If
        awb.Activate
        Sheets("Cir 1").Select
        Range("H" & i).Value = CE
        c = 1
        i = i + 1
    Loop
    
    Workbooks(filename).Close Savechanges:=False
    Application.ScreenUpdating = True
    MsgBox ("OK!")
    
End Sub
Sub extract_data_2()
    Dim awb As Workbook
    Dim loc As Variant, fil As Variant, filename As Variant, dat As Variant, A As Variant, B As Variant
    Dim CuEf() As Double
    Dim location As Variant
    Dim month As String
    Dim n As Integer, m As Integer, i As Integer, c As Integer, j As Integer, k As Integer
    Dim vol As Double, CE As Double, sum As Double, avg As Double
    On Error Resume Next
    loc = InputBox("Please enter the location of the file")
    fil = InputBox("Please enter the name of the file")
    Err.Clear
    filename = fil + ".xlsx"
    location = loc + "\" + fil + ".xlsx"
    Set awb = ThisWorkbook
    

    month = Mid(fil, 5, Len(fil) - 7)
    
    If month = "Jan" Or month = "Mar" Or month = "May" Or month = "July" Or month = "Aug" Or month = "Oct" Or month = "Dec" Then
        m = 33
    ElseIf month = "April" Or month = "Jun" Or month = "Sept" Or month = "Nov" Then
        m = 32
    Else: m = 30
    End If
    
    Application.ScreenUpdating = False
    Workbooks.Open filename:=location, UpdateLinks:=0
    Windows(filename).Activate
    Sheets("Elec composition").Select
    Range("A3:A" & m).Select
    Selection.Copy
    awb.Activate
    Range("B3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("D3:D" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("C3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("G3:G" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("D3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("J3:J30").Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("E3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("BO3:BO30").Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("F3").Select
    ActiveSheet.Paste
    
    Windows(filename).Activate
    Sheets("Vol Cir 2").Select
    i = 1
    A = Format(Range("A" & i + 7), "mmm")
    B = Left(month, 3)
    Do While A = B
        If Range("A" & i + 7).Value <> Range("A" & i + 8).Value Then
        dat = Range("A" & i + 7).Value
        Range("N" & i + 7).Select
        Application.CutCopyMode = False
        Selection.Copy
        awb.Activate
        For n = 3 To Range("B" & Rows.Count).End(xlUp).Row
            If Range("B" & n).Value = dat Then
            Range("G" & n).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            End If
        Next
        End If
        Windows(filename).Activate
        Sheets("Vol Cir 2").Select
        i = i + 1
        A = Format(Range("A" & i + 7), "mmm")
        B = Left(month, 3)
    Loop
    
    awb.Activate
    Sheets("Cir 2").Select
    i = 3
    j = 8
    c = 1
    Do While i < Range("B" & Rows.Count).End(xlUp).Row + 1
        dat = Range("B" & i).Value
        Windows(filename).Activate
        Sheets("CE").Select
        For j = 8 To Range("B" & Rows.Count).End(xlUp).Row
            If dat = Range("B" & j).Value And Range("C" & j).Value > 9 And Range("C" & j).Value < 21 Then
            ReDim Preserve CuEf(c) As Double
            CuEf(c) = Range("M" & j).Value
            c = c + 1
            End If
        Next
        sum = 0
        For k = 1 To c
            On Error Resume Next
            sum = sum + CuEf(k)
        Next
        CE = sum / (c - 1)
        If CE > 1000 Then
        CE = 0
        End If
        awb.Activate
        Sheets("Cir 2").Select
        Range("H" & i).Value = CE
        c = 1
        i = i + 1
    Loop
    
    Workbooks(filename).Close Savechanges:=False
    Application.ScreenUpdating = True
    MsgBox ("OK!")
    
End Sub
Sub extract_data_3()
    Dim awb As Workbook
    Dim loc As Variant, fil As Variant, filename As Variant, dat As Variant, A As Variant, B As Variant
    Dim CuEf() As Double
    Dim location As Variant
    Dim month As String
    Dim n As Integer, m As Integer, i As Integer, c As Integer, j As Integer, k As Integer
    Dim vol As Double, CE As Double, sum As Double, avg As Double
    On Error Resume Next
    loc = InputBox("Please enter the location of the file")
    fil = InputBox("Please enter the name of the file")
    Err.Clear
    On Error GoTo 0
    filename = fil + ".xlsx"
    location = loc + "\" + fil + ".xlsx"
    Set awb = ThisWorkbook
    

    month = Mid(fil, 5, Len(fil) - 7)
    
    If month = "Jan" Or month = "Mar" Or month = "May" Or month = "July" Or month = "Aug" Or month = "Oct" Or month = "Dec" Then
        m = 33
    ElseIf month = "April" Or month = "Jun" Or month = "Sept" Or month = "Nov" Then
        m = 32
    Else: m = 30
    End If
    
    Application.ScreenUpdating = False
    Workbooks.Open filename:=location, UpdateLinks:=0
    Windows(filename).Activate
    Sheets("Elec composition").Select
    Range("A3:A" & m).Select
    Selection.Copy
    awb.Activate
    Range("B3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("E3:E" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("C3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("H3:H" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("D3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("K3:K" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("E3").Select
    ActiveSheet.Paste
    Windows(filename).Activate
    Range("BP3:BP" & m).Select
    Application.CutCopyMode = False
    Selection.Copy
    awb.Activate
    Range("F3").Select
    ActiveSheet.Paste
    
    Windows(filename).Activate
    Sheets("Vol Cir 3").Select
    i = 1
    A = Format(Range("A" & i + 7), "mmm")
    B = Left(month, 3)
    Do While A = B
        If Range("A" & i + 7).Value <> Range("A" & i + 8).Value Then
        dat = Range("A" & i + 7).Value
        Range("N" & i + 7).Select
        Application.CutCopyMode = False
        Selection.Copy
        awb.Activate
        For n = 3 To Range("B" & Rows.Count).End(xlUp).Row
            If Range("B" & n).Value = dat Then
            Range("G" & n).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            End If
        Next
        End If
        Windows(filename).Activate
        Sheets("Vol Cir 3").Select
        i = i + 1
        A = Format(Range("A" & i + 7), "mmm")
        B = Left(month, 3)
    Loop
    
    awb.Activate
    Sheets("Cir 3").Select
    i = 3
    j = 8
    c = 1
    Do While i < Range("B" & Rows.Count).End(xlUp).Row + 1
        dat = Range("B" & i).Value
        Windows(filename).Activate
        Sheets("CE").Select
        For j = 8 To Range("B" & Rows.Count).End(xlUp).Row
            If dat = Range("B" & j).Value And Range("C" & j).Value > 20 Then
            ReDim Preserve CuEf(c) As Double
            CuEf(c) = Range("M" & j).Value
            c = c + 1
            End If
        Next
        sum = 0
        For k = 1 To c
            On Error Resume Next
            sum = sum + CuEf(k)
        Next
        CE = sum / (c - 1)
        If CE > 1000 Then
        CE = 0
        End If
        awb.Activate
        Sheets("Cir 1").Select
        Range("H" & i).Value = CE
        c = 1
        i = i + 1
    Loop
    
    Workbooks(filename).Close Savechanges:=False
    Application.ScreenUpdating = True
    MsgBox ("OK!")
    
End Sub
