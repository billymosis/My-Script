Sub y()
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Filters.Clear
        .Filters.Add "Picture Files", "*.bmp;*.jpg;*.gif;*.png;*.jpeg"
        .ButtonName = "Select"
        .AllowMultiSelect = False
        .Title = "Choose Photo"
        .InitialView = msoFileDialogViewDetails
        .Show
    End With
        
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .UserPicture fd.SelectedItems(1)
        .TextureTile = msoFalse
    End With

End Sub

Sub Copy_Form()
    ActiveCell.Range("A1:T55").Select
    Selection.Copy
    ActiveCell.Offset(55, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Application.CutCopyMode = False
    ActiveCell.Select
End Sub

Sub PageBreak()
    Dim i As Integer
    On Error GoTo ErrorHandler

    For i = 1 To 50
        ActiveSheet.Rows((i * 55) + 1).PageBreak = xlPageBreakManual
            With ActiveSheet.Rows(i * 55).Borders(xlEdgeBottom)
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next i

ErrorHandler:
    Exit Sub
End Sub

Sub PageNumbering()
    Dim i As Integer, j As Integer, X As Integer

    With ActiveSheet
        For i = 1 To .HPageBreaks.Count
        ActiveSheet.Cells(i, 1).Value = .HPageBreaks(i).Location
        Next i
        For j = 1 To .VPageBreaks.Count
        ActiveSheet.Cells(j, 2).Value = .VPageBreaks(j).Location
        Next j
    End With

    For X = 1 To i
        ActiveSheet.Cells(8 + 55 * (X - 1), "I").Value = "Lembar                        : " & X & "/" & i
    Next X

End Sub

Sub PrintPDF()
    Dim FileName1 As String
    Dim FileName As String
    Dim Directory As String
    Directory = "D:\Mrican\GIS\Inventarisasi\"

    With ActiveSheet.PageSetup
    '   .Zoom = 88
    End With

    Dim i As Integer, j As Integer, X As Integer

    With ActiveSheet
        For i = 1 To .HPageBreaks.Count
        ActiveSheet.Cells(i, 1).Value = .HPageBreaks(i).Location
        Next i
    End With

    For X = 1 To i
        FileName1 = Replace(ActiveSheet.Cells(8 + 55 * (X - 1), "N").Value, ":", "")
        FileName = Replace(FileName1, "/", "@")
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Directory & FileName & ".pdf", From:=X, To:=X, Quality:=xlQualityStandard
    Next X

End Sub

Sub Test()
    Dim FileName1 As String
    Dim FileName As String
    FileName1 = Replace(ActiveSheet.Cells(8, "N").Value, ":", "")
    FileName = Replace(FileName1, "/", "")
    MsgBox FileName
End Sub

Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call PrintSemua
    Next
    Application.ScreenUpdating = True
End Sub

Sub PrintBangunan()
    Dim Ws As Worksheet
    Dim fname As String
    Dim Directory As String
    Dim rg As Range
    Dim LastRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim P As Integer
    Dim Q As Integer
    Dim O As Integer
    Dim Bgn As String
    Set Ws = ActiveSheet
    Directory = "D:\PDFX\"
    fname = ".pdf"
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row + 6

    For i = 1 To LastRow / 54
        j = 54 * i
        P = j - 53
        Q = 9 + i * 54
        O = Q - 54
        Bgn = Replace(Replace(Ws.Cells(O, 14).Value, ":   ", ""), " / ", " ")
        If Bgn = "" Then
            Bgn = "Blank"
        End If
        Set rg = Ws.Range(Cells(P, 1), Cells(j, 20))
        rg.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Directory & "Bangunan " & Bgn & fname
        'Debug.Print I
        'Debug.Print rg.Address
        Debug.Print Directory & "Bangunan " & Bgn & fname
    Next i

    '    Range("A55:T108").PrintOut Copies:=1, PrintToFile:=True, Collate:=True, PrToFileName:=Directory & fname
End Sub
Sub PrintSaluran()
    Dim Ws As Worksheet
    Dim fname As String
    Dim Directory As String
    Dim rg As Range
    Dim LastRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim P As Integer
    Dim Q As Integer
    Dim O As Integer
    Dim Bgn As String
    Set Ws = ActiveSheet
    Directory = "D:\PDFX\"
    fname = ".pdf"
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row + 6

    For i = 1 To LastRow / 54
        j = 54 * i
        P = j - 53
        Q = 8 + i * 54
        O = Q - 54
        Bgn = Replace(Replace(Ws.Cells(O, 36).Value, ":   ", ""), "/ ", "")
        If Bgn = "" Then
            Bgn = "Blank"
        End If
        Set rg = Ws.Range(Cells(P, 21), Cells(j, 41))
        rg.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Directory & "Saluran " & Bgn & fname
        'Debug.Print I
        'Debug.Print rg.Address
        Debug.Print Directory & "Saluran " & Bgn & fname
    Next i

    '    Range("A55:T108").PrintOut Copies:=1, PrintToFile:=True, Collate:=True, PrToFileName:=Directory & fname
End Sub

Sub PrintSaluranNanang()
    Dim Ws As Worksheet
    Dim fname As String
    Dim Directory As String
    Dim rg As Range
    Dim LastRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim P As Integer
    Dim Q As Integer
    Dim O As Integer
    Dim Bgn As String
    Set Ws = ActiveSheet
    Directory = "D:\PDFX\"
    fname = ".pdf"
    Bgn = Ws.Name
    If Bgn = "" Then
        Bgn = "Blank"
    End If
    Set rg = Ws.Range(Cells(1, 1), Cells(53, 20))
    rg.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Directory & "Saluran " & Bgn & fname
    'Debug.Print I
    'Debug.Print rg.Address
    Debug.Print Directory & "Saluran " & Bgn & fname

    '    Range("A55:T108").PrintOut Copies:=1, PrintToFile:=True, Collate:=True, PrToFileName:=Directory & fname
End Sub


Sub PrintSemua()
    Call PrintSaluran
    Call PrintBangunan
End Sub

Sub RubahNama()
    Dim i As Long
    Dim j As Long

    For j = 2 To 10
        For i = 20 To 40
        'Debug.Print J & " " & I
            If Mid(Cells(j, 5).Value, 10, 8) Like "*" & Cells(i, 3).Value & "*" Then
                Debug.Print Cells(j, 5).Value & " sama dengan " & Cells(i, 3).Value
                Cells(j, 6) = Cells(i, 1)
            Else
        End If
        Next i
    Next j

End Sub


Sub RubahNamaX()
    If Mid(Cells(2, 5).Value, 10, 8) Like "*" & Cells(25, 3).Value & "*" Then
        Debug.Print Cells(2, 1).Address
    Else
        Debug.Print "Salah"
    End If
End Sub

Sub Range_End_Method()
    'Finds the last non-blank cell in a single row or column
    Dim lRow As Integer
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 2).End(xlUp).Row + 6
    MsgBox "Last Row: " & lRow
End Sub

Sub TestY()
    Dim Ws As Worksheet
    Set Ws = Sheets("B28 (SS.C)")
    Debug.Print Ws.Name
    End Sub
    Sub Pindah_DATA()
    Dim WbFrom, WbTo As Workbook
    Dim RowFrom, RowTo As Long
    Set WbFrom = Workbooks("Hidrolis sekunder kiri bimo050220.xlsm")
    Set WbTo = Workbooks("25.SS T D.txt")

    RowTo = InputBox("RowTo")
    RowFrom = InputBox("RowFrom")
    WbTo.ActiveSheet.Cells(RowTo, "C") = WbFrom.ActiveSheet.Cells(RowFrom + 1, "C") & "-" & WbFrom.ActiveSheet.Cells(RowFrom, "D") 'nama
    WbTo.ActiveSheet.Cells(RowTo, "D") = WbFrom.ActiveSheet.Cells(RowFrom, "R") 'A
    WbTo.ActiveSheet.Cells(RowTo, "E") = WbFrom.ActiveSheet.Cells(RowFrom, "AG") 'Q
    WbTo.ActiveSheet.Cells(RowTo, "F") = WbFrom.ActiveSheet.Cells(RowFrom, "T") 'b
    WbTo.ActiveSheet.Cells(RowTo, "G") = WbFrom.ActiveSheet.Cells(RowFrom, "Y") 'h
    WbTo.ActiveSheet.Cells(RowTo, "H") = WbFrom.ActiveSheet.Cells(RowFrom, "Z") 'W
    WbTo.ActiveSheet.Cells(RowTo, "I") = WbFrom.ActiveSheet.Cells(RowFrom, "V") 'm
    WbTo.ActiveSheet.Cells(RowTo, "J") = WbFrom.ActiveSheet.Cells(RowFrom, "W") 'm
    WbTo.ActiveSheet.Cells(RowTo, "L") = WbFrom.ActiveSheet.Cells(RowFrom, "AF") 'v
    WbTo.ActiveSheet.Cells(RowTo, "M") = WbFrom.ActiveSheet.Cells(RowFrom, "X") 'I
    WbTo.ActiveSheet.Cells(RowTo, "P") = WbFrom.ActiveSheet.Cells(RowFrom, "V") 'm luar
    WbTo.ActiveSheet.Cells(RowTo, "AB") = WbFrom.ActiveSheet.Cells(RowFrom, "K") 'el Akhir
    WbTo.ActiveSheet.Cells(RowTo, "AC") = WbFrom.ActiveSheet.Cells(RowFrom + 1, "J") 'el Awal

    'Debug.Print WbFrom.ActiveSheet.Cells(RowFrom, "R")
    'Debug.Print WbTo.Name
    'Debug.Print WbTo.ActiveSheet.Name
End Sub

Sub Negativize()
    Dim cel As Range
    For Each cel In Selection
        If cel.Value > 0 Then
            cel.Value = -Abs(cel.Value)
        ElseIf cel.Value < 0 Then
            cel.Value = Abs(cel.Value)
        End If
    Next cel
End Sub

Sub yellow_Cells()
    Dim MyRange As Range
    Dim LastRow As Long

    Set MyRange = ActiveSheet.Range("E1:E4135")
    Cells(2, 5).EntireRow.Insert
    For Each c In MyRange
        If c.Interior.ColorIndex = 6 Then
            c.EntireRow.Insert
            Debug.Print
        End If
    Next
End Sub
Sub Macro2()
    Rows("4133:4133").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub Labeling()
    ActiveCell.Offset(2, -2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "T1"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "L"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "D1L"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.ClearContents
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "D1L"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "L"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "T1"
    ActiveCell.Offset(1, -4).Range("A1").Select
End Sub
Sub Formula()
    Debug.Print ActiveCell.Address
    '    ActiveCell.FormulaR1C1 = _
    '        "=(((R[-1]C[-6] - R[0]C[-6])^2+(R[-1]C[-5]-R[0]C[-5])^2)^0.5)*IF(R[-1]C[-5]<R[0]C[-5],-1,1)"
    '    ActiveCell.Offset(0, -1).Range("A1").Formula2R1C1 = _
    '        "=(R[0]C[-4])"
    ActiveCell.Formula = "=(((" & ActiveCell.Offset(-1, -6).Address(True, True) & "-" & ActiveCell.Offset(0, -6).Address(False, False) & ")^2+(" _
                        & ActiveCell.Offset(-1, -5).Address(True, True) & "-" & ActiveCell.Offset(0, -5).Address(False, False) & ")^2)^0.5)*IF(" _
                        & ActiveCell.Offset(-1, -5).Address(True, True) & "<" & ActiveCell.Offset(0, -5).Address(False, False) & ",1,-1)"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "=R[0]C[-5]"
    ActiveCell.Offset(0, 2).FormulaR1C1 = "=R[0]C[-5]"
End Sub
Sub FlipColumns()
    Dim rng As Range
    Dim WorkRng As Range
    Dim Arr As Variant
    Dim i As Integer, j As Integer, k As Integer
 
    On Error Resume Next
 
    xTitleId = "Flip columns vertically"
    Set WorkRng = Selection
   ' If WorkRng <> Null Then
    'Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    'End If
    Arr = WorkRng.Formula
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
 
    For j = 1 To UBound(Arr, 2)
        k = UBound(Arr, 1)
            For i = 1 To UBound(Arr, 1) / 2
                xTemp = Arr(i, j)
                Arr(i, j) = Arr(k, j)
                Arr(k, j) = xTemp
                k = k - 1
            Next
    Next
 
    WorkRng.Formula = Arr
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
 
End Sub

Sub CopyFlipTranspose()
    Selection.Copy
    ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Call FlipColumns
    Selection.Copy
    ActiveCell.Offset(0, 14).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=True
    ActiveCell.Offset(-1, -24).Range("A1:C1").Select
        Selection.Copy
    ActiveCell.Offset(1, 15).Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues
    ActiveCell.Offset(-1, -12).Range("A1").Select
        Selection.Copy
    ActiveCell.Offset(1, 11).Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues
End Sub

Sub Oye()

Dim col As New Collection
Dim rng As Range
Dim rng2 As Range
Dim i As Integer
Set i = 0

    For Each c In Selection
        If Left(c.Value, 3) = "B.1" Then
            Set rng = c
            col.Add (rng.Address)
        End If
    Next
    
    For Each X In col
    Set rng2 = Range(X)

    rng2(-i, 0).EntireRow.Insert Shift:=xlShiftDown, CopyOrigin:=xlInsertFormatOriginConstant
    i = i + 1
    Next
        
End Sub

Sub MinusElevasiPatok()
	ActiveCell.Value = ActiveCell.Value - 0.25
	ActiveCell.Offset(3, 0).Select
End Sub

Sub SwapRow()
    Selection.Cut
    ActiveCell.Offset(-1, 0).Range("A1:L1").Select
    Selection.Insert Shift:=xlDown
End Sub

Sub WriteCSVFile()
Dim i As Integer
Dim WS_Count As Integer

WS_Count = ActiveWorkbook.Worksheets.Count
For i = 1 To WS_Count
Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets(i)
     PathName = "" & ThisWorkbook.Path & "\" & ws.Name & ".csv"
    ws.Copy
    ActiveWorkbook.SaveAs Filename:=PathName, _
        FileFormat:=xlCSV, CreateBackup:=False
Next i

End Sub

Sub AutoFill()
    Dim rng As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rng1 As Range
    Dim i As Integer
    Dim s As String

    Set rng1 = ActiveCell.Range("A1", "C1")
    Set rng2 = ActiveCell.Offset(0, -3).Range("A1", "A20").Find("B16")
    i = rng2.Row - ActiveCell.Row
    Set rng = ActiveCell.Offset(0, 0).Range("A1", "C" & i)
    Set rng3 = ActiveCell.Offset(-1, 0).Range("A1", "C1")
    Debug.Print (rng3.Address)
    Selection.Copy
    rng3.PasteSpecial
    rng1.AutoFill Destination:=rng, Type:=xlFillDefault
   
    
End Sub


