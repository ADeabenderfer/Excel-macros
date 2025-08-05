Function IsWorkBookOpen(Name As String) As Boolean
'Check if workbook is already open
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function

Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

Sub compAddLocToReport()
'created by Amber Deabenderfer 08/2022
'Update naming convention to match different data pulls
    Columns("F:F").Select
    Selection.Replace What:=" Resident Address Change", Replacement:= _
        "Resident Address Change", LookAt:=xlWhole, SearchOrder:=xlByRows, _
        MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, _
        FormulaVersion:=xlReplaceFormula2

'Merge EDO data to main report
Call EDOtoComp

'Select and Copy data from Pulled Report
    Range("A2:AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

Dim dateInput As String
Dim inputMonth As Double
Dim inputYear As Double
Dim inputQtr As String

'Get date report was pulled from user
dateInput = InputBox("Enter Report Pull Date", "Report Date", "MM/DD/YYYY")
    If StrPtr(dateInput) = 0 Then
        MsgBox ("Cancelation confirmed")
        Exit Sub
    Else
    End If

'extract month and year
inputMonth = Left(dateInput, 2)
inputYear = Right(dateInput, 4)

'Get applicable quarter
If inputMonth < 4 Then
    inputQtr = "Q1"
    Else
        If inputMonth < 7 Then
        inputQtr = "Q2"
        Else
            If inputMonth < 10 Then
            inputQtr = "Q3"
            Else
                If inputMonth > 9 Then
                inputQtr = "Q4"
                End If
            End If
        End If
End If

'Verify if workbook exists - used with IsFile function
Dim userInfo As String
userInfo = Environ("Username")

    Dim xRet As Boolean
    xRet = IsFile("C:\Users\" & userInfo & "\OneDrive - Liberty Mutual\Tax\" & _
        "Pay Period Reports\Comprehensive Resident and Location update report\" & _
        inputYear & "\" & inputYear & " " & inputQtr & _
        " Comprehensive Resident Address Report.xlsx")
    If xRet Then
    Else
    MsgBox "Report Year/Quarter workbook doesn't exist, please create and re-run"
    Exit Sub
    End If

'Verify if workbook is open - used with IsWorkBookOpen function
    xRet = IsWorkBookOpen(inputYear & " " & inputQtr & " Comprehensive Resident Address Report.xlsx")
    If xRet Then
    Else
        'Opens workbook if closed
        Workbooks.Open "C:\Users\" & userInfo & "\OneDrive - Liberty Mutual\" & _
            "Tax\Pay Period Reports\Comprehensive Resident and Location update report\" & _
            inputYear & "\" & inputYear & " " & inputQtr & _
            " Comprehensive Resident Address Report.xlsx"
    End If

'Clear filters on new report if filters set
On Error Resume Next
    ActiveSheet.ShowAllData
On Error GoTo 0
        
'Find
Dim Col As String, cfind As Range

Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim nStart As Range
Dim nEnd As Range

'Count columns between Report effective date and EE name
Sheets("Comprehensive").Select

Col = "Report p_effective_date"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
b = cfind.column

Col = "Employee Name"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
a = cfind.column
cfind.Select

c = b - a

    If Selection.Offset(1, 0) = "" Then
        Selection.Offset(1, 0).Select
    Else
    Selection.End(xlDown).Offset(1, 0).Select
    End If

Set nStart = Selection.Offset(0, c)

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Col = "Employee Name"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.End(xlDown).Select

Set nEnd = Selection.Offset(0, c)

Range(nStart, nEnd).Select

Selection.Value = dateInput

'Refresh all - FWA/PBI
ThisWorkbook.RefreshAll

End Sub

Sub EDOtoComp()
    Dim EDO As Double
    Dim formStart As Range, formEnd As Range
    Dim PStart As Range, PEnd As Range

    Sheets("Comprehensive Address details").Select
    Columns("N:U").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("EDO Details").Select
    
'EDO?
    Range("A2").Select
    Selection.Copy
    EDO = Selection

If EDO > 0 Then

    Range("B1:I1").Select
    Selection.Copy
    Sheets("Comprehensive Address details").Select
    Range("N1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      
    'Count rows
    Col = "Employee Name "
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    Set formStart = cfind.Offset(1, 0)
    Set formEnd = cfind.End(xlDown)
    
    'EDO line 1
    Set PStart = formStart.Offset(0, 13)
    Set PEnd = formEnd.Offset(0, 13)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-12],'EDO Details'!C[-13],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 2
    Set PStart = formStart.Offset(0, 14)
    Set PEnd = formEnd.Offset(0, 14)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-13],'EDO Details'!C[-14],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 3
    Set PStart = formStart.Offset(0, 15)
    Set PEnd = formEnd.Offset(0, 15)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-14],'EDO Details'!C[-15],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 4
    Set PStart = formStart.Offset(0, 16)
    Set PEnd = formEnd.Offset(0, 16)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-15],'EDO Details'!C[-16],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 5
    Set PStart = formStart.Offset(0, 17)
    Set PEnd = formEnd.Offset(0, 17)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-16],'EDO Details'!C[-17],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 6
    Set PStart = formStart.Offset(0, 18)
    Set PEnd = formEnd.Offset(0, 18)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-17],'EDO Details'!C[-18],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 7
    Set PStart = formStart.Offset(0, 19)
    Set PEnd = formEnd.Offset(0, 19)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-18],'EDO Details'!C[-19],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

    'EDO line 8
    Set PStart = formStart.Offset(0, 20)
    Set PEnd = formEnd.Offset(0, 20)

    PStart.Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX('EDO Details'!C[-12],MATCH('Comprehensive Address details'!RC[-19],'EDO Details'!C[-20],0)),"""")"
    Selection.AutoFill Destination:=Range(PStart, PEnd)

        Else
        Range("B1:I1").Select
        Selection.Copy
        Sheets("Comprehensive Address details").Select
        Range("N1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
End Sub

