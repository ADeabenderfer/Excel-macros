Function IsWorkBookOpen(Name As String) As Boolean
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

Sub LienLoad()
'
' LienLoad Macro

'get user info
Dim userInfo As String
userInfo = Environ("Username")

'Unmerge all cells
Range("A1:L5000").UnMerge
    
'Add in column for comments and column for Report Date
    Range("M10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16643047
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Comments"
    Range("N10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16643047
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Date"

'Add in formula to auto fill Comment on "Sucess" Message
    Range("M11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=""Success"",""Success ran with interrogatory, no action needed"","""")"
    Range("M11").Select
    Selection.AutoFill Destination:=Range("M11:M5000"), Type:=xlFillDefault
    Range("M11:M5000").Select

'Add in formula to copy date to row 5000 if data in Last name
    Range("N12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-11]="""","""",R11C14)"
    Range("N12").Select
    Selection.AutoFill Destination:=Range("N12:N5000"), Type:=xlFillDefault
    Range("N12:N5000").Select

'Add in report date
 Dim year As String
 Dim month As String
 Dim day As String
 Dim d As String
 Dim mName As String
 Dim wBK As Workbook
 
Set wBK = ActiveWorkbook
 strWBName = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
 year = Right(strWBName, 4)
strWBName = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 5)
 day = Right(strWBName, 2)
strWBName = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 7)
 month = Right(strWBName, 2)
  d = month & "/" & day & "/" & year
  Range("N11").Value = d

'Sort blanks to bottom
Range("A11:L5000").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("C11:C5000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A11:L5000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("N:N").Select
    Selection.NumberFormat = "m/d/yyyy"

'Save to folder
mName = monthName(month)

    ActiveWorkbook.SaveAs Filename:= _
        "https://company.sharepoint.com/teams/Subfolder1/Subfolder2/Subfolder3/Subfolder4/" & year & "/" & month & "%20" & mName & "%20Lien%20" & year & "/" & ActiveWorkbook.Name _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

'Select all data
Dim LastRow As Integer

LastRow = Cells(Rows.Count, "L").End(xlUp).Row
Range("A11:N" & LastRow).Select
Selection.Copy

'Verify if workbook exists - used with IsFile function
    Dim xRet As Boolean
    xRet = IsFile("C:\Users\" & userInfo & "\OneDrive - Company\Subfolder1\Subfolder2\Subfolder3\" & year & _
        "\" & month & " " & mName & " Lien " & year & "\ADP_Lien " & mName & " Report.xlsx")
                
    If xRet Then
    Else
    MsgBox "Report Year/mName workbook doesn't exist, please create and re-run"
    Exit Sub
    End If

'Verify if workbook is open - used with IsWorkBookOpen function
    xRet = IsWorkBookOpen("ADP_Lien " & mName & " Report.xlsx")
    If xRet Then
    Else
        'Opens workbook if closed
        Workbooks.Open "C:\Users\" & userInfo & "\OneDrive - Company\Subfolder1\Subfolder2\Subfolder3\" & year & _
        "\" & month & " " & mName & " Lien " & year & "\ADP_Lien " & mName & " Report.xlsx"
    End If
    
'Clear any filter(s)
If ActiveSheet.AutoFilterMode Then 'autofilter is 'on'
   On Error Resume Next   'turn off error reporting
   ActiveSheet.ShowAllData
   On Error GoTo 0   'turn error reporting back on
End If
    
'Select first open cell in data
LastRow = Cells(Rows.Count, "L").End(xlUp).Row
Range("A" & LastRow).Offset(1, 0).Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Application.DisplayAlerts = False
wBK.Close
Application.DisplayAlerts = True

End Sub
