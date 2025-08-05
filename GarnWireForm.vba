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

'Sub sendEmail()
''Working in Excel 2000-2016
''For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'    Dim Source As Range
'    Dim Dest As Workbook
'    Dim wb As Workbook
'    Dim TempFilePath As String
'    Dim TempFileName As String
'    Dim FileExtStr As String
'    Dim FileFormatNum As Long
'    Dim OutApp As Object
'    Dim OutMail As Object
'
'    Set Source = Nothing
'    Set emailTo = Worksheets("Form")
'    On Error Resume Next
'    Set Source = Range("A1:O75")
'    On Error GoTo 0
'
'    If Source Is Nothing Then
'        MsgBox "The source is not a range or the sheet is protected, please correct and try again.", vbOKOnly
'        Exit Sub
'    End If
'
'    With Application
'        .ScreenUpdating = False
'        .EnableEvents = False
'    End With
'
'    Set wb = ActiveWorkbook
'    Set Dest = Workbooks.Add(xlWBATWorksheet)
'
'    Source.Copy
'    With Dest.Sheets(1)
'        .Cells(1).PasteSpecial Paste:=8
'        .Cells(1).PasteSpecial Paste:=xlPasteValues
'        .Cells(1).PasteSpecial Paste:=xlPasteFormats
'        .Cells(1).Select
'        Application.CutCopyMode = False
'        ActiveWindow.DisplayGridlines = False
'    End With
'
'    TempFilePath = Environ$("temp") & "\"
'    TempFileName = "ADP Breakdown"
'
'    If Val(Application.Version) < 12 Then
'        'You use Excel 97-2003
'        FileExtStr = ".xls": FileFormatNum = -4143
'    Else
'        'You use Excel 2007-2016
'        FileExtStr = ".xlsx": FileFormatNum = 51
'    End If
'
'    Set OutApp = CreateObject("Outlook.Application")
'    Set OutMail = OutApp.CreateItem(0)
'
'    With Dest
'        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
'        On Error Resume Next
'        With OutMail
'    .To = emailTo.Range("C7").Value
'    .Subject = "Garnishment Wire Breakdown – LYM2 $" & emailTo.Range("I27").Value
'    .Body = emailTo.Range("F7").Value
'            .Attachments.Add Dest.FullName
'            .Display   'or use .Send
'        End With
'        On Error GoTo 0
'        .Close SaveChanges:=False
'    End With
'
'    Kill TempFilePath & TempFileName & FileExtStr
'
'    Set OutMail = Nothing
'    Set OutApp = Nothing
'
'    With Application
'        .ScreenUpdating = True
'        .EnableEvents = True
'    End With
'
'End Sub

Sub ADPgarnWireFormOne()

Dim userInfo As String
userInfo = Environ("Username")

'Verify if workbook exists - used with IsFile function
    Dim xRet As Boolean
    xRet = IsFile("C:\Users\" & userInfo & "\OneDrive - Liberty Mutual\" & _
            "Tax\Garnishments\Balancing-Payments\ADP Breakdown Template.xlsm")
    If xRet Then
    Else
    MsgBox "Workbook doesn't exist, please create and re-run"
    Exit Sub
    End If

'Verify if workbook is open - used with IsWorkBookOpen function
    xRet = IsWorkBookOpen("ADP Breakdown Template.xlsm")
    If xRet Then
    Else
        'Opens workbook if closed
        Workbooks.Open "C:\Users\" & userInfo & "\OneDrive - Liberty Mutual\" & _
            "Tax\Garnishments\Balancing-Payments\ADP Breakdown Template.xlsm"
    End If


''Move sheet to ADP form
'' Declare iteration object variable
'Dim iWorkbook As Workbook
'
'For Each iWorkbook In Application.Workbooks
'    If InStr(1, iWorkbook.Name, "Invoice_Report_03142024", StringCompareMethodConstant) > 0 Then
'        iWorkbook.Activate
'    End If
'Next iWorkbook
'
'    Sheets("Export Data").Select
'    Sheets("Export Data").Move After:=Workbooks("ADP Breakdown Template.xlsx").Sheets( _
'        2)
'
'' Format amounts to number
'    Columns("J:J").Select
'    Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
'        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
'        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub

'Sub ADPgarnWireFormTwo()
'
''Obtain processor name
'Dim eeName As String
'
'Sheets("Form").Select
'Range("G17:M17").Select
'Selection.UnMerge
'Range("G17").Select
'eeName = Selection
'
'    Range("G17:M17").Select
'    With Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    Selection.Merge
'    With Selection
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = True
'    End With
'
'eeName = InputBox("Verify/Update your name: If below is incorrect, update as needed", , eeName)
'
'If StrPtr(eeName) = 0 Then
'    MsgBox "Macro Cancelled"
'   Exit Sub
'End If
'
'Range("G17").Select
'Selection.FormulaR1C1 = eeName
'
''Duplicate Form page
'    Sheets("Form").Select
'    Sheets("Form").Copy
'    Range("G20:M20").Select
'    Selection.Copy
'
' 'remove formulas
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("D22:G22").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    ActiveWindow.SmallScroll Down:=8
'    Range("C33:D44").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("G33:H44").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("J33:M44").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("I27:M27").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("I24:M24").Select
'    Range("G47:H47").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
''Add to email
'Call sendEmail
'
''Close temporary sheet
'Application.DisplayAlerts = False
'ActiveWindow.Close
'
''Activate template and remove Sheet
'Workbooks("ADP Breakdown Template.xlsx").Activate
'    Sheets("Sheet1").Select
'    ActiveWindow.SelectedSheets.Delete
'Application.DisplayAlerts = True
'    Sheets("Form").Select
'
'End Sub
'
