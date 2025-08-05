Sub FWASplitReport()

' FWASplitReport Macro
' created by Amber Deabenderfer

  'Format the report
    Columns("A:A").Select
    Range("A2").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Selection.UnMerge
    Columns("A:A").Select
  'Deliminate    
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'    Range("A3:T7").Select

' Save the formatted report
    ActiveWorkbook.SaveAs Filename:= _
        "https://company.sharepoint.com/sites/Shared/Shared%20Documents/SubFolder1/SubFolder2/SubFolder3/Subfolder4/FWA%20Split%20Report_FWA%20Template1.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False 'update folder names and website as needed

'Close workbook
ActiveWorkbook.Close

End Sub
