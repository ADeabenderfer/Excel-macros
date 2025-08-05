Sub MWSQuarterFilePowerTool()
'
' MWSQuarterFilePowerTool Macro

'Delete excess tabs
    Sheets("Box_B_Records").Delete
    Sheets("Wage_Fed_W_Records").Delete
'    Sheets("X Records").Delete

'Select Wage_W_Records
    Sheets("Wage_W_Records").Select
    
'Delete extra columns
Dim currentColumn As Integer
Dim columnHeading As String

For currentColumn = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
    columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

'Check whether to preserve the column
    Select Case columnHeading
    'Insert name of columns to preserve
        Case "Client ID", "Employee Id", "Tax Code", "QTD Total Subject Wages", "Month-1 Employee Worked", "Month-2 Employee Worked", "Month-3 Employee Worked"
            'Do nothing
        Case Else
            'Delete the column
            ActiveSheet.Columns(currentColumn).Delete
        End Select
    Next
    
'Add filter to columns
    Range("A1:G1").Select
    Selection.AutoFilter
    
'Filter to SUI_ER
    ActiveSheet.Range("A1:G1").AutoFilter Field:=3, Criteria1:="*" & "SUI_ER" & "*"

'copy all data and paste to new tab
    Range("A1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'New tab
    Sheets.Add.Name = "SUI_ER"
    Sheets("SUI_ER").Select
    ActiveSheet.Paste
    
'Replace Y and N
Dim Col As String, cfind As Range

    Col = "Month-1 Employee Worked"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="Y", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="N", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Col = "Month-2 Employee Worked"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="Y", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="N", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Col = "Month-3 Employee Worked"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="Y", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="N", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
'Select ER data for Pivot
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
        
'Add Pivot Tab
    Sheets.Add.Name = "Pivot"
    Sheets("Pivot").Select
        
'Add Pivot
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "SUI_ER!R1C1:R190000C7", Version:=8).CreatePivotTable TableDestination:= _
        "Pivot!R3C1", TableName:="PivotTable1", DefaultVersion:=8
    Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    Sheets("Pivot").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Client ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Tax Code")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Month-1 Employee Worked"), _
        "Sum of Month-1 Employee Worked", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Month-2 Employee Worked"), _
        "Sum of Month-2 Employee Worked", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Month-3 Employee Worked"), _
        "Sum of Month-3 Employee Worked", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("QTD Total Subject Wages"), _
        "Sum of QTD Total Subject Wages", xlSum
    
 ' Format number to whole #
    Col = "Sum of QTD Total Subject Wages"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.Select
    cfind.EntireColumn.Select
    Selection.NumberFormat = "0"
    
'Add slicers and format
    Range("A4").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Client ID").Slicers.Add ActiveSheet, , "Client ID", "Client ID", 171, 489.75, _
        144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Tax Code").Slicers.Add ActiveSheet, , "Tax Code", "Tax Code", 208.5, 527.25, _
        144, 198.75
    ActiveSheet.Shapes.Range(Array("Tax Code")).Select
    Rows("1:1").RowHeight = 65.25
    ActiveSheet.Shapes.Range(Array("Client ID")).Select
    ActiveSheet.Shapes("Client ID").IncrementLeft -489.75
    ActiveSheet.Shapes("Client ID").IncrementTop -221.25
    ActiveWorkbook.SlicerCaches("Slicer_Client_ID").Slicers("Client ID"). _
        NumberOfColumns = 2
    ActiveWorkbook.SlicerCaches("Slicer_Client_ID").Slicers("Client ID"). _
        NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches("Slicer_Client_ID").Slicers("Client ID"). _
        NumberOfColumns = 4
    ActiveSheet.Shapes("Client ID").ScaleWidth 1.8229166667, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Client ID").ScaleHeight 0.3283018868, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Tax Code")).Select
    ActiveSheet.Shapes("Tax Code").IncrementLeft 239.25
    ActiveSheet.Shapes("Tax Code").IncrementTop -258.75
    ActiveWorkbook.SlicerCaches("Slicer_Tax_Code").Slicers("Tax Code"). _
        NumberOfColumns = 2
    ActiveSheet.Shapes("Tax Code").ScaleHeight 2.5320754717, msoFalse, _
        msoScaleFromTopLeft
    ActiveWorkbook.SlicerCaches("Slicer_Tax_Code").Slicers("Tax Code"). _
        NumberOfColumns = 3
    ActiveSheet.Shapes("Tax Code").IncrementLeft -160.5
    ActiveSheet.Shapes("Tax Code").ScaleWidth 1.765625, msoFalse, _
        msoScaleFromTopLeft

'Add Calculation formulas
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "BLS"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "Qtr File"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Variance"
    Range("H10").Select
    ActiveCell.FormulaR1C1 = "large office"
    Range("H11").Select
    ActiveCell.FormulaR1C1 = "new off #'s"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "Grand Total"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "month 1"
    Range("I7").Select
    Selection.NumberFormat = "#,##0"
    
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[-3]C,C[-8]:C[-7],2,FALSE)"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C-R[-2]C"
    Range("I11").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+R[-3]C"
    Range("J5").Select
    ActiveCell.FormulaR1C1 = "month 2"
    Range("J7").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[-3]C[-1],C[-9]:C[-7],3,FALSE)"
    Range("J8").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C-R[-2]C"
    Range("J11").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+R[-3]C"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "month 3"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[-3]C[-2],C[-10]:C[-7],4,FALSE)"
    Range("K8").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C-R[-2]C"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+R[-3]C"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "wages"
    Range("L7").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[-3]C[-3],C[-11]:C[-7],5,FALSE)"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C-R[-2]C"
    Range("L11").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+R[-3]C"
    Range("M6").Select
    ActiveCell.FormulaR1C1 = "from BLS website"
    Range("M7").Select
    ActiveCell.FormulaR1C1 = "this is pulled from pivot on left"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "this is the variance"
    Range("M11").Select
    ActiveCell.FormulaR1C1 = _
        "this is the new numbers that you will put into the BLS website for the large office"
    Range("M12").Select
    
    Range("I7:L7").Select
    Selection.NumberFormat = "#,##0"
    
    Range("I6:L6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("I10:L10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Columns("H:M").Select
    Columns("H:M").EntireColumn.AutoFit
    
End Sub


