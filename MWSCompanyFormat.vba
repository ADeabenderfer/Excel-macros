Sub MWSCompanyFormat()
'
'Add in rounded calculation row
    Range("M1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]="""","""",IF(RC[-3]=""Quarterly Wages"","""",ROUND(RC[-3],0)))"
    Range("M1").Select
    Selection.AutoFill Destination:=Range("M1:M3000")
    Range("M1:M3000").Select

'Remove bold and set Account # as bold
    Columns("A:J").Select
    Range("A3").Activate
    Selection.Font.Bold = False


    Columns("A:A").Select
    Range("A3").Activate
    With Application.ReplaceFormat.Font
        .FontStyle = "Bold"
        .Subscript = False
        .TintAndShade = 0
    End With
    Selection.Replace What:="SUI Account Number", Replacement:= _
        "SUI Account Number", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=True, FormulaVersion:= _
        xlReplaceFormula2

'Group each section by state
Dim r As Integer, LR As Integer, startr As Integer, endr As Integer

LR = Cells(Rows.Count, "A").End(xlUp).Row

For r = 1 To LR

    If Range("A" & r).Font.Bold = True Then
        If startr = 0 Then
            startr = r + 1
        Else
            endr = r + -1
            Range("A" & startr & ":A" & endr).Rows.Group
            startr = r + 1
        End If
    End If

Next r

Range("A" & startr & ":A" & LR).Rows.Group

    Dim b As Worksheet
    For Each b In Worksheets
        b.Outline.ShowLevels ColumnLevels:=1, RowLevels:=1
    Next b
       
End Sub
