Attribute VB_Name = "count_qd"


Sub qdqd1()
'
' qdqd1 Macro
'
    Dim usedRowNum As Integer
    Dim nowTime, sheetName As String
    nowTime = Format(Time, "hh_mm")
    sheetName = "huizong" & nowTime

    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\sheji111\Desktop\" & "qdauto" & Format(Date, "MMDD") & ".xlsm"

    ActiveSheet.Name = "sheet1"
    Range("H1").Select

    ' let the chinese cell to english cell
    Range("A1").value = "ID"
    Range("B1").value = "type"
    Range("C1").value = "contract"
    Range("D1").value = "gift"
    Range("E1").value = "mon"
    Range("F1").value = "year"
    Range("G1").value = "salesman"

    ActiveCell.FormulaR1C1 = "qdqd"
    Range("H2").Select
    ''get an value like String, which is used for the formula in R1C1
    ActiveCell.FormulaR1C1 = _
        "=IF(MID(RC[-7],3,2)=""LL"",MID(RC[-7],1,2)&MID(RC[-7],5,2),MID(RC[-7],1,4))"
    Range("H2").Select

    ''search for the not_empty cells
    usedRowNum = Worksheets("sheet1").UsedRange.Rows.Count
    Range("H2:" & "H" & usedRowNum).Select

    Selection.FillDown
    Range("F6").Select
    ''the usedcell for AutoFilter
    ' Selection.AutoFilter
    ' Selection.AutoFilter
    ActiveSheet.UsedRange.Select
    Range("C3").Activate

    ''add PivotTables
    ActiveWorkbook.Sheets.Add
    ActiveSheet.Name = sheetName
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets("sheet1").UsedRange, Version:=xlPivotTableVersion12) _
        .CreatePivotTable TableDestination:=sheetName & "!R3C1", TableName:="Pivottable1", _
        DefaultVersion:=xlPivotTableVersion12
    Sheets(sheetName).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Pivottable1").PivotFields("ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ''add a count
    ActiveSheet.PivotTables("Pivottable1").AddDataField ActiveSheet.PivotTables("Pivottable1" _
        ).PivotFields("ID"), "count:ID", xlCount
    With ActiveSheet.PivotTables("Pivottable1").PivotFields("qdqd")
        .Orientation = xlRowField
        .Position = 1
    End With
    ''add three sum
    ActiveSheet.PivotTables("Pivottable1").AddDataField ActiveSheet.PivotTables("Pivottable1" _
        ).PivotFields("gift"), "sum:gift", xlSum
    ActiveSheet.PivotTables("Pivottable1").AddDataField ActiveSheet.PivotTables("Pivottable1" _
        ).PivotFields("mon"), "sum:mon", xlSum
    ActiveSheet.PivotTables("Pivottable1").AddDataField ActiveSheet.PivotTables("Pivottable1" _
        ).PivotFields("year"), "sum:year", xlSum
End Sub
