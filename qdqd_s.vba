' 用vba控制表格，已成功
' 希望能实现，将该vba宏，用py置入对应表格


Sub qdqd1()
'
' qdqd1 Macro
'
    Dim usedRowNum As Integer
    Dim nowTime, sheetName As String
    nowTime = Format(Time, "hh_mm")
    sheetName = "huizong" & nowTime

    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\sheji111\Desktop\新桌面\IT开发相关\qd测试\test_make\" & "qd自动生成" & Format(Date, "MMDD") & ".xlsm"

    ActiveSheet.Name = "sheet1"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "qdqd"
    Range("H2").Select
    ''获取或设置一个String值, 该值代表使用 R1C1 样式表示法的对象的公式
    ActiveCell.FormulaR1C1 = _
        "=IF(MID(RC[-7],3,2)=""LL"",MID(RC[-7],1,2)&MID(RC[-7],5,2),MID(RC[-7],1,4))"
    Range("H2").Select

    ''循环查找最下方的非空单元格
    usedRowNum = Worksheets("sheet1").UsedRange.Rows.Count
    Range("H2:" & "H" & usedRowNum).Select

    Selection.FillDown
    Range("F6").Select
    ''选中单元格的自动筛选
    ' Selection.AutoFilter
    ' Selection.AutoFilter
    ActiveSheet.UsedRange.Select
    Range("C3").Activate

    ''添加数据透视表
    ActiveWorkbook.Sheets.Add
    ActiveSheet.Name = sheetName
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets("sheet1").UsedRange, Version:=xlPivotTableVersion12) _
        .CreatePivotTable TableDestination:=sheetName & "!R3C1", TableName:="数据透视表1", _
        DefaultVersion:=xlPivotTableVersion12
    Sheets(sheetName).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("会员帐号")
        .Orientation = xlRowField
        .Position = 1
    End With
    ''添加一个计数项
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("会员帐号"), "计数项:会员帐号", xlCount
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("qdqd")
        .Orientation = xlRowField
        .Position = 1
    End With
    ''添加三个求和项
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("赠菜量"), "求和项:赠菜量", xlSum
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("月卡量"), "求和项:月卡量", xlSum
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("年卡量"), "求和项:年卡量", xlSum
End Sub
