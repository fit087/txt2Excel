Attribute VB_Name = "Módulo5"
Sub importarEx()
Attribute importarEx.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' importarEx Macro
'
' Atalho do teclado: Ctrl+k
'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\adolfo.correa\Google Drive\1. Dissertação\5. ModeFrontier\D16_LA50_R150Relax2_r.txt" _
        , Destination:=Range("$A$1"))
        .Name = "D16_LA50_R150Relax2_r"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\adolfo.correa\Google Drive\1. Dissertação\5. ModeFrontier\D16_LA50_R150Relax2_r.MAP" _
        , Destination:=Range("$A$1"))
        .Name = "D16_LA50_R150Relax2_r_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A32452:B32452").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Cells(1)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Selection.End(xlToRight).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, -8).Columns("A:G").EntireColumn.Select
    ActiveCell.Offset(0, -2).Range("A1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, -13).Columns("A:F").EntireColumn.Select
    ActiveCell.Offset(0, -8).Range("A1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveCell.Offset(0, -13).Columns("A:F").EntireColumn.Select
    ActiveCell.Offset(0, -8).Range("A1").Activate
    Selection.Delete Shift:=xlToLeft
End Sub
