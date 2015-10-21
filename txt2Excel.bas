Attribute VB_Name = "Módulo1"
Sub Importtxt()
Attribute Importtxt.VB_Description = "importar aquivo de texto"
Attribute Importtxt.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' Importtxt Macro
' importar aquivo de texto
'
' Atalho do teclado: Ctrl+Shift+T
'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\Laila\Desktop\D16_LA50_R150Relax2_r.MAP", Destination:=Range( _
        "$A$1"))
        .CommandType = 0
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
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
Sub endoffile()
Attribute endoffile.VB_Description = "Final do arquivo"
Attribute endoffile.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' endoffile Macro
' Final do arquivo
'
' Atalho do teclado: Ctrl+Shift+F
'
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlToLeft).Select
End Sub
Sub inicio_de_tabela()
Attribute inicio_de_tabela.VB_Description = "Inicio da tabela referencias relativas"
Attribute inicio_de_tabela.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' inicio_de_tabela Macro
' Inicio da tabela referencias relativas
'
' Atalho do teclado: Ctrl+Shift+S
'
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(-2, 0).Range("A1").Select
End Sub
Sub eliminar()
Attribute eliminar.VB_Description = "Eliminar demas matrizes"
Attribute eliminar.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' eliminar Macro
' Eliminar demas matrizes
'
' Atalho do teclado: Ctrl+Shift+E
'
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Cells(1)).Select
End Sub
Sub eliminar1()
Attribute eliminar1.VB_Description = "Eliminar matrizes"
Attribute eliminar1.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' eliminar1 Macro
' Eliminar matrizes
'
' Atalho do teclado: Ctrl+Shift+D
'
    ActiveCell.Offset(-1, 0).Range("A1:B1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Cells(1)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub
