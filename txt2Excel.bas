Attribute VB_Name = "Módulo1"
Sub txtimport()
Attribute txtimport.VB_Description = "import txt"
Attribute txtimport.VB_ProcData.VB_Invoke_Func = " \n14"
'
' txtimport Macro
' import txt
'
'
' "\D16_LA50_R150Relax2_r.MAP"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & ActiveWorkbook.Path & "\Resultado.MAP" _
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
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

'    With ActiveSheet.QueryTables.Add(Connection:= _
'        "TEXT;C:\Users\adolfo.correa\Google Drive\1. Dissertação\5. ModeFrontier\D16_LA50_R150Relax2_r.MAP" _
'        , Destination:=Range("$A$1"))
'        .Name = "D16_LA50_R150Relax2_r"
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .TextFilePromptOnRefresh = False
'        .TextFilePlatform = 437
'        .TextFileStartRow = 1
'        .TextFileParseType = xlDelimited
'        .TextFileTextQualifier = xlTextQualifierDoubleQuote
'        .TextFileConsecutiveDelimiter = True
'        .TextFileTabDelimiter = False
'        .TextFileSemicolonDelimiter = False
'        .TextFileCommaDelimiter = False
'        .TextFileSpaceDelimiter = True
'        .TextFileColumnDataTypes = Array(2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'        1, 1, 1, 1, 1, 1, 1)
'        .TextFileTrailingMinusNumbers = True
'        .Refresh BackgroundQuery:=False
'    End With
    
End Sub
'Sub ult_matrix()
''
'' ult_matrix Macro
'' matrix ultima
''
''
'    ActiveCell.Offset(0, 1).Range("A1").Select
'    Selection.End(xlDown).Select
'    ActiveCell.Offset(0, -1).Range("A1").Select
'    Selection.End(xlUp).Select
'    Selection.End(xlUp).Select
'    ActiveCell.Offset(-3, 0).Range("A1").Select
'End Sub

Sub ult_matrix()
'    ActiveCell.SpecialCells(xlLastCell).Select  'va para a ultima celula do arquivo
'    Selection.End(xlToLeft).Select              'Retorna ao inicio da ultima linha
'    Selection.End(xlUp).Select                  'sobe até o inicio da coluna ctrl+up
'    Selection.End(xlUp).Select                  '(ctrl+up)
'    ActiveCell.Offset(-1, 0).Range("A1").Select 'sobe uma celula para cima (up) usa referencia relativa
    
    Range("B1").Select                              'va para a celula B1
    Selection.End(xlDown).Select                    '(ctrl+Down)
    ActiveCell.Offset(0, -1).Range("A1").Select
    'ActiveCell.Offset(0, -1).Range("B1").Select     'Regresa 1 celula (LEFT) usa referencia relativa
    Selection.End(xlUp).Select                      'sobe até o inicio da coluna ctrl+up
    Selection.End(xlUp).Select                      '(ctrl+up)
    ActiveCell.Offset(-3, 0).Range("A1").Select     'sobe 3 celulas para cima (up) usa referencia relativa

End Sub
Sub test()
    'Call txtimport
    'Call ult_matrix
    '...ActiveCell.SpecialCells(xlLastCell).Select
    '...ActiveCell.Offset(0, 1).Range("A1").Select
    'Range("B1").Select
    'Selection.End(xlDown).Select
    'ActiveCell.Offset(0, -1).Range("A1").Select
    'Selection.End(xlUp).Select                  'sobe até o inicio da coluna ctrl+up
    'Selection.End(xlUp).Select                  '(ctrl+up)
    'ActiveCell.Offset(-3, 0).Range("A1").Select
    'Path = Worksheets
    
    'MsgBox "Path: " & "1"
    'ActiveWindow.ScrollColumn = 2
    
    MsgBox "The name of the active workbook is " & ActiveWorkbook.Path
End Sub
Sub suprim()
Attribute suprim.VB_ProcData.VB_Invoke_Func = " \n14"
'
' suprim Macro
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Cells(1)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub
Sub an_map()
    Call txtimport
    Call ult_matrix
    Call suprim
    Call VME_DTE_max
End Sub
Sub VME_DTE_max()
Attribute VME_DTE_max.VB_Description = "Von Mises e Deformação maximo"
Attribute VME_DTE_max.VB_ProcData.VB_Invoke_Func = " \n14"
'
' VME_DTE_max Macro
' Von Mises e Deformação maximo
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 3
'    ActiveWindow.ScrollColumn = 4
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[5]C:R[854]C)"
    Range("S1").Select
    Selection.copy
'    ActiveWindow.ScrollColumn = 5
'    ActiveWindow.ScrollColumn = 6
'    ActiveWindow.ScrollColumn = 7
    Range("AA1").Select
    ActiveSheet.paste
    'Range("AA1").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=MAX(R[5]C:R[854]C)"
    Range("A1").Select
End Sub
Sub ult_matrix1()
'
' ult_matrix1 Macro
'

'
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A32453").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("B31601").Select
    Selection.End(xlUp).Select
    Range("A1").Select
End Sub
Sub ult_matrix2()
'
' ult_matrix2 Macro
'

'
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    'Range("A32452").Select
End Sub
Sub subir()
'
' subir Macro
'

'
    ActiveCell.Offset(-1, 0).Range("A1").Select
End Sub
Sub endereço()
'
' endereço Macro
'

'
    ActiveSheet.paste
End Sub
Sub txtimport1(ByVal file_name As String)
'
' txtimport Macro
' import txt
'
'
' "\D16_LA50_R150Relax2_r.MAP"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & ActiveWorkbook.Path & "\" & file_name _
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
       
End Sub

Public Sub prueva()
    'Teste_Mauro_RD_Linha_1_r.SNP
    'txtimport1 ("TesteRodando3_1_1_RD_Linha_1_r.SNP")
    txtimport1 ("Teste_Mauro_RD_Linha_1_r.SNP")
    Call ult_copy
    paste (426)
    'Call paste_test
    Derivada (426)
    tan_min (426)
    Call encontrar_indice
    Call coor_pto_inflex
End Sub
Sub ult_copy()
'
' ult_desc Macro
' ultimas deformações
'

'
    Range("E3").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(-425, -3).Range("A1:C426").Select
    Selection.copy
    'ActiveCell.Offset(0, -1).Range("A1").Activate
End Sub
Sub copy()
'
' copy Macro
'

'
    Selection.copy
End Sub
Sub paste(n)
'
' paste Macro
'

'
    Range("G4").Select
    ActiveSheet.paste
    Range("G3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "X"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "Y"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Z"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "Y"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "Z"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "=RC[-8]+RC[-3]"
    Range("J4").Select
    Selection.AutoFill Destination:=Range("J4:L4"), Type:=xlFillDefault
    
    
    Range("J4:L4").Select
    'Selection.AutoFill Destination:=Range("J4:L3412")
    'esticar ("A1:C426")
    esticar ("A1:C" & n)
    
    
    'Range("J4:L3412").Select
    Columns("L:L").EntireColumn.AutoFit
    Range("J3:L3").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub paste_test()
'
' paste Macro
'

'
'    Range("G4").Select
'    ActiveSheet.paste
'    Range("G3").Select
'    Application.CutCopyMode = False
'    ActiveCell.FormulaR1C1 = "X"
'    Range("H3").Select
'    ActiveCell.FormulaR1C1 = "Y"
'    Range("I3").Select
'    ActiveCell.FormulaR1C1 = "Z"
'    Range("J3").Select
'    ActiveCell.FormulaR1C1 = "X"
'    Range("K3").Select
'    ActiveCell.FormulaR1C1 = "Y"
'    Range("L3").Select
'    ActiveCell.FormulaR1C1 = "Z"
'    Range("J4").Select
'    ActiveCell.FormulaR1C1 = "=RC[-8]+RC[-3]"
'    Range("J4").Select
'    Selection.AutoFill Destination:=Range("J4:L4"), Type:=xlFillDefault
    
    
    
    'Range("J4:L4").Select
    'esticar ("A1:C426")
    
    'Selection.AutoFill Destination:=Range("J4:L430")
    'Range("J4:L3412").Select
    
    Columns("L:L").EntireColumn.AutoFit
    Range("J3:L3").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub Derivada(n)
'
' Derivada Macro
' Derivada de Z
'

'
    Range("M3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Z'"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]-R[-1]C[-1])/(RC[-2]-R[-1]C[-2])"
    Range("M5").Select
    'Dim n As Integer
    'ActiveCell.Value
    'n = 425
    Dim nn As String
    nn = "A1:A" & n - 1
    esticar (nn)
    'Range("M5:M206187").Select
End Sub
Sub esticar(ByVal matriz As String)
'
' esticar Macro
' esticar formulas
'

'
    Selection.AutoFill Destination:=ActiveCell.Range(matriz), Type:= _
        xlFillDefault
    ActiveCell.Range(matriz).Select
End Sub
Sub tan_min(ByVal n As Integer)
'
' tan_min Macro
' o minimo da tangente (derivada)
'

'
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[3]C:R[" & n + 1 & "]C)"
    Range("M3").Select
    Selection.End(xlUp).Select
End Sub
Sub encontrar_indice()
'
' encontrar_indice Macro
' Encontra o Indice do valor procurado
'

'
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=MATCH(R[1]C,R[3]C:R[428]C,0)"
    Range("M1").Select
End Sub
Sub coor_pto_inflex()
'
' coor_pto_inflex Macro
' coordenadas do ponto de inflexão
'

'
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R4C10:R429C12,R1C13,1)"
    Range("O3").Select
    Selection.AutoFill Destination:=Range("O3:Q3"), Type:=xlFillDefault
    Range("O3:Q3").Select
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R4C10:R429C12,R1C13,2)"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R4C10:R429C12,R1C13,3)"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "Y"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "Z"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Pto. De Inflexão"
    Range("O1:Q1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("O1:Q2").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.Font.Bold = True
End Sub


