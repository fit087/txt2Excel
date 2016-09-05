Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub automatic()
    'Application.Calculation = xlCalculationManual
    'Application.Calculation = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    
    'Application.Calculation = xlCalculationAutomatic
    'Application.Calculation = True
    'Application.ScreenUpdating = True
    'Application.EnableEvents = True


    Dim no As Integer
    'Call DeleteSheet
    'Call NewSheet
    Call ClearSheet1
    no = an_snp()
    an_map (no)
    'Application.Calculation = xlCalculationAutomatic
    'Application.Calculation = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
Sub testando()
    Dim a As String
    a = Range("A1").Value
    MsgBox ("Rodou")
    MsgBox (a)
    MsgBox (ActiveWorkbook.Path)
End Sub

'Sub an_map(no As Integer)
''
'' an_map Macro
'' Análise dos Resultados do MAP File
''
'' Atalho do teclado: Not Registered
''
'    Call txtimport
'    Call ult_matrix
'    Call suprim
'    Call VME_DTE_max
'End Sub
Private Sub test1()
    ' an_map (317)
'    Range("R5").Select
'    ActiveCell.FormulaR1C1 = ChrW(916) ' " \n916"
'    Range("S5") = ChrW(916)

' Take the file name of the workbook with extension and add an ending
    Dim file_name As String
    file_name = ActiveWorkbook.Name
    MsgBox (Left(file_name, InStrRev(file_name, ".") - 1) + "_d_0.map")
End Sub
Sub an_map(no As Integer)
'
' an_map Macro
' Análise dos Resultados do MAP File
'
' Atalho do teclado: Not Registered
'
    Dim file_name, root_name  As String
    file_name = ActiveWorkbook.Name
    'MsgBox (Left(file_name, InStrRev(file_name, ".") - 1) + "_d_0.map")
    root_name = Left(file_name, InStrRev(file_name, ".") - 1)
    
    Call txtimport(root_name)
    'Call txtimport
    Call ult_matrix
    Call suprim
    'Call VME_DTE_max
    VME_DTE_max1 (no)
End Sub
Public Function an_snp() As Integer
'
' an_snp Macro
' Análise dos Resultados do SNP File
'
' Atalho do teclado: Not Registered
'
    'txtimport1 ("TesteRodando3_1_1_RD_Linha_1_r.SNP")
    'txtimport1 ("Teste_Mauro_RD_Linha_1_r.SNP")
    txtimport1 ("TesteRodandoLR1_d_0.SNP")
    Call ult_copy
    'paste (426)
    'Call paste
    Dim line_pnt As Integer
    line_pnt = paste
    'Call paste_test
    'Derivada (426)
    Derivada (line_pnt)
    'tan_min (426)
    tan_min (line_pnt)
    Call encontrar_indice
    Call coor_pto_inflex
    an_snp = Range("M1").Value
    Call delete_snp_data
End Function
Private Sub txtimport(root_name As String)
'Private Sub txtimport()
'
' txtimport Macro
' import the text file called Resultado.MAP
' The 2 first columns are taken with text the anothers how
'
'
' "\D16_LA50_R150Relax2_r.MAP"
' "\TesteRodandoLR1_d_0.MAP"
' "\Resultado.MAP"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & ActiveWorkbook.Path & "\TesteRodandoLR1_d_0.MAP" _
        , Destination:=Range("$A$1"))
        .name = "D16_LA50_R150Relax2_r"
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
End Sub
Private Sub ult_matrix()
'
' ult_matrix Macro
' Find the last matrix on the *.MAP document
' Posiciona a selecao na ultima celda da penultima matriz (uma fila antes do inicio da ultima matriz)
'
' Atalho do teclado: Not Registered
'

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
Private Sub suprim()
'
' suprim Macro
' Delete the other matrices on the *.map file

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Cells(1)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub
Private Sub VME_DTE_max()
'
' VME_DTE_max Macro
' Toma o Von Mises e Deformação maxima from the *.MAP file
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[5]C:R[854]C)"
    Range("S1").Select
    Selection.copy
    Range("AA1").Select
    ActiveSheet.paste
    'Range("AA1").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=MAX(R[5]C:R[854]C)"
    Range("A1").Select
End Sub
Private Sub VME_DTE_max1(inflex_pto As Integer)
'
' VME_DTE_max Macro
' Toma o Von Mises e Deformação maxima from the *.MAP file
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[5]C:R[854]C)"
    Range("S1").Select
    Selection.copy
    Range("AA1").Select
    ActiveSheet.paste
    'Range("AA1").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=MAX(R[5]C:R[854]C)"
    
    ' Sagbend
    ' VME
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[6]C:R[" & 2 * inflex_pto + 5 & "]C)"
    Range("R1") = "VME_SagBend"
    
    ' DTE
    Range("S1").Select
    Selection.copy
    Range("AA1").Select
    ActiveSheet.paste
    Range("Z1") = "DTE_SagBend"
    
    ' Overbend
    ' VME
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[" & 2 * inflex_pto + 6 & "]C:R[856]C)"
    Range("R1") = "VME_OverBend"

    ' DTE
    Range("S1").Select
    Selection.copy
    Range("AA1").Select
    ActiveSheet.paste
    Range("Z1") = "DTE_OverBend"

    
    Range("A1").Select
End Sub

Private Sub txtimport1(ByVal file_name As String)
'
' txtimport Macro
' Import txt for *.SNP file
'
'
' "\D16_LA50_R150Relax2_r.MAP"
' TesteRodandoLR1_d_0.MAP
    'MsgBox (ActiveWorkbook.Path)
    'On Error Resume Next
    'On Error GoTo 0
    On Error GoTo didnt_import
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & ActiveWorkbook.Path & "\" & file_name _
        , Destination:=Range("$A$1"))
        .name = "TesteRodandoLR1_d_0.MAP"
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
    Exit Sub
    
didnt_import:
       MsgBox ("Arquivo '" & file_name & "' não encontrado em '" & ActiveWorkbook.Path & "' ")
       Exit Sub
    
End Sub
Private Sub ult_copy()
'
' ult_desc Macro
' Copia as ultimas deformações from *.SNP file
'

'
    Range("E3").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    
    ' Atention esta sendo usando 426 que es un valor definido de pontos da linha
    ActiveCell.Offset(-425, -3).Range("A1:C426").Select
    Selection.copy
    'ActiveCell.Offset(0, -1).Range("A1").Activate
End Sub
Private Sub tes()
    Range("E4").PasteSpecial
End Sub

'Private Sub paste(n)
'Private Sub paste()
Private Function paste() As Integer

'
' paste Macro
' Cola as deformações de n pontos do lado da configuração original (SNP file)

'
    'Dim n As Integer
    ' ---------------------------
    Range("G4").Select
    ActiveSheet.paste
    'n = Selection.Rows.Count
    paste = Selection.Rows.Count
    Range("G3").Select
    'MsgBox (n)
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ChrW(916) & "X" 'ActiveCell.FormulaR1C1 = "X"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = ChrW(916) & "Y"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = ChrW(916) & "Z"
    ' ----------------------------
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
    '-------------------------------
    
    Range("J4:L4").Select
    'Selection.AutoFill Destination:=Range("J4:L3412")
    'esticar ("A1:C426")
    'esticar ("A1:C" & n)
    esticar ("A1:C" & paste)
    
    
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
End Function
Private Sub Derivada(n)
'
' Derivada Macro
' Derivada do dz/dy por diferenças finitas
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
Private Sub tan_min(ByVal n As Integer)
'
' tan_min Macro
' Acha o minimo da tangente (derivada) à curva z=f(y)
'

'
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[3]C:R[" & n + 1 & "]C)"
    Range("M3").Select
    Selection.End(xlUp).Select
End Sub
Private Sub encontrar_indice()
'
' encontrar_indice Macro
' Encontra o Indice do valor procurado nesse caso a tangente minima que corresponde ao ponto de inflexão.
'

'
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=MATCH(R[1]C,R[3]C:R[428]C,0)"
    Range("M1").Select
End Sub
Private Sub coor_pto_inflex()
'
' coor_pto_inflex Macro
' ubica na matriz de pontos as coordenadas do ponto de inflexão com o conhecimento do indice do pto de minimo.
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
Private Sub esticar(ByVal matriz As String)
'
' esticar Macro
' esticar formulas até fim dos dados de entrada.
'

'
    Selection.AutoFill Destination:=ActiveCell.Range(matriz), Type:= _
        xlFillDefault
    ActiveCell.Range(matriz).Select
End Sub
Private Sub delete_snp_data()
'
' delete_snp_data Macro
'

'
    Range("J3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, -9).Range("A1").Select
'    ActiveCell.SpecialCells(xlLastCell).Select
'    Selection.End(xlUp).Select
'    Selection.End(xlToLeft).Select
'    Selection.End(xlToLeft).Select
'    Selection.End(xlToLeft).Select
'    ActiveCell.Offset(2, 9).Range("A1").Select
'    Selection.End(xlDown).Select
'    ActiveCell.Offset(1, -9).Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub

Private Sub delete_all()
'
' delete_all Macro
'

'
    Cells.Select
    'Range("Q1").Activate
    'Selection.QueryTable.Delete
    Selection.QueryTable.Delete
    Selection.ClearContents
End Sub
Private Sub delete__all_cells()
'
' delete__all_cells Macro
'

'
    ActiveCell.Cells.Select
    Selection.Delete Shift:=xlUp
End Sub
Sub DeleteSheet()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Sheets("Plan1").Delete
    'Sheets(1).Delete
    'ActiveSheet.Delete
    'ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = False
    
'    Option Explicit
'Sub DelSht()
'    Application.DisplayAlerts = False
'    Sheets("Sheet1").Delete
'    Application.DisplayAlerts = True
'End Sub

End Sub

Private Static Sub NewSheet()
    'ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count), Count:=1
    'Sheets.Add After:=Sheets(Sheets.Count)
    Sheets.Add.name = "Plan1"
End Sub
Private Sub ClearSheet1()
    Dim wb As Workbook, ws As Worksheet, flag As Boolean
    flag = False
        
    Set wb = ActiveWorkbook     ' Atribuição de Variável Objeto
    
    For Each ws In wb.Worksheets
    
        If ws.name = "Plan1" Then
            'Do something here
            Call DeleteSheet
            Call NewSheet
            flag = True
        End If
    
    Next
    
    If flag = False Then
        Call NewSheet
    End If
End Sub


