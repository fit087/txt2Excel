Attribute VB_Name = "Módulo4"
Option Explicit
Dim oCode As Variant
Dim oFile As Variant
Dim oSheet As Variant
Dim oCol As Integer
Dim oBundesland As Variant
Dim oBland As Variant
Dim oRange As Range
 
Sub open_txtfile()
    'Clearing the sheet
    Plan1.Cells.Clear
    
    Application.EnableEvents = False
    oFile = ".\D16.txt"
    oSheet = "PLZ"
    oCol = 2
    Workbooks.Open Filename:=oFile
    Set oRange = Workbooks(oFile).Worksheets(oSheet).Range("A2:C29390")
    oBundesland = Application.WorksheetFunction.VLookup(oCode, oRange, oCol, False)
    Application.EnableEvents = True
    MsgBox "VLookup result is : " & oBundesland
End Sub

'Option Explicit
'Dim oCode As Variant
'Dim oFile As Variant
'Dim oSheet As Variant
'Dim oCol As Integer
'Dim oBundesland As Variant
'Dim oBland As Variant
'Dim oRange As Range
'
'Sub getPLZnetData(oCode)
'    Application.EnableEvents = False
'    oFile = "F:\Marketing\Add In\XLS\lib\PLZ.xls"
'    oSheet = "PLZ"
'    oCol = 2
'    Workbooks.Open Filename:=oFile
'    Set oRange = Workbooks(oFile).Worksheets(oSheet).Range("A2:C29390")
'    oBundesland = Application.WorksheetFunction.VLookup(oCode, oRange, oCol, False)
'    Application.EnableEvents = True
'    MsgBox "VLookup result is : " & oBundesland
'End Sub

