Attribute VB_Name = "Módulo2"
Sub soma()


Dim xlApp As Excel.Application
Dim x, y As Integer

Set xlApp = New Excel.Application
x = ActiveDocument.ContentControls(1).Range.Text
y = ActiveDocument.ContentControls(2).Range.Text

ActiveDocument.ContentControls(3).Range.Text = _
    xlApp.WorksheetFunction.Sum(x, y)

End Sub

