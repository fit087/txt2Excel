
'Sheets(“INSERT WORKSHEET NAME HERE”).Cells.ClearContents

'The VBA for clearing the entire sheet is nice and straight-forward.  It can be achieved a few ways, like most things in Excel, but the common practice is to do the following:

Sub ClearSheet()

    Sheets(“Sheet1”).Cells.ClearContents

End Sub


'The clear contents method is useful but sometimes there will be additional formatting on a worksheet, table borders, cell shadings etc.  Clear contents does not remove these formats, it simply clears the entire sheet of data, not formats.

'To clear your entire sheet of data and formats you need to use a delete option and to do this we simply change our VBA script to:

Sub ClearSheet()

Sheets(“Sheet1”).Cells.Delete

End Sub

'There may be times when deleting the entire sheet is the best option rather than just clearing the entire contents.  To do this we can again change our existing code to the following:

Sub ClearSheet()

Sheets(“Sheet1”).Delete

End Sub


' Save
'With this Method we can Save a Workbook as it's existing name or as another name and path.  We can trick Excel into thinking a Workbook is already had it's changes Saved.  This may also be a good time to introduce the two methods available to refer to the active Workbook.  The first one is:
Sub SaveActiveWorkbook()
    ActiveWorkbook.Save
End Sub

' The second is:
Sub SaveThisWorkbook()
    ThisWorkbook.Save
End Sub

Sub sbVBS_To_SAVE_ActiveWorkbook()
    ActiveWorkbook.Save
End Sub

Sub ActivateAnotherWorkbookViaName()
    Workbooks("Book2").Activate
End Sub

Sub ActivateAnotherWorkbookViaIndex()
    Workbooks(3).Activate
End Sub


' Save as
'  expression .SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)

' expression A variable that represents a Workbook object.

Set NewBook = Workbooks.Add 
Do 
    fName = Application.GetSaveAsFilename 
Loop Until fName <> False 
NewBook.SaveAs Filename:=fName


' NewBook

Sub AddOne()
    Workbooks.Add
End Sub
		

'A better way to create a new workbook is to assign it to an object variable. In the following example, the Workbook object returned by the Add method is assigned to an object variable, newBook. Next, several properties of newBook are set. You can easily control the new workbook using the object variable.

Sub AddNew()
Set NewBook = Workbooks.Add
    With NewBook
        .Title = "All Sales"
        .Subject = "Sales"
        .SaveAs Filename:="Allsales.xls"
    End With
End Sub
		

