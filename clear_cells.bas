
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