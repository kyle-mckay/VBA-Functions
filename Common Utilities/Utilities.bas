Attribute VB_Name = "Utilities"
Sub UnhideAll(Optional wb As Workbook)
' Function: Loops through all worksheets within a (optional) specified workbook and toggles visibility to visible regardless of state
Dim ws As Worksheet
If wb Is Nothing Then Set wb = ThisWorkbook

For Each ws In wb.Worksheets
    If ws.Visible = xlSheetHidden Then ws.Visible = xlSheetVisible
Next
End Sub
Sub ClearSheet(Optional ws As Worksheet, Optional paste As Boolean)
' Function: Clears the specified worksheet of cell contents and styalizes them back to normal.
' If specified will also input "Paste Here" in cell A1
If ws Is Nothing Then Set ws = ActiveSheet
Cells.ClearContents
Cells.Style = "Normal"
Range("A1").Select
If paste = True Then Range("A1").Value = "Paste Here"
End Sub
Sub Screen(b As Boolean)
' Function: Enables or disables screen updating based on input bool value
Application.ScreenUpdating = b
End Sub
