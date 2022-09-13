VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkbookSelection 
   Caption         =   "Workbook Selection"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "WorkbookSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkbookSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Workbook Selection
''' Purpose:
''' When user form activated, all combo boxes are filled with instances of each open (editable) workbook.

Private Sub Execute_Click()
'
    If cb1.Value = "" Or cb2.Value = "" Or cb3.Value = "" Then
        MsgBox ("At least one combo box is empty. Will not run.")
    End If
    BulkExport cb1.Value, cb2.Value, cb3.Value, closewb, check2
    If check2 = True Then MsgBox ("Do something with check2")
End Sub
Private Sub UserForm_Activate()
'When activated, centers on screen and triggers combo boxes to contain WB names
ClearForm
With WBSelection
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With
cbAdditem
End Sub
Sub cbAdditem()
'For each WB that is not this workbook, add to all combo boxes in this form
Dim wb As Workbook
Dim ctrl As Control, cbox As MSForms.ComboBox

For Each wb In Application.Workbooks
    If wb.Name <> ThisWorkbook.Name Then
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "ComboBox" Then
                Set cbox = cb
                cbox.AddItem wb.Name
            End If
        Next
    End If
Next
End Sub
Private Sub UserForm_Deactivate()
' Call function when form is deactivated
ClearForm
End Sub
Sub ClearForm()
' Clears all data from the user form when called
    With WBSelection
        .cb1.Clear
        .cb2.Clear
        .cb3.Clear
        .closewb.Value = False
        .check2.Value = False
    End With
End Sub
Sub SelectRange()
Dim c As Range
Cells(Cells.CurrentRegion.Rows.Count, Cells.CurrentRegion.Columns.Count).Select
Set c = Selection
Range("A1", c).Select
End Sub
Sub BulkExport(cb1 As String, cb2 As String, cb3 As String, closewb As Boolean, check2 As Boolean)
Application.ScreenUpdating = False
Application.EnableEvents = False
purgesheets
Dim wb As Workbook
    If cb1 <> "skip" Then
        Set wb = Workbooks(cb1)
        wb.Activate
        SelectRange
        Selection.Copy
        ThisWorkbook.Activate
        Sheets("MB51-1").Select ' Change this to your sheets import location
        Range("A1").Select
        ActiveSheet.Paste
        wb.Application.CutCopyMode = False
        If closewb = True Then wb.Close
        Me.cb1.Value = "Complete"
    End If
    If cb2 <> "skip" Then
        Set wb = Workbooks(cb2)
        wb.Activate
        SelectRange
        Selection.Copy
        ThisWorkbook.Activate
        Sheets("MB51-2").Select ' Change this to your sheets import location
        Range("A1").Select
        ActiveSheet.Paste
        wb.Application.CutCopyMode = False
        If closewb = True Then wb.Close
        Me.cb2.Value = "Complete"
    End If
    If cb3 <> "skip" Then
        Set wb = Workbooks(cb3)
        wb.Activate
        SelectRange
        Selection.Copy
        ThisWorkbook.Activate
        Sheets("Err").Select ' Change this to your sheets import location
        Range("A1").Select
        ActiveSheet.Paste
        wb.Application.CutCopyMode = False
        If closewb = True Then wb.Close
        Me.cb3.Value = "Complete"
    End If
Application.ScreenUpdating = True
Application.EnableEvents = True
If closewb = True Then
    ThisWorkbook.Save
'    ThisWorkbook.Close 'Uncomment if you want to also close the WB when this is complete
End If
Me.hide
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub cb1_AfterUpdate()
'cbRemoveItem cb1.Value, cb1.Name
'End Sub

'Sub cbRemoveItem(wb As String, cbname As String)
'Dim ctrl As Control, cbox As MSForms.ComboBox
'
'For Each ctrl In Me.Controls
'    If TypeName(ctrl) = "ComboBox" And ctrl.Name <> cbname Then
'        Set cbox = ctrl
'
'        cbox.Items.Remove wb
'    End If
'Next
'
'End Sub
