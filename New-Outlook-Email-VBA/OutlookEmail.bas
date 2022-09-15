Attribute VB_Name = "OutlookEmail"
Sub AddComponent()
' Attempts to assign references in order to run the code
    On Error Resume Next
    Dim MSR, VBAE As String
    MSO = "C:\Program Files\Common Files\microsoft shared\OFFICE16\MSO.dll" ' Microsoft Office 16.0 Object Library
    MSOUTL = "C:\Program Files\Microsoft Office\root\Office16\MSOUTL.OLB" ' Microsoft Office Outlook 16.0 Object Library
    ARR = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb" ' Microsoft Core Library

    With ThisWorkbook.VBProject.References
        .AddFromFile (MSO)
        .AddFromFile (MSOUTL)
        .AddFromFile (ARR)
    End With
End Sub
Sub asdasad()
GenerateEmail
End Sub
Sub GenerateEmail(Optional automation As String, Optional bSendNow As Boolean)
'ScStop
Dim ws As Excel.Worksheet
Dim SaveToDirectory As String, ChartName As String, TempDirectory As String, SheetDirectory As String, ExportDirectory As String, CopyName As String, plant As String, Recipients As String, CarbonCopy As String, Subject As String
Dim hsheet As Worksheet
Dim objChrt As ChartObject
Dim myChart As Chart
Dim ChartNameNumber As Integer
Dim bExportCharts As Boolean, bExportRange As Boolean, bKillDirectory As Boolean
Dim FileList As ArrayList
TempDirectory = Environ("temp") & "\"
''''' User Defined Settings
    SheetName = "ExportSheet.xlsx"
    Set hsheet = Sheets("Sheet1")
    Recipients = "" ' Seperate each email address with ;
    CarbonCopy = "" ' Seperate each email address with ;
    Subject = "Email Subject"
    ExportDirectory = Environ("Userprofile") & "\" & "Downloads\Repair and Returns\"
    
    bKillDirectory = False ' Delete files from export directory on run
        If bKillDirectory = True Then KillFiles ExportDirectory
    
    bExportRange = True ' Convert ranges to range in exportrange:
    bExportCharts = True ' Convert ranges to range in ExportCharts:
    bSendNow = False ' Sends email immediately upong generation
    
'''''''''''''''''''''''''''
SheetDirectory = TempDirectory & SheetName ' Set export sheet directory

' Save copy of workbook
Application.DisplayAlerts = False
hsheet.Select
ActiveWorkbook.SaveCopyAs Filename:=SheetDirectory
Application.DisplayAlerts = True

hsheet.Select
If bExportCharts = True Then
ExportCharts:
    ExportCharts (ExportDirectory & "Charts\")
End If

Dim outlookApp As Outlook.Application
Dim myMail As Outlook.MailItem

    ActiveWorkbook.Save
    On Error GoTo 0
    'Generate Email
    Set outlookApp = CreateObject("outlook.application")
    Set myMail = outlookApp.CreateItem(olMailItem)
        If bExportCharts = True Then
        ' Import exported charts as attachments if true
            Set FileList = LoopThroughFiles(ExportDirectory & "Charts\") ' Obtain files
            For Each f In FileList ' Loop through files and add as attachments
                source_file = f
                myMail.Attachments.Add source_file
            Next
        End If
        If bExportRange = True Then
        ' Convert ranges to images and attach
            Dim rgp As String
            Dim i As Integer
            Dim rg As Range
            '''' This code must run per worksheet
ExportRange:
                For i = 1 To 1 ' The 1 on the right being the number of ranges you wish to import to email
                    If i = 1 Then Set rg = Sheets("Sheet1").Range("A1:B6")
                    rg.Worksheet.Select
                    rgp = ExportRange(rg, i, ExportDirectory & "Range\", True)
                    myMail.Attachments.Add rgp
                Next
            ''''
        End If
        
        myMail.Attachments.Add SheetDirectory
        
        myMail.To = Recipients
        myMail.CC = CarbonCopy
        myMail.Subject = Subject
        myMail.HTMLBody = "<span LANG=EN>" _
                & "Good Morning All,<br> " _
                & "<br>" _
                & "This is an email" _
                & "<br></font></span>"
                ' Note - If you wish to add images into the body follow the following format for a new line
                '"<img src=""cid:Chart-Volume.png"" width=60%>"
                ' Where Chart-Volume.png = the file name

        myMail.Display
        If bSendNow = True Then myMail.Send
        
        hsheet.Select
'ScResume
End Sub
Sub zoomOut(Optional Charts As Worksheet)
    If Charts Is Nothing Then Set Charts = ActiveSheet
    Charts.Activate
    ActiveWindow.Zoom = 90
End Sub
Sub zoomIn(Optional Charts As Worksheet)
    If Charts Is Nothing Then Set Charts = ActiveSheet
    Charts.Activate
    ActiveWindow.Zoom = 250
End Sub
Public Function ExportRange(rg As Range, c As Integer, Optional ExportDirectory, Optional btotext As Boolean) As String
     
Dim ws As Worksheet, ns As Worksheet
Dim chartO As ChartObject
Dim lWidth As Long, lHeight As Long
If ExportDirectory = "" Then ExportDirectory = Environ("Userprofile") & "\Downloads\Ranges\"
KillFiles (ExportDirectory)
Set ws = ActiveSheet

If btotext = True Then
    RangeToText rg
    Set rg = Selection
    Set ns = ActiveSheet
End If
rg.CopyPicture xlScreen, xlPicture
lWidth = rg.Width
lHeight = rg.Height

If btotext = True Then
    Set chartO = ns.ChartObjects.Add(Left:=0, Top:=0, Width:=lWidth, Height:=lHeight)
Else
    Set chartO = ws.ChartObjects.Add(Left:=0, Top:=0, Width:=lWidth, Height:=lHeight)
End If

chartO.Activate
With chartO.Chart
 .Paste
 .Export Filename:=ExportDirectory & "Case" & c & ".png", Filtername:="PNG"
End With

chartO.Delete
If btotext = True Then
    Application.DisplayAlerts = False
    ns.Delete
    Application.DisplayAlerts = True
End If
   ExportRange = ExportDirectory & "Case" & c & ".png"
End Function
Sub createJpg(SheetName As String, xRgPic As Range, i As Integer)
    'Dim xRgPic As Range
    ThisWorkbook.Activate
    Worksheets(SheetName).Activate
    'Set xRgPic = ThisWorkbook.Worksheets(SheetName).Range(xRgAddrss)
    xRgPic.CopyPicture
    With ThisWorkbook.Worksheets(SheetName).ChartObjects.Add(xRgPic.Left, xRgPic.Top, xRgPic.Width, xRgPic.Height)
        .Activate
        .Chart.Paste
        .Chart.Export Environ$("temp") & "\" & i & ".png", "PNG"
    End With
    Worksheets(SheetName).ChartObjects(Worksheets(SheetName).ChartObjects.count).Delete
Set xRgPic = Nothing
End Sub
Sub ExportCharts(Optional ExportDirectory As String)
Dim TempDirectory As String
TempDirectory = Environ("temp") & "\Excel Chart Export\"
If ExportDirectory = "" Then
    ExportDirectory = Environ("Userprofile") & "\Downloads\Charts\"
End If
KillFiles ExportDirectory
    SaveToDirectory = ActiveWorkbook.path & "\"
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = False Then ws.Visible = True
        ws.Activate 'go there
        For Each objChrt In ws.ChartObjects
            objChrt.Activate
            Set myChart = objChrt.Chart

            myFileName = ExportDirectory & myChart.Name & ".png"
                        
            On Error Resume Next
            MkDir ExportDirectory
            Kill myFileName ' Delete file if already exists
            myChart.Export Filename:=myFileName, Filtername:="PNG"
            On Error GoTo 0

        Next
    Next
End Sub
Sub testsda()
GenerateEmail
End Sub
Public Function LoopThroughFiles(ExportDirectory As String) As ArrayList

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Set LoopThroughFiles = New ArrayList

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFolder = oFSO.GetFolder(ExportDirectory)

For Each oFile In oFolder.Files
    LoopThroughFiles.Add oFile
Next oFile

End Function
Sub KillFiles(path As String)
    On Error Resume Next
    Kill path & Chr(92) & "*.*" ' Delete all files in path
'    RmDir path ' Remove folder if empty
    MkDir path ' Make path after remove
    On Error GoTo 0
End Sub
Sub RangeToText(rg As Range)
Dim ns As Worksheet
    Sheets.Add After:=rg.Worksheet
    Set ns = ActiveSheet
    rg.Copy
    ns.Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Columns.AutoFit
End Sub

