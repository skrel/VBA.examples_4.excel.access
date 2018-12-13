Attribute VB_Name = "Module1"
Sub PDFActiveSheet()
'www.contextures.com
'for Excel 2010 and later
Dim wsA As Worksheet
Dim wbA As Workbook
Dim strTime As String
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim strPathFile As String
Dim myFile As Variant
On Error GoTo errHandler

Set wbA = ActiveWorkbook
Set wsA = ActiveSheet
strTime = Format(Now(), "yyyymmdd\_hhmm")

'get active workbook folder, if saved
strPath = wbA.Path
If strPath = "" Then
  strPath = Application.DefaultFilePath
End If
strPath = strPath & "\"

'replace spaces and periods in sheet name
strName = Replace(wsA.Name, " ", "")
strName = Replace(strName, ".", "_")

'create default name for savng file
strFile = strName & "_" & strTime & ".pdf"
strPathFile = strPath & strFile

'use can enter name and
' select folder for file
myFile = Application.GetSaveAsFilename _
    (InitialFileName:=strPathFile, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Select Folder and FileName to save")

'export to PDF if a folder was selected
If myFile <> "False" Then
    wsA.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=myFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    'confirmation message with file info
    MsgBox "PDF file has been created: " _
      & vbCrLf _
      & myFile
End If

exitHandler:
    Exit Sub
errHandler:
    MsgBox "Could not create PDF file"
    Resume exitHandler
End Sub


