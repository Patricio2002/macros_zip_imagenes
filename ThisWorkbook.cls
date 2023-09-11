VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function

Sub Zip_All_Files_in_Folder()
    Dim FileNameZip, folderName
    Dim strDate As String, DefPath As String
    Dim oApp As Object
    Dim saveToFolder As String
    Dim GetBook As String
    
    DefPath = Application.ActiveWorkbook.Path
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If
    
    saveToFolder = Application.ActiveWorkbook.Path
    GetBook = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    saveToFolder = saveToFolder & "\" & GetBook
    
    If Right(saveToFolder, 1) <> "\" Then
        saveToFolder = saveToFolder & "\"
    End If
    
    folderName = saveToFolder   '<< Change

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = DefPath & GetBook & ".zip"

    'Create empty Zip File
    NewZip (FileNameZip)

    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(folderName).Items

    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace(FileNameZip).Items.Count = _
       oApp.Namespace(folderName).Items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0
End Sub

