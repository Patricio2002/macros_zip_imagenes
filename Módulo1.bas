Attribute VB_Name = "Módulo1"
Option Explicit


Sub ExportImages()


    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "No worksheet is active!", vbExclamation
        Exit Sub
    End If
    
    Dim GetBook As String
    Dim saveToFolder As String
    Dim saveAsFilename As String
    Dim currentShape As Shape
    Dim shapeCount As Long
    Dim BookName As String
    
    Application.ScreenUpdating = False
    
    saveToFolder = Application.ActiveWorkbook.Path
    GetBook = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    MkDir saveToFolder & "\" & GetBook
    saveToFolder = saveToFolder & "\" & GetBook
    If Right(saveToFolder, 1) <> "\" Then
        saveToFolder = saveToFolder & "\"
    End If
    
    shapeCount = 0
    For Each currentShape In ActiveSheet.Shapes
        If currentShape.Type = msoPicture Then
            If Not Intersect(Columns("A:A"), currentShape.TopLeftCell) Is Nothing Then
                saveAsFilename = Cells(currentShape.TopLeftCell.Row, "C").Value & ".jpg"
                ExportImage saveToFolder, saveAsFilename, currentShape
                shapeCount = shapeCount + 1
            End If
        End If
    Next currentShape
    
    Application.ScreenUpdating = True
    
     
    Kill saveToFolder & "*.jpg"
    RmDir saveToFolder
    
End Sub


Sub ExportImage(ByVal saveToFolder As String, ByVal saveAsFilename As String, ByVal shapeToExport As Shape)


    Dim ws As Worksheet
    
    Set ws = shapeToExport.Parent
    
    With ws.ChartObjects.Add(Left:=0, Top:=0, Width:=shapeToExport.Width, Height:=shapeToExport.Height)
        .Activate
        With .Chart
            .ChartArea.Format.Line.Visible = msoFalse
            shapeToExport.Copy
            .Paste
            .Export Filename:=saveToFolder & saveAsFilename, filtername:="JPG"
        End With
        .Delete
    End With
    Call ThisWorkbook.Zip_All_Files_in_Folder
    
End Sub
