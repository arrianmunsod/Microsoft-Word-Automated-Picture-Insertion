Attribute VB_Name = "Module1"
Option Explicit

Public Sub AutoExec()

    InitializeWordEvents

End Sub


Public Sub InitializeWordEvents()

    ThisDocument.InitializeEvents

End Sub


Public Sub FitSelectedImageInTableCell(ByVal Sel As Selection)

    Dim shp As InlineShape
    Dim tblCell As Cell
    
    Dim cellW As Single
    Dim cellH As Single
    
    Dim availW As Single
    Dim availH As Single
    
    Dim margin As Single
    
    Dim ratio As Double
    Dim newW As Single
    Dim newH As Single
    
    Dim cellRange As Range
    
    ' =========================================================
    ' 2-pixel margin
    ' 2 pixels ˜ 1.5 points at 96 DPI
    ' =========================================================
    margin = 1.5
    
    ' Get selected picture
    Set shp = Sel.InlineShapes(1)
    
    ' Get containing table cell
    Set tblCell = shp.Range.Cells(1)
    
    ' Get cell dimensions
    cellW = tblCell.Width
    cellH = tblCell.Height
    
    ' Available area after margins
    availW = cellW - (margin * 2)
    availH = cellH - (margin * 2)
    
    If availW <= 0 Or availH <= 0 Then Exit Sub
    
    ' Preserve aspect ratio
    ratio = shp.Width / shp.Height
    
    ' Determine limiting dimension
    If (availW / availH) < ratio Then
        
        ' Width is limiting
        newW = availW
        newH = newW / ratio
        
    Else
        
        ' Height is limiting
        newH = availH
        newW = newH * ratio
        
    End If
    
    ' Resize picture
    With shp
        
        .LockAspectRatio = msoTrue
        
        .Width = newW
        .Height = newH
        
    End With
    
    ' Center horizontally
    Set cellRange = shp.Range.Paragraphs(1).Range
    
    cellRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    ' Center vertically
    tblCell.VerticalAlignment = wdCellAlignVerticalCenter

End Sub


