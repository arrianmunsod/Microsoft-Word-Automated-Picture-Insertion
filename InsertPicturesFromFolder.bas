Attribute VB_Name = "Module2"

Option Explicit

Public Sub InsertPicturesFromFolder()

    Dim doc As Document
    Dim fd As FileDialog
    Dim folderPath As String

    Dim rootStory As Range
    Dim story As Range
    Dim searchRange As Range

    Dim placeholder As String
    Dim imagePath As String

    Dim pic As InlineShape
    Dim insertedCount As Long
    Dim missingCount As Long
    Dim missingList As String
    Dim result As String

    Set doc = ActiveDocument

    ' Select the picture folder each time.
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Select the folder containing your pictures"
        .AllowMultiSelect = False

        If .Show <> -1 Then Exit Sub

        folderPath = .SelectedItems(1)
    End With

    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    Application.ScreenUpdating = False

    On Error GoTo HandleError

    ' Search document stories, including headers and footers.
    For Each rootStory In doc.StoryRanges

        Set story = rootStory

        Do While Not story Is Nothing

            Set searchRange = story.Duplicate

            Do
                With searchRange.Find
                    .ClearFormatting
                    .Text = "\{\{*\}\}"
                    .MatchWildcards = True
                    .Forward = True
                    .Wrap = wdFindStop
                End With

                If Not searchRange.Find.Execute Then Exit Do

                placeholder = searchRange.Text

                imagePath = FindMatchingImage( _
                                folderPath, placeholder)

                If Len(imagePath) > 0 Then

                    ' Remove the placeholder.
                    searchRange.Text = vbNullString

                    ' Insert the picture at that location.
                    Set pic = searchRange.InlineShapes.AddPicture( _
                        fileName:=imagePath, _
                        LinkToFile:=False, _
                        SaveWithDocument:=True)

                    ' AUTOMATICALLY FIT THE IMAGE IN ITS CELL.
                    FitImageInTableCell pic

                    insertedCount = insertedCount + 1

                    ' Continue after the inserted image.
                    searchRange.SetRange _
                        Start:=pic.Range.End, _
                        End:=pic.Range.End

                Else

                    missingCount = missingCount + 1
                    missingList = missingList & _
                                  placeholder & vbCrLf

                    ' Keep unmatched placeholders unchanged.
                    searchRange.Collapse wdCollapseEnd

                End If

            Loop

            ' Process the next linked story, if any.
            Set story = story.NextStoryRange

        Loop

    Next rootStory

    Application.ScreenUpdating = True

    result = "Picture insertion finished." & vbCrLf & vbCrLf & _
             "Pictures inserted and fitted: " & insertedCount & _
             vbCrLf & _
             "Missing matches: " & missingCount

    If missingCount > 0 Then
        result = result & vbCrLf & vbCrLf & _
                 "Placeholders without matching pictures:" & _
                 vbCrLf & missingList
    End If

    MsgBox result, vbInformation, "Word Picture Automation"
    Exit Sub

HandleError:

    Application.ScreenUpdating = True

    MsgBox "An error occurred: " & Err.Description, _
           vbExclamation, "Word Picture Automation"

End Sub


Private Sub FitImageInTableCell(ByVal shp As InlineShape)

    Dim tblCell As Cell
    Dim cellRange As Range

    Dim cellW As Single
    Dim cellH As Single
    Dim availW As Single
    Dim availH As Single

    Dim margin As Single
    Dim ratio As Double
    Dim newW As Single
    Dim newH As Single

    ' 1.5-point margin on each side.
    margin = 1.5

    ' Only fit images located inside table cells.
    If shp.Range.Cells.Count = 0 Then Exit Sub

    Set tblCell = shp.Range.Cells(1)

    ' Get the containing cell dimensions.
    cellW = tblCell.Width
    cellH = tblCell.Height

    ' Calculate usable space after margins.
    availW = cellW - (margin * 2)
    availH = cellH - (margin * 2)

    If availW <= 0 Or availH <= 0 Then Exit Sub
    If shp.Height <= 0 Then Exit Sub

    ' Preserve the image's aspect ratio.
    ratio = shp.Width / shp.Height

    ' Fit within both available dimensions.
    If (availW / availH) < ratio Then

        newW = availW
        newH = newW / ratio

    Else

        newH = availH
        newW = newH * ratio

    End If

    ' Resize without distortion.
    With shp
        .LockAspectRatio = msoTrue
        .Width = newW
        .Height = newH
    End With

    ' Center the image horizontally.
    Set cellRange = shp.Range.Paragraphs(1).Range

    cellRange.ParagraphFormat.Alignment = _
        wdAlignParagraphCenter

    ' Center the image vertically.
    tblCell.VerticalAlignment = wdCellAlignVerticalCenter

End Sub


Private Function FindMatchingImage( _
    ByVal folderPath As String, _
    ByVal placeholder As String) As String

    Dim key As String
    Dim fileName As String
    Dim baseName As String
    Dim bestMatch As String

    Dim score As Long
    Dim bestScore As Long
    Dim ambiguous As Boolean

    key = NormalizeName(placeholder)

    fileName = Dir$(folderPath & "*.*")

    Do While Len(fileName) > 0

        If IsSupportedImage(fileName) Then

            baseName = NormalizeName( _
                Left$(fileName, InStrRev(fileName, ".") - 1))

            score = MatchScore(key, baseName)

            If score > bestScore Then

                bestScore = score
                bestMatch = folderPath & fileName
                ambiguous = False

            ElseIf score > 0 And score = bestScore Then

                ' Do not guess between equally strong matches.
                ambiguous = True

            End If

        End If

        fileName = Dir$()

    Loop

    If ambiguous Then
        FindMatchingImage = vbNullString
    Else
        FindMatchingImage = bestMatch
    End If

End Function


Private Function NormalizeName(ByVal value As String) As String

    value = UCase$(value)

    value = Replace(value, "{{", "")
    value = Replace(value, "}}", "")
    value = Replace(value, " ", "")
    value = Replace(value, "_", "")
    value = Replace(value, "-", "")
    value = Replace(value, ".", "")

    NormalizeName = value

End Function


Private Function MatchScore( _
    ByVal key As String, _
    ByVal candidate As String) As Long

    If Len(key) = 0 Or Len(candidate) = 0 Then Exit Function

    If key = candidate Then

        MatchScore = 100

    ElseIf InStr(1, candidate, key, vbTextCompare) > 0 Then

        MatchScore = 80

    ElseIf InStr(1, key, candidate, vbTextCompare) > 0 Then

        MatchScore = 60

    End If

End Function


Private Function IsSupportedImage( _
    ByVal fileName As String) As Boolean

    Dim extension As String

    extension = LCase$(Mid$( _
        fileName, InStrRev(fileName, ".") + 1))

    Select Case extension

        Case "jpg", "jpeg", "png", "bmp", _
             "gif", "tif", "tiff"

            IsSupportedImage = True

    End Select

End Function


