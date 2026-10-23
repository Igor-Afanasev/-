Attribute VB_Name = "Module1"
Option Explicit
Private IsUpdating As Boolean
Private Const FIRST_ROW As Long = 13
Private Const LAST_ROW As Long = 547
Private Const PADDING As Double = 6
Private Const MAX_IMAGE_HEIGHT As Double = 390
Private Const IMAGE_PREFIX As String = "SearchImage_"
Private Const IMAGE_TAG As String = "PARAM_IMAGE|"
Private Const RENDER_VERSION As String = "2"

Public Sub UpdateParameterImages()
    RefreshImages True
End Sub

Public Sub EnsureParameterImages()
    RefreshImages False
End Sub

Private Function QueryKey(ByVal ui As Worksheet) As String
    QueryKey = CStr(ui.Range("D4").Value2) & "|" & CStr(ui.Range("D5").Value2) & "|" & CStr(ui.Range("D6").Value2)
End Function

Private Function StartsWith(ByVal value As String, ByVal prefix As String) As Boolean
    StartsWith = (Left$(value, Len(prefix)) = prefix)
End Function

Private Function IsOutputImage(ByVal item As Shape) As Boolean
    Dim alt As String
    On Error GoTo NotOutput
    If item.Type = msoPicture Or item.Type = msoLinkedPicture Then
        IsOutputImage = True
        Exit Function
    End If
    alt = item.AlternativeText
    If StartsWith(item.Name, IMAGE_PREFIX) Or StartsWith(item.Name, "RefCard_") Or StartsWith(alt, IMAGE_TAG) Or StartsWith(alt, "RefCard_") Then
        IsOutputImage = True
        Exit Function
    End If
NotOutput:
End Function

Private Sub ClearOutputImages(ByVal ui As Worksheet)
    Dim i As Long, item As Shape
    For i = ui.Shapes.Count To 1 Step -1
        Set item = ui.Shapes(i)
        If IsOutputImage(item) Then
            On Error Resume Next
            item.Visible = msoFalse
            item.Delete
            On Error GoTo 0
        End If
    Next i
End Sub

Private Sub ClearUnclaimedCopies(ByVal ui As Worksheet)
    Dim i As Long, item As Shape
    For i = ui.Shapes.Count To 1 Step -1
        Set item = ui.Shapes(i)
        If StartsWith(item.Name, "RefCard_") Or StartsWith(item.AlternativeText, "RefCard_") Then item.Delete
    Next i
End Sub

Private Function ImagesAreCurrent(ByVal ui As Worksheet, ByVal key As String, ByVal expectedCount As Long) As Boolean
    Dim item As Shape, frame As Range, count As Long, rowNumber As Long
    On Error GoTo NotCurrent
    For Each item In ui.Shapes
        If IsOutputImage(item) Then
            If Not StartsWith(item.Name, IMAGE_PREFIX) Then Exit Function
            rowNumber = CLng(Mid$(item.Name, Len(IMAGE_PREFIX) + 1))
            If rowNumber < FIRST_ROW Or rowNumber > LAST_ROW Then Exit Function
            If item.Visible <> msoTrue Then Exit Function
            If item.AlternativeText <> IMAGE_TAG & key & "|row=" & CStr(rowNumber) Then Exit Function
            If item.TopLeftCell.Row <> rowNumber Or item.TopLeftCell.Column <> 7 Then Exit Function
            Set frame = ui.Cells(rowNumber, 7)
            If frame.MergeCells Then Exit Function
            If item.Left < frame.Left + PADDING - 0.2 Then Exit Function
            If item.Left + item.Width > frame.Left + frame.Width - PADDING + 0.2 Then Exit Function
            If item.Top < frame.Top + PADDING - 0.2 Then Exit Function
            If item.Top + item.Height > frame.Top + frame.Height - PADDING + 0.2 Then Exit Function
            count = count + 1
        End If
    Next item
    ImagesAreCurrent = (count = expectedCount)
NotCurrent:
End Function

Private Function TryPasteCard(ByVal original As Shape, ByVal library As Worksheet, ByVal ui As Worksheet, ByVal rowNumber As Long, ByVal key As String, ByRef lastError As String) As Shape
    Dim pasted As Object, item As Shape, errorNumber As Long
    On Error GoTo Failed
    library.Activate
    original.Copy
    ui.Activate
    Set pasted = ui.Pictures.Paste
    If pasted Is Nothing Then Err.Raise vbObjectError + 210, , "Excel did not return the pasted picture."
    Set item = ui.Shapes(pasted.Name)
    item.Name = IMAGE_PREFIX & CStr(rowNumber)
    item.AlternativeText = IMAGE_TAG & key & "|row=" & CStr(rowNumber)
    Application.CutCopyMode = False
    Set TryPasteCard = item
    Exit Function
Failed:
    errorNumber = Err.Number
    lastError = Err.Description
    On Error Resume Next
    ClearUnclaimedCopies ui
    Application.CutCopyMode = False
    On Error GoTo 0
    If errorNumber = 18 Then Err.Raise 18, "TryPasteCard", lastError
End Function

Private Function PasteCard(ByVal original As Shape, ByVal library As Worksheet, ByVal ui As Worksheet, ByVal rowNumber As Long, ByVal key As String) As Shape
    Dim attempt As Long, item As Shape, lastError As String
    For attempt = 1 To 3
        Set item = TryPasteCard(original, library, ui, rowNumber, key, lastError)
        If Not item Is Nothing Then
            Set PasteCard = item
            Exit Function
        End If
    Next attempt
    Err.Raise vbObjectError + 211, "PasteCard", "Cannot paste a reference card: " & lastError
End Function

Private Sub FitInsideColumnG(ByVal item As Shape, ByVal ui As Worksheet, ByVal rowNumber As Long)
    Dim frame As Range, scaleFactor As Double, availableWidth As Double
    Dim imageWidth As Double, imageHeight As Double
    Set frame = ui.Cells(rowNumber, 7)
    availableWidth = frame.Width - 2 * PADDING
    If availableWidth < 10 Or item.Width <= 0 Or item.Height <= 0 Then Err.Raise vbObjectError + 212, , "Column G is too narrow, or the image has invalid dimensions."
    scaleFactor = 1
    If item.Width > availableWidth Then scaleFactor = availableWidth / item.Width
    If item.Height * scaleFactor > MAX_IMAGE_HEIGHT Then scaleFactor = MAX_IMAGE_HEIGHT / item.Height
    imageWidth = item.Width * scaleFactor
    imageHeight = item.Height * scaleFactor
    item.LockAspectRatio = msoFalse
    item.Width = imageWidth
    item.Height = imageHeight
    item.LockAspectRatio = msoTrue
    ui.Rows(rowNumber).RowHeight = Application.Max(105, imageHeight + 2 * PADDING)
    item.Left = frame.Left + PADDING
    item.Top = frame.Top + PADDING
    item.Placement = xlMoveAndSize
End Sub

Private Sub RefreshImages(ByVal force As Boolean)
    Dim ui As Worksheet, library As Worksheet, catalog As Worksheet
    Dim picture As Shape, original As Shape, frame As Range
    Dim matchRow As Variant, catalogRow As Long, resultCount As Long
    Dim r As Long, pictureCount As Long, libraryLastRow As Long
    Dim location As String, pictureName As String, key As String, failure As String
    Dim oldScreen As Boolean, oldCancel As Long

    If IsUpdating Then Exit Sub
    oldScreen = Application.ScreenUpdating
    oldCancel = Application.EnableCancelKey
    On Error GoTo Failed
    Set ui = ThisWorkbook.Worksheets(1)
    key = QueryKey(ui)
    If Not force Then
        If CStr(ui.Range("P4").Value2) = key And Len(CStr(ui.Range("P5").Value2)) = 0 And CStr(ui.Range("P7").Value2) = RENDER_VERSION Then
            If ImagesAreCurrent(ui, key, CLng(ui.Range("P6").Value2)) Then Exit Sub
        End If
    End If
    IsUpdating = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler

    Set library = ThisWorkbook.Worksheets("_ImageLibrary")
    Set catalog = ThisWorkbook.Names("SearchCatalog").RefersToRange.Worksheet
    ' Invalidate and remove the previous render before any operation that can fail.
    ui.Range("P4:P7").ClearContents
    ClearOutputImages ui
    ui.Range("G13:J547").UnMerge
    ui.Range("G13:J547").ClearContents
    ui.Rows("13:547").RowHeight = 105

    Application.Calculate
    catalog.Calculate
    ui.Calculate
    If QueryKey(ui) <> key Then Err.Raise vbObjectError + 213, , "Search inputs changed during refresh. Press Refresh images."
    resultCount = CLng(ui.Range("D8").Value2)
    If resultCount < 0 Or resultCount > LAST_ROW - FIRST_ROW + 1 Then Err.Raise vbObjectError + 214, , "Invalid search result count."
    libraryLastRow = library.Cells(library.Rows.Count, 1).End(xlUp).Row
    library.Visible = xlSheetVisible

    For r = FIRST_ROW To FIRST_ROW + resultCount - 1
        If Not IsNumeric(ui.Cells(r, 6).Value2) Or Len(CStr(ui.Cells(r, 6).Value2)) = 0 Then Err.Raise vbObjectError + 215, , "Missing catalog row."
        catalogRow = CLng(ui.Cells(r, 6).Value2) + 1
        If catalogRow < 2 Or catalogRow > catalog.Cells(catalog.Rows.Count, 1).End(xlUp).Row Then Err.Raise vbObjectError + 216, , "Catalog index is outside its range."
        If StrComp(CStr(catalog.Cells(catalogRow, 1).Value2), CStr(ui.Range("D4").Value2), vbTextCompare) <> 0 Or CStr(catalog.Cells(catalogRow, 2).Value2) <> CStr(ui.Cells(r, 2).Value2) Then Err.Raise vbObjectError + 217, , "Search results do not match the selected stage."
        location = CStr(catalog.Cells(catalogRow, 11).Value2)
        Set frame = ui.Cells(r, 7)
        If Len(location) > 0 Then
            matchRow = Application.Match(location, library.Range("A2:A" & CStr(libraryLastRow)), 0)
            If IsError(matchRow) Then Err.Raise vbObjectError + 218, , "The reference card is missing from the picture library."
            pictureName = CStr(library.Cells(CLng(matchRow) + 1, 2).Value2)
            Set original = library.Shapes(pictureName)
            Set picture = PasteCard(original, library, ui, r, key)
            FitInsideColumnG picture, ui, r
            pictureCount = pictureCount + 1
        Else
            frame.Value2 = "No reference card for this parameter"
        End If
    Next r

    If QueryKey(ui) <> key Then Err.Raise vbObjectError + 219, , "Search inputs changed during refresh. Press Refresh images."
    If Not ImagesAreCurrent(ui, key, pictureCount) Then Err.Raise vbObjectError + 220, , "The image set or its bounds failed validation."
    ui.Range("G4").Value2 = CStr(ui.Range("D4").Value2) & " / " & CStr(ui.Range("D6").Value2)
    ui.Range("P4").Value2 = key
    ui.Range("P6").Value2 = pictureCount
    ui.Range("P7").Value2 = RENDER_VERSION
    GoTo Finish

Failed:
    failure = "Image refresh error " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Not ui Is Nothing Then
        ui.Range("P4:P7").ClearContents
        ClearOutputImages ui
        ui.Range("G13:J547").ClearContents
        ui.Range("P5").Value2 = failure
    End If

Finish:
    On Error Resume Next
    If Not ui Is Nothing Then ui.Activate
    If Not library Is Nothing Then library.Visible = xlSheetVeryHidden
    Application.CutCopyMode = False
    Application.ScreenUpdating = oldScreen
    Application.EnableCancelKey = oldCancel
    ' The manual refresh button also recovers automatic updates after disabled events.
    Application.EnableEvents = True
    IsUpdating = False
End Sub
