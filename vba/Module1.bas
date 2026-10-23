Attribute VB_Name = "Module1"
Option Explicit
Private IsUpdating As Boolean
Private Const FIRST_ROW As Long = 13
Private Const LAST_ROW As Long = 547
Private Const PADDING As Double = 6
Private Const MAX_IMAGE_HEIGHT As Double = 390
Private Const IMAGE_PREFIX As String = "SearchImage_"
Private Const IMAGE_TAG As String = "PARAM_IMAGE|"
Private Const RENDER_VERSION As String = "3"

Public Sub UpdateParameterImages()
    ' Manual button: refresh and explicitly recover automatic Excel events.
    RefreshImages True, True
End Sub

Public Sub RefreshParameterImages()
    ' Workbook events: preserve the caller's application settings.
    RefreshImages True, False
End Sub

Public Sub EnsureParameterImages()
    RefreshImages False, False
End Sub

Private Function FindSheet(ByVal codeName As String) As Worksheet
    Dim ws As Worksheet, anchor As Range, binding As String
    Select Case codeName
        Case "Sheet1": binding = "ParameterSearchInputs"
        Case "ImageLibrarySheet": binding = "ParameterPictureIndex"
    End Select
    If Len(binding) > 0 Then
        On Error Resume Next
        Set anchor = ThisWorkbook.Names(binding).RefersToRange
        Err.Clear
        On Error GoTo 0
        If Not anchor Is Nothing Then
            Set ws = anchor.Worksheet
            If Not ws.Parent Is ThisWorkbook Then Err.Raise vbObjectError + 244, "FindSheet", "A worksheet binding points outside this workbook."
            Set FindSheet = ws
            Exit Function
        End If
    End If
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.CodeName, codeName, vbBinaryCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 230, "FindSheet", "A required worksheet was not found: " & codeName
End Function

Public Function IsParameterSearchSheet(ByVal candidate As Object) As Boolean
    Dim ui As Worksheet
    On Error GoTo NotSearch
    If candidate Is Nothing Then Exit Function
    If Not TypeOf candidate Is Worksheet Then Exit Function
    If Not candidate.Parent Is ThisWorkbook Then Exit Function
    Set ui = FindSheet("Sheet1")
    IsParameterSearchSheet = (candidate Is ui)
NotSearch:
End Function

Private Function ControlKey(ByVal cell As Range) As String
    Dim value As Variant, token As String
    value = cell.Value2
    If IsError(value) Then
        token = "[cell-error:" & CStr(cell.Text) & "]"
    Else
        token = CStr(value)
    End If
    ' Keys remain within Excel's cell length limit even for invalid pasted inputs.
    If Len(token) > 2048 Then token = Left$(token, 2048) & "[length=" & CStr(Len(token)) & "]"
    ControlKey = token
End Function

Private Function QueryKey(ByVal ui As Worksheet) As String
    QueryKey = ControlKey(ui.Range("D4")) & "|" & ControlKey(ui.Range("D5")) & "|" & ControlKey(ui.Range("D6"))
End Function

Private Sub ValidateControls(ByVal ui As Worksheet)
    Dim cell As Range, stage As String
    For Each cell In ui.Range("D4:D6")
        If IsError(cell.Value2) Then Err.Raise vbObjectError + 231, "ValidateControls", "An input cell contains an Excel error: " & cell.Address(False, False)
    Next cell
    stage = CStr(ui.Range("D4").Value2)
    If StrComp(stage, "1ST", vbTextCompare) <> 0 And StrComp(stage, "2ST", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 232, "ValidateControls", "Choose 1ST or 2ST in D4."
End Sub

Private Function ReadCount(ByVal value As Variant, ByVal maximum As Long, ByVal label As String) As Long
    Dim number As Double
    If IsError(value) Then Err.Raise vbObjectError + 233, "ReadCount", label & " contains an Excel error."
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 234, "ReadCount", label & " is not a number."
    number = CDbl(value)
    If number < 0 Or number > maximum Then Err.Raise vbObjectError + 235, "ReadCount", label & " is outside its permitted range."
    If number <> Fix(number) Then Err.Raise vbObjectError + 236, "ReadCount", label & " must be an integer."
    ReadCount = CLng(number)
End Function

Private Function CatalogRowFor(ByVal ui As Worksheet, ByVal catalog As Worksheet, ByVal rowNumber As Long, ByVal lastCatalogRow As Long) As Long
    Dim index As Long, value As Variant
    value = ui.Cells(rowNumber, 6).Value2
    index = ReadCount(value, lastCatalogRow - 1, "Catalog index")
    If index < 1 Then Err.Raise vbObjectError + 237, "CatalogRowFor", "A result has no catalog index."
    CatalogRowFor = index + 1
    If StrComp(CStr(catalog.Cells(index + 1, 1).Value2), CStr(ui.Range("D4").Value2), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 238, "CatalogRowFor", "A result belongs to a different stage."
    If CStr(catalog.Cells(index + 1, 2).Value2) <> CStr(ui.Cells(rowNumber, 2).Value2) Then Err.Raise vbObjectError + 239, "CatalogRowFor", "A result does not match its catalog entry."
End Function

Private Function ResultKey(ByVal ui As Worksheet, ByVal catalog As Worksheet, ByVal resultCount As Long) As String
    Dim r As Long, index As Long, lastCatalogRow As Long, token As String
    lastCatalogRow = catalog.Cells(catalog.Rows.Count, 1).End(xlUp).Row
    token = CStr(resultCount) & ";"
    For r = FIRST_ROW To FIRST_ROW + resultCount - 1
        index = CatalogRowFor(ui, catalog, r, lastCatalogRow)
        token = token & CStr(index) & ":" & CStr(catalog.Cells(index, 11).Value2) & ";"
    Next r
    If Len(token) > 30000 Then Err.Raise vbObjectError + 240, "ResultKey", "The result description is too long."
    ResultKey = token
End Function

Private Function FailedForSameInput(ByVal ui As Worksheet, ByVal key As String) As Boolean
    On Error GoTo NotFailed
    FailedForSameInput = (Len(CStr(ui.Range("P5").Value2)) > 0 And CStr(ui.Range("P9").Value2) = key)
NotFailed:
End Function

Private Function CacheIsCurrent(ByVal ui As Worksheet, ByVal catalog As Worksheet, ByVal key As String) As Boolean
    Dim count As Long, images As Long
    On Error GoTo NotCurrent
    If CStr(ui.Range("P4").Value2) <> key Then Exit Function
    If Len(CStr(ui.Range("P5").Value2)) <> 0 Then Exit Function
    If CStr(ui.Range("P7").Value2) <> RENDER_VERSION Then Exit Function
    count = ReadCount(ui.Range("D8").Value2, LAST_ROW - FIRST_ROW + 1, "Result count")
    images = ReadCount(ui.Range("P6").Value2, count, "Cached image count")
    If CStr(ui.Range("P8").Value2) <> ResultKey(ui, catalog, count) Then Exit Function
    CacheIsCurrent = ImagesAreCurrent(ui, key, images)
NotCurrent:
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

Private Sub DeletePictureQuietly(ByVal item As Shape)
    On Error Resume Next
    If item Is Nothing Then Exit Sub
    item.Visible = msoFalse
    item.Delete
End Sub

Private Sub ClearOutputImages(ByVal ui As Worksheet, ByVal strict As Boolean)
    Dim i As Long, item As Shape, failure As String
    For i = ui.Shapes.Count To 1 Step -1
        Set item = ui.Shapes(i)
        If IsOutputImage(item) Then
            On Error Resume Next
            item.Visible = msoFalse
            Err.Clear
            item.Delete
            If Err.Number <> 0 And Len(failure) = 0 Then failure = item.Name & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    If strict And Len(failure) > 0 Then Err.Raise vbObjectError + 241, "ClearOutputImages", "Cannot delete a previous picture: " & failure
End Sub

Private Sub ClearPasteArtifacts(ByVal ui As Worksheet, ByVal key As String)
    Dim i As Long, item As Shape
    For i = ui.Shapes.Count To 1 Step -1
        Set item = ui.Shapes(i)
        If IsOutputImage(item) Then
            If Not StartsWith(item.Name, IMAGE_PREFIX) Or Not StartsWith(item.AlternativeText, IMAGE_TAG & key & "|row=") Then DeletePictureQuietly item
        End If
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

Private Function TryPasteCard(ByVal original As Shape, ByVal library As Worksheet, ByVal ui As Worksheet, ByVal rowNumber As Long, ByVal key As String, ByVal attempt As Long, ByRef lastError As String) As Shape
    Dim pasted As Object, item As Shape, errorNumber As Long
    On Error GoTo Failed
    Application.CutCopyMode = False
    library.Activate
    If attempt = 1 Then
        original.Copy
    Else
        original.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    End If
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
    DeletePictureQuietly item
    If Not pasted Is Nothing Then pasted.Delete
    ClearPasteArtifacts ui, key
    Application.CutCopyMode = False
    On Error GoTo 0
    If errorNumber = 18 Then Err.Raise 18, "TryPasteCard", lastError
End Function

Private Function PasteCard(ByVal original As Shape, ByVal library As Worksheet, ByVal ui As Worksheet, ByVal rowNumber As Long, ByVal key As String) As Shape
    Dim attempt As Long, item As Shape, lastError As String
    For attempt = 1 To 3
        Set item = TryPasteCard(original, library, ui, rowNumber, key, attempt, lastError)
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

Private Sub RefreshImages(ByVal force As Boolean, ByVal recoverEvents As Boolean)
    Dim ui As Worksheet, library As Worksheet, catalog As Worksheet
    Dim picture As Shape, original As Shape, frame As Range
    Dim previousSheet As Object, previousSelection As Object
    Dim matchRow As Variant, catalogRow As Long, resultCount As Long
    Dim r As Long, pictureCount As Long, libraryLastRow As Long, lastCatalogRow As Long
    Dim location As String, pictureName As String, key As String, results As String, failure As String
    Dim oldScreen As Boolean, oldEvents As Boolean, oldCancel As Long, didMutate As Boolean

    If IsUpdating Then Exit Sub
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCancel = Application.EnableCancelKey
    IsUpdating = True
    On Error GoTo Failed
    On Error Resume Next
    Set previousSheet = Application.ActiveSheet
    If TypeName(Application.Selection) = "Range" Then Set previousSelection = Application.Selection
    Err.Clear
    On Error GoTo Failed
    Set ui = FindSheet("Sheet1")
    key = QueryKey(ui)
    If Not force Then
        ' Do not repeat the same failed clipboard/validation operation on every Calculate event.
        If FailedForSameInput(ui, key) Then GoTo Finish
    End If
    Set library = FindSheet("ImageLibrarySheet")
    Set catalog = ThisWorkbook.Names("SearchCatalog").RefersToRange.Worksheet
    If Not catalog.Parent Is ThisWorkbook Then Err.Raise vbObjectError + 242, , "The catalog must belong to this workbook."
    If Not force Then
        If CacheIsCurrent(ui, catalog, key) Then GoTo Finish
    End If
    didMutate = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler

    ui.Range("P4:P9").ClearContents
    ClearOutputImages ui, True
    ui.Range("G13:J547").UnMerge
    ui.Range("G13:J547").ClearContents
    ui.Rows("13:547").RowHeight = 105
    ValidateControls ui

    Application.Calculate
    catalog.Calculate
    ui.Calculate
    If QueryKey(ui) <> key Then Err.Raise vbObjectError + 213, , "Search inputs changed during refresh. Press Refresh images."
    resultCount = ReadCount(ui.Range("D8").Value2, LAST_ROW - FIRST_ROW + 1, "Result count")
    results = ResultKey(ui, catalog, resultCount)
    libraryLastRow = library.Cells(library.Rows.Count, 1).End(xlUp).Row
    lastCatalogRow = catalog.Cells(catalog.Rows.Count, 1).End(xlUp).Row
    If libraryLastRow < 2 Then Err.Raise vbObjectError + 243, , "The picture library is empty."
    library.Visible = xlSheetVisible

    For r = FIRST_ROW To FIRST_ROW + resultCount - 1
        catalogRow = CatalogRowFor(ui, catalog, r, lastCatalogRow)
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

    If ReadCount(ui.Range("D8").Value2, LAST_ROW - FIRST_ROW + 1, "Result count") <> resultCount Then Err.Raise vbObjectError + 245, , "The number of search results changed during refresh."
    If QueryKey(ui) <> key Or ResultKey(ui, catalog, resultCount) <> results Then Err.Raise vbObjectError + 219, , "Search results changed during refresh. Press Refresh images."
    If Not ImagesAreCurrent(ui, key, pictureCount) Then Err.Raise vbObjectError + 220, , "The image set or its bounds failed validation."
    ui.Range("G4").Value2 = CStr(ui.Range("D4").Value2) & " / " & CStr(ui.Range("D6").Value2)
    ui.Range("P4").Value2 = key
    ui.Range("P6").Value2 = pictureCount
    ui.Range("P7").Value2 = RENDER_VERSION
    ui.Range("P8").Value2 = results
    GoTo Finish

Failed:
    failure = "Image refresh error " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Not ui Is Nothing Then
        didMutate = True
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        ui.Range("P4:P9").ClearContents
        ClearOutputImages ui, False
        ui.Range("G13:J547").ClearContents
        ui.Range("P5").Value2 = failure
        ui.Range("P9").Value2 = key
    ElseIf recoverEvents Then
        MsgBox failure, vbExclamation, "Parameter images"
    End If

Finish:
    On Error Resume Next
    If Not didMutate Then
        IsUpdating = False
        If recoverEvents Then Application.EnableEvents = True
        Exit Sub
    End If
    ' Keep events disabled while restoring focus; Activate can itself raise workbook events.
    Application.EnableEvents = False
    If Not library Is Nothing Then
        If Not ui Is Nothing Then ui.Activate
        library.Visible = xlSheetVeryHidden
    End If
    If Not previousSheet Is Nothing Then
        If previousSheet Is library Then
            If Not ui Is Nothing Then ui.Activate
        Else
            previousSheet.Activate
            If Not previousSelection Is Nothing Then previousSelection.Select
        End If
    End If
    Application.CutCopyMode = False
    Application.ScreenUpdating = oldScreen
    Application.EnableCancelKey = oldCancel
    IsUpdating = False
    If recoverEvents Then
        Application.EnableEvents = True
    Else
        Application.EnableEvents = oldEvents
    End If
End Sub
