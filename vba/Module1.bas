Attribute VB_Name = "Module1"
Option Explicit
Private IsUpdating As Boolean

Public Sub UpdateParameterImages()
    Dim ui As Worksheet, library As Worksheet, catalog As Worksheet
    Dim picture As Shape, original As Shape
    Dim frame As Range, matchRow As Variant, catalogRow As Long
    Dim r As Long, i As Long, lastResult As Long
    Dim location As String, pictureName As String
    Dim oldEvents As Boolean, oldScreen As Boolean
    Dim failure As String

    If IsUpdating Then Exit Sub
    IsUpdating = True
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ui = ThisWorkbook.Worksheets(1)
    Set library = ThisWorkbook.Worksheets("_ImageLibrary")
    Set catalog = ThisWorkbook.Names("SearchCatalog").RefersToRange.Worksheet
    ui.Range("P5").ClearContents
    Application.Calculate

    For i = ui.Shapes.Count To 1 Step -1
        If Left$(ui.Shapes(i).Name, 12) = "SearchImage_" Then ui.Shapes(i).Delete
    Next i
    ui.Range("G13:J547").ClearContents
    ui.Rows("13:547").RowHeight = 105
    library.Visible = xlSheetVisible

    For r = 13 To 547
        If Len(CStr(ui.Cells(r, 2).Value2)) = 0 Then Exit For
        Set frame = ui.Range(ui.Cells(r, 7), ui.Cells(r, 10))
        If IsNumeric(ui.Cells(r, 6).Value2) Then
            catalogRow = CLng(ui.Cells(r, 6).Value2) + 1
            location = CStr(catalog.Cells(catalogRow, 11).Value2)
        Else
            location = vbNullString
        End If
        matchRow = Application.Match(location, library.Range("A2:A171"), 0)
        If Len(location) > 0 And Not IsError(matchRow) Then
            pictureName = CStr(library.Cells(CLng(matchRow) + 1, 2).Value2)
            Set original = library.Shapes(pictureName)
            library.Activate
            original.Copy
            ui.Activate
            ui.Paste
            Set picture = ui.Shapes(ui.Shapes.Count)
            picture.Name = "SearchImage_" & CStr(r)
            picture.LockAspectRatio = msoTrue
            picture.Width = frame.Width - 12
            If picture.Height > 390 Then picture.Height = 390
            ui.Rows(r).RowHeight = Application.Max(120, picture.Height + 12)
            picture.Left = frame.Left + (frame.Width - picture.Width) / 2
            picture.Top = frame.Top + 6
            picture.Placement = xlMoveAndSize
            picture.AlternativeText = CStr(ui.Range("D4").Value2) & " / " & CStr(ui.Cells(r, 2).Value2)
            Application.CutCopyMode = False
        Else
            frame.Cells(1, 1).Value2 = "No reference card for this parameter"
        End If
    Next r
    ui.Range("G4").Value2 = CStr(ui.Range("D4").Value2) & " / " & CStr(ui.Range("D6").Value2)
    ui.Range("P4").Value2 = CStr(ui.Range("D4").Value2) & "|" & CStr(ui.Range("D5").Value2) & "|" & CStr(ui.Range("D6").Value2)
    GoTo Finish

Failed:
    failure = "Image refresh error " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Not ui Is Nothing Then ui.Range("P5").Value2 = failure

Finish:
    On Error Resume Next
    If Not ui Is Nothing Then ui.Activate
    If Not library Is Nothing Then library.Visible = xlSheetVeryHidden
    Application.CutCopyMode = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    IsUpdating = False
End Sub
