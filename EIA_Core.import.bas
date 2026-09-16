Attribute VB_Name = "EIA_Core"
Option Explicit

Public Const EIA_APP As String = "ExcelImageAssistant_Mohamed"
Public Const EIA_TITLE As String = "Excel Image Assistant"

' URL downloads use late-bound XMLHTTP + ADODB.Stream for 32/64-bit compatibility.


' ============================================================
' Ribbon callbacks
' ============================================================
Public Sub ImgSingle(ByVal control As IRibbonControl)
    PicInsertOne
End Sub

Public Sub ImgAll(ByVal control As IRibbonControl)
    InAllImages
End Sub

Public Sub ImgDel(ByVal control As IRibbonControl)
    DeletePictures
End Sub

Public Sub CommSingle(ByVal control As IRibbonControl)
    Call AddCommentImage(ActiveCell)
End Sub

Public Sub CommAll(ByVal control As IRibbonControl)
    AddAllCommentImages
End Sub

Public Sub CommDel(ByVal control As IRibbonControl)
    ClearComments
End Sub

Public Sub rbnDownInsertFile(ByVal control As IRibbonControl)
    DownInsertFile
End Sub

Public Sub rbnDownInsertFilet(ByVal control As IRibbonControl)
    DownInsertFile
End Sub

Public Sub rbnDownInsertFiles(ByVal control As IRibbonControl)
    DownInsertFiles
End Sub

Public Sub Compress(ByVal control As IRibbonControl)
    CompressPictures
End Sub

Public Sub rbnFileName(ByVal control As IRibbonControl)
    InsertFileName
End Sub

Public Sub rbnAllFileNames(ByVal control As IRibbonControl)
    InsertAllFileNames
End Sub

Public Sub ShowPicc(ByVal control As IRibbonControl)
    frmEIAMain.Show
End Sub

Public Sub ShowPicc1(ByVal control As IRibbonControl)
    frmEIASettings.Show
End Sub

Public Sub ShowAbout(ByVal control As IRibbonControl)
    frmEIAAbout.Show
End Sub

' ============================================================
' Build-time smoke test. Build-Addin.ps1 runs this before saving.
' If this procedure cannot run, the add-in is not released.
' ============================================================
Public Sub EIA_SelfTest()
    Dim s As String
    Dim n As Long
    Dim d As Double

    s = GetInsertMode()
    s = GetDirection()
    s = GetImageSide()
    n = GetOffsetRows()
    n = GetOffsetCols()
    n = GetPlacement()
    d = GetCommentHeight()

    If Len(EIA_TITLE) = 0 Then
        Err.Raise vbObjectError + 4100, "EIA_SelfTest", "Invalid application title."
    End If
End Sub

' ============================================================
' Settings
' ============================================================
Public Function GetPicPath() As String
    GetPicPath = GetSetting(EIA_APP, "Settings", "PicPath", "")
End Function

Public Sub SetPicPath(ByVal s As String)
    SaveSetting EIA_APP, "Settings", "PicPath", s
End Sub

Public Function GetOffsetRows() As Long
    GetOffsetRows = CLng(Val(GetSetting(EIA_APP, "Settings", "OffsetRows", "0")))
End Function

Public Sub SetOffsetRows(ByVal n As Long)
    SaveSetting EIA_APP, "Settings", "OffsetRows", CStr(n)
End Sub

Public Function GetOffsetCols() As Long
    GetOffsetCols = CLng(Val(GetSetting(EIA_APP, "Settings", "OffsetCols", "0")))
End Function

Public Sub SetOffsetCols(ByVal n As Long)
    SaveSetting EIA_APP, "Settings", "OffsetCols", CStr(n)
End Sub

Public Function GetInsertMode() As String
    GetInsertMode = GetSetting(EIA_APP, "Settings", "InsertMode", "Cells")
End Function

Public Sub SetInsertMode(ByVal s As String)
    SaveSetting EIA_APP, "Settings", "InsertMode", s
End Sub

Public Function GetCommentHeight() As Double
    Dim v As Double

    v = CDbl(Val(GetSetting(EIA_APP, "Settings", "CommentHeight", "120")))
    If v <= 0 Then
        v = 120
    End If
    GetCommentHeight = v
End Function

Public Sub SetCommentHeight(ByVal n As Double)
    SaveSetting EIA_APP, "Settings", "CommentHeight", CStr(n)
End Sub

Public Function GetPlacement() As Long
    Dim v As Long

    v = CLng(Val(GetSetting(EIA_APP, "Settings", "Placement", "1")))
    If v < 1 Or v > 3 Then
        v = 1
    End If
    GetPlacement = v
End Function

Public Sub SetPlacement(ByVal n As Long)
    SaveSetting EIA_APP, "Settings", "Placement", CStr(n)
End Sub

Public Function GetDirection() As String
    GetDirection = GetSetting(EIA_APP, "Settings", "Direction", "Down")
End Function

Public Sub SetDirection(ByVal s As String)
    SaveSetting EIA_APP, "Settings", "Direction", s
End Sub

Public Function GetImageSide() As String
    Dim s As String

    s = UCase$(Trim$(GetSetting(EIA_APP, "Settings", "ImageSide", "Right")))
    If s <> "LEFT" And s <> "RIGHT" Then
        s = "RIGHT"
    End If
    GetImageSide = s
End Function

Public Sub SetImageSide(ByVal s As String)
    s = UCase$(Trim$(s))
    If s <> "LEFT" And s <> "RIGHT" Then
        s = "RIGHT"
    End If
    SaveSetting EIA_APP, "Settings", "ImageSide", s
End Sub

Public Function GetPadding() As Double
    GetPadding = 4
End Function

Public Function PickFolder(Optional ByVal startPath As String = "") As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select A Folder"
        .AllowMultiSelect = False
        If Len(startPath) > 0 Then
            .InitialFileName = startPath
        End If
        If .Show <> -1 Then
            Exit Function
        End If
        PickFolder = .SelectedItems(1)
    End With
End Function

' ============================================================
' Picture insertion
' ============================================================
Public Sub PicInsertOne()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cell with valid file name!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    If Len(CellLookupKey(ActiveCell)) = 0 Then
        MsgBox "Please select a cell containing an image code / file name!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    If UCase$(GetInsertMode()) = "COMMENTS" Then
        Call AddCommentImage(ActiveCell)
    Else
        InsertImageForCell ActiveCell
    End If
End Sub

Public Sub PicInsert()
    PicInsertOne
End Sub

Public Sub InAllImages()
    Dim sourceCells As Collection
    Dim item As Variant
    Dim c As Range
    Dim inserted As Long
    Dim failures As Long
    Dim skipped As Long
    Dim ok As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cell(s) with valid file name(s)!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    Set sourceCells = BuildInsertAllSourceCells()
    If sourceCells Is Nothing Then
        MsgBox "Please select cell(s) with valid file name(s)!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    If sourceCells.Count = 0 Then
        MsgBox "Please select cell(s) with valid file name(s)!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo FatalError

    For Each item In sourceCells
        Set c = item
        If Len(CellLookupKey(c)) > 0 Then
            If UCase$(GetInsertMode()) = "COMMENTS" Then
                ok = AddCommentImage(c, False)
            Else
                ok = InsertImageForCell(c, False)
            End If

            If ok Then
                inserted = inserted + 1
            Else
                failures = failures + 1
            End If
        Else
            skipped = skipped + 1
        End If
    Next item

CleanExit:
    Application.ScreenUpdating = True

    If failures > 0 Then
        MsgBox CStr(inserted) & " image(s) inserted." & vbCrLf & _
               CStr(failures) & " file(s) could not be inserted.", _
               vbExclamation, EIA_TITLE
    ElseIf inserted = 0 Then
        MsgBox "No valid image files were found.", vbExclamation, EIA_TITLE
    End If
    Exit Sub

FatalError:
    failures = failures + 1
    Resume CleanExit
End Sub

Private Function BuildInsertAllSourceCells() As Collection
    Dim result As New Collection
    Dim c As Range
    Dim scanRange As Range
    Dim startCell As Range
    Dim ws As Worksheet
    Dim lastIndex As Long
    Dim i As Long
    Dim seen As Object
    Dim key As String
    Dim topLeft As Range

    Set seen = CreateObject("Scripting.Dictionary")

    If TypeName(Selection) <> "Range" Then
        Set BuildInsertAllSourceCells = result
        Exit Function
    End If

    Set ws = ActiveSheet

    ' If the user selected one or more entire columns/rows, only scan the
    ' actually used portion of the worksheet. This avoids iterating through
    ' more than one million blank cells.
    If Selection.Rows.CountLarge = ws.Rows.Count Or _
       Selection.Columns.CountLarge = ws.Columns.Count Then
        On Error Resume Next
        Set scanRange = Intersect(Selection, ws.UsedRange)
        On Error GoTo 0

        If scanRange Is Nothing Then
            Set BuildInsertAllSourceCells = result
            Exit Function
        End If

        For Each c In scanRange.Cells
            Set topLeft = c.MergeArea.Cells(1, 1)
            key = CStr(topLeft.Row) & ":" & CStr(topLeft.Column)
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add topLeft
            End If
        Next c

        Set BuildInsertAllSourceCells = result
        Exit Function
    End If

    ' A normal multi-cell selection is used exactly as the reference range.
    ' Merged cells are de-duplicated so one code creates only one picture.
    If Selection.Cells.CountLarge > 1 Then
        For Each c In Selection.Cells
            Set topLeft = c.MergeArea.Cells(1, 1)
            key = CStr(topLeft.Row) & ":" & CStr(topLeft.Column)
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add topLeft
            End If
        Next c
        Set BuildInsertAllSourceCells = result
        Exit Function
    End If

    ' With a single starting cell, continue in the configured insertion
    ' direction until the last used cell in that row/column.
    Set startCell = ActiveCell.MergeArea.Cells(1, 1)
    Set ws = startCell.Worksheet

    If UCase$(GetDirection()) = "RIGHT" Then
        lastIndex = ws.Cells(startCell.Row, ws.Columns.Count).End(xlToLeft).Column
        If lastIndex < startCell.Column Then
            lastIndex = startCell.Column
        End If

        For i = startCell.Column To lastIndex
            Set c = ws.Cells(startCell.Row, i)
            Set topLeft = c.MergeArea.Cells(1, 1)
            key = CStr(topLeft.Row) & ":" & CStr(topLeft.Column)
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add topLeft
            End If
        Next i
    Else
        lastIndex = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).Row
        If lastIndex < startCell.Row Then
            lastIndex = startCell.Row
        End If

        For i = startCell.Row To lastIndex
            Set c = ws.Cells(i, startCell.Column)
            Set topLeft = c.MergeArea.Cells(1, 1)
            key = CStr(topLeft.Row) & ":" & CStr(topLeft.Column)
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add topLeft
            End If
        Next i
    End If

    Set BuildInsertAllSourceCells = result
End Function

Public Function InsertImageForCell(ByVal sourceCell As Range, Optional ByVal showErr As Boolean = True) As Boolean
    Dim p As String
    Dim tgt As Range
    Dim lookupKey As String

    lookupKey = CellLookupKey(sourceCell)
    p = ResolveImagePath(lookupKey)
    If Len(p) = 0 Then
        If showErr Then
            MsgBox "No matching image was found for code: " & lookupKey, vbExclamation, EIA_TITLE
        End If
        Exit Function
    End If

    On Error GoTo EH
    Set tgt = GetPictureDestinationCell(sourceCell)
    RemoveEIAImagesFromCell tgt
    InsertPicture p, tgt
    InsertImageForCell = True
    Exit Function

EH:
    InsertImageForCell = False
    If showErr Then
        MsgBox "There was an error inserting the image." & vbCrLf & _
               "Code: " & lookupKey & vbCrLf & _
               "File: " & p, vbExclamation, EIA_TITLE
    End If
End Function

Public Function GetPictureDestinationCell(ByVal sourceCell As Range) As Range
    Dim src As Range
    Dim targetColumn As Long

    Set src = sourceCell.MergeArea.Cells(1, 1)

    If UCase$(GetImageSide()) = "LEFT" Then
        targetColumn = src.Column - 1
        If targetColumn < 1 Then
            Err.Raise vbObjectError + 4110, "GetPictureDestinationCell", _
                      "Cannot insert to the left of column A."
        End If
    Else
        targetColumn = src.Column + src.MergeArea.Columns.Count
        If targetColumn > src.Worksheet.Columns.Count Then
            Err.Raise vbObjectError + 4111, "GetPictureDestinationCell", _
                      "Cannot insert to the right of the last worksheet column."
        End If
    End If

    Set GetPictureDestinationCell = src.Worksheet.Cells(src.Row, targetColumn)
End Function

Private Sub RemoveEIAImagesFromCell(ByVal tgt As Range)
    Dim i As Long
    Dim shp As Shape
    Dim box As Range

    Set box = tgt.MergeArea
    For i = tgt.Worksheet.Shapes.Count To 1 Step -1
        Set shp = tgt.Worksheet.Shapes(i)
        If InStr(1, shp.AlternativeText, "Excel Image Assistant", vbTextCompare) > 0 Then
            If ShapeCenterInsideRange(shp, box) Then
                shp.Delete
            End If
        End If
    Next i
End Sub

Private Function ShapeCenterInsideRange(ByVal shp As Shape, ByVal box As Range) As Boolean
    Dim cx As Double
    Dim cy As Double

    cx = shp.Left + (shp.Width / 2)
    cy = shp.Top + (shp.Height / 2)
    ShapeCenterInsideRange = (cx >= box.Left And cx <= box.Left + box.Width And _
                              cy >= box.Top And cy <= box.Top + box.Height)
End Function

Private Function CellLookupKey(ByVal c As Range) As String
    Dim src As Range
    Dim t As String

    Set src = c.MergeArea.Cells(1, 1)
    On Error Resume Next
    t = Trim$(CStr(src.Text))
    On Error GoTo 0

    If Len(t) = 0 Or InStr(1, t, "#", vbBinaryCompare) > 0 Then
        t = Trim$(CStr(src.Value2))
    End If

    CellLookupKey = t
End Function

Private Function SafeOffset(ByVal c As Range, ByVal r As Long, ByVal col As Long) As Range
    On Error GoTo BadOffset
    Set SafeOffset = c.Offset(r, col)
    Exit Function

BadOffset:
    MsgBox "The selected offset value is beyond Excel range." & vbCrLf & _
           "Please select a different value.", vbExclamation, EIA_TITLE
    Err.Raise vbObjectError + 4101, "SafeOffset", "Offset is outside the worksheet."
End Function

Public Sub InsertPicture(ByVal filePath As String, ByVal tgt As Range)
    Dim img As Shape
    Dim box As Range
    Dim maxW As Double
    Dim maxH As Double
    Dim pad As Double
    Dim scaleFactor As Double

    Set box = tgt.MergeArea
    pad = GetPadding()

    Set img = tgt.Worksheet.Shapes.AddPicture( _
        Filename:=filePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=box.Left + pad, _
        Top:=box.Top + pad, _
        Width:=-1, _
        Height:=-1)

    img.LockAspectRatio = msoTrue

    maxW = box.Width - (pad * 2)
    maxH = box.Height - (pad * 2)
    If maxW < 2 Then maxW = 2
    If maxH < 2 Then maxH = 2

    scaleFactor = 1
    If img.Width > maxW Then
        scaleFactor = maxW / img.Width
    End If
    If (img.Height * scaleFactor) > maxH Then
        scaleFactor = maxH / img.Height
    End If

    If scaleFactor < 1 Then
        img.Width = img.Width * scaleFactor
    End If

    img.Left = box.Left + ((box.Width - img.Width) / 2)
    img.Top = box.Top + ((box.Height - img.Height) / 2)

    Select Case GetPlacement()
        Case 1
            img.Placement = xlMoveAndSize
        Case 2
            img.Placement = xlMove
        Case Else
            img.Placement = xlFreeFloating
    End Select

    img.AlternativeText = "Excel Image Assistant | " & filePath
End Sub

' ============================================================
' Comment image insertion
' ============================================================
Public Function AddCommentImage(ByVal sourceCell As Range, Optional ByVal showErr As Boolean = True) As Boolean
    Dim p As String
    Dim tgt As Range
    Dim cm As Comment
    Dim h As Double

    p = ResolveImagePath(CellLookupKey(sourceCell))
    If Len(p) = 0 Then
        If showErr Then
            MsgBox "Please select cell with valid file name!", vbExclamation, EIA_TITLE
        End If
        Exit Function
    End If

    On Error GoTo EH
    Set tgt = GetPictureDestinationCell(sourceCell)

    On Error Resume Next
    If Not tgt.Comment Is Nothing Then
        tgt.Comment.Delete
    End If
    On Error GoTo EH

    Set cm = tgt.AddComment("")
    cm.Shape.Fill.UserPicture p
    cm.Shape.LockAspectRatio = msoTrue
    h = GetCommentHeight()
    cm.Shape.Height = h
    cm.Shape.Width = h

    AddCommentImage = True
    Exit Function

EH:
    AddCommentImage = False
    If showErr Then
        MsgBox "There was an error in some of the files!" & vbCrLf & p, vbExclamation, EIA_TITLE
    End If
End Function

Public Sub AddAllCommentImages()
    Dim sourceCells As Collection
    Dim item As Variant
    Dim c As Range
    Dim inserted As Long
    Dim failures As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cell(s) with valid file name(s)!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    Set sourceCells = BuildInsertAllSourceCells()

    Application.ScreenUpdating = False
    On Error GoTo FatalError

    For Each item In sourceCells
        Set c = item
        If Len(CellLookupKey(c)) > 0 Then
            If AddCommentImage(c, False) Then
                inserted = inserted + 1
            Else
                failures = failures + 1
            End If
        End If
    Next item

CleanExit:
    Application.ScreenUpdating = True
    If failures > 0 Then
        MsgBox CStr(inserted) & " comment image(s) inserted." & vbCrLf & _
               CStr(failures) & " file(s) could not be inserted.", _
               vbExclamation, EIA_TITLE
    End If
    Exit Sub

FatalError:
    failures = failures + 1
    Resume CleanExit
End Sub

' ============================================================
' Delete tools
' ============================================================
Public Sub DeletePictures()
    Dim ans As VbMsgBoxResult
    Dim i As Long

    ans = MsgBox("Are you sure you want to delete all images ?", vbYesNo + vbQuestion, EIA_TITLE)
    If ans <> vbYes Then
        Exit Sub
    End If

    For i = ActiveSheet.Shapes.Count To 1 Step -1
        If ActiveSheet.Shapes(i).Type = msoPicture Or _
           InStr(1, ActiveSheet.Shapes(i).AlternativeText, "Excel Image Assistant", vbTextCompare) > 0 Then
            ActiveSheet.Shapes(i).Delete
        End If
    Next i
End Sub

Public Sub ClearComments()
    Dim ans As VbMsgBoxResult
    Dim c As Range

    ans = MsgBox("Are you sure you want to delete all comments ?", vbYesNo + vbQuestion, EIA_TITLE)
    If ans <> vbYes Then
        Exit Sub
    End If

    On Error Resume Next
    For Each c In ActiveSheet.UsedRange.Cells
        If Not c.Comment Is Nothing Then
            c.Comment.Delete
        End If
    Next c
    On Error GoTo 0
End Sub

Public Sub DeleteAllImagesComments()
    Dim ans As VbMsgBoxResult
    Dim i As Long
    Dim c As Range

    ans = MsgBox("Are you sure you want to delete all images/comments?", vbYesNo + vbQuestion, EIA_TITLE)
    If ans <> vbYes Then
        Exit Sub
    End If

    For i = ActiveSheet.Shapes.Count To 1 Step -1
        If ActiveSheet.Shapes(i).Type = msoPicture Or _
           InStr(1, ActiveSheet.Shapes(i).AlternativeText, "Excel Image Assistant", vbTextCompare) > 0 Then
            ActiveSheet.Shapes(i).Delete
        End If
    Next i

    On Error Resume Next
    For Each c In ActiveSheet.UsedRange.Cells
        If Not c.Comment Is Nothing Then
            c.Comment.Delete
        End If
    Next c
    On Error GoTo 0
End Sub

' ============================================================
' File name tools
' ============================================================
Public Sub InsertFileName()
    Dim fd As FileDialog
    Dim i As Long
    Dim c As Range
    Dim ans As VbMsgBoxResult
    Dim firstPath As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select pictures"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Image Files", "*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.png;*.tif;*.tiff;*.webp"
        If .Show <> -1 Then
            Exit Sub
        End If
    End With

    ans = MsgBox("The file name(s) will overwrite cells. Do you want to continue ?", _
                 vbYesNo + vbQuestion, EIA_TITLE)
    If ans <> vbYes Then
        Exit Sub
    End If

    Set c = ActiveCell
    For i = 1 To fd.SelectedItems.Count
        c.Value = GetFileNameOnly(fd.SelectedItems(i))
        If i = 1 Then
            firstPath = fd.SelectedItems(i)
        End If
        Set c = NextCell(c)
    Next i

    If Len(firstPath) > 0 Then
        SetPicPath GetParentFolder(firstPath)
    End If
End Sub

Public Sub InsertAllFileNames()
    Dim folder As String
    Dim fso As Object
    Dim fld As Object
    Dim fil As Object
    Dim c As Range
    Dim ans As VbMsgBoxResult

    folder = PickFolder(GetPicPath())
    If Len(folder) = 0 Then
        MsgBox "You didn't select a folder!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    ans = MsgBox("The file name(s) will overwrite cells. Do you want to continue ?", _
                 vbYesNo + vbQuestion, EIA_TITLE)
    If ans <> vbYes Then
        Exit Sub
    End If

    SetPicPath folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(folder)
    Set c = ActiveCell

    For Each fil In fld.Files
        If IsImageExt(fso.GetExtensionName(fil.Name)) Then
            c.Value = fil.Name
            Set c = NextCell(c)
        End If
    Next fil
End Sub

' ============================================================
' URL tools
' ============================================================
Public Sub DownInsertFile()
    If LCase$(Left$(Trim$(CStr(ActiveCell.Value2)), 4)) <> "http" Then
        MsgBox "Please select cell with valid URL address!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    Call InsertUrlImage(ActiveCell)
End Sub

Public Sub DownInsertFiles()
    Dim sourceCells As Collection
    Dim item As Variant
    Dim c As Range
    Dim inserted As Long
    Dim failures As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cell(s) with valid URL address(es)!", vbExclamation, EIA_TITLE
        Exit Sub
    End If

    Set sourceCells = BuildInsertAllSourceCells()

    Application.ScreenUpdating = False
    On Error GoTo FatalError

    For Each item In sourceCells
        Set c = item
        If Len(Trim$(CStr(c.Value2))) > 0 Then
            If LCase$(Left$(Trim$(CStr(c.Value2)), 4)) = "http" Then
                If InsertUrlImage(c, False) Then
                    inserted = inserted + 1
                Else
                    failures = failures + 1
                End If
            End If
        End If
    Next item

CleanExit:
    Application.ScreenUpdating = True
    If failures > 0 Then
        MsgBox CStr(inserted) & " URL image(s) inserted." & vbCrLf & _
               CStr(failures) & " URL(s) could not be inserted.", _
               vbExclamation, EIA_TITLE
    End If
    Exit Sub

FatalError:
    failures = failures + 1
    Resume CleanExit
End Sub

Public Function InsertUrlImage(ByVal sourceCell As Range, Optional ByVal showErr As Boolean = True) As Boolean
    Dim url As String
    Dim tmp As String
    Dim ext As String
    Dim tgt As Range

    url = Trim$(CStr(sourceCell.Value2))
    If LCase$(Left$(url, 4)) <> "http" Then
        If showErr Then
            MsgBox "Please select cell with valid URL address!", vbExclamation, EIA_TITLE
        End If
        Exit Function
    End If

    ext = UrlExtension(url)
    Randomize
    tmp = Environ$("TEMP") & "\EIA_" & Format$(Now, "yyyymmdd_hhnnss") & _
          "_" & CStr(Int(Rnd() * 1000000)) & ext

    If Not DownloadFileHttp(url, tmp) Then
        If showErr Then
            MsgBox "There was an error downloading the image URL!", vbExclamation, EIA_TITLE
        End If
        Exit Function
    End If

    On Error GoTo EH
    Set tgt = GetPictureDestinationCell(sourceCell)
    RemoveEIAImagesFromCell tgt
    InsertPicture tmp, tgt
    InsertUrlImage = True

Cleanup:
    On Error Resume Next
    If Len(tmp) > 0 Then
        Kill tmp
    End If
    On Error GoTo 0
    Exit Function

EH:
    InsertUrlImage = False
    If showErr Then
        MsgBox "There was an error in some of the files!", vbExclamation, EIA_TITLE
    End If
    Resume Cleanup
End Function

Private Function DownloadFileHttp(ByVal url As String, ByVal outputPath As String) As Boolean
    Dim http As Object
    Dim stm As Object

    On Error GoTo EH

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 Excel Image Assistant"
    http.send

    If http.readyState <> 4 Then
        Exit Function
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        Exit Function
    End If

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile outputPath, 2
    stm.Close

    DownloadFileHttp = FileExists(outputPath)
    Exit Function

EH:
    DownloadFileHttp = False
    On Error Resume Next
    If Not stm Is Nothing Then
        stm.Close
    End If
    On Error GoTo 0
End Function

Public Sub CompressPictures()
    On Error GoTo TryLegacy
    Application.CommandBars.ExecuteMso "PicturesCompress"
    Exit Sub

TryLegacy:
    On Error Resume Next
    Application.CommandBars.ExecuteMso "PictureCompress"
    On Error GoTo 0
End Sub

' ============================================================
' Helpers
' ============================================================
Public Function ResolveImagePath(ByVal rawValue As String) As String
    Dim v As String
    Dim folder As String
    Dim p As String
    Dim ext As Variant
    Dim fso As Object
    Dim fld As Object
    Dim fil As Object
    Dim baseName As String

    v = Trim$(rawValue)
    If Len(v) = 0 Then
        Exit Function
    End If

    On Error Resume Next
    If Len(Dir$(v, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0 Then
        ResolveImagePath = v
        Exit Function
    End If
    On Error GoTo 0

    folder = GetPicPath()
    If Len(folder) = 0 Then
        Exit Function
    End If

    If Right$(folder, 1) <> "\" Then
        folder = folder & "\"
    End If

    ' 1) Exact file name if the cell already contains an extension.
    p = folder & v
    If FileExists(p) Then
        ResolveImagePath = p
        Exit Function
    End If

    ' 2) Exact code / base-name with any supported image extension.
    If InStrRev(v, ".") = 0 Then
        For Each ext In Array(".jpg", ".jpeg", ".png", ".bmp", ".gif", ".ico", ".tif", ".tiff", ".webp")
            p = folder & v & CStr(ext)
            If FileExists(p) Then
                ResolveImagePath = p
                Exit Function
            End If
        Next ext
    End If

    ' 3) Case-insensitive base-name fallback. Useful when Windows/Office returns
    '    the cell code as text but the image extension/case differs.
    On Error GoTo CleanFail
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(Left$(folder, Len(folder) - 1))
    For Each fil In fld.Files
        If IsImageExt(fso.GetExtensionName(fil.Name)) Then
            baseName = fso.GetBaseName(fil.Name)
            If StrComp(baseName, v, vbTextCompare) = 0 Then
                ResolveImagePath = fil.Path
                Exit Function
            End If
        End If
    Next fil

CleanFail:
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(p, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0)
    On Error GoTo 0
End Function

Public Function NextCell(ByVal c As Range) As Range
    If UCase$(GetDirection()) = "RIGHT" Then
        Set NextCell = c.Offset(0, 1)
    Else
        Set NextCell = c.Offset(1, 0)
    End If
End Function

Public Function IsImageExt(ByVal ext As String) As Boolean
    Select Case LCase$(ext)
        Case "jpg", "jpeg", "bmp", "gif", "ico", "png", "tif", "tiff", "webp"
            IsImageExt = True
        Case Else
            IsImageExt = False
    End Select
End Function

Private Function GetFileNameOnly(ByVal fullPath As String) As String
    Dim pos As Long

    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileNameOnly = Mid$(fullPath, pos + 1)
    Else
        GetFileNameOnly = fullPath
    End If
End Function

Private Function GetParentFolder(ByVal fullPath As String) As String
    Dim pos As Long

    pos = InStrRev(fullPath, "\")
    If pos > 1 Then
        GetParentFolder = Left$(fullPath, pos - 1)
    Else
        GetParentFolder = ""
    End If
End Function

Private Function UrlExtension(ByVal url As String) As String
    Dim cleanUrl As String
    Dim q As Long
    Dim slashPos As Long
    Dim dotPos As Long
    Dim candidate As String

    cleanUrl = url
    q = InStr(1, cleanUrl, "?", vbBinaryCompare)
    If q > 0 Then
        cleanUrl = Left$(cleanUrl, q - 1)
    End If

    slashPos = InStrRev(cleanUrl, "/")
    dotPos = InStrRev(cleanUrl, ".")

    candidate = ".jpg"
    If dotPos > slashPos Then
        candidate = LCase$(Mid$(cleanUrl, dotPos))
        Select Case candidate
            Case ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"
                ' valid
            Case Else
                candidate = ".jpg"
        End Select
    End If

    UrlExtension = candidate
End Function

' ============================================================
' Fallback UI (Add-ins tab/menu) in case RibbonX is blocked/cached.
' ============================================================
Public Sub Auto_Open()
    CreateEIAFallbackMenu
End Sub

Public Sub Auto_Close()
    RemoveEIAFallbackMenu
End Sub

Public Sub CreateEIAFallbackMenu()
    Dim bar As CommandBar
    Dim pop As CommandBarControl
    Dim btn As CommandBarControl

    On Error Resume Next
    RemoveEIAFallbackMenu
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    If bar Is Nothing Then Exit Sub

    Set pop = bar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    pop.Caption = "Excel Image Assistant"
    pop.Tag = "EIA_Mohamed_17"

    Set btn = pop.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Insert Image"
    btn.OnAction = "'" & ThisWorkbook.Name & "'!PicInsertOne"

    Set btn = pop.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Insert All Images"
    btn.OnAction = "'" & ThisWorkbook.Name & "'!InAllImages"

    Set btn = pop.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Settings"
    btn.OnAction = "'" & ThisWorkbook.Name & "'!ShowSettingsDirect"

    Set btn = pop.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Info"
    btn.OnAction = "'" & ThisWorkbook.Name & "'!ShowAboutDirect"
    On Error GoTo 0
End Sub

Public Sub RemoveEIAFallbackMenu()
    Dim bar As CommandBar
    Dim ctl As CommandBarControl
    On Error Resume Next
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    If bar Is Nothing Then Exit Sub
    For Each ctl In bar.Controls
        If ctl.Tag = "EIA_Mohamed_17" Or ctl.Caption = "Excel Image Assistant" Then
            If ctl.Tag = "EIA_Mohamed_17" Then ctl.Delete
        End If
    Next ctl
    On Error GoTo 0
End Sub

Public Sub ShowSettingsDirect()
    frmEIASettings.Show
End Sub

Public Sub ShowAboutDirect()
    frmEIAAbout.Show
End Sub
