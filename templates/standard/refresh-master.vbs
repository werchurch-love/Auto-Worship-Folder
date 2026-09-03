Option Explicit

' Refreshes automatically generated PPTX hyperlinks on master.pptx.
'
' Required master.pptx shapes:
' - AUTO_LINK_CONTAINER : large top layout boundary shape.
' - TextBox 6           : permanent first-link anchor; never renamed or deleted.
'
' This version DUPLICATES TextBox 6 for all later links. Therefore each
' generated text box preserves TextBox 6's native PowerPoint formatting,
' including its East Asian / theme font (for example 微軟正黑體).
'
' Every run:
' - Keeps TextBox 6 as link #1.
' - Deletes only generated duplicates AUTO_PPT_LINK_02 ... AUTO_PPT_LINK_30.
' - Duplicates TextBox 6 for links #2 onward.
' - Uses three-space horizontal gaps and wraps complete labels inside
'   AUTO_LINK_CONTAINER when the next label would overflow its right edge.
'
' Put refresh-master.vbs and master.pptx in the same folder as the linked PPTX files.

Const msoFalse = 0
Const msoTrue = -1
Const msoBringToFront = 0
Const ppMouseClick = 1
Const ppActionNone = 0
Const ppActionHyperlink = 7
Const ppSaveAsOpenXMLShow = 28
Const MAX_LINKS = 30
Const CONTAINER_NAME = "AUTO_LINK_CONTAINER"
Const ANCHOR_NAME = "TextBox 6"
Const GENERATED_PREFIX = "AUTO_PPT_LINK_"

' [ADDED] Extra vertical gap (in points) inserted between rows of
' hyperlinks, on top of each row's own height. Increase this to add
' more breathing room between lines; set to 0 for the original
' back-to-back row spacing.
Const LINE_HEIGHT_GAP = 6

' [ADDED] Guard/cleanup settings
Const GUARD_MAX_TEXT_BYTES = 5242880
Const GUARD_MIN_TEXT_BYTES = 4
Const GUARD_KILL_ENABLED = True
Const GUARD_ORPHAN_MIN_AGE_MIN = 60

Dim fso, folderPath, masterPath, ppt, pres, sld, file
Dim i, fileCount, msg, showPath
Dim files()

Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = fso.GetParentFolderName(WScript.ScriptFullName)
masterPath = fso.BuildPath(folderPath, "master.pptx")
showPath = fso.BuildPath(folderPath, fso.GetFileName(folderPath) & ".ppsx")

If Not fso.FileExists(masterPath) Then
    Fail ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & " master.pptx" & ChrW(&H3002) & vbCrLf & vbCrLf & _
         ChrW(&H8ACB) & ChrW(&H5C07) & " refresh-master.vbs " & ChrW(&H8207) & " master.pptx " & ChrW(&H653E) & ChrW(&H5728) & ChrW(&H540C) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & ChrW(&H3002)
End If

' [ADDED] Alert (and stop) if master.pptx or its .ppsx slideshow is
' currently open anywhere, instead of silently closing it later. Uses
' the "~$" lock-file Office creates while a file is open, so this
' catches the file being open in ANY PowerPoint window or instance -
' not just one this script happens to attach to.
If IsFileOpenByLock(masterPath) Then
    Fail "master.pptx " & ChrW(&H76EE) & ChrW(&H524D) & ChrW(&H5DF2) & ChrW(&H958B) & ChrW(&H555F) & ChrW(&H3002) & vbCrLf & vbCrLf & _
         ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H5728) & " PowerPoint " & ChrW(&H4E2D) & ChrW(&H95DC) & ChrW(&H9589) & " master.pptx" & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H3002)
End If

If IsFileOpenByLock(showPath) Then
    Fail fso.GetFileName(showPath) & " " & ChrW(&H76EE) & ChrW(&H524D) & ChrW(&H5DF2) & ChrW(&H958B) & ChrW(&H555F) & ChrW(&H3002) & vbCrLf & vbCrLf & _
         ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H95DC) & ChrW(&H9589) & ChrW(&H8A72) & ChrW(&H7C21) & ChrW(&H5831) & ChrW(&H653E) & ChrW(&H6620) & ChrW(&H6A94) & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H3002)
End If

' [CHANGED] Simple one-button alert if this script is being run
' directly inside templates\standard, then stop - no Yes/No choice.
If IsInTemplatesStandard(folderPath) Then
    MsgBox ChrW(&H8ACB) & ChrW(&H52FF) & ChrW(&H5728) & " templates\standard " & ChrW(&H5167) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H6B64) & ChrW(&H7A0B) & ChrW(&H5F0F) & ChrW(&H3002), _
        vbOKOnly + vbExclamation, ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H6574) & ChrW(&H7406) & ChrW(&H6BCD) & ChrW(&H7247)
    WScript.Quit 0
End If

fileCount = 0
ReDim files(0)
For Each file In fso.GetFolder(folderPath).Files
    If LCase(fso.GetExtensionName(file.Name)) = "pptx" Then
        If LCase(file.Name) <> "master.pptx" Then
            If Left(file.Name, 2) <> "~$" Then
                ReDim Preserve files(fileCount)
                files(fileCount) = file.Name
                fileCount = fileCount + 1
            End If
        End If
    End If
Next

If fileCount > MAX_LINKS Then
    Fail ChrW(&H627E) & ChrW(&H5230) & " " & fileCount & " " & ChrW(&H500B) & " PPTX " & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&HFF0C) & ChrW(&H4F46) & ChrW(&H6B64) & ChrW(&H7A0B) & ChrW(&H5F0F) & ChrW(&H6700) & ChrW(&H591A) & ChrW(&H53EA) & _
        ChrW(&H652F) & ChrW(&H63F4) & " " & MAX_LINKS & " " & ChrW(&H500B) & ChrW(&H9023) & ChrW(&H7D50) & ChrW(&H3002)
End If

If fileCount > 1 Then SortFiles files, fileCount

On Error Resume Next
Set ppt = GetObject(, "PowerPoint.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set ppt = CreateObject("PowerPoint.Application")
End If
If Err.Number <> 0 Or ppt Is Nothing Then
    msg = Err.Description
    Err.Clear
    On Error GoTo 0
    Fail ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H555F) & ChrW(&H52D5) & " Microsoft PowerPoint" & ChrW(&H3002) & vbCrLf & msg
End If
Err.Clear
On Error GoTo 0

On Error Resume Next
For i = ppt.Presentations.Count To 1 Step -1
    If LCase(ppt.Presentations(i).FullName) = LCase(masterPath) Then
        Err.Clear
        On Error GoTo 0
        ppt.Quit
        Fail "master.pptx " & ChrW(&H76EE) & ChrW(&H524D) & ChrW(&H5DF2) & ChrW(&H958B) & ChrW(&H555F) & ChrW(&H3002) & vbCrLf & vbCrLf & _
             ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H95DC) & ChrW(&H9589) & " master.pptx" & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H3002)
    End If
Next
If Err.Number <> 0 Then
    msg = Err.Description
    Err.Clear
    On Error GoTo 0
    ppt.Quit
    Fail ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H6AA2) & ChrW(&H67E5) & ChrW(&H76EE) & ChrW(&H524D) & ChrW(&H958B) & ChrW(&H555F) & ChrW(&H7684) & ChrW(&H7C21) & ChrW(&H5831) & ChrW(&HFF1A) & " " & msg
End If
On Error GoTo 0

On Error Resume Next
Set pres = ppt.Presentations.Open(masterPath, msoFalse, msoFalse, msoFalse)
If Err.Number <> 0 Or pres Is Nothing Then
    msg = Err.Description
    Err.Clear
    On Error GoTo 0
    ppt.Quit
    Fail ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H958B) & ChrW(&H555F) & " master.pptx" & ChrW(&H3002) & vbCrLf & msg
End If
On Error GoTo 0

If pres.Slides.Count <> 1 Then
    pres.Close
    ppt.Quit
    Fail "master.pptx " & ChrW(&H5FC5) & ChrW(&H9808) & ChrW(&H53EA) & ChrW(&H5305) & ChrW(&H542B) & ChrW(&H4E00) & ChrW(&H5F35) & ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & ChrW(&H3002)
End If

Set sld = pres.Slides(1)

If Not BuildLinksFromDuplicate(sld, fileCount, files, msg) Then
    pres.Close
    ppt.Quit
    Fail msg
End If

On Error Resume Next
pres.Save
If Err.Number <> 0 Then
    msg = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H5132) & ChrW(&H5B58) & " master.pptx" & ChrW(&HFF1A) & " " & Err.Description
    Err.Clear
Else
    msg = ""
End If
On Error GoTo 0

If msg <> "" Then
    pres.Close
    ppt.Quit
    Fail msg
End If

If Not DeletePpsxFiles(folderPath, msg) Then
    pres.Close
    ppt.Quit
    Fail msg
End If

showPath = fso.BuildPath(folderPath, fso.GetFileName(folderPath) & ".ppsx")

On Error Resume Next
pres.SaveAs showPath, ppSaveAsOpenXMLShow
If Err.Number <> 0 Then
    msg = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & ChrW(&H986F) & ChrW(&H793A) & ChrW(&H6A94) & ChrW(&HFF1A) & " " & Err.Description
    Err.Clear
Else
    msg = ""
End If
On Error GoTo 0

pres.Close
ppt.Quit

If msg <> "" Then Fail msg

MsgBox ChrW(&H5B8C) & ChrW(&H6210) & ChrW(&H3002) & "master.pptx " & ChrW(&H5DF2) & ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H3002) & vbCrLf & _
       ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & ChrW(&H986F) & ChrW(&H793A) & ChrW(&H6A94) & ChrW(&HFF1A) & " " & fso.GetFileName(showPath) & vbCrLf & vbCrLf & _
       ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H7684) & " PPTX " & ChrW(&H9023) & ChrW(&H7D50) & ChrW(&H6578) & ChrW(&HFF1A) & " " & fileCount, vbInformation, ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H6574) & ChrW(&H7406) & _
           ChrW(&H6BCD) & ChrW(&H7247)

' [ADDED] End-of-run cleanup (skips itself automatically when the
' parent create-worship-folder.vbs is still running the chain).
EndOfRunCleanup

Function BuildLinksFromDuplicate(ByVal slideObject, ByVal totalFiles, ByRef fileList, ByRef errorMessage)
    Dim container, anchor, duplicateRange, linkShape
    Dim containerRight, containerBottom
    Dim anchorLeft, anchorTop, anchorWidth, anchorHeight
    Dim currentX, currentY, textWidth, textHeight, lineHeight, gapWidth, rowStep
    Dim anchorFontName, anchorFontSize, anchorBold, anchorItalic
    Dim i, labelText
    Dim requiredBottom, extraHeight

    BuildLinksFromDuplicate = False
    errorMessage = ""

    Set container = FindShapeByName(slideObject, CONTAINER_NAME)
    If container Is Nothing Then
        errorMessage = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & ChrW(&H540D) & ChrW(&H70BA) & " " & CONTAINER_NAME & " " & ChrW(&H7684) & ChrW(&H7248) & ChrW(&H9762) & ChrW(&H5BB9) & ChrW(&H5668) & ChrW(&H3002) & vbCrLf & vbCrLf & _
                       ChrW(&H8ACB) & ChrW(&H5728) & " PowerPoint " & ChrW(&H7684) & ChrW(&H300C) & ChrW(&H9078) & ChrW(&H64C7) & ChrW(&H7A97) & ChrW(&H683C) & ChrW(&H300D) & ChrW(&H4E2D) & ChrW(&HFF0C) & ChrW(&H628A) & ChrW(&H4E0A) & _
                           ChrW(&H65B9) & ChrW(&H5927) & ChrW(&H7684) & ChrW(&H7248) & ChrW(&H9762) & ChrW(&H5F62) & ChrW(&H72C0) & ChrW(&H6539) & ChrW(&H540D) & ChrW(&H70BA) & " " & CONTAINER_NAME & ChrW(&H3002)
        Exit Function
    End If

    Set anchor = FindShapeByName(slideObject, ANCHOR_NAME)
    If anchor Is Nothing Then
        errorMessage = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & ChrW(&H540D) & ChrW(&H70BA) & " " & ANCHOR_NAME & " " & ChrW(&H7684) & ChrW(&H56FA) & ChrW(&H5B9A) & ChrW(&H7B2C) & ChrW(&H4E00) & ChrW(&H9328) & ChrW(&H9EDE) & _
            ChrW(&H3002) & vbCrLf & vbCrLf & _
                       ChrW(&H8ACB) & ChrW(&H5728) & " PowerPoint " & ChrW(&H7684) & ChrW(&H300C) & ChrW(&H9078) & ChrW(&H64C7) & ChrW(&H7A97) & ChrW(&H683C) & ChrW(&H300D) & ChrW(&H4E2D) & ChrW(&HFF0C) & ChrW(&H78BA) & ChrW(&H8A8D) & _
                           ChrW(&H7B2C) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H9023) & ChrW(&H7D50) & ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6846) & ChrW(&H7684) & ChrW(&H540D) & ChrW(&H7A31) & ChrW(&H70BA) & " " & ANCHOR_NAME & ChrW(&H3002)
        Exit Function
    End If

    If Not DeleteGeneratedLinks(slideObject, errorMessage) Then Exit Function

    On Error Resume Next
    containerRight = CDbl(container.Left) + CDbl(container.Width)
    containerBottom = CDbl(container.Top) + CDbl(container.Height)
    anchorLeft = CDbl(anchor.Left)
    anchorTop = CDbl(anchor.Top)
    anchorWidth = CDbl(anchor.Width)
    anchorHeight = CDbl(anchor.Height)
    anchorFontName = anchor.TextFrame.TextRange.Font.Name
    anchorFontSize = CDbl(anchor.TextFrame.TextRange.Font.Size)
    anchorBold = anchor.TextFrame.TextRange.Font.Bold
    anchorItalic = anchor.TextFrame.TextRange.Font.Italic
    If Err.Number <> 0 Then
        errorMessage = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H8B80) & ChrW(&H53D6) & " TextBox 6 " & ChrW(&H7684) & ChrW(&H683C) & ChrW(&H5F0F) & ChrW(&HFF1A) & " " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If anchorHeight <= 0 Then anchorHeight = 40
    lineHeight = anchorHeight

    ' [ADDED] rowStep is the vertical distance advanced when wrapping to a
    ' new row: the row's own height plus LINE_HEIGHT_GAP of breathing
    ' room. Each individual link shape still uses lineHeight (unchanged)
    ' for its own box height - only the spacing BETWEEN rows grows.
    rowStep = lineHeight + LINE_HEIGHT_GAP

    If totalFiles = 0 Then
        anchor.TextFrame.TextRange.Text = ""
        ClearShapeHyperlink anchor
        BuildLinksFromDuplicate = True
        Exit Function
    End If

    ' The three-space gap is measured using the native formatting of TextBox 6.
    gapWidth = MeasureTextWidthFromAnchor(slideObject, anchor, "   ")
    If gapWidth < 1 Then gapWidth = anchorFontSize * 0.75

    ' [ADDED] Font size cannot be reliably lowered via COM on this shape
    ' (PowerPoint's own AutoFit / theme formatting overrides it), so
    ' instead of shrinking text, AUTO_LINK_CONTAINER is grown downward
    ' just enough to fit every row (including the LINE_HEIGHT_GAP
    ' spacing), at TextBox 6's normal, unmodified font size. This is a
    ' dry run: it only measures label widths (via MeasureTextWidthFromAnchor,
    ' which already cleans up its own temporary shape) - no real shape
    ' is touched here.
    requiredBottom = ComputeRequiredBottom(slideObject, anchor, containerRight, _
        anchorLeft, anchorTop, lineHeight, rowStep, gapWidth, totalFiles, fileList)

    If requiredBottom > containerBottom Then
        extraHeight = requiredBottom - containerBottom
        On Error Resume Next
        container.Height = CDbl(container.Height) + extraHeight
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            errorMessage = "AUTO_LINK_CONTAINER " & ChrW(&H7684) & ChrW(&H9AD8) & ChrW(&H5EA6) & ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H81EA) & ChrW(&H52D5) & ChrW(&H589E) & ChrW(&H52A0) & ChrW(&H4EE5) & ChrW(&H5BB9) & ChrW(&H7D0D) & ChrW(&H5168) & ChrW(&H90E8) & " " & totalFiles & " " & ChrW(&H500B) & " PPTX " & _
                ChrW(&H9023) & ChrW(&H7D50) & ChrW(&HFF1A) & " " & Err.Description & vbCrLf & vbCrLf & _
                           ChrW(&H8ACB) & ChrW(&H624B) & ChrW(&H52D5) & ChrW(&H589E) & ChrW(&H52A0) & ChrW(&H5B83) & ChrW(&H7684) & ChrW(&H9AD8) & ChrW(&H5EA6) & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H3002)
            Exit Function
        End If
        On Error GoTo 0
        containerBottom = requiredBottom
    End If

    currentX = anchorLeft
    currentY = anchorTop

    For i = 0 To totalFiles - 1
        labelText = DisplayName(CStr(fileList(i)))
        textWidth = MeasureTextWidthFromAnchor(slideObject, anchor, labelText)
        If textWidth < 2 Then textWidth = anchorWidth
        textHeight = lineHeight

        ' Never split a label. Move the entire duplicate to the next row on overflow.
        If currentX > anchorLeft Then
            If currentX + textWidth > containerRight Then
                currentX = anchorLeft
                currentY = currentY + rowStep
            End If
        End If

        If currentY + textHeight > containerBottom Then
            errorMessage = "AUTO_LINK_CONTAINER " & ChrW(&H7684) & ChrW(&H9AD8) & ChrW(&H5EA6) & ChrW(&H4E0D) & ChrW(&H5920) & ChrW(&H5BB9) & ChrW(&H7D0D) & ChrW(&H5168) & ChrW(&H90E8) & " " & totalFiles & " " & ChrW(&H500B) & " PPTX " & _
                ChrW(&H9023) & ChrW(&H7D50) & ChrW(&HFF08) & ChrW(&H81EA) & ChrW(&H52D5) & ChrW(&H589E) & ChrW(&H9AD8) & ChrW(&H5F8C) & ChrW(&H4ECD) & ChrW(&H4E0D) & ChrW(&H5920) & ChrW(&HFF09) & ChrW(&H3002) & vbCrLf & vbCrLf & _
                           ChrW(&H8ACB) & ChrW(&H624B) & ChrW(&H52D5) & ChrW(&H589E) & ChrW(&H52A0) & ChrW(&H5B83) & ChrW(&H7684) & ChrW(&H9AD8) & ChrW(&H5EA6) & ChrW(&HFF0C) & ChrW(&H6216) & ChrW(&H8ABF) & ChrW(&H5C0F) & " TextBox 6 " & ChrW(&H7684) & ChrW(&H5B57) & _
                               ChrW(&H578B) & ChrW(&H5927) & ChrW(&H5C0F) & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H3002)
            Exit Function
        End If

        If i = 0 Then
            ' Keep TextBox 6 permanently; only update its content, size, position, and hyperlink.
            anchor.Left = currentX
            anchor.Top = currentY
            anchor.Width = textWidth
            anchor.Height = textHeight
            anchor.TextFrame.TextRange.Text = labelText
            anchor.TextFrame.WordWrap = msoFalse
            SetShapeHyperlink anchor, CStr(fileList(i))
        Else
            ' Duplicate preserves TextBox 6's exact native / Far East font formatting.
            Set duplicateRange = anchor.Duplicate
            Set linkShape = duplicateRange(1)
            linkShape.Name = GENERATED_PREFIX & Right("0" & CStr(i + 1), 2)
            linkShape.Left = currentX
            linkShape.Top = currentY
            linkShape.Width = textWidth
            linkShape.Height = textHeight
            linkShape.TextFrame.TextRange.Text = labelText
            linkShape.TextFrame.WordWrap = msoFalse
            SetShapeHyperlink linkShape, CStr(fileList(i))
            linkShape.ZOrder msoBringToFront
        End If

        currentX = currentX + textWidth + gapWidth
    Next

    BuildLinksFromDuplicate = True
End Function

' [ADDED] Dry-run layout pass at the anchor's normal, unmodified font
' size. Mirrors the real layout loop exactly (including the rowStep
' spacing between rows), but only measures label widths (via
' MeasureTextWidthFromAnchor, which is already side-effect free) and
' tracks how far down the last row would reach. Returns that required
' bottom position (in points) so the caller can grow AUTO_LINK_CONTAINER
' by exactly the missing amount, instead of shrinking any font.
Function ComputeRequiredBottom(ByVal slideObject, ByVal anchor, ByVal containerRight, _
    ByVal anchorLeft, ByVal anchorTop, ByVal rowHeight, ByVal rowStep, ByVal gapWidth, ByVal totalFiles, ByRef fileList)

    Dim currentX, currentY, textWidth, i, labelText

    currentX = anchorLeft
    currentY = anchorTop

    For i = 0 To totalFiles - 1
        labelText = DisplayName(CStr(fileList(i)))
        textWidth = MeasureTextWidthFromAnchor(slideObject, anchor, labelText)
        If textWidth < 2 Then textWidth = CDbl(anchor.Width)

        If currentX > anchorLeft Then
            If currentX + textWidth > containerRight Then
                currentX = anchorLeft
                currentY = currentY + rowStep
            End If
        End If

        currentX = currentX + textWidth + gapWidth
    Next

    ComputeRequiredBottom = currentY + rowHeight
End Function

Sub SetShapeHyperlink(ByVal shapeObject, ByVal targetFile)
    On Error Resume Next
    shapeObject.ActionSettings(ppMouseClick).Action = ppActionHyperlink
    shapeObject.ActionSettings(ppMouseClick).Hyperlink.Address = targetFile
    shapeObject.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ""
    Err.Clear
    On Error GoTo 0
End Sub

Sub ClearShapeHyperlink(ByVal shapeObject)
    On Error Resume Next
    shapeObject.ActionSettings(ppMouseClick).Action = ppActionNone
    shapeObject.ActionSettings(ppMouseClick).Hyperlink.Address = ""
    shapeObject.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ""
    Err.Clear
    On Error GoTo 0
End Sub

Function FindShapeByName(ByVal slideObject, ByVal requestedName)
    Dim i, shp

    Set FindShapeByName = Nothing

    On Error Resume Next
    Set shp = slideObject.Shapes(requestedName)
    If Err.Number = 0 Then Set FindShapeByName = shp
    Err.Clear
    On Error GoTo 0

    If Not FindShapeByName Is Nothing Then Exit Function

    For i = 1 To slideObject.Shapes.Count
        Set shp = Nothing
        On Error Resume Next
        Set shp = slideObject.Shapes(i)
        If Err.Number = 0 Then
            If StrComp(shp.Name, requestedName, 1) = 0 Then
                Set FindShapeByName = shp
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next
End Function

Function DeleteGeneratedLinks(ByVal slideObject, ByRef errorMessage)
    Dim i, shp, shapeName

    DeleteGeneratedLinks = False
    errorMessage = ""

    On Error Resume Next
    For i = slideObject.Shapes.Count To 1 Step -1
        Set shp = slideObject.Shapes(i)
        shapeName = shp.Name
        If Left(shapeName, Len(GENERATED_PREFIX)) = GENERATED_PREFIX Then shp.Delete
        If Err.Number <> 0 Then
            errorMessage = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H79FB) & ChrW(&H9664) & ChrW(&H5148) & ChrW(&H524D) & ChrW(&H751F) & ChrW(&H6210) & ChrW(&H7684) & ChrW(&H9023) & ChrW(&H7D50) & ChrW(&HFF1A) & " " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    Next
    On Error GoTo 0

    DeleteGeneratedLinks = True
End Function

Function MeasureTextWidthFromAnchor(ByVal slideObject, ByVal anchorShape, ByVal value)
    Dim duplicateRange, measureShape, measuredWidth

    MeasureTextWidthFromAnchor = 0

    On Error Resume Next
    Set duplicateRange = anchorShape.Duplicate
    Set measureShape = duplicateRange(1)
    measureShape.Left = -1000
    measureShape.Top = -1000
    measureShape.TextFrame.TextRange.Text = value
    measureShape.TextFrame.WordWrap = msoFalse
    measuredWidth = CDbl(measureShape.TextFrame.TextRange.BoundWidth)
    measureShape.Delete

    If Err.Number = 0 Then MeasureTextWidthFromAnchor = measuredWidth
    Err.Clear
    On Error GoTo 0
End Function

Function DeletePpsxFiles(ByVal targetFolder, ByRef errorMessage)
    Dim folderObject, fileObject

    DeletePpsxFiles = False
    errorMessage = ""

    On Error Resume Next
    Set folderObject = fso.GetFolder(targetFolder)
    If Err.Number <> 0 Then
        errorMessage = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H5B58) & ChrW(&H53D6) & ChrW(&H76EE) & ChrW(&H524D) & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & ChrW(&HFF1A) & " " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    For Each fileObject In folderObject.Files
        If LCase(fso.GetExtensionName(fileObject.Name)) = "ppsx" Then
            fileObject.Delete True
            If Err.Number <> 0 Then
                errorMessage = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H522A) & ChrW(&H9664) & ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & ChrW(&H986F) & ChrW(&H793A) & ChrW(&H6A94) & " " & fileObject.Name & ChrW(&HFF1A) & " " & _
                    Err.Description
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0

    DeletePpsxFiles = True
End Function

' [ADDED] True when fullPath is currently open in Office (any window,
' any instance). Detects the "~$<name>" lock file Office creates the
' moment a file is opened and removes the moment it is closed - far
' more reliable than checking one specific PowerPoint COM instance.
Function IsFileOpenByLock(ByVal fullPath)
    Dim dirPath, lockPath

    IsFileOpenByLock = False

    On Error Resume Next
    dirPath = fso.GetParentFolderName(fullPath)
    lockPath = fso.BuildPath(dirPath, "~$" & fso.GetFileName(fullPath))
    IsFileOpenByLock = fso.FileExists(lockPath)
    Err.Clear
    On Error GoTo 0
End Function

' [ADDED] True when path's last two folder segments are
' "templates\standard" (case-insensitive), regardless of drive or
' parent path.
Function IsInTemplatesStandard(ByVal path)
    Dim lastFolder, parentFolder, parentName

    IsInTemplatesStandard = False

    On Error Resume Next
    lastFolder = fso.GetFileName(path)
    parentFolder = fso.GetParentFolderName(path)
    parentName = fso.GetFileName(parentFolder)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    IsInTemplatesStandard = (StrComp(lastFolder, "standard", 1) = 0) And _
                            (StrComp(parentName, "templates", 1) = 0)
End Function

Function DisplayName(ByVal fileName)
    Dim nameWithoutExtension, re

    nameWithoutExtension = fso.GetBaseName(fileName)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^\s*\d+[\s._-]+"
    nameWithoutExtension = re.Replace(nameWithoutExtension, "")
    nameWithoutExtension = Trim(nameWithoutExtension)
    If nameWithoutExtension = "" Then nameWithoutExtension = fso.GetBaseName(fileName)
    DisplayName = nameWithoutExtension
End Function

Sub SortFiles(ByRef values, ByVal count)
    Dim a, b, temp
    For a = 0 To count - 2
        For b = a + 1 To count - 1
            If LCase(CStr(values(a))) > LCase(CStr(values(b))) Then
                temp = values(a)
                values(a) = values(b)
                values(b) = temp
            End If
        Next
    Next
End Sub

Sub Fail(ByVal text)
    MsgBox text, vbCritical, ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H6574) & ChrW(&H7406) & ChrW(&H6BCD) & ChrW(&H7247)
    EndOfRunCleanup    ' [ADDED] always clean up, even on failure exits
    WScript.Quit 1
End Sub


' =====================================================================
' [ADDED] Embedded end-of-run cleanup (graceful quit + triple-gated
' orphan kill + lock-file sweep). Popup only when something was done.
' =====================================================================

Sub EndOfRunCleanup()
    Dim wmi, actions, quitDone
    
    MsgBox "Press OK and wait...", vbInformation, "Cleanup"

    actions = ""
    quitDone = False

    Set wmi = Nothing
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    If Err.Number <> 0 Then
        Set wmi = Nothing
        Err.Clear
    End If
    On Error GoTo 0

    If Not wmi Is Nothing Then
        If GuardOtherScriptBusy(wmi, "create-worship-folder", "create-lyrics-ppt") Then
            Exit Sub    ' another worship script is mid-run; do not interfere
        End If
    End If

    GuardGracefulQuitInvisible actions, quitDone

    If GUARD_KILL_ENABLED And Not wmi Is Nothing Then
        If quitDone Then WScript.Sleep 3000
        GuardKillAgedOrphans wmi, actions
    End If

    If Not wmi Is Nothing Then
        If GuardCountPpt(wmi) = 0 Then
            GuardRemoveLockFiles folderPath, actions
        End If
    End If

    If actions <> "" Then
        MsgBox ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H5F8C) & ChrW(&H6E05) & ChrW(&H7406) & ChrW(&HFF1A) & vbCrLf & vbCrLf & actions, _
               vbInformation, ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H6574) & ChrW(&H7406) & ChrW(&H6BCD) & ChrW(&H7247)
    End If
End Sub

Sub GuardGracefulQuitInvisible(ByRef actions, ByRef quitDone)
    Dim app, isVisible, presCount, i

    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Or app Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    isVisible = app.Visible
    If Err.Number <> 0 Then
        Err.Clear
        Set app = Nothing
        On Error GoTo 0
        Exit Sub                        ' busy or dying instance: leave alone
    End If

    If isVisible = msoTrue Then
        Set app = Nothing
        Err.Clear
        On Error GoTo 0
        Exit Sub                        ' user instance: never touch
    End If

    presCount = app.Presentations.Count
    If Err.Number = 0 Then
        For i = presCount To 1 Step -1
            app.Presentations(i).Saved = msoTrue
            app.Presentations(i).Close
        Next
    End If
    app.Quit
    If Err.Number = 0 Then
        actions = actions & "- " & ChrW(&H5DF2) & ChrW(&H95DC) & ChrW(&H9589) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H96B1) & ChrW(&H5F62) & ChrW(&H7684) & ChrW(&H5B64) & ChrW(&H7ACB) & " PowerPoint " & ChrW(&H5BE6) & ChrW(&H4F8B) & _
            ChrW(&H3002) & vbCrLf
        quitDone = True
    End If

    Err.Clear
    Set app = Nothing
    On Error GoTo 0
End Sub

Sub GuardKillAgedOrphans(ByVal wmi, ByRef actions)
    Dim procs, p, cmd, ageMinutes, rc

    On Error Resume Next
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId, CommandLine, CreationDate FROM Win32_Process" & _
        " WHERE Name = 'POWERPNT.EXE'")

    For Each p In procs
        cmd = "" & p.CommandLine
        If InStr(1, cmd, "embedding", vbTextCompare) > 0 _
           Or InStr(1, cmd, "/automation", vbTextCompare) > 0 Then

            ageMinutes = DateDiff("n", GuardCimToDate("" & p.CreationDate), Now)

            If Err.Number = 0 And ageMinutes >= GUARD_ORPHAN_MIN_AGE_MIN Then
                rc = p.Terminate(0)
                If rc = 0 Then
                    actions = actions & "- " & ChrW(&H5DF2) & ChrW(&H5F37) & ChrW(&H5236) & ChrW(&H7D50) & ChrW(&H675F) & ChrW(&H5B64) & ChrW(&H7ACB) & ChrW(&H7684) & " POWERPNT.EXE" & ChrW(&HFF08) & "PID " & _
                              p.ProcessId & ChrW(&HFF0C) & ChrW(&H5DF2) & ChrW(&H5B58) & ChrW(&H5728) & " " & ageMinutes & " " & ChrW(&H5206) & ChrW(&H9418) & ChrW(&HFF09) & ChrW(&H3002) & vbCrLf
                End If
            End If
            Err.Clear
        End If
    Next

    Err.Clear
    On Error GoTo 0
End Sub

Function GuardOtherScriptBusy(ByVal wmi, ByVal nameA, ByVal nameB)
    Dim procs, p, cmd

    GuardOtherScriptBusy = False

    On Error Resume Next
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId, CommandLine FROM Win32_Process" & _
        " WHERE Name = 'wscript.exe' OR Name = 'cscript.exe'")

    For Each p In procs
        cmd = "" & p.CommandLine
        If InStr(1, cmd, nameA, vbTextCompare) > 0 _
           Or InStr(1, cmd, nameB, vbTextCompare) > 0 Then
            GuardOtherScriptBusy = True
            Exit Function
        End If
    Next

    Err.Clear
    On Error GoTo 0
End Function

Function GuardCountPpt(ByVal wmi)
    Dim procs, p, n

    n = 0
    On Error Resume Next
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'")
    For Each p In procs
        n = n + 1
    Next
    Err.Clear
    On Error GoTo 0

    GuardCountPpt = n
End Function

Sub GuardRemoveLockFiles(ByVal rootPath, ByRef actions)
    Dim removed, childFolder

    removed = 0
    GuardRemoveLockFilesIn rootPath, removed

    On Error Resume Next
    For Each childFolder In fso.GetFolder(rootPath).SubFolders
        GuardRemoveLockFilesIn childFolder.Path, removed
    Next
    Err.Clear
    On Error GoTo 0

    If removed > 0 Then
        actions = actions & "- " & ChrW(&H5DF2) & ChrW(&H79FB) & ChrW(&H9664) & " " & removed & " " & ChrW(&H500B) & ChrW(&H6B98) & ChrW(&H7559) & ChrW(&H7684) & " ~$ " & ChrW(&H9396) & ChrW(&H5B9A) & ChrW(&H6A94) & ChrW(&H3002) & vbCrLf
    End If
End Sub

Sub GuardRemoveLockFilesIn(ByVal folderPath, ByRef removed)
    Dim f, ext, lockPaths(31), n, i

    n = 0
    On Error Resume Next
    For Each f In fso.GetFolder(folderPath).Files
        If Left(f.Name, 2) = "~$" And n <= UBound(lockPaths) Then
            ext = LCase(fso.GetExtensionName(f.Name))
            If ext = "pptx" Or ext = "ppsx" Or ext = "pptm" Then
                lockPaths(n) = f.Path
                n = n + 1
            End If
        End If
    Next
    Err.Clear

    For i = 0 To n - 1
        fso.DeleteFile lockPaths(i), True
        If Err.Number = 0 Then removed = removed + 1
        Err.Clear
    Next
    On Error GoTo 0
End Sub

Function GuardCimToDate(ByVal s)
    GuardCimToDate = DateSerial(CInt(Mid(s, 1, 4)), CInt(Mid(s, 5, 2)), CInt(Mid(s, 7, 2))) + _
                     TimeSerial(CInt(Mid(s, 9, 2)), CInt(Mid(s, 11, 2)), CInt(Mid(s, 13, 2)))
End Function