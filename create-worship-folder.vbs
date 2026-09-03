Option Explicit

Const sequence_call_response = "02_"
Const sequence_songName = "20_"
Const sequence_scriptures = "30_"
Const sequence_reading = "40_"

Const msoFalse = 0
Const msoTrue = -1

Const ppAutoSizeNone = 0
Const msoAutoSizeNone = 0

Const MAX_CHARS_PER_SLIDE = 110
Const MAX_LINES_PER_SLIDE = 6

Const MAX_SEARCH_YEAR = 2100
Const MIN_SEARCH_YEAR = 2000

Const FILE_ATTRIBUTE_REPARSE_POINT = 1024
Const msoTabStopLeft = 1

Const GUARD_MAX_TEXT_BYTES = 5242880
Const GUARD_MIN_TEXT_BYTES = 4
Const GUARD_KILL_ENABLED = True
Const GUARD_ORPHAN_MIN_AGE_MIN = 60

Dim fso, shell, rootFolder, templatesFolder, dynamicFolder, standardFolder
Dim outputFolder, folderName

Dim masterTemplate, refreshScript
Dim readingTemplate, scripturesTemplate, callResponseTemplate, songFoldersFile
Dim readingTextFile, scripturesTextFile, callResponseTextFile
Dim readingOutputFile, scripturesOutputFile, callResponseOutputFile

Dim songsInput, songNames, songName
Dim songRootPaths(), songRootCount, songFilePath
Dim missingSongs, i

Dim ppt, errText

On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If Err.Number <> 0 Or fso Is Nothing Or shell Is Nothing Then
    MsgBox ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H555F) & ChrW(&H52D5) & " Windows " & ChrW(&H8173) & ChrW(&H672C) & ChrW(&H5143) & ChrW(&H4EF6) & ChrW(&H3002), _
           vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H5D07) & ChrW(&H62DC) & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E)
    WScript.Quit 1
End If

Err.Clear
On Error GoTo 0

rootFolder = fso.GetParentFolderName(WScript.ScriptFullName)
templatesFolder = fso.BuildPath(rootFolder, "templates")
dynamicFolder = fso.BuildPath(templatesFolder, "dynamic")
standardFolder = fso.BuildPath(templatesFolder, "standard")

readingTextFile = txtPrefix() & CNRead() & ".txt"
scripturesTextFile = txtPrefix() & CNScriptures() & ".txt"
callResponseTextFile = txtPrefix() & CNCallResponse() & ".txt"

readingOutputFile = CNRead() & ".pptx"
scripturesOutputFile = CNScriptures() & ".pptx"
callResponseOutputFile = CNCallResponse() & ".pptx"

readingTemplate = fso.BuildPath(dynamicFolder, CNRead() & "-template.pptx")
scripturesTemplate = fso.BuildPath(dynamicFolder, CNScriptures() & "-template.pptx")
callResponseTemplate = fso.BuildPath(dynamicFolder, CNCallResponse() & "-template.pptx")

masterTemplate = fso.BuildPath(standardFolder, "master.pptx")
refreshScript = fso.BuildPath(standardFolder, "refresh-master.vbs")

songFoldersFile = fso.BuildPath(rootFolder, "song-folders.txt")

' [ADDED] Embedded pre-run guard runs before any user input.
PreRunGuard

' ---------------------------------------------------------------
' Validate required structure before asking the user for input.
' ---------------------------------------------------------------

If Not fso.FolderExists(templatesFolder) Then
    Fail MsgtemplatesMissing()
End If

If Not fso.FolderExists(dynamicFolder) Then
    Fail MsgDynamicFolderMissing()
End If

If Not fso.FolderExists(standardFolder) Then
    Fail MsgStandardFolderMissing()
End If

If Not fso.FileExists(masterTemplate) Then
    Fail MsgFileMissing("templates\standard\master.pptx")
End If

If Not fso.FileExists(refreshScript) Then
    Fail MsgFileMissing("templates\standard\refresh-master.vbs")
End If

If Not fso.FileExists(readingTemplate) Then
    Fail MsgFileMissing("templates\dynamic\" & CNRead() & "-template.pptx")
End If

If Not fso.FileExists(scripturesTemplate) Then
    Fail MsgFileMissing("templates\dynamic\" & CNScriptures() & "-template.pptx")
End If

If Not fso.FileExists(callResponseTemplate) Then
    Fail MsgFileMissing("templates\dynamic\" & CNCallResponse() & "-template.pptx")
End If

If Not fso.FileExists(songFoldersFile) Then
    Fail MsgFileMissing("song-folders.txt")
End If

If Not fso.FileExists(fso.BuildPath(rootFolder, readingTextFile)) Then
    Fail MsgFileMissing(readingTextFile)
End If

If Not fso.FileExists(fso.BuildPath(rootFolder, scripturesTextFile)) Then
    Fail MsgFileMissing(scripturesTextFile)
End If

If Not fso.FileExists(fso.BuildPath(rootFolder, callResponseTextFile)) Then
    Fail MsgFileMissing(callResponseTextFile)
End If

' ---------------------------------------------------------------
' Ask for new folder name.
' ---------------------------------------------------------------

folderName = InputBox(MsgEnterFolderName(), MsgAppTitle())
folderName = Trim(folderName)

If folderName = "" Then
    WScript.Quit 0
End If

If Not IsValidFolderName(folderName) Then
    Fail MsgInvalidFolderName()
End If

outputFolder = fso.BuildPath(rootFolder, folderName)

If fso.FolderExists(outputFolder) Then
    Fail MsgFolderExists() & vbCrLf & outputFolder & vbCrLf & vbCrLf & _
         MsgNoChanges()
End If

' ---------------------------------------------------------------
' Ask for requested songs.
' ---------------------------------------------------------------

songsInput = InputBox( _
    MsgEnterSongs() & vbCrLf & MsgSongExample(), _
    MsgAppTitle() _
)

If Len(songsInput) = 0 Then
    WScript.Quit 0
End If

songNames = Split(songsInput, ",")

If Not ReadSongFolders(songFoldersFile, songRootPaths, songRootCount, errText) Then
    Fail MsgCannotReadSongFolders() & vbCrLf & errText
End If

If songRootCount = 0 Then
    Fail MsgNoValidSongFolders() & vbCrLf & songFoldersFile
End If

' ---------------------------------------------------------------
' Start PowerPoint.
' ---------------------------------------------------------------

On Error Resume Next

Set ppt = GetObject(, "PowerPoint.Application")

If Err.Number <> 0 Then
    Err.Clear
    Set ppt = CreateObject("PowerPoint.Application")
End If

If Err.Number <> 0 Or ppt Is Nothing Then
    errText = Err.Description
    Err.Clear
    On Error GoTo 0

    Fail MsgCannotStartPowerPoint() & vbCrLf & errText
End If

On Error GoTo 0

' ---------------------------------------------------------------
' Create output folder and copy direct fixed files.
' ---------------------------------------------------------------

errText = ""

On Error Resume Next

fso.CreateFolder outputFolder

If Err.Number <> 0 Then
    errText = Err.Description
    Err.Clear
Else
    CopyDirectFiles standardFolder, outputFolder, errText
End If

On Error GoTo 0

If errText <> "" Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCopyFixedFiles() & errText
End If

' ---------------------------------------------------------------
' Create call-and-response PPTX file.
'
' Template Selection Pane names required:
'   Slide 1: CALL_SCRIPTURE_1
'   Slide 2: CALL_1 and RESPONSE_1
' ---------------------------------------------------------------

If Not CreateCallResponsePpt( _
    ppt, _
    callResponseTemplate, _
    fso.BuildPath(rootFolder, callResponseTextFile), _
    fso.BuildPath(outputFolder, sequence_call_response & callResponseOutputFile), _
    errText _
) Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCreateFile(sequence_call_response & callResponseOutputFile) & _
         vbCrLf & errText
End If

' ---------------------------------------------------------------
' Create scripture PPTX files.
' ---------------------------------------------------------------

If Not CreateScripturePpt( _
    ppt, _
    readingTemplate, _
    fso.BuildPath(rootFolder, readingTextFile), _
    fso.BuildPath(outputFolder, sequence_reading & readingOutputFile), _
    errText _
) Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCreateFile(sequence_reading & readingOutputFile) & _
         vbCrLf & errText
End If

If Not CreateScripturePpt( _
    ppt, _
    scripturesTemplate, _
    fso.BuildPath(rootFolder, scripturesTextFile), _
    fso.BuildPath(outputFolder, sequence_scriptures & scripturesOutputFile), _
    errText _
) Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCreateFile(sequence_scriptures & scripturesOutputFile) & _
         vbCrLf & errText
End If

' ---------------------------------------------------------------
' Copy song PPTX files.
' Output naming e.g.:
'   20_<song title>.pptx
'
' ---------------------------------------------------------------

missingSongs = ""

For i = 0 To UBound(songNames)
    songName = Trim(songNames(i))

    If songName <> "" Then
        songFilePath = FindSongPptInRoots(songRootPaths, songRootCount, songName)

        If songFilePath <> "" Then
            errText = ""

            On Error Resume Next

            fso.CopyFile _
                songFilePath, _
                fso.BuildPath(outputFolder, sequence_songName & songName & ".pptx"), _
                True

            If Err.Number <> 0 Then
                errText = Err.Description
                Err.Clear
            End If

            On Error GoTo 0

            If errText <> "" Then
                CleanupFolder outputFolder
                SafeQuitPowerPoint ppt

                Fail MsgCannotCopySong() & errText
            End If

        Else
            If Not CreateMissingSongMarker( _
                outputFolder, songName, errText _
            ) Then
                CleanupFolder outputFolder
                SafeQuitPowerPoint ppt

                Fail MsgCannotCreateMarker() & errText
            End If

            If missingSongs <> "" Then
                missingSongs = missingSongs & vbCrLf
            End If

            missingSongs = missingSongs & songName
        End If
    End If
Next

RemoveTxtFilesWithMatchingPptx outputFolder
SafeQuitPowerPoint ppt
Set ppt = Nothing

'================================================================
' Refresh master after all PPTX files are in the output folder.
'================================================================

errText = ""

On Error Resume Next

shell.Run _
    "wscript.exe """ & _
    fso.BuildPath(outputFolder, "refresh-master.vbs") & _
    """", _
    0, _
    True

If Err.Number <> 0 Then
    errText = Err.Description
    Err.Clear
End If

On Error GoTo 0

If errText <> "" Then
    MsgBox MsgCreatedRefreshFailed() & vbCrLf & _
           errText & vbCrLf & vbCrLf & _
           MsgRunRefreshManually(), _
           vbExclamation, MsgAppTitle()

ElseIf missingSongs <> "" Then
    MsgBox MsgDone() & vbCrLf & vbCrLf & _
           MsgCreatedFolder() & vbCrLf & _
           outputFolder & vbCrLf & vbCrLf & _
           MsgMissingSongMarkers() & vbCrLf & _
           missingSongs, _
           vbExclamation, MsgAppTitle()

Else
    MsgBox MsgDone() & vbCrLf & vbCrLf & _
           MsgCreatedFolder() & vbCrLf & _
           outputFolder, _
           vbInformation, MsgAppTitle()
End If

' [ADDED] End-of-run cleanup after the normal finish.
EndOfRunCleanup

Function ReadSongFolders(ByVal filePath, ByRef folderPaths, ByRef folderCount, ByRef errorMessage)

    Dim stream
    Dim text
    Dim rawLines
    Dim i
    Dim candidatePath

    ReadSongFolders = False
    errorMessage = ""
    folderCount = 0
    ReDim folderPaths(0)

    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    text = Replace(text, ChrW(&HFEFF), "")
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    rawLines = Split(text, vbLf)

    For i = 0 To UBound(rawLines)
        candidatePath = Trim(rawLines(i))

        If candidatePath <> "" Then
            If fso.FolderExists(candidatePath) Then
                ReDim Preserve folderPaths(folderCount)
                folderPaths(folderCount) = candidatePath
                folderCount = folderCount + 1
            End If
        End If
    Next

    ReadSongFolders = True
End Function

Function FindSongPptInRoots(ByRef rootPaths, ByVal rootCount, ByVal songName)

    Dim i
    Dim foundPath

    FindSongPptInRoots = ""

    For i = 0 To rootCount - 1
        foundPath = FindSongPpt(CStr(rootPaths(i)), songName)

        If foundPath <> "" Then
            FindSongPptInRoots = foundPath
            Exit Function
        End If
    Next
End Function

' ===============================================================
' Song search
' ===============================================================
' Search order for each path listed in song-folders.txt:
' 1. Direct .pptx files inside the listed path itself.
' 2. Only when nothing matches there, every subfolder of that root
'    is enumerated in the exact order Windows returns them
'    (no re-sorting, no year-folder priority).
' 3. Each subfolder is searched direct .pptx files first; only when
'    nothing matches does the search recurse one level deeper,
'    repeating the same files-first-then-subfolders rule.
' 4. The first .pptx whose file name contains the requested song
'    name wins and the search stops immediately.

Function FindSongPpt(ByVal worshipDataRoot, ByVal songName)

    Dim rootFolder
    Dim childFolder
    Dim foundPath

    FindSongPpt = ""

    ' Step 1: direct .pptx files inside the listed path itself.
    foundPath = FindMatchingPptxDirect( _
        worshipDataRoot, _
        songName _
    )

    If foundPath <> "" Then
        FindSongPpt = foundPath
        Exit Function
    End If

    ' Steps 2-5: enumerate every subfolder in the order Windows
    ' returns them; FindMatchingPptxRecursive searches each one
    ' direct files first, then recurses downward with the same rule.
    On Error Resume Next
    Set rootFolder = fso.GetFolder(worshipDataRoot)

    If Err.Number <> 0 Or rootFolder Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    For Each childFolder In rootFolder.SubFolders

        foundPath = FindMatchingPptxRecursive( _
            childFolder.Path, _
            songName _
        )

        If foundPath <> "" Then
            FindSongPpt = foundPath
            Exit Function
        End If
    Next
End Function

' Finds the first direct .pptx file whose base name contains songName.
Function FindMatchingPptxDirect(ByVal folderPath, ByVal songName)

    Dim folderObject
    Dim fileObject

    FindMatchingPptxDirect = ""

    On Error Resume Next
    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Or folderObject Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    For Each fileObject In folderObject.Files
        If IsMatchingSongPptx(fileObject.Name, songName) Then
            FindMatchingPptxDirect = fileObject.Path
            Exit Function
        End If
    Next
End Function

' Recursively finds the first .pptx file whose base name contains songName.
' Skips junctions, symbolic links, and other reparse-point folders.
Function FindMatchingPptxRecursive(ByVal folderPath, ByVal songName)

    Dim folderObject
    Dim fileObject
    Dim subFolderObject
    Dim foundPath
    Dim attributes

    FindMatchingPptxRecursive = ""

    On Error Resume Next
    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Or folderObject Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    For Each fileObject In folderObject.Files
        If IsMatchingSongPptx(fileObject.Name, songName) Then
            FindMatchingPptxRecursive = fileObject.Path
            Exit Function
        End If
    Next

    For Each subFolderObject In folderObject.SubFolders

        On Error Resume Next
        attributes = subFolderObject.Attributes

        If Err.Number = 0 Then
            If (attributes And FILE_ATTRIBUTE_REPARSE_POINT) = 0 Then

                On Error GoTo 0

                foundPath = FindMatchingPptxRecursive( _
                    subFolderObject.Path, _
                    songName _
                )

                If foundPath <> "" Then
                    FindMatchingPptxRecursive = foundPath
                    Exit Function
                End If

            Else
                Err.Clear
                On Error GoTo 0
            End If
        Else
            Err.Clear
            On Error GoTo 0
        End If
    Next
End Function

' True only when fileName is a PPTX and its name without extension
' contains songName anywhere, case-insensitively.
Function IsMatchingSongPptx(ByVal fileName, ByVal songName)

    Dim baseName

    IsMatchingSongPptx = False

    If LCase(fso.GetExtensionName(fileName)) <> "pptx" Then
        Exit Function
    End If

    baseName = fso.GetBaseName(fileName)

    If InStr(1, baseName, songName, vbTextCompare) > 0 Then
        IsMatchingSongPptx = True
    End If
End Function

' Returns True only if the complete folder name consists of decimal digits.
' Examples:
'   "2026"  -> True
'   "0026"  -> True
'   "2026a" -> False
'   "2026 " -> False
Function IsIntegerFolderName(ByVal folderName)

    Dim re

    IsIntegerFolderName = False

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = "^\d+$"

    IsIntegerFolderName = re.Test(CStr(folderName))
End Function

' Recursively searches only PPTX files for an exact case-insensitive filename.
' It skips junctions, symbolic links, and other reparse-point folders.
Function FindFileRecursivePptxOnly(ByVal folderPath, ByVal requestedFileName)

    Dim folderObject
    Dim fileObject
    Dim subFolderObject
    Dim foundPath
    Dim attributes

    FindFileRecursivePptxOnly = ""

    On Error Resume Next

    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Or folderObject Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    ' Check PPTX files directly inside this folder first.
    For Each fileObject In folderObject.Files

        If LCase(fso.GetExtensionName(fileObject.Name)) = "pptx" Then
            If StrComp(fileObject.Name, requestedFileName, 1) = 0 Then
                FindFileRecursivePptxOnly = fileObject.Path
                Exit Function
            End If
        End If
    Next

    ' Then recursively search ordinary child folders only.
    For Each subFolderObject In folderObject.SubFolders

        On Error Resume Next

        attributes = subFolderObject.Attributes

        If Err.Number = 0 Then

            If (attributes And FILE_ATTRIBUTE_REPARSE_POINT) = 0 Then

                On Error GoTo 0

                foundPath = FindFileRecursivePptxOnly( _
                    subFolderObject.Path, _
                    requestedFileName _
                )

                If foundPath <> "" Then
                    FindFileRecursivePptxOnly = foundPath
                    Exit Function
                End If

            Else
                Err.Clear
                On Error GoTo 0
            End If

        Else
            Err.Clear
            On Error GoTo 0
        End If
    Next
End Function

Function FindFileRecursive(ByVal folderPath, ByVal requestedFileName)

    Dim folderObject
    Dim fileObject
    Dim subFolderObject
    Dim foundPath
    Dim attributes

    FindFileRecursive = ""

    On Error Resume Next

    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    ' Check direct files in this folder first.
    For Each fileObject In folderObject.Files
        If StrComp(fileObject.Name, requestedFileName, 1) = 0 Then
            FindFileRecursive = fileObject.Path
            Exit Function
        End If
    Next

    ' Search child folders.
    ' Skip reparse points, junctions, and symbolic links.
    For Each subFolderObject In folderObject.SubFolders

        On Error Resume Next
        attributes = subFolderObject.Attributes

        If Err.Number = 0 Then
            If (attributes And FILE_ATTRIBUTE_REPARSE_POINT) = 0 Then
                On Error GoTo 0

                foundPath = FindFileRecursive( _
                    subFolderObject.Path, _
                    requestedFileName _
                )

                If foundPath <> "" Then
                    FindFileRecursive = foundPath
                    Exit Function
                End If

            Else
                Err.Clear
                On Error GoTo 0
            End If

        Else
            Err.Clear
            On Error GoTo 0
        End If
    Next
End Function

' ===============================================================
' Copy fixed files
' ===============================================================

Sub CopyDirectFiles(ByVal sourceFolder, ByVal destinationFolder, ByRef errorMessage)

    Dim fileObject

    errorMessage = ""

    On Error Resume Next

    For Each fileObject In fso.GetFolder(sourceFolder).Files
        fso.CopyFile _
            fileObject.Path, _
            fso.BuildPath(destinationFolder, fileObject.Name), _
            True

        If Err.Number <> 0 Then
            errorMessage = Err.Description
            Err.Clear
            Exit For
        End If
    Next

    On Error GoTo 0
End Sub

' ===============================================================
' Scripture PPTX generation
' ===============================================================

Function CreateScripturePpt( _
    ByVal pptApp, _
    ByVal templatePath, _
    ByVal textPath, _
    ByVal destinationPath, _
    ByRef errorMessage _
)

    Dim title
    Dim lines, lineCount
    Dim pageStart(), pageEnd(), pageCount
    Dim pres, pageIndex, slideObject
    Dim duplicateRange, pageText
    Dim indentPoints

    CreateScripturePpt = False
    errorMessage = ""

    If Not ReadUtf8ScriptureFile( _
        textPath, title, lines, lineCount, errorMessage _
    ) Then
        Exit Function
    End If

    If Not MakePageBreaks( _
        lines, lineCount, _
        pageStart, pageEnd, pageCount, _
        errorMessage _
    ) Then
        Exit Function
    End If

    On Error Resume Next

    fso.CopyFile templatePath, destinationPath, True

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set pres = pptApp.Presentations.Open( _
        destinationPath, _
        msoFalse, _
        msoFalse, _
        msoFalse _
    )

    If Err.Number <> 0 Or pres Is Nothing Then
        errorMessage = MsgCannotOpenTemplateCopy() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    If pres.Slides.Count <> 1 Then
        pres.Close
        errorMessage = MsgTemplateSlideCount()
        Exit Function
    End If

    For pageIndex = 2 To pageCount
        Set duplicateRange = pres.Slides(1).Duplicate
        duplicateRange(1).MoveTo pres.Slides.Count
    Next

    For pageIndex = 1 To pageCount
        Set slideObject = pres.Slides(pageIndex)

        pageText = JoinLines( _
            lines, _
            pageStart(pageIndex), _
            pageEnd(pageIndex) _
        )

        ' Replace every exact {{SCRIPTURE_TITLE}} placeholder.
        ' Title lines get no hanging indent.
        If ReplaceAllExactText( _
            slideObject, _
            "{{SCRIPTURE_TITLE}}", _
            title, _
            0 _
        ) = 0 Then
            pres.Close
            errorMessage = MsgTemplateMissingTitle()
            Exit Function
        End If

        ' Replace every exact {{SCRIPTURE_TEXT}} placeholder and apply
        ' the dynamic hanging indent that matches the widest verse
        ' number on this slide.
        indentPoints = IndentPointsForVerseDigits( _
            MaxVerseDigits( _
                lines, _
                pageStart(pageIndex), _
                pageEnd(pageIndex) _
            ) _
        )

        If ReplaceAllExactText( _
            slideObject, _
            "{{SCRIPTURE_TEXT}}", _
            pageText, _
            indentPoints _
        ) = 0 Then
            pres.Close
            errorMessage = MsgTemplateMissingText()
            Exit Function
        End If
    Next

    On Error Resume Next

    pres.Save

    If Err.Number <> 0 Then
        errorMessage = MsgCannotSavePpt() & Err.Description
        Err.Clear
        pres.Close
        On Error GoTo 0
        Exit Function
    End If

    pres.Close

    On Error GoTo 0

    CreateScripturePpt = True
End Function

Function ReplaceAllExactText( _
    ByVal slideObject, _
    ByVal placeholder, _
    ByVal replacement, _
    ByVal hangingIndentPoints _
)

    Dim i
    Dim shapeObject
    Dim replacementCount
    Dim isScriptureText
    Dim paragraphFormat2
    Dim tabIndex, tabStopCount

    replacementCount = 0
    isScriptureText = (placeholder = "{{SCRIPTURE_TEXT}}")

    For i = 1 To slideObject.Shapes.Count

        On Error Resume Next

        Set shapeObject = slideObject.Shapes(i)

        If shapeObject.HasTextFrame = msoTrue Then

            If Trim(shapeObject.TextFrame.TextRange.Text) = placeholder Then

                ' For scripture text, stop PowerPoint from shrinking
                ' the text box to the original placeholder height.
                If isScriptureText Then
                    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
                    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
                    shapeObject.TextFrame.VerticalAnchor = 1
                End If

                shapeObject.TextFrame.TextRange.Text = replacement

                ' PowerPoint may reset AutoFit after assigning text.
                If isScriptureText Then
                    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
                    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
                    shapeObject.TextFrame.VerticalAnchor = 1

                    ' Apply the dynamic hanging indent (Before text =
                    ' Hanging = hangingIndentPoints, in points). A left
                    ' tab stop is placed at the same position, and the
                    ' verse text is assembled as "[n]<tab>text", so
                    ' every verse's text starts at the same point.
                    If hangingIndentPoints > 0 Then
                        Set paragraphFormat2 = _
                            shapeObject.TextFrame2.TextRange.ParagraphFormat

                        paragraphFormat2.LeftIndent = hangingIndentPoints
                        paragraphFormat2.FirstLineIndent = -hangingIndentPoints

                        ' Keep default tab stops from landing before the
                        ' indent position.
                        paragraphFormat2.TabStops.DefaultSpacing = hangingIndentPoints

                        tabStopCount = paragraphFormat2.TabStops.Count

                        For tabIndex = 1 To tabStopCount
                            paragraphFormat2.TabStops.Item(1).Clear
                        Next

                        paragraphFormat2.TabStops.Add msoTabStopLeft, hangingIndentPoints
                    End If
                End If

                replacementCount = replacementCount + 1
            End If
        End If

        Err.Clear
        On Error GoTo 0
    Next

    ReplaceAllExactText = replacementCount
End Function

Function MakePageBreaks( _
    ByRef lines, _
    ByVal lineCount, _
    ByRef pageStart, _
    ByRef pageEnd, _
    ByRef pageCount, _
    ByRef errorMessage _
)

    Dim currentLine
    Dim pageFirst
    Dim pageChars
    Dim lineLength

    MakePageBreaks = False
    errorMessage = ""
    pageCount = 0

    ReDim pageStart(1)
    ReDim pageEnd(1)

    If lineCount = 0 Then
        pageCount = 1
        pageStart(1) = 0
        pageEnd(1) = -1

        MakePageBreaks = True
        Exit Function
    End If

    currentLine = 0

    Do While currentLine < lineCount
        pageFirst = currentLine
        pageChars = 0

        Do While currentLine < lineCount
            lineLength = Len(lines(currentLine))

            If currentLine > pageFirst Then
                lineLength = lineLength + 1
            End If

            If currentLine > pageFirst Then
                If pageChars + lineLength > MAX_CHARS_PER_SLIDE Or _
                   currentLine - pageFirst >= MAX_LINES_PER_SLIDE Then
                    Exit Do
                End If
            End If

            pageChars = pageChars + lineLength
            currentLine = currentLine + 1
        Loop

        ' Protect against a single unusually long line.
        If currentLine = pageFirst Then
            currentLine = currentLine + 1
        End If

        pageCount = pageCount + 1

        ReDim Preserve pageStart(pageCount)
        ReDim Preserve pageEnd(pageCount)

        pageStart(pageCount) = pageFirst
        pageEnd(pageCount) = currentLine - 1
    Loop

    MakePageBreaks = True
End Function

' Wraps a leading verse number in square brackets:
' Lines that already start with "[" are returned unchanged.
Function WrapVerseNumber(ByVal lineValue)

    Dim trimmedLine
    Dim i
    Dim ch
    Dim digitRun

    trimmedLine = Trim(lineValue)
    WrapVerseNumber = trimmedLine

    If Left(trimmedLine, 1) = "[" Then
        Exit Function
    End If

    ' Collect the leading digits.
    digitRun = ""

    For i = 1 To Len(trimmedLine)
        ch = Mid(trimmedLine, i, 1)

        If ch >= "0" And ch <= "9" Then
            digitRun = digitRun & ch
        Else
            Exit For
        End If
    Next

    If digitRun <> "" Then
        WrapVerseNumber = "[" & digitRun & "]" & _
                          Mid(trimmedLine, Len(digitRun) + 1)
    End If
End Function

Function ReadUtf8ScriptureFile( _
    ByVal filePath, _
    ByRef title, _
    ByRef bodyLines, _
    ByRef lineCount, _
    ByRef errorMessage _
)

    Dim stream
    Dim text
    Dim rawLines
    Dim i
    Dim lineValue

    ReadUtf8ScriptureFile = False

    errorMessage = ""
    title = ""
    lineCount = 0

    ReDim bodyLines(0)

    On Error Resume Next

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        errorMessage = MsgCannotReadText() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    text = Replace(text, ChrW(&HFEFF), "")
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    rawLines = Split(text, vbLf)

    title = Trim(rawLines(0))

    If title = "" Then
        errorMessage = MsgEmptyTitle()
        Exit Function
    End If

    ReDim bodyLines(0)

    For i = 1 To UBound(rawLines)
        lineValue = rawLines(i)

        If Trim(lineValue) <> "" Then
            ReDim Preserve bodyLines(lineCount)
            bodyLines(lineCount) = WrapVerseNumber(lineValue)
            lineCount = lineCount + 1
        End If
    Next

    ReadUtf8ScriptureFile = True
End Function

Function JoinLines(ByRef values, ByVal firstIndex, ByVal lastIndex)

    Dim i
    Dim result
    Dim currentLine
    Dim closingBracketPosition
    Dim afterBracket
    Dim verseNumberText
    Dim maxDigits
    Dim padding

    result = ""

    If lastIndex < firstIndex Then
        JoinLines = ""
        Exit Function
    End If

    maxDigits = MaxVerseDigits(values, firstIndex, lastIndex)

    For i = firstIndex To lastIndex

        currentLine = values(i)

        ' Expected scripture line format:
        ' [1] text
        ' [16] text
        '
        ' Pad the verse-number block with spaces so every verse number
        ' is as wide as the widest one on the slide. The verse text
        ' then starts at the same column on every line.
        If Left(currentLine, 1) = "[" Then

            closingBracketPosition = InStr(currentLine, "]")

            If closingBracketPosition > 2 Then

                verseNumberText = Mid( _
                    currentLine, _
                    2, _
                    closingBracketPosition - 2 _
                )

                afterBracket = Mid( _
                    currentLine, _
                    closingBracketPosition + 1 _
                )

                ' Strip ALL leading spaces and tabs after "]", so the
                ' text joins the verse number with clean padding only.
                Do While Len(afterBracket) > 0 And _
                         (Left(afterBracket, 1) = " " Or _
                          Left(afterBracket, 1) = vbTab)

                    afterBracket = Mid(afterBracket, 2)
                Loop

                ' One separator space, plus enough extra spaces so the
                ' number block matches the widest number on the slide:
                '   max 2 digits -> [9]   [16]
                '   max 3 digits -> [9]    [99]   [100]
                padding = Space( _
                    maxDigits - Len(verseNumberText) + 1 _
                )

                currentLine = Left( _
                    currentLine, _
                    closingBracketPosition _
                ) & padding & afterBracket
            End If
        End If

        If result <> "" Then
            result = result & vbCrLf
        End If

        result = result & currentLine
    Next

    JoinLines = result
End Function

' ===============================================================
' Dynamic scripture indent
' ===============================================================

' Returns the number of digits of the widest verse number on a slide.
' Verse lines are expected to start with "[<number>]".
Function MaxVerseDigits(ByRef values, ByVal firstIndex, ByVal lastIndex)

    Dim i
    Dim currentLine
    Dim closingBracketPosition
    Dim verseNumberText
    Dim digits

    MaxVerseDigits = 1

    For i = firstIndex To lastIndex

        currentLine = Trim(values(i))

        If Left(currentLine, 1) = "[" Then

            closingBracketPosition = InStr(currentLine, "]")

            If closingBracketPosition > 2 Then

                verseNumberText = Trim(Mid( _
                    currentLine, _
                    2, _
                    closingBracketPosition - 2 _
                ))

                If IsNumeric(verseNumberText) Then

                    digits = Len(verseNumberText)

                    If digits > MaxVerseDigits Then
                        MaxVerseDigits = digits
                    End If
                End If
            End If
        End If
    Next
End Function

' Maps the widest verse number's digit count to the hanging-indent
' width (in points). Both "Before text" and "Hanging by" are set to
' this value, and a left tab stop is placed at the same position.
'
'   1 digit   -> 2.82 cm
'   2 digits  -> 3.53 cm
'   3 digits  -> 4.23 cm
Function IndentPointsForVerseDigits(ByVal digits)

    Const CM_TO_POINTS = 28.3464567

    Dim indentCm

    Select Case digits
        Case 1
            indentCm = 2.82
        Case 2
            indentCm = 3.53
        Case Else
            indentCm = 4.23
    End Select

    IndentPointsForVerseDigits = indentCm * CM_TO_POINTS
End Function

' ===============================================================
' Call-and-response PPTX generation
' ===============================================================

' ===============================================================
' PART 1
' Replace these two existing functions:
'   CreateCallResponsePpt
'   SetTextInNamedShape
' ===============================================================

Function CreateCallResponsePpt( _
    ByVal pptApp, _
    ByVal templatePath, _
    ByVal textPath, _
    ByVal destinationPath, _
    ByRef errorMessage _
)

    Dim callScripture
    Dim callLines(), responseLines(), pairCount
    Dim hasCallScripture, hasCallResponse
    Dim pres, pairIndex, slideObject
    Dim duplicateRange
    Dim detailError
    Dim callResponseText

    CreateCallResponsePpt = False
    errorMessage = ""

    If Not ReadCallResponseFile( _
        textPath, _
        callScripture, _
        hasCallScripture, _
        callLines, _
        responseLines, _
        pairCount, _
        hasCallResponse, _
        errorMessage _
    ) Then
        Exit Function
    End If

    On Error Resume Next

    fso.CopyFile templatePath, destinationPath, True

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set pres = pptApp.Presentations.Open( _
        destinationPath, _
        msoFalse, _
        msoFalse, _
        msoFalse _
    )

    If Err.Number <> 0 Or pres Is Nothing Then
        errorMessage = MsgCannotOpenTemplateCopy() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    If pres.Slides.Count <> 2 Then
        pres.Close
        errorMessage = MsgCallTemplateSlideCount()
        Exit Function
    End If

    If hasCallResponse Then
        For pairIndex = 2 To pairCount
            Set duplicateRange = pres.Slides(2).Duplicate
            duplicateRange(1).MoveTo pres.Slides.Count
        Next
    End If

    If hasCallScripture Then
        detailError = ""

        If Not SetTextInNamedShape( _
            pres.Slides(1), _
            "CALL_SCRIPTURE_1", _
            callScripture, _
            detailError _
        ) Then
            pres.Close
            errorMessage = MsgTemplateMissingNamedShape( _
                "CALL_SCRIPTURE_1" _
            ) & vbCrLf & detailError
            Exit Function
        End If
    End If

    If hasCallResponse Then
        For pairIndex = 1 To pairCount

            Set slideObject = pres.Slides(pairIndex + 1)

            callResponseText = _
                CNCallPrefix() & callLines(pairIndex) & vbCrLf & vbCrLf & _
                CNResponsePrefix() & responseLines(pairIndex)

            detailError = ""

            If Not SetCallResponseText( _
                slideObject, _
                "CALL_RESPONSE_1", _
                callResponseText, _
                Len(CNCallPrefix() & callLines(pairIndex) & vbCrLf), _
                detailError _
            ) Then
                pres.Close
                errorMessage = MsgTemplateMissingNamedShape( _
                    "CALL_RESPONSE_1" _
                ) & vbCrLf & detailError
                Exit Function
            End If
        Next
    End If

    ' Remove unused template slides.
    If Not hasCallResponse Then
        pres.Slides(2).Delete
    End If

    If Not hasCallScripture Then
        pres.Slides(1).Delete
    End If

    On Error Resume Next

    pres.Save

    If Err.Number <> 0 Then
        errorMessage = MsgCannotSavePpt() & Err.Description
        Err.Clear
        pres.Close
        On Error GoTo 0
        Exit Function
    End If

    pres.Close

    On Error GoTo 0

    CreateCallResponsePpt = True
End Function

' Gets a text box by its exact PowerPoint Selection Pane name and
' replaces all its text. It does not search for placeholder strings.
Function SetTextInNamedShape( _
    ByVal slideObject, _
    ByVal shapeName, _
    ByVal replacement, _
    ByRef errorMessage _
)

    Dim shapeObject

    SetTextInNamedShape = False
    errorMessage = ""

    On Error Resume Next

    Set shapeObject = slideObject.Shapes(shapeName)

    If Err.Number <> 0 Or shapeObject Is Nothing Then
        errorMessage = MsgNamedShapeNotFound(shapeName)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If shapeObject.HasTextFrame <> msoTrue Then
        errorMessage = MsgNamedShapeNoText(shapeName)
        On Error GoTo 0
        Exit Function
    End If

    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
    shapeObject.TextFrame.VerticalAnchor = 1

    shapeObject.TextFrame.TextRange.Text = replacement

    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
    shapeObject.TextFrame.VerticalAnchor = 1

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    SetTextInNamedShape = True
End Function

Function SetCallResponseText( _
    ByVal slideObject, _
    ByVal shapeName, _
    ByVal replacement, _
    ByVal responseStartPosition, _
    ByRef errorMessage _
)

    Const PURPLE_RED = 112
    Const PURPLE_GREEN = 48
    Const PURPLE_BLUE = 160

    Dim shapeObject
    Dim responseLength

    SetCallResponseText = False
    errorMessage = ""

    On Error Resume Next

    Set shapeObject = slideObject.Shapes(shapeName)

    If Err.Number <> 0 Or shapeObject Is Nothing Then
        errorMessage = MsgNamedShapeNotFound(shapeName)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If shapeObject.HasTextFrame <> msoTrue Then
        errorMessage = MsgNamedShapeNoText(shapeName)
        On Error GoTo 0
        Exit Function
    End If

    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
    shapeObject.TextFrame.VerticalAnchor = 1

    shapeObject.TextFrame.TextRange.Text = replacement

    ' responseStartPosition includes the line break, therefore the
    ' response line starts at this character position.
    responseLength = Len(replacement) - responseStartPosition + 1

    shapeObject.TextFrame.TextRange.Characters( _
        responseStartPosition, _
        responseLength _
    ).Font.Color.RGB = RGB( _
        PURPLE_RED, _
        PURPLE_GREEN, _
        PURPLE_BLUE _
    )

    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
    shapeObject.TextFrame.VerticalAnchor = 1

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    SetCallResponseText = True
End Function

' ===============================================================
' PART 2
' Replace the existing ReadCallResponseFile function with this.
'
' Also add the four MsgCall... functions at the bottom of this part
' if they do not already exist in your VBS.
' ===============================================================

' ===============================================================
' Replace ReadCallResponseFile with this version
' ===============================================================

Function ReadCallResponseFile( _
    ByVal filePath, _
    ByRef callScripture, _
    ByRef hasCallScripture, _
    ByRef callLines, _
    ByRef responseLines, _
    ByRef pairCount, _
    ByRef hasCallResponse, _
    ByRef errorMessage _
)

    Dim stream
    Dim text
    Dim rawLines
    Dim i
    Dim rawLine
    Dim trimmedLine
    Dim section
    Dim sawCallSection, sawResponseSection
    Dim pendingCall, hasPendingCall
    Dim pendingBlankLine
    Dim pendingBlankLine2

    ReadCallResponseFile = False

    errorMessage = ""
    callScripture = ""
    hasCallScripture = False
    pairCount = 0
    hasCallResponse = False

    section = 0
    sawCallSection = False
    sawResponseSection = False
    pendingCall = ""
    hasPendingCall = False
    pendingBlankLine = False
    pendingBlankLine2 = False

    ReDim callLines(0)
    ReDim responseLines(0)

    On Error Resume Next

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        errorMessage = MsgCannotReadText() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    text = Replace(text, ChrW(&HFEFF), "")
    text = Replace(text, vbTab, "")
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    rawLines = Split(text, vbLf)

    ' section:
    ' 0 = before / outside sections
    For i = 0 To UBound(rawLines)

        rawLine = rawLines(i)
        trimmedLine = Trim(rawLine)

        If trimmedLine = CNCallSection() Then

            If sawCallSection Then
                errorMessage = MsgCallSectionRepeated(CNCallSection())
                Exit Function
            End If

            section = 1
            sawCallSection = True
            pendingBlankLine = False

        ElseIf trimmedLine = CNResponseSection() Then

            If sawResponseSection Then
                errorMessage = MsgCallSectionRepeated(CNResponseSection())
                Exit Function
            End If

            section = 2
            sawResponseSection = True
            pendingBlankLine2 = False

        ElseIf section = 1 Then

            ' A blank source line is remembered, then converted into
            ' TWO line breaks before the next non-empty line. This
            ' produces one visible empty paragraph in PowerPoint.
            If trimmedLine = "" Then

                If callScripture <> "" Then
                    pendingBlankLine = True
                End If

            Else

                If callScripture <> "" Then

                    If pendingBlankLine Then
                        ' One normal paragraph break + one empty paragraph.
                        callScripture = callScripture & _
                                        vbCrLf & vbCrLf
                    Else
                        callScripture = callScripture & vbCrLf
                    End If
                End If

                callScripture = callScripture & Trim(rawLine)
                pendingBlankLine = False
            End If

        ElseIf section = 2 Then

            If trimmedLine = "" Then

                ' A blank source line inside [啟應] is only meaningful when
                ' there is already an accumulated 啟：/應： line to attach
                ' trailing free text to (see the Else branch below). It is
                ' remembered the same way as in the [宣召] section and
                ' converted into TWO line breaks before the next non-empty
                ' line, producing one visible empty paragraph in PowerPoint.
                If pairCount > 0 Or hasPendingCall Then
                    pendingBlankLine2 = True
                End If

            Else

                If Left(trimmedLine, 2) = CNCallPrefix() Or _
                   Left(trimmedLine, 2) = CNCallPrefixHalf() Then

                    If hasPendingCall Then
                        errorMessage = MsgCallPairMismatch()
                        Exit Function
                    End If

                    pendingCall = Trim(Mid(trimmedLine, 3))

                    If pendingCall = "" Then
                        errorMessage = MsgCallEmptyLine(CNCallPrefix())
                        Exit Function
                    End If

                    hasPendingCall = True
                    pendingBlankLine2 = False

                ElseIf Left(trimmedLine, 2) = CNResponsePrefix() Or _
                       Left(trimmedLine, 2) = CNResponsePrefixHalf() Then

                    If Not hasPendingCall Then
                        errorMessage = MsgCallPairMismatch()
                        Exit Function
                    End If

                    If Trim(Mid(trimmedLine, 3)) = "" Then
                        errorMessage = MsgCallEmptyLine(CNResponsePrefix())
                        Exit Function
                    End If

                    pairCount = pairCount + 1

                    ReDim Preserve callLines(pairCount)
                    ReDim Preserve responseLines(pairCount)

                    callLines(pairCount) = pendingCall
                    responseLines(pairCount) = Trim(Mid(trimmedLine, 3))

                    pendingCall = ""
                    hasPendingCall = False
                    pendingBlankLine2 = False

                Else

                    ' A line that does not start with 啟：/應： (for example
                    ' a trailing scripture citation such as "(詩篇一〇三
                    ' 2-3,8-12)") is treated the same way free-form lines are
                    ' treated in the [宣召] section: it is appended to
                    ' whichever 啟：/應： text was most recently active,
                    ' with the same blank-line-becomes-double-break rule.
                    If pairCount > 0 Then

                        If pendingBlankLine2 Then
                            responseLines(pairCount) = responseLines(pairCount) & _
                                                        vbCrLf & vbCrLf & Trim(rawLine)
                        Else
                            responseLines(pairCount) = responseLines(pairCount) & _
                                                        vbCrLf & Trim(rawLine)
                        End If

                        pendingBlankLine2 = False

                    ElseIf hasPendingCall Then

                        If pendingBlankLine2 Then
                            pendingCall = pendingCall & vbCrLf & vbCrLf & Trim(rawLine)
                        Else
                            pendingCall = pendingCall & vbCrLf & Trim(rawLine)
                        End If

                        pendingBlankLine2 = False

                    Else
                        errorMessage = MsgCallBadLine()
                        Exit Function
                    End If
                End If
            End If
        End If
    Next

    hasCallScripture = (Trim(callScripture) <> "")

    If hasPendingCall Then
        errorMessage = MsgCallPairMismatch()
        Exit Function
    End If

    hasCallResponse = (pairCount > 0)

    If Not hasCallScripture And Not hasCallResponse Then
        errorMessage = MsgCallNoUsableSection()
        Exit Function
    End If

    ReadCallResponseFile = True
End Function

Function MsgCallNoUsableSection()
    MsgCallNoUsableSection = CNCallResponse() & ".txt " & _
                             ChrW(&H5167) & ChrW(&H6C92) & ChrW(&H6709) & _
                             ChrW(&H53EF) & ChrW(&H7528) & ChrW(&H5167) & _
                             ChrW(&H5BB9) & ChrW(&H3002) & _
                             ChrW(&H8ACB) & ChrW(&H586B) & ChrW(&H5BEB) & _
                             " " & CNCallSection() & " " & _
                             ChrW(&H6216) & " " & CNResponseSection() & _
                             ChrW(&H3002)
End Function


Function MsgCallSectionRepeated(ByVal sectionName)
    MsgCallSectionRepeated = sectionName & " " & _
                             ChrW(&H6BB5) & ChrW(&H843D) & _
                             ChrW(&H91CD) & ChrW(&H8907) & ChrW(&H51FA) & _
                             ChrW(&H73FE) & ChrW(&H3002)
End Function


Function MsgCallEmptyLine(ByVal prefix)
    MsgCallEmptyLine = prefix & " " & _
                       ChrW(&H5F8C) & ChrW(&H5FC5) & ChrW(&H9808) & _
                       ChrW(&H586B) & ChrW(&H5BEB) & ChrW(&H6587) & _
                       ChrW(&H5B57) & ChrW(&H3002)
End Function


Function MsgCallBadLine()
    MsgCallBadLine = CNResponseSection() & " " & _
                     ChrW(&H6BB5) & ChrW(&H843D) & _
                     ChrW(&H5167) & ChrW(&HFF0C) & _
                     ChrW(&H6BCF) & ChrW(&H884C) & _
                     ChrW(&H5FC5) & ChrW(&H9808) & _
                     ChrW(&H4EE5) & " " & _
                     CNCallPrefix() & " " & _
                     ChrW(&H6216) & " " & _
                     CNResponsePrefix() & " " & _
                     ChrW(&H958B) & ChrW(&H982D) & ChrW(&H3002)
End Function

' ===============================================================
' Missing-song marker
' ===============================================================

Function CreateMissingSongMarker( _
    ByVal destinationFolder, _
    ByVal songTitle, _
    ByRef errorMessage _
)

    Dim markerFile
    Dim textStream

    CreateMissingSongMarker = False
    errorMessage = ""

    markerFile = fso.BuildPath( _
        destinationFolder, _
        MsgMissingPrefix() & songTitle & ".txt" _
    )

    On Error Resume Next

    Set textStream = fso.CreateTextFile(markerFile, True, True)

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    textStream.WriteLine songTitle
    textStream.Close

    On Error GoTo 0

    CreateMissingSongMarker = True
End Function

' ===============================================================
' Path and validation helpers
' ===============================================================

Function IsValidFolderName(ByVal value)

    Dim invalidCharacters, i

    invalidCharacters = "\/:*?""<>|"

    IsValidFolderName = True

    If value = "" Or value = "." Or value = ".." Then
        IsValidFolderName = False
        Exit Function
    End If

    For i = 1 To Len(invalidCharacters)
        If InStr(value, Mid(invalidCharacters, i, 1)) > 0 Then
            IsValidFolderName = False
            Exit Function
        End If
    Next
End Function

Function MinNumber(ByVal firstValue, ByVal secondValue)

    If firstValue < secondValue Then
        MinNumber = firstValue
    Else
        MinNumber = secondValue
    End If
End Function

Sub SafeQuitPowerPoint(ByVal pptApp)

    On Error Resume Next

    If Not pptApp Is Nothing Then
        pptApp.Quit
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Sub CleanupFolder(ByVal path)

    On Error Resume Next

    If fso.FolderExists(path) Then
        fso.DeleteFolder path, True
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Sub Fail(ByVal text)

    MsgBox text, vbCritical, MsgAppTitle()
    EndOfRunCleanup    ' [ADDED] always clean up, even on failure exits
    WScript.Quit 1
End Sub

' ===============================================================
' Remove copied TXT files when a same-name PPTX exists
' ===============================================================

Sub RemoveTxtFilesWithMatchingPptx(ByVal folderPath)

    Dim folderObject
    Dim fileObject
    Dim pptxPath

    On Error Resume Next

    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo 0

    ' The folder is newly created, so deletion is safe:
    For Each fileObject In folderObject.Files

        If LCase(fso.GetExtensionName(fileObject.Name)) = "txt" Then

            pptxPath = fso.BuildPath( _
                folderPath, _
                fso.GetBaseName(fileObject.Name) & ".pptx" _
            )

            On Error Resume Next

            If fso.FileExists(pptxPath) Then
                fso.DeleteFile fileObject.Path, True
            End If

            Err.Clear
            On Error GoTo 0
        End If
    Next
End Sub

' ===============================================================
' Chinese filenames and Chinese messages
' ===============================================================
Function txtPrefix()
    txtPrefix = ChrW(&H5167) & ChrW(&H5BB9) & "-"
End Function

Function CNRead()
    CNRead = ChrW(&H8B80) & ChrW(&H7D93)
End Function

Function CNScriptures()
    CNScriptures = ChrW(&H7D93) & ChrW(&H8A13)
End Function

Function CNSongs()
    CNSongs = ChrW(&H8A69) & ChrW(&H6B4C)
End Function

Function CNCallResponse()
    CNCallResponse = ChrW(&H5BA3) & ChrW(&H53EC) & _
                     ChrW(&H53CA) & ChrW(&H555F) & _
                     ChrW(&H61C9)
End Function

Function CNCallSection()
    CNCallSection = "[" & ChrW(&H5BA3) & ChrW(&H53EC) & "]"
End Function

Function CNResponseSection()
    CNResponseSection = "[" & ChrW(&H555F) & ChrW(&H61C9) & "]"
End Function

Function CNCallPrefix()
    CNCallPrefix = ChrW(&H555F) & ChrW(&HFF1A)
End Function

Function CNResponsePrefix()
    CNResponsePrefix = ChrW(&H61C9) & ChrW(&HFF1A)
End Function

Function CNCallPrefixHalf()
    CNCallPrefixHalf = ChrW(&H555F) & ":"
End Function

Function CNResponsePrefixHalf()
    CNResponsePrefixHalf = ChrW(&H61C9) & ":"
End Function

Function MsgAppTitle()
    MsgAppTitle = ChrW(&H5EFA) & ChrW(&H7ACB) & _
                  ChrW(&H5D07) & ChrW(&H62DC) & _
                  ChrW(&H8CC7) & ChrW(&H6599) & _
                  ChrW(&H593E)
End Function

Function MsgtemplatesMissing()
    MsgtemplatesMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                        " templates " & _
                        ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                        ChrW(&H3002)
End Function

Function MsgCopyFolderMissing()
    MsgCopyFolderMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                           " templates\worship-files " & _
                           ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                           ChrW(&H3002)
End Function

Function MsgFileMissing(ByVal f)
    MsgFileMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                     ChrW(&H6A94) & ChrW(&H6848) & ChrW(&HFF1A) & f
End Function

Function MsgEnterFolderName()
    MsgEnterFolderName = ChrW(&H8ACB) & ChrW(&H8F38) & ChrW(&H5165) & _
                         ChrW(&H65B0) & ChrW(&H7684) & _
                         ChrW(&H5D07) & ChrW(&H62DC) & _
                         ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                         ChrW(&H540D) & ChrW(&H7A31) & ChrW(&HFF1A)
End Function

Function MsgInvalidFolderName()
    MsgInvalidFolderName = ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                           ChrW(&H540D) & ChrW(&H7A31) & _
                           ChrW(&H7121) & ChrW(&H6548) & ChrW(&H3002) & _
                           ChrW(&H4E0D) & ChrW(&H53EF) & ChrW(&H4F7F) & _
                           ChrW(&H7528) & ChrW(&HFF1A) & _
                           "\ / : * ? "" < > |"
End Function

Function MsgFolderExists()
    MsgFolderExists = ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                      ChrW(&H5DF2) & ChrW(&H7D93) & ChrW(&H5B58) & _
                      ChrW(&H5728) & ChrW(&HFF1A)
End Function

Function MsgNoChanges()
    MsgNoChanges = ChrW(&H6C92) & ChrW(&H6709) & _
                   ChrW(&H4F5C) & ChrW(&H51FA) & _
                   ChrW(&H4EFB) & ChrW(&H4F55) & _
                   ChrW(&H8B8A) & ChrW(&H66F4) & _
                   ChrW(&H3002)
End Function

Function MsgEnterSongs()
    MsgEnterSongs = ChrW(&H8ACB) & ChrW(&H8F38) & ChrW(&H5165) & _
                    ChrW(&H8A69) & ChrW(&H6B4C) & _
                    ChrW(&H540D) & ChrW(&H7A31) & _
                    ChrW(&HFF0C) & ChrW(&H4EE5) & _
                    ChrW(&H9017) & ChrW(&H865F) & _
                    ChrW(&H5206) & ChrW(&H9694) & _
                    ChrW(&HFF1A)
End Function

Function MsgSongExample()
    MsgSongExample = ChrW(&H4F8B) & ChrW(&H5982) & ChrW(&HFF1A) & _
                     ChrW(&H8A69) & ChrW(&H6B4C) & "A," & _
                     ChrW(&H8A69) & ChrW(&H6B4C) & "B," & _
                     ChrW(&H8A69) & ChrW(&H6B4C) & "C"
End Function

Function MsgDynamicFolderMissing()
    MsgDynamicFolderMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                              " templates\dynamic " & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & ChrW(&H3002)
End Function

Function MsgStandardFolderMissing()
    MsgStandardFolderMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                               " templates\standard " & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & ChrW(&H3002)
End Function

Function MsgCannotReadSongFolders()
    MsgCannotReadSongFolders = ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H8B80) & ChrW(&H53D6) & _
                                " song-folders.txt" & ChrW(&HFF1A)
End Function

Function MsgNoValidSongFolders()
    MsgNoValidSongFolders = "song-folders.txt " & ChrW(&H5167) & ChrW(&H6C92) & ChrW(&H6709) & _
                             ChrW(&H4EFB) & ChrW(&H4F55) & ChrW(&H6709) & ChrW(&H6548) & ChrW(&H7684) & _
                             ChrW(&H8A69) & ChrW(&H6B4C) & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                             ChrW(&H8DEF) & ChrW(&H5F91) & ChrW(&H3002)
End Function

Function MsgSongRootMissing()
    MsgSongRootMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                         " songs-folder.lnk " & _
                         ChrW(&H6240) & ChrW(&H6307) & ChrW(&H5B9A) & _
                         ChrW(&H7684) & ChrW(&H5D07) & ChrW(&H62DC) & _
                         ChrW(&H7528) & ChrW(&H8CC7) & ChrW(&H6599) & _
                         ChrW(&H6839) & ChrW(&H76EE) & ChrW(&H9304) & _
                         ChrW(&H3002)
End Function

Function MsgCannotStartPowerPoint()
    MsgCannotStartPowerPoint = ChrW(&H7121) & ChrW(&H6CD5) & _
                               ChrW(&H555F) & ChrW(&H52D5) & _
                               " Microsoft PowerPoint" & ChrW(&H3002)
End Function

Function MsgCannotCopyFixedFiles()
    MsgCannotCopyFixedFiles = ChrW(&H7121) & ChrW(&H6CD5) & _
                              ChrW(&H8907) & ChrW(&H88FD) & _
                              " templates\standard " & _
                              ChrW(&H5167) & ChrW(&H7684) & _
                              ChrW(&H6A94) & ChrW(&H6848) & _
                              ChrW(&HFF1A)
End Function

Function MsgCannotCreateFile(ByVal f)
    MsgCannotCreateFile = ChrW(&H7121) & ChrW(&H6CD5) & _
                          ChrW(&H5EFA) & ChrW(&H7ACB) & _
                          " " & f & ChrW(&HFF1A)
End Function

Function MsgCannotCopySong()
    MsgCannotCopySong = ChrW(&H7121) & ChrW(&H6CD5) & _
                        ChrW(&H8907) & ChrW(&H88FD) & _
                        ChrW(&H8A69) & ChrW(&H6B4C) & _
                        " PPTX" & ChrW(&HFF1A)
End Function

Function MsgCannotCreateMarker()
    MsgCannotCreateMarker = ChrW(&H7121) & ChrW(&H6CD5) & _
                            ChrW(&H5EFA) & ChrW(&H7ACB) & _
                            ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                            ChrW(&H8A69) & ChrW(&H6B4C) & _
                            ChrW(&H7684) & ChrW(&H63D0) & ChrW(&H793A) & _
                            ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                            ChrW(&HFF1A)
End Function

Function MsgCreatedRefreshFailed()
    MsgCreatedRefreshFailed = ChrW(&H5D07) & ChrW(&H62DC) & _
                              ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                              ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & _
                              ChrW(&HFF0C) & ChrW(&H4F46) & _
                              ChrW(&H7121) & ChrW(&H6CD5) & _
                              ChrW(&H81EA) & ChrW(&H52D5) & _
                              ChrW(&H66F4) & ChrW(&H65B0) & _
                              " master.pptx" & ChrW(&HFF1A)
End Function

Function MsgRunRefreshManually()
    MsgRunRefreshManually = ChrW(&H8ACB) & ChrW(&H5728) & _
                            ChrW(&H65B0) & ChrW(&H8CC7) & _
                            ChrW(&H6599) & ChrW(&H593E) & _
                            ChrW(&H5167) & ChrW(&H624B) & _
                            ChrW(&H52D5) & ChrW(&H57F7) & _
                            ChrW(&H884C) & _
                            " refresh-master.vbs" & ChrW(&H3002)
End Function

Function MsgDone()
    MsgDone = ChrW(&H5B8C) & ChrW(&H6210) & ChrW(&H3002)
End Function

Function MsgCreatedFolder()
    MsgCreatedFolder = ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & _
                       ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                       ChrW(&HFF1A)
End Function

Function MsgMissingSongMarkers()
    MsgMissingSongMarkers = ChrW(&H4EE5) & ChrW(&H4E0B) & _
                            ChrW(&H8A69) & ChrW(&H6B4C) & _
                            ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                            ChrW(&HFF0C) & ChrW(&H5DF2) & _
                            ChrW(&H5EFA) & ChrW(&H7ACB) & _
                            ChrW(&H63D0) & ChrW(&H793A) & _
                            ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                            ChrW(&HFF1A)
End Function

Function MsgMissingPrefix()
    MsgMissingPrefix = "(" & _
                       ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                       ") "
End Function

Function MsgCannotOpenTemplateCopy()
    MsgCannotOpenTemplateCopy = ChrW(&H7121) & ChrW(&H6CD5) & _
                                ChrW(&H958B) & ChrW(&H555F) & _
                                ChrW(&H8907) & ChrW(&H88FD) & _
                                ChrW(&H5F8C) & ChrW(&H7684) & _
                                ChrW(&H7BC4) & ChrW(&H672C) & ChrW(&HFF1A)
End Function

Function MsgTemplateSlideCount()
    MsgTemplateSlideCount = ChrW(&H7BC4) & ChrW(&H672C) & _
                            ChrW(&H5FC5) & ChrW(&H9808) & _
                            ChrW(&H53EA) & ChrW(&H5305) & ChrW(&H542B) & _
                            " 1 " & ChrW(&H5F35) & _
                            ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & _
                            ChrW(&H3002)
End Function

Function MsgTemplateMissingTitle()
    MsgTemplateMissingTitle = ChrW(&H7BC4) & ChrW(&H672C) & _
                              ChrW(&H7F3A) & ChrW(&H5C11) & _
                              " {{SCRIPTURE_TITLE}}" & ChrW(&H3002)
End Function

Function MsgTemplateMissingText()
    MsgTemplateMissingText = ChrW(&H7BC4) & ChrW(&H672C) & _
                             ChrW(&H7F3A) & ChrW(&H5C11) & _
                             " {{SCRIPTURE_TEXT}}" & ChrW(&H3002)
End Function

Function MsgCannotSavePpt()
    MsgCannotSavePpt = ChrW(&H7121) & ChrW(&H6CD5) & _
                       ChrW(&H5132) & ChrW(&H5B58) & _
                       ChrW(&H8F38) & ChrW(&H51FA) & ChrW(&H7684) & _
                       " PPTX" & ChrW(&HFF1A)
End Function

Function MsgCannotReadText()
    MsgCannotReadText = ChrW(&H7121) & ChrW(&H6CD5) & _
                        ChrW(&H8B80) & ChrW(&H53D6) & _
                        " UTF-8 " & _
                        ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                        ChrW(&HFF1A)
End Function

Function MsgEmptyTitle()
    MsgEmptyTitle = ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                    ChrW(&H7B2C) & ChrW(&H4E00) & ChrW(&H884C) & _
                    ChrW(&H5FC5) & ChrW(&H9808) & ChrW(&H662F) & _
                    ChrW(&H7D93) & ChrW(&H6587) & _
                    ChrW(&H6A19) & ChrW(&H984C) & ChrW(&H3002)
End Function

Function MsgCallTemplateSlideCount()
    MsgCallTemplateSlideCount = CNCallResponse() & ChrW(&H7BC4) & ChrW(&H672C) & _
                                ChrW(&H5FC5) & ChrW(&H9808) & _
                                ChrW(&H53EA) & ChrW(&H5305) & ChrW(&H542B) & _
                                " 2 " & ChrW(&H5F35) & _
                                ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & _
                                ChrW(&H3002)
End Function

Function MsgTemplateMissingNamedShape(ByVal shapeName)
    MsgTemplateMissingNamedShape = ChrW(&H7BC4) & ChrW(&H672C) & _
                                   ChrW(&H7F3A) & ChrW(&H5C11) & _
                                   ChrW(&H300C) & ChrW(&H9078) & ChrW(&H64C7) & ChrW(&H7A97) & ChrW(&H683C) & ChrW(&H300D) & _
                                   ChrW(&H540D) & ChrW(&H7A31) & " " & _
                                   shapeName & ChrW(&H3002)
End Function

Function MsgNamedShapeNotFound(ByVal shapeName)
    MsgNamedShapeNotFound = ChrW(&H300C) & ChrW(&H9078) & ChrW(&H64C7) & ChrW(&H7A97) & ChrW(&H683C) & ChrW(&H300D) & _
                            ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                            " " & shapeName & ChrW(&H3002)
End Function

Function MsgNamedShapeNoText(ByVal shapeName)
    MsgNamedShapeNoText = ChrW(&H300C) & ChrW(&H9078) & ChrW(&H64C7) & ChrW(&H7A97) & ChrW(&H683C) & ChrW(&H300D) & _
                          ChrW(&H540D) & ChrW(&H7A31) & " " & _
                          shapeName & " " & _
                          ChrW(&H4E0D) & ChrW(&H662F) & _
                          ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H65B9) & _
                          ChrW(&H584A) & ChrW(&H3002)
End Function

Function MsgCallSectionMissing()
    MsgCallSectionMissing = CNCallResponse() & ".txt " & _
                            ChrW(&H7F3A) & ChrW(&H5C11) & " " & _
                            CNCallSection() & " " & _
                            ChrW(&H6216) & " " & _
                            CNResponseSection() & " " & _
                            ChrW(&H6BB5) & ChrW(&H843D) & ChrW(&H3002)
End Function

Function MsgCallEmptyScripture()
    MsgCallEmptyScripture = CNCallSection() & " " & _
                            ChrW(&H6BB5) & ChrW(&H843D) & _
                            ChrW(&H6C92) & ChrW(&H6709) & _
                            ChrW(&H5167) & ChrW(&H5BB9) & ChrW(&H3002)
End Function

Function MsgCallNoPairs()
    MsgCallNoPairs = CNResponseSection() & " " & _
                     ChrW(&H6BB5) & ChrW(&H843D) & _
                     ChrW(&H6C92) & ChrW(&H6709) & _
                     ChrW(&H4EFB) & ChrW(&H4F55) & " " & _
                     CNCallPrefix() & "/" & CNResponsePrefix() & " " & _
                     ChrW(&H914D) & ChrW(&H5C0D) & ChrW(&H3002)
End Function

Function MsgCallPairMismatch()
    MsgCallPairMismatch = ChrW(&H6BCF) & ChrW(&H53E5) & " " & _
                          CNCallPrefix() & " " & _
                          ChrW(&H4E4B) & ChrW(&H5F8C) & _
                          ChrW(&H5FC5) & ChrW(&H9808) & _
                          ChrW(&H8DDF) & ChrW(&H96A8) & _
                          ChrW(&H4E00) & ChrW(&H53E5) & " " & _
                          CNResponsePrefix() & ChrW(&H3002)
End Function


' =====================================================================
' [ADDED] Embedded pre-run guard: PowerPoint state + input file sanity.
' Blocks the run (via Fail) on unsaved user work or crash-trigger files.
' =====================================================================

Sub PreRunGuard()
    Dim notes
    notes = ""

    GuardCheckPowerPointState notes
    GuardCheckTextFile CNRead() & ".txt"
    GuardCheckTextFile CNScriptures() & ".txt"
    GuardCheckTextFile CNCallResponse() & ".txt"

    If notes = "" Then
        notes = "- PowerPoint " & ChrW(&H72C0) & ChrW(&H614B) & ChrW(&H6B63) & ChrW(&H5E38) & ChrW(&H3002) & vbCrLf
    End If

    MsgBox ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H524D) & ChrW(&H6AA2) & ChrW(&H67E5) & ChrW(&H901A) & ChrW(&H904E) & ChrW(&HFF1A) & vbCrLf & vbCrLf & notes & vbCrLf & _
           ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&H5DF2) & ChrW(&H9A57) & ChrW(&H8B49) & ChrW(&H3002) & ChrW(&H73FE) & ChrW(&H5728) & ChrW(&H958B) & ChrW(&H59CB) & ChrW(&H3002), _
           vbInformation, MsgAppTitle()
End Sub

Sub GuardCheckPowerPointState(ByRef notes)
    Dim app, isVisible, presCount, i, unsavedList

    Set app = Nothing
    unsavedList = ""

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
        Fail "PowerPoint " & ChrW(&H6B63) & ChrW(&H5728) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H4F46) & ChrW(&H6C92) & ChrW(&H6709) & ChrW(&H56DE) & ChrW(&H61C9) & ChrW(&HFF08) & ChrW(&H53EF) & ChrW(&H80FD) & ChrW(&H5F48) & ChrW(&H51FA) & _
            ChrW(&H5C0D) & ChrW(&H8A71) & ChrW(&H65B9) & ChrW(&H584A) & ChrW(&HFF09) & ChrW(&H3002) & _
             ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H624B) & ChrW(&H52D5) & ChrW(&H8655) & ChrW(&H7406) & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H3002)
    End If

    If isVisible = msoTrue Then
        presCount = app.Presentations.Count
        If Err.Number = 0 Then
            For i = 1 To presCount
                If app.Presentations(i).Saved <> msoTrue Then
                    unsavedList = unsavedList & "    " & _
                                  app.Presentations(i).Name & vbCrLf
                End If
                If Err.Number <> 0 Then Exit For
            Next
        End If
        Err.Clear
        Set app = Nothing
        On Error GoTo 0

        If unsavedList <> "" Then
            Fail "PowerPoint " & ChrW(&H6709) & ChrW(&H5C1A) & ChrW(&H672A) & ChrW(&H5132) & ChrW(&H5B58) & ChrW(&H7684) & ChrW(&H7C21) & ChrW(&H5831) & ChrW(&HFF1A) & vbCrLf & _
                 unsavedList & vbCrLf & _
                 ChrW(&H9019) & ChrW(&H6B21) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H6703) & ChrW(&H95DC) & ChrW(&H9589) & " PowerPoint" & ChrW(&HFF0C) & ChrW(&H90A3) & ChrW(&H4E9B) & ChrW(&H672A) & ChrW(&H5132) & ChrW(&H5B58) & _
                     ChrW(&H7684) & ChrW(&H5DE5) & ChrW(&H4F5C) & ChrW(&H6703) & _
                 ChrW(&H907A) & ChrW(&H5931) & ChrW(&H3002) & ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H5132) & ChrW(&H5B58) & ChrW(&H4E26) & ChrW(&H95DC) & ChrW(&H9589) & " PowerPoint" & ChrW(&H3002) & ChrW(&H672A) & ChrW(&H505A) & _
                     ChrW(&H4EFB) & ChrW(&H4F55) & ChrW(&H66F4) & ChrW(&H6539) & ChrW(&H3002)
        End If

        notes = notes & "- PowerPoint " & ChrW(&H5DF2) & ChrW(&H958B) & ChrW(&H555F) & ChrW(&HFF08) & ChrW(&H5168) & ChrW(&H90E8) & ChrW(&H5DF2) & ChrW(&H5132) & ChrW(&H5B58) & ChrW(&HFF09) & ChrW(&HFF1B) & _
                ChrW(&H672C) & ChrW(&H6B21) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H5B8C) & ChrW(&H6210) & ChrW(&H5F8C) & ChrW(&H6703) & ChrW(&H628A) & ChrW(&H5B83) & ChrW(&H95DC) & ChrW(&H9589) & ChrW(&H3002) & vbCrLf
    Else
        presCount = app.Presentations.Count
        If Err.Number = 0 Then
            For i = presCount To 1 Step -1
                app.Presentations(i).Saved = msoTrue
                app.Presentations(i).Close
            Next
        End If
        app.Quit
        Err.Clear
        Set app = Nothing
        On Error GoTo 0
        notes = notes & "- " & ChrW(&H5DF2) & ChrW(&H79FB) & ChrW(&H9664) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H96B1) & ChrW(&H5F62) & ChrW(&H7684) & ChrW(&H5B64) & ChrW(&H7ACB) & " PowerPoint" & _
                " " & ChrW(&H5BE6) & ChrW(&H4F8B) & ChrW(&H3002) & vbCrLf
    End If
End Sub

Sub GuardCheckTextFile(ByVal fileName)
    Dim p, sz

    p = fso.BuildPath(rootFolder, fileName)
    If Not fso.FileExists(p) Then Exit Sub    ' existence checked by main flow

    On Error Resume Next
    sz = fso.GetFile(p).Size
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Fail ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H8B80) & ChrW(&H53D6) & ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&HFF1A) & " " & fileName
    End If
    On Error GoTo 0

    If sz < GUARD_MIN_TEXT_BYTES Then
        Fail ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&H662F) & ChrW(&H7A7A) & ChrW(&H7684) & ChrW(&HFF08) & ChrW(&H6216) & ChrW(&H53EA) & ChrW(&H6709) & " BOM" & ChrW(&HFF09) & ChrW(&HFF0C) & ChrW(&H6703) & _
            ChrW(&H5C0E) & ChrW(&H81F4) & ChrW(&H9019) & ChrW(&H6B21) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H4E2D) & ChrW(&H65B7) & ChrW(&HFF1A) & " " & _
             fileName & vbCrLf & ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H586B) & ChrW(&H5165) & ChrW(&H5167) & ChrW(&H5BB9) & ChrW(&H3002) & ChrW(&H672A) & ChrW(&H505A) & ChrW(&H4EFB) & ChrW(&H4F55) & ChrW(&H66F4) & ChrW(&H6539) & ChrW(&H3002)
    ElseIf sz > GUARD_MAX_TEXT_BYTES Then
        Fail ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&H9AD4) & ChrW(&H7A4D) & ChrW(&H7570) & ChrW(&H5E38) & ChrW(&H5927) & ChrW(&HFF08) & sz & " " & ChrW(&H4F4D) & ChrW(&H5143) & ChrW(&H7D44) & ChrW(&HFF09) & _
            ChrW(&HFF1A) & " " & _
             fileName & vbCrLf & ChrW(&H662F) & ChrW(&H4E0D) & ChrW(&H662F) & ChrW(&H653E) & ChrW(&H932F) & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&HFF1F) & ChrW(&H672A) & ChrW(&H505A) & ChrW(&H4EFB) & ChrW(&H4F55) & ChrW(&H66F4) & _
                 ChrW(&H6539) & ChrW(&H3002)
    End If
End Sub

' =====================================================================
' [ADDED] Embedded end-of-run cleanup (graceful quit + triple-gated
' orphan kill + lock-file sweep). Popup only when something was done.
' =====================================================================

Sub EndOfRunCleanup()
    Dim wmi, actions, quitDone

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
        If GuardOtherScriptBusy(wmi, "refresh-master", "create-lyrics-ppt") Then
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
            GuardRemoveLockFiles rootFolder, actions
        End If
    End If

    If actions <> "" Then
        MsgBox ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H5F8C) & ChrW(&H6E05) & ChrW(&H7406) & ChrW(&HFF1A) & vbCrLf & vbCrLf & actions, _
               vbInformation, MsgAppTitle()
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