Option Explicit

' PackNGo.bas - SolidWorks VBA macro
'
' Prompts for a job number, locates the source SolidWorks drawing
' referenced from the AutoCAD job folder's "Eng Ref" Word doc, then
' runs Pack-and-Go (flattened) into the SolidWorks job folder.

Private Const SW_ROOT      As String = "Z:\Solidworks\Current\JOBS\"
Private Const ACAD_ROOT    As String = "Z:\AUTOCAD\CURRENT\JOBS\"
Private Const DOC_MARKER   As String = "See file path below for original files."
Private Const SHORTCUT_BAT As String = "Z:\DAG\SOLIDWORKS-AUTOCAD JOB FOLDER\RunJobShortcut.bat"

Private Const LOG_DIR      As String = "Z:\DAG\SOLIDWORKS MACRO\Pack'n'Go\Log\"
Private Const LOG_XLSX     As String = "Z:\DAG\SOLIDWORKS MACRO\Pack'n'Go\Log\PackNGo_Log.xlsx"
Private Const LOG_OVERFLOW As String = "Z:\DAG\SOLIDWORKS MACRO\Pack'n'Go\Log\PackNGo_Log_Overflow.csv"
Private Const HEADER_ROW   As Long = 3
Private Const DATA_START   As Long = 4

' SolidWorks folder name -> AutoCAD folder name
Private Function MapAcadFolder(swType As String) As String
    Select Case UCase$(swType)
        Case "GENERAL LINE": MapAcadFolder = "GENERAL LINE"
        Case "HD-PFD":       MapAcadFolder = "HD-PFD-IAF"
        Case "HDX":          MapAcadFolder = "HDX"
        Case "AXIAL":        MapAcadFolder = "AXIAL"
        Case Else:           MapAcadFolder = ""
    End Select
End Function

' AutoCAD folder name -> SolidWorks folder name
Private Function MapSwFolder(acadType As String) As String
    Select Case UCase$(acadType)
        Case "GENERAL LINE": MapSwFolder = "GENERAL LINE"
        Case "HD-PFD-IAF":   MapSwFolder = "HD-PFD"
        Case "HDX":          MapSwFolder = "HDX"
        Case "AXIAL":        MapSwFolder = "AXIAL"
        Case Else:           MapSwFolder = ""
    End Select
End Function

' Range folder name based on first 3 digits, in groups of 5.
' Special case: 401-405 is rolled into "400-405".
Private Function ComputeRangeFolder(jobNum As String) As String
    Dim prefix As Long: prefix = CLng(Left$(jobNum, 3))
    Dim n As Long:      n = -Int(-prefix / 5)               ' ceil(prefix / 5)
    Dim startN As Long: startN = 5 * (n - 1) + 1
    Dim endN As Long:   endN = 5 * n
    If startN = 401 And endN = 405 Then
        ComputeRangeFolder = "400-405"
    Else
        ComputeRangeFolder = startN & "-" & endN
    End If
End Function

' AutoCAD intermediate: HDX -> range folder; everyone else -> first 3 digits.
Private Function ComputeAcadIntermediate(acadType As String, jobNum As String) As String
    If UCase$(acadType) = "HDX" Then
        ComputeAcadIntermediate = ComputeRangeFolder(jobNum)
    Else
        ComputeAcadIntermediate = Left$(jobNum, 3)
    End If
End Function

' SolidWorks intermediate: HD-PFD lives in a single "40XXXX" bucket;
' everyone else uses a range folder on the first 3 digits.
Private Function ComputeSwIntermediate(swType As String, jobNum As String) As String
    If UCase$(swType) = "HD-PFD" Then
        ComputeSwIntermediate = "40XXXX"
    Else
        ComputeSwIntermediate = ComputeRangeFolder(jobNum)
    End If
End Function

Private Function FolderExists(p As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function FileExists(p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(p)) > 0)
    On Error GoTo 0
End Function

Private Function NormalizeFolder(p As String) As String
    NormalizeFolder = Trim$(p)
    If Right$(NormalizeFolder, 1) <> "\" Then NormalizeFolder = NormalizeFolder & "\"
End Function

' Probes every AutoCAD job-type folder; returns the type that contains
' <jobNum> and writes the matching AutoCAD job folder path to acadJobFolder.
Private Function FindAcadJobType(jobNum As String, ByRef acadJobFolder As String) As String
    Dim acadTypes As Variant
    acadTypes = Array("GENERAL LINE", "HD-PFD-IAF", "HDX", "AXIAL")
    Dim i As Long, candidate As String, intermediate As String
    For i = LBound(acadTypes) To UBound(acadTypes)
        intermediate = ComputeAcadIntermediate(CStr(acadTypes(i)), jobNum)
        candidate = ACAD_ROOT & acadTypes(i) & "\" & intermediate & "\" & jobNum & "\"
        If FolderExists(candidate) Then
            FindAcadJobType = CStr(acadTypes(i))
            acadJobFolder = candidate
            Exit Function
        End If
    Next i
    FindAcadJobType = ""
End Function

' Recursively creates a folder (and any missing parents).
Private Sub EnsureFolder(p As String)
    Dim folder As String: folder = NormalizeFolder(p)
    If FolderExists(folder) Then Exit Sub

    Dim trimmed As String: trimmed = Left$(folder, Len(folder) - 1)
    Dim lastSlash As Long: lastSlash = InStrRev(trimmed, "\")
    If lastSlash > 0 Then EnsureFolder Left$(trimmed, lastSlash)

    On Error Resume Next
    MkDir folder
    On Error GoTo 0
End Sub

' Hands the AutoCAD job folder to the shortcut-creator batch file
' (the same way the user does it manually by drag-and-drop).
Private Function RunShortcutBat(folderPath As String) As Boolean
    If Not FileExists(SHORTCUT_BAT) Then
        MsgBox "Shortcut helper not found:" & vbCrLf & SHORTCUT_BAT, vbExclamation
        RunShortcutBat = False
        Exit Function
    End If

    Dim arg As String: arg = folderPath
    If Right$(arg, 1) = "\" Then arg = Left$(arg, Len(arg) - 1)

    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & SHORTCUT_BAT & """ """ & arg & """", 1, True
    If Err.Number <> 0 Then
        MsgBox "Failed to run shortcut helper:" & vbCrLf & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        RunShortcutBat = False
        Exit Function
    End If
    On Error GoTo 0
    RunShortcutBat = True
End Function

' Opens the Eng Ref doc, finds the marker line, returns the next non-empty paragraph.
Private Function ReadSourcePathFromDoc(docPath As String) As String
    Dim wdApp As Object, wdDoc As Object, para As Object
    Dim found As Boolean, txt As String
    On Error GoTo Cleanup
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(Filename:=docPath, ReadOnly:=True, AddToRecentFiles:=False)
    For Each para In wdDoc.Paragraphs
        txt = para.Range.Text
        txt = Replace(txt, Chr$(13), "")
        txt = Replace(txt, Chr$(11), "")
        txt = Replace(txt, Chr$(7), "")
        txt = Trim$(txt)
        If found And Len(txt) > 0 Then
            ReadSourcePathFromDoc = txt
            GoTo Cleanup
        End If
        If InStr(1, txt, DOC_MARKER, vbTextCompare) > 0 Then found = True
    Next para
Cleanup:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Function

' Looks for <jobNum>-01.SLDDRW first, then <jobNum>-02.SLDDRW.
Private Function FindSourceDrawing(sourceFolder As String, jobNum As String, _
                                   ByRef drawingBase As String) As String
    Dim folder As String: folder = NormalizeFolder(sourceFolder)
    Dim cand As String
    cand = folder & jobNum & "-01.SLDDRW"
    If FileExists(cand) Then
        drawingBase = jobNum & "-01"
        FindSourceDrawing = cand
        Exit Function
    End If
    cand = folder & jobNum & "-02.SLDDRW"
    If FileExists(cand) Then
        drawingBase = jobNum & "-02"
        FindSourceDrawing = cand
        Exit Function
    End If
    FindSourceDrawing = ""
End Function

Private Function HasSwFiles(folder As String) As Boolean
    HasSwFiles = (Len(Dir$(NormalizeFolder(folder) & "*.SLD*")) > 0)
End Function

' If the SW job folder is empty of SW files, returns it as the destination.
' Otherwise prompts for a sub-folder name (defaulting to <drawingBase>_(N))
' or returns "" if the user cancels.
Private Function ResolveDestination(swJobFolder As String, drawingBase As String) As String
    Dim folder As String: folder = NormalizeFolder(swJobFolder)
    If Not HasSwFiles(folder) Then
        ResolveDestination = folder
        Exit Function
    End If

    Dim n As Long, defaultName As String, candidate As String
    n = 2
    Do
        defaultName = drawingBase & "_(" & n & ")"
        candidate = folder & defaultName & "\"
        If Not FolderExists(candidate) Then Exit Do
        n = n + 1
    Loop

    Dim userName As String
    userName = InputBox( _
        "Job folder already contains SolidWorks files." & vbCrLf & vbCrLf & _
        "Enter a sub-folder name to place this Pack-and-Go in," & vbCrLf & _
        "or click Cancel to abort.", _
        "Pack-n-Go: sub-folder name", defaultName)
    userName = Trim$(userName)
    If Len(userName) = 0 Then
        ResolveDestination = ""
        Exit Function
    End If
    candidate = folder & userName & "\"
    If Not FolderExists(candidate) Then MkDir candidate
    ResolveDestination = candidate
End Function

' Logs one Pack-n-Go run to PackNGo_Log.xlsx, falling through to a CSV
' overflow file if the workbook is locked or Excel is unavailable.
Private Sub LogPackNGo(ByVal jobNum As String, _
                       ByVal jobType As String, _
                       ByVal drawing As String, _
                       ByVal destination As String, _
                       ByVal usedSubfolder As String, _
                       ByVal shortcutRun As String, _
                       ByVal timeSavedMin As Long)
    Dim headers As Variant
    headers = Array("Date", "Time", "User", "Job Number", "Job Type", _
                    "Drawing", "Destination", "Used Subfolder", _
                    "Shortcut Run", "Time Saved (min)")
    Const TIME_SAVED_COL As Long = 10  ' index in headers (1-based)

    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0

    EnsureFolder LOG_DIR

    Dim xlApp As Object, wb As Object, ws As Object
    Dim isNew As Boolean

    On Error GoTo Overflow
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False

    If fso.FileExists(LOG_XLSX) Then
        Set wb = xlApp.Workbooks.Open(Filename:=LOG_XLSX, ReadOnly:=False)
        If wb.ReadOnly Then
            wb.Close SaveChanges:=False
            Set wb = Nothing
            xlApp.Quit
            Set xlApp = Nothing
            GoTo Overflow
        End If
        Set ws = wb.Sheets(1)
    Else
        Set wb = xlApp.Workbooks.Add
        Set ws = wb.Sheets(1)
        ws.Name = "PackNGo Log"
        isNew = True
    End If

    ' Summary row (always re-asserted to upgrade old hardcoded values)
    ws.Cells(1, 1).Value = "Total Runs"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Formula = "=COUNTA(A" & DATA_START & ":A1048576)"
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 3).Value = "Total Minutes Saved"
    ws.Cells(1, 3).Font.Bold = True
    ws.Cells(1, 4).Formula = "=SUM(" & ColLetter(TIME_SAVED_COL) & DATA_START & ":" & _
                             ColLetter(TIME_SAVED_COL) & "1048576)"
    ws.Cells(1, 4).Font.Bold = True

    ' Headers in row 3 - write missing ones, leave existing ones alone
    Dim cell As Object, i As Long
    For i = 0 To UBound(headers)
        Set cell = ws.Cells(HEADER_ROW, i + 1)
        If Len(CStr(cell.Value)) = 0 Then
            cell.Value = headers(i)
            cell.Font.Bold = True
        End If
    Next i

    ' Next data row
    Dim nextRow As Long
    If isNew Then
        nextRow = DATA_START
    Else
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  ' xlUp
        If lastRow < DATA_START Then nextRow = DATA_START Else nextRow = lastRow + 1
    End If

    ws.Cells(nextRow, 1).Value = Format$(Now, "YYYY-MM-DD")
    ws.Cells(nextRow, 2).Value = Format$(Now, "HH:MM:SS")
    ws.Cells(nextRow, 3).Value = Environ$("USERNAME")
    ws.Cells(nextRow, 4).Value = jobNum
    ws.Cells(nextRow, 5).Value = jobType
    ws.Cells(nextRow, 6).Value = drawing
    ws.Cells(nextRow, 7).Value = destination
    ws.Cells(nextRow, 8).Value = usedSubfolder
    ws.Cells(nextRow, 9).Value = shortcutRun
    ws.Cells(nextRow, 10).Value = timeSavedMin

    ws.UsedRange.Columns.AutoFit

    If isNew Then
        wb.SaveAs Filename:=LOG_XLSX, FileFormat:=51   ' xlOpenXMLWorkbook
    Else
        wb.Save
    End If

    wb.Close SaveChanges:=False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing: Set fso = Nothing
    Exit Sub

Overflow:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing

    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    If Not fso.FileExists(LOG_OVERFLOW) Then
        Set ts = fso.CreateTextFile(LOG_OVERFLOW, True)
        ts.WriteLine Join(headers, ",")
        ts.Close
    End If
    Set ts = fso.OpenTextFile(LOG_OVERFLOW, 8, True)   ' ForAppending
    ts.WriteLine Format$(Now, "YYYY-MM-DD") & "," & _
                 Format$(Now, "HH:MM:SS") & "," & _
                 Environ$("USERNAME") & "," & _
                 CsvEscape(jobNum) & "," & _
                 CsvEscape(jobType) & "," & _
                 CsvEscape(drawing) & "," & _
                 CsvEscape(destination) & "," & _
                 CsvEscape(usedSubfolder) & "," & _
                 CsvEscape(shortcutRun) & "," & _
                 timeSavedMin
    ts.Close
    Set ts = Nothing: Set fso = Nothing
End Sub

Private Function ColLetter(colNum As Long) As String
    Dim n As Long: n = colNum
    Do While n > 0
        ColLetter = Chr$(65 + ((n - 1) Mod 26)) & ColLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function CsvEscape(s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CsvEscape = """" & Replace(s, """", """""") & """"
    Else
        CsvEscape = s
    End If
End Function

Public Sub main()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim jobNum As String
    jobNum = Trim$(InputBox("Enter job number:", "Pack-n-Go"))
    If Len(jobNum) = 0 Then Exit Sub
    If Not IsNumeric(jobNum) Or Len(jobNum) < 3 Then
        MsgBox "Job number must be numeric and at least 3 digits.", vbExclamation
        Exit Sub
    End If

    ' AutoCAD folder is the authoritative source; probe it to determine type.
    Dim acadJobFolder As String, acadType As String
    acadType = FindAcadJobType(jobNum, acadJobFolder)
    If Len(acadType) = 0 Then
        MsgBox "No AutoCAD job folder found for job " & jobNum & "." & vbCrLf & _
               "Searched all type folders under " & ACAD_ROOT, vbExclamation
        Exit Sub
    End If

    ' Build SW job folder path (creating it if missing).
    Dim swJobFolder As String, swType As String
    swType = MapSwFolder(acadType)
    swJobFolder = SW_ROOT & swType & "\" & _
                  ComputeSwIntermediate(swType, jobNum) & "\" & jobNum & "\"
    EnsureFolder swJobFolder
    If Not FolderExists(swJobFolder) Then
        MsgBox "Could not create SolidWorks job folder:" & vbCrLf & swJobFolder, vbExclamation
        Exit Sub
    End If

    Dim docPath As String
    docPath = acadJobFolder & "ENG REF\" & jobNum & " Eng Ref.docx"
    If Not FileExists(docPath) Then
        MsgBox "Engineering Reference doc not found:" & vbCrLf & docPath, vbExclamation
        Exit Sub
    End If

    Dim sourceFolder As String
    sourceFolder = ReadSourcePathFromDoc(docPath)
    If Len(sourceFolder) = 0 Then
        MsgBox "Could not find a path under the marker '" & DOC_MARKER & "' in:" & vbCrLf & docPath, vbExclamation
        Exit Sub
    End If
    If Not FolderExists(sourceFolder) Then
        MsgBox "Source folder from Eng Ref doc does not exist:" & vbCrLf & sourceFolder, vbExclamation
        Exit Sub
    End If

    Dim drawingPath As String, drawingBase As String
    drawingPath = FindSourceDrawing(sourceFolder, jobNum, drawingBase)
    If Len(drawingPath) = 0 Then
        MsgBox "No drawing named " & jobNum & "-01 or " & jobNum & "-02 found in:" & vbCrLf & sourceFolder, vbExclamation
        Exit Sub
    End If

    Dim destFolder As String
    destFolder = ResolveDestination(swJobFolder, drawingBase)
    If Len(destFolder) = 0 Then Exit Sub  ' user cancelled

    Dim errors As Long, warnings As Long
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_ReadOnly, "", errors, warnings)
    If swModel Is Nothing Then
        MsgBox "Failed to open drawing:" & vbCrLf & drawingPath, vbExclamation
        Exit Sub
    End If

    Dim swPnG As SldWorks.PackAndGo
    Set swPnG = swModel.Extension.GetPackAndGo
    swPnG.SetSaveToName True, destFolder
    swPnG.FlattenToSingleFolder = True
    swPnG.IncludeDrawings = True

    Dim statuses As Variant
    statuses = swModel.Extension.SavePackAndGo(swPnG)

    Dim title As String
    title = swModel.GetTitle
    swApp.CloseDoc title

    Dim shortcutOk As Boolean
    shortcutOk = RunShortcutBat(acadJobFolder)

    Dim usedSub As String
    usedSub = IIf(StrComp(NormalizeFolder(destFolder), NormalizeFolder(swJobFolder), vbTextCompare) = 0, "No", "Yes")

    LogPackNGo jobNum, swType, drawingBase & ".SLDDRW", destFolder, _
               usedSub, IIf(shortcutOk, "Yes", "No"), 4

    Shell "explorer.exe """ & destFolder & """", vbNormalFocus

    MsgBox "Pack-and-Go complete." & vbCrLf & _
           "Drawing: " & drawingBase & ".SLDDRW" & vbCrLf & _
           "Destination: " & destFolder, vbInformation + vbSystemModal
End Sub
