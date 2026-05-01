Option Explicit

' PackNGo.bas - SolidWorks VBA macro
'
' Prompts for a job number, locates the source SolidWorks drawing
' referenced from the AutoCAD job folder's "Eng Ref" Word doc, then
' runs Pack-and-Go (flattened) into the SolidWorks job folder.

Private Const SW_ROOT    As String = "Z:\Solidworks\Current\JOBS\"
Private Const ACAD_ROOT  As String = "Z:\AUTOCAD\CURRENT\JOBS\"
Private Const DOC_MARKER As String = "See file path below for original files."

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

Private Function ComputeIntermediate(jobType As String, jobNum As String) As String
    Dim prefix As Long
    prefix = CLng(Left$(jobNum, 3))
    If UCase$(jobType) = "HDX" Then
        Dim n As Long, startN As Long, endN As Long
        n = -Int(-prefix / 5)               ' ceil(prefix / 5)
        startN = 5 * (n - 1) + 1
        endN = 5 * n
        ' 401-405 is rolled into the 400-405 folder
        If startN = 401 And endN = 405 Then
            ComputeIntermediate = "400-405"
        Else
            ComputeIntermediate = startN & "-" & endN
        End If
    Else
        ComputeIntermediate = CStr(prefix)
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
        intermediate = ComputeIntermediate(CStr(acadTypes(i)), jobNum)
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
                  ComputeIntermediate(acadType, jobNum) & "\" & jobNum & "\"
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

    MsgBox "Pack-and-Go complete." & vbCrLf & _
           "Drawing: " & drawingBase & ".SLDDRW" & vbCrLf & _
           "Destination: " & destFolder, vbInformation
End Sub
