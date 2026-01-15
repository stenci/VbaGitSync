Attribute VB_Name = "GitSync"
Option Explicit

' Requires reference to:
'  - Microsoft Scripting Runtime: C:\Windows\SysWOW64\scrrun.dll
'  - Microsoft Visual Basic for Applications Extensibility: C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'
' Enable programmatic access to the VBA project object model:
' File > Options > Trust Center > Trust Center Settings > Macro Settings >
'   Enable "Trust access to the VBA project object model"

Private ExportPath As String
Private Const MetaDataFile As String = ".GitSync"
Private FSO As New Scripting.FileSystemObject
Private WB As Workbook
Private NonAnsiWarnings As Collection

Public Function GitSync(WorkbookToExport As Workbook, Optional ExportToPath As String, Optional ShowReport As Boolean = True) As Collection
  Set WB = WorkbookToExport
  ExportPath = IIf(ExportToPath <> "", ExportToPath, WB.Path)
  If Right(ExportPath, 1) <> "\" Then ExportPath = ExportPath & "\"
  
  ' Check if the workbook is stored on SharePoint or another web-based location; local file operations are not supported
  If Left(LCase(ExportPath), 7) = "http://" Or Left(LCase(ExportPath), 8) = "https://" Then Exit Function

  Dim MetaDataPath As String
  MetaDataPath = ExportPath & MetaDataFile

  If Not FSO.FileExists(MetaDataPath) Then
    If MsgBox("No "".GitSync"" file found in this folder:" & vbLf & ExportPath & vbLf & vbLf & _
      "This usually means this is not a VBA repository folder." & vbLf & vbLf & _
      "Do you want to export and create a synced folder here?" & vbLf & vbLf & _
      "If this is not the folder you want to sync, click No to cancel.", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Sync") <> vbYes Then Exit Function
  End If

  Set NonAnsiWarnings = New Collection

  Dim VBProj As VBIDE.VBProject
  Set VBProj = WB.VBProject

  DeleteAllConflictArtifacts VBProj

  ' load & prune metadata
  Dim MetaData As New Dictionary, FileContent As String, Lines() As String, I As Long, LineItem, Parts() As String
  If FSO.FileExists(MetaDataPath) Then
    FileContent = ReadAllText(MetaDataPath)
    Lines = Split(FileContent, vbCrLf)
    For Each LineItem In Lines
      Parts = Split(LineItem, "/")
      If UBound(Parts) = 2 Then MetaData(Parts(0)) = Parts(1) & "/" & Parts(2)
    Next LineItem

    Dim ValidFiles As New Dictionary, Comp As VBIDE.VBComponent, Ext As String
    For Each Comp In VBProj.VBComponents
      Ext = GetComponentExtension(Comp)
      If Ext <> "" Then ValidFiles(Comp.Name & Ext) = True
    Next Comp

    Dim RemovedFromMetadata As New Collection, Key
    For Each Key In MetaData.Keys
      If Not ValidFiles.Exists(Key) Then
        RemovedFromMetadata.Add Key
      End If
    Next Key
    
    For Each Key In RemovedFromMetadata
      MetaData.Remove Key
    Next Key
  End If

  ' sync loop
  Dim FName As String, FilePath As String
  For Each Comp In VBProj.VBComponents
    Ext = GetComponentExtension(Comp)
    If Ext = "" Then GoTo NextComp

    FName = Comp.Name & Ext
    FilePath = ExportPath & FName

    ' read external if exists
    Dim ExternalExists As Boolean, ExternalCode As String, ExternalHash As String, ExternalTime As Date
    ExternalExists = FSO.FileExists(FilePath)
    If ExternalExists Then
      ExternalCode = StripTrailingEmptyLines(ReadAllText(FilePath))
      ExternalHash = GetTextHash(ExternalCode)
      ExternalTime = FileDateTime(FilePath)
    End If

    ' read internal (for all modules, including document modules)
    Dim InternalCode As String, InternalHash As String
    InternalCode = GetComponentCode(Comp)
    InternalHash = GetTextHash(InternalCode)

    If MetaData.Exists(FName) And ExternalExists Then
      Dim KnownHash As String, LastSync As Date
      KnownHash = Split(MetaData(FName), "/")(0)
      LastSync = CDate(Split(MetaData(FName), "/")(1))

      ' 1. internal changed only -> export
      Dim ExportedFiles As New Collection
      If InternalHash <> KnownHash And ExternalHash = KnownHash Then
        WriteAllText FilePath, InternalCode
        ExportedFiles.Add FName
        MetaData(FName) = InternalHash & "/" & Format(Now, "yyyy-mm-dd hh:nn:ss")
        GoTo NextComp
      End If

      ' 2. external changed
      If ExternalHash <> KnownHash Then
        If InternalHash <> KnownHash Then
          If InternalHash = ExternalHash Then
            ' Both changed, but now identical: just update metadata
            Dim MetaUpdatedFiles As New Collection
            MetaData(FName) = InternalHash & "/" & Format(Now, "yyyy-mm-dd hh:nn:ss")
            MetaUpdatedFiles.Add FName
          Else
            ' true conflict: hashes differ
            Dim ConflictedFiles As New Collection
            WriteAllText ExportPath & Comp.Name & "-vba" & Ext, InternalCode
            ConflictedFiles.Add FName
          End If
        ElseIf ExternalTime > LastSync Then
          ' external only -> import
          Dim ImportedFiles As New Collection, NormalizedExternalCode As String
          NormalizedExternalCode = StripTrailingEmptyLines(EnsureFileCrlf(FilePath))
          If NormalizedExternalCode <> ExternalCode Then
            ExternalCode = NormalizedExternalCode
            ExternalHash = GetTextHash(ExternalCode)
          End If
          Set Comp = ImportComponentCode(VBProj, Comp, ExternalCode, FilePath)
          ImportedFiles.Add FName
          If Comp.Type = vbext_ct_Document Then
            InternalCode = GetComponentCode(Comp)
            InternalHash = GetTextHash(InternalCode)
            MetaData(FName) = InternalHash & "/" & Format(Now, "yyyy-mm-dd hh:nn:ss")
          Else
            MetaData(FName) = ExternalHash & "/" & Format(Now, "yyyy-mm-dd hh:nn:ss")
          End If
        End If
      End If
    Else
      ' new module -> export
      WriteAllText FilePath, InternalCode
      ExportedFiles.Add FName
      MetaData(FName) = InternalHash & "/" & Format(Now, "yyyy-mm-dd hh:nn:ss")
    End If

NextComp:
  Next Comp

  ExportAllSheetsToCSV ExportedFiles

  ' save metadata
  Dim OutputText As String
  For Each LineItem In MetaData
    OutputText = OutputText & LineItem & "/" & MetaData(LineItem) & vbCrLf
  Next LineItem
  WriteAllText MetaDataPath, OutputText

  ' Now delete unused files (after metadata is up-to-date)
  Dim DeletedFiles As Collection
  Set DeletedFiles = DeleteUnusedFiles

  Set GitSync = New Collection
  Set GitSync = GenerateReportLines(ExportedFiles, ImportedFiles, ConflictedFiles, DeletedFiles, MetaUpdatedFiles, OnlyModified:=True)
  If ShowReport Then ShowReportSummary GitSync

  FSO.CopyFile WB.FullName, ExportPath & WB.Name
End Function

Private Function DeleteUnusedFiles() As Collection
  Dim Deleted As New Collection
  If Not FSO.FileExists(ExportPath & MetaDataFile) Then
    Set DeleteUnusedFiles = Deleted
    Exit Function
  End If

  Dim Folder As Scripting.Folder, TrackedFiles As Dictionary, CsvFiles As Dictionary
  Set Folder = FSO.GetFolder(ExportPath)
  Set TrackedFiles = GetAllTrackedFiles
  Set CsvFiles = New Dictionary

  Dim WS As Worksheet, FileName As String
  For Each WS In WB.Worksheets
    FileName = GetWorksheetCsvFileName(WS)
    CsvFiles(FileName) = True
  Next WS

  Dim File As Scripting.File, Ext As String
  For Each File In Folder.Files
    Ext = LCase(FSO.GetExtensionName(File.Name))
    If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Then
      ' Only process base files (not -vba)
      If Not Right(File.Name, Len("-vba." & Ext)) = "-vba." & Ext Then
        If Not TrackedFiles.Exists(File.Name) Then
          Deleted.Add File.Name
          ' Also delete -vba version if it exists
          Dim BaseName As String, VbaFile As String
          BaseName = Left(File.Name, Len(File.Name) - Len(Ext) - 1)
          VbaFile = BaseName & "-vba." & Ext
          On Error Resume Next
          File.Delete
          If FSO.FileExists(Folder.Path & "\" & VbaFile) Then
            Kill Folder.Path & "\" & VbaFile
            Deleted.Add VbaFile
          End If
          On Error GoTo 0
        End If
      End If
    ElseIf Ext = "csv" Then
      If Not CsvFiles.Exists(File.Name) Then
        Deleted.Add File.Name
        On Error Resume Next
        File.Delete
        On Error GoTo 0
      End If
    End If
  Next File

  Set DeleteUnusedFiles = Deleted
End Function

Private Function GetComponentCode(Comp As VBIDE.VBComponent) As String
  If Right(Comp.Name, 11) = "__to_delete" Then Stop
  Dim TempPath As String, CM As VBIDE.CodeModule
  If Comp.Type = vbext_ct_ClassModule Or Comp.Type = vbext_ct_StdModule Then
    TempPath = ExportPath & "__temp" & GetComponentExtension(Comp)
    Comp.Export TempPath
    GetComponentCode = ReadAllText(TempPath)
    Kill TempPath
    GetComponentCode = StripTrailingEmptyLines(GetComponentCode)
  Else
    Set CM = Comp.CodeModule
    GetComponentCode = StripTrailingEmptyLines(CM.Lines(1, CM.CountOfLines))
  End If
  WarnIfNonAnsi GetComponentCode, Comp.Name
End Function

Private Function ImportComponentCode(VBProj As VBIDE.VBProject, Comp As VBIDE.VBComponent, CodeText As String, FilePath As String) As VBIDE.VBComponent
  If Comp.Type = vbext_ct_ClassModule Or Comp.Type = vbext_ct_StdModule Then
    Dim CompName As String
    CompName = Comp.Name
    ' sometimes removing a component fails silently (vba bug?), so we add a suffix and then we
    ' try to delete it. if the deletion fails, the module will be deleted at the next run
    Comp.Name = Left(CompName, 21) & "__to_delete" 'this fails when two modules with names sharing the first 21 letters are externally modified at the same time... and i'm ok with it!
    WriteAllText FilePath, CodeText
    VBProj.VBComponents.Remove Comp
    VBProj.VBComponents.Import FilePath
    Set ImportComponentCode = VBProj.VBComponents(CompName)
  Else
    Dim CM As VBIDE.CodeModule
    Set CM = Comp.CodeModule
    CM.DeleteLines 1, CM.CountOfLines
    If Len(CodeText) > 0 Then CM.InsertLines 1, CodeText
    Set ImportComponentCode = Comp
  End If
End Function

Private Function ReadAllText(FilePath As String) As String
  ReadAllText = FSO.OpenTextFile(FilePath, ForReading).ReadAll
End Function

Private Sub WriteAllText(FilePath As String, Text As String)
  FSO.CreateTextFile(FilePath, True).Write Text
End Sub

Private Function EnsureFileCrlf(FilePath As String) As String
  Dim FileContent As String, Normalized As String
  FileContent = ReadAllText(FilePath)
  Normalized = NormalizeLineEndingsToCrlf(FileContent)
  If Normalized <> FileContent Then WriteAllText FilePath, Normalized
  EnsureFileCrlf = Normalized
End Function

Private Function NormalizeLineEndingsToCrlf(Text As String) As String
  Dim Normalized As String
  Normalized = Replace(Text, vbCrLf, vbLf)
  Normalized = Replace(Normalized, vbCr, vbLf)
  NormalizeLineEndingsToCrlf = Replace(Normalized, vbLf, vbCrLf)
End Function

Private Function GetComponentExtension(Comp As VBIDE.VBComponent) As String
  Select Case Comp.Type
    Case vbext_ct_StdModule: GetComponentExtension = ".bas"
    Case vbext_ct_ClassModule: GetComponentExtension = ".cls"
    Case vbext_ct_MSForm: GetComponentExtension = ".frm"
    Case vbext_ct_Document: GetComponentExtension = ".cls"
    Case Else: GetComponentExtension = ""
  End Select
End Function

Private Function IsDocumentModule(Comp As VBIDE.VBComponent) As Boolean
  IsDocumentModule = Comp.Type = vbext_ct_Document
End Function

Private Function StripTrailingEmptyLines(Text As String) As String
  Dim Lines() As String, I As Long
  Lines = Split(Text, vbCrLf)
  For I = UBound(Lines) To 0 Step -1
    If Trim(Lines(I)) <> "" Then Exit For
  Next I
  ReDim Preserve Lines(0 To I)
  StripTrailingEmptyLines = VBA.Join(Lines, vbCrLf)
End Function

Private Function GetTextHash(Text As String) As String
  Dim MD5 As Object
  Set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
  Dim Bytes() As Byte: Bytes = StrConv(Text, vbFromUnicode)
  Dim Hash() As Byte: Hash = MD5.ComputeHash_2(Bytes)
  Dim I As Long
  For I = 0 To UBound(Hash)
    GetTextHash = GetTextHash & LCase(Right("0" & Hex(Hash(I)), 2))
  Next I
End Function

Private Function GetAllTrackedFiles() As Dictionary
  Dim Result As New Dictionary, MetaDataPath As String
  MetaDataPath = ExportPath & MetaDataFile
  If Not FSO.FileExists(MetaDataPath) Then
    Set GetAllTrackedFiles = Result
    Exit Function
  End If

  Dim FileContent As String, Line, Parts() As String
  FileContent = ReadAllText(MetaDataPath)
  For Each Line In Split(FileContent, vbCrLf)
    Parts = Split(Line, "/")
    If UBound(Parts) = 2 Then Result(Parts(0)) = True
  Next Line

  Set GetAllTrackedFiles = Result
End Function

Private Sub QuickSort(Arr() As Variant, ByVal First As Long, ByVal Last As Long)
  Dim Low As Long, High As Long, Mid As Variant, Temp As Variant
  If UBound(Arr) <= 0 Then Exit Sub
  Low = First: High = Last: Mid = Arr((First + Last) \ 2)
  Do While Low <= High
    Do While Arr(Low) < Mid: Low = Low + 1: Loop
    Do While Arr(High) > Mid: High = High - 1: Loop
    If Low <= High Then
      Temp = Arr(Low): Arr(Low) = Arr(High): Arr(High) = Temp
      Low = Low + 1: High = High - 1
    End If
  Loop
  If First < High Then QuickSort Arr, First, High
  If Low < Last Then QuickSort Arr, Low, Last
End Sub

Private Sub ExportAllSheetsToCSV(ExportedFiles As Collection)
  Dim FilePath As String, LastRow As Long, LastCol As Long
  Dim RowValues() As String, CellValue As String, I As Long, J As Long
  Dim NewCsvContent As String, OldCsvContent As String, WS As Worksheet
  For Each WS In WB.Worksheets
    FilePath = ExportPath & GetWorksheetCsvFileName(WS)
    LastRow = WS.UsedRange.Rows.Count
    LastCol = WS.UsedRange.Columns.Count
    NewCsvContent = ""
    For I = 1 To LastRow
      ReDim RowValues(1 To LastCol)
      For J = 1 To LastCol
        CellValue = "*** Something just went wrong! ***"
        On Error Resume Next
        CellValue = WS.Cells(I, J).Formula
        If Err.Number Then CellValue = WS.Cells(I, J).Value
        On Error GoTo 0
        If InStr(CellValue, """") > 0 Then
          CellValue = Replace(CellValue, """", """""")
        End If
        If InStr(CellValue, vbLf) > 0 Or InStr(CellValue, vbCr) > 0 Or InStr(CellValue, """") > 0 Or InStr(CellValue, ",") > 0 Then
          CellValue = """" & CellValue & """"
        End If
        RowValues(J) = CellValue
      Next J
      NewCsvContent = NewCsvContent & VBA.Join(RowValues, ",") & vbCrLf
    Next I

    If FSO.FileExists(FilePath) Then
      OldCsvContent = ReadAllText(FilePath)
      If NewCsvContent = OldCsvContent Then GoTo NextSheet
    End If

    WriteAllText FilePath, NewCsvContent
    ExportedFiles.Add GetWorksheetCsvFileName(WS)

NextSheet:
  Next WS
End Sub

Private Function GenerateReportLines(Exported As Collection, Imported As Collection, Conflicts As Collection, Deleted As Collection, MetaUpdated As Collection, Optional OnlyModified As Boolean = False) As Collection
  Dim Lines As New Collection, FileStatus As Dictionary, FName, I As Long
  Set FileStatus = New Dictionary

  For Each FName In Exported
    FileStatus(FName) = "** EXPORTED **"
  Next FName

  For Each FName In Imported
    FileStatus(FName) = "** IMPORTED **"
  Next FName

  For Each FName In Conflicts
    FileStatus(FName) = "** CONFLICT **"
  Next FName

  For Each FName In Deleted
    FileStatus(FName) = "** DELETED **"
  Next FName

  For Each FName In MetaUpdated
    FileStatus(FName) = "** METADATA UPDATED (hashes now match, metadata refreshed) **"
  Next FName

  If Not OnlyModified Then
    Dim AllTracked As Dictionary: Set AllTracked = GetAllTrackedFiles
    For Each FName In AllTracked.Keys
      If Not FileStatus.Exists(FName) Then
        FileStatus(FName) = "unchanged"
      End If
    Next FName
  End If

  Dim Keys() As Variant
  Keys = FileStatus.Keys
  Call QuickSort(Keys, 0, UBound(Keys))

  For I = 0 To UBound(Keys)
    Lines.Add Keys(I) & ": " & FileStatus(Keys(I))
  Next I
  
  If Lines.Count = 0 Then Lines.Add "No changes"

  If Not NonAnsiWarnings Is Nothing Then
    If NonAnsiWarnings.Count > 0 Then
      Lines.Add "-----------------------------"
      For I = 1 To NonAnsiWarnings.Count
        Lines.Add NonAnsiWarnings(I)
      Next I
    End If
  End If

  Set GenerateReportLines = Lines
End Function

Private Sub ShowReportSummary(Lines As Collection)
  Dim Msg As String, I As Long, HasConflicts As Boolean, Style As VbMsgBoxStyle
  For I = 1 To Lines.Count
    Msg = Msg & vbCrLf & Lines(I)
    If InStr(Lines(I), ": ** CONFLICT **") > 0 Then HasConflicts = True
  Next I

  If HasConflicts Then
    Msg = Msg & vbCrLf & vbCrLf & _
      "Conflicts were found." & vbCrLf & _
      "Please diff and merge the internal ""-vba"" file with the external one, and run GitSync again."
    Style = vbCritical
  Else
    Style = vbInformation
  End If

  MsgBox Trim(Msg), Style, WB.Name & " Sync Summary"
End Sub

Private Function GetWorksheetCsvFileName(WS As Worksheet) As String
  GetWorksheetCsvFileName = WS.CodeName & " (" & WS.Name & ").csv"
End Function

Private Sub DeleteAllConflictArtifacts(VBProj As VBIDE.VBProject)
  Dim Folder As Scripting.Folder, File As Scripting.File, NameLower As String
  Set Folder = FSO.GetFolder(ExportPath)
  For Each File In Folder.Files
    NameLower = LCase(File.Name)
    If InStr(NameLower, "-vba.") > 0 Then File.Delete
  Next File

  Dim Comp As VBIDE.VBComponent, I As Integer
  For I = VBProj.VBComponents.Count To 1 Step -1
    Set Comp = VBProj.VBComponents(I)
    If Right(Comp.Name, 11) = "__to_delete" Then VBProj.VBComponents.Remove Comp
  Next I
End Sub

Private Sub WarnIfNonAnsi(Text As String, Optional FileName As String = "")
  Dim I As Long, Ch As String
  For I = 1 To Len(Text)
    Ch = Mid(Text, I, 1)
    If AscW(Ch) > 127 Then
      NonAnsiWarnings.Add "Warning: Non-ANSI characters found in module """ & FileName & """:" & vbLf & _
             """..." & Mid(Text, IIf(I > 10, I - 10, 1), 30) & "...""" & vbLf & _
             "Import / export of VBA module supports only ANSI characters."
      Exit Sub
    End If
  Next I
End Sub