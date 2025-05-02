Option Explicit

Private Const ExportFolder = "C:\workspace\VbaGitSync"
Private SavingForTheSecondTime As Boolean

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Debug.Print "Workbook_BeforeSave", SavingForTheSecondTime
  If SavingForTheSecondTime Then Exit Sub

  Cancel = True

  If SaveAsUI Then
    ' This is required because of the Excel bug mentioned below
    MsgBox "Save As is not supported by GitSync", vbCritical, Title:="GitSync"
    Exit Sub
  End If

  GitSync.GitSync ThisWorkbook, ExportFolder

  ' Excel bug: GitSync uses UsedRange, and since accessing UsedRange in this
  ' event will cause Excel not save, we need to save again
  SavingForTheSecondTime = True
  ThisWorkbook.Save
  SavingForTheSecondTime = False
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  If Not Success Then Exit Sub
  Debug.Print "Workbook_AfterSave", SavingForTheSecondTime

  MsgBox ThisWorkbook.Name & " saved", vbInformation, "GitSync"

  Dim FSO As New Scripting.FileSystemObject
  FSO.CopyFile ThisWorkbook.FullName, ExportFolder & "\" & ThisWorkbook.Name
End Sub