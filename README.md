# GitSync for VBA

**GitSync for VBA** is an Excel add-in that automates the synchronization between VBA code in macro-enabled workbooks or add-ins and plain text files stored in a local folder (typically a Git repository). It enables source control for VBA development by exporting and importing code automatically on save, making it easier to track changes and collaborate.

---

## Features

- 🔁 **Automatic Export/Import**\
  When a macro workbook is saved:

  - The VBA code is automatically exported to a folder.
  - If the code in the exported files has changed, it is automatically imported back into the workbook.
  - If both workbook and files have changed, a `-vba` suffix file is exported to help you perform a diff manually.

- ⚖️ **Diff Support**\
  When a conflict is detected (both sides changed), GitSync generates a `ModuleName-vba.bas` file with the current VBA code so you can visually compare and decide whether to:

  - Delete the `.bas`/`.cls` file and save again to export the VBA code, or
  - Update the file to match the code in the workbook.
  - Replace the code in the workbook to match the file.

- 📌 **Dual Import Strategy**\
  GitSync uses two different methods to import modules, based on their type:

  - For **worksheet, workbook and form modules**, GitSync uses `CodeModule.InsertLines`, since these components cannot be imported with `.bas` or `.cls` files.
  - For **standard and class modules**, GitSync uses `VBComponents.Import` to load the entire file. This method preserves module-level attributes, such as:
    ```vba
    Attribute Item.VB_UserMemId = 0
    ```
    These attributes are commonly used in class modules and are lost if lines are inserted manually. Since `Import` cannot be used on document modules and `InsertLines` loses metadata, both techniques are used where appropriate.

- ⚠️ **Unicode Handling**\
  Unicode characters are not preserved across export/import operations. A warning is shown if such characters are detected.

- 🥾 **Silent Module Deletion Workaround**\
  Sometimes removing a VBA component fails silently due to an Excel bug. GitSync works around this by:

  - Renaming the module temporarily.
  - Attempting deletion again.
  - If it still fails, deletion is retried on the next run.

- ❓ **Sync Confirmation Outside Repository**\
  When saving a workbook, GitSync checks for the presence of a `.GitSync` file in the target folder. If the file does not exist (indicating the folder is not a VBA repository), you will be prompted to confirm whether to perform the sync. This prevents accidental exports when saving copies of workbooks outside your repository.


---

## How to Use

### 1. Install the Add-In

Load the `VbaGitSync.xlam` add-in into your development Excel instance.

### 2. Add This to Your Workbook's `ThisWorkbook` Module

Basic setup to export VBA modules to the same folder as the workbook:

```vba
Private IsGitSyncAvailable As Boolean

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  On Error Resume Next
  IsGitSyncAvailable = Not AddIns("VbaGitSync") Is Nothing
  On Error GoTo 0
  If IsGitSyncAvailable Then Application.Run "VbaGitSync.xlam!GitSync.GitSync", ThisWorkbook
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  If Not IsGitSyncAvailable Then Exit Sub
  If Not Success Then
    MsgBox "The save failed"
  Else
    MsgBox ThisWorkbook.Name & " saved"
  End If
End Sub
```

### 3. (Optional) Save to a Different Folder

To export code to a separate location:

```vba
Private Const ExportFolder = "C:\Workspace\MyAddIn"

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  [...]
  If IsGitSyncAvailable Then Application.Run "VbaGitSync.xlam!GitSync.GitSync", ThisWorkbook, ExportFolder
End Sub
```

---

## Excel Save Bug Workaround

In some cases, executing `UsedRange` during `Workbook_BeforeSave` causes Excel to silently fail to save the workbook. This workaround ensures reliable saving:

```vba
Private Const ExportFolder = "C:\Workspace\MyAddIn"
Private SavingForTheSecondTime As Boolean

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Debug.Print "Workbook_BeforeSave", SavingForTheSecondTime
  If SavingForTheSecondTime Then Exit Sub

  Cancel = True

  If SaveAsUI Then
    MsgBox "Save As is not supported by GitSync"
    Exit Sub
  End If

  CleanupBeforeSave

  Dim A As AddIn
  On Error Resume Next
  Set A = AddIns("VbaGitSync")
  On Error GoTo 0
  If Not A Is Nothing Then Application.Run "VbaGitSync.xlam!GitSync.GitSync", ThisWorkbook, ExportFolder

  SavingForTheSecondTime = True
  ThisWorkbook.Save
  SavingForTheSecondTime = False
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  If Not Success Then Exit Sub
  Debug.Print "Workbook_AfterSave", SavingForTheSecondTime
  MsgBox ThisWorkbook.Name & " saved"
End Sub
```

---

## License

This project is licensed under the MIT License.  
See the [LICENSE](LICENSE) file for details.
