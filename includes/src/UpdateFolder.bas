Attribute VB_Name = "UpdateFolder"
Sub SelectExportFolder()
    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
            .Title = "Select a Folder"
            .AllowMultiSelect = False
            .InitialFileName = ActiveWorkbook.Path
            If .Show <> -1 Then GoTo NextCode
            sItem = .SelectedItems(1)
        End With
NextCode:
        Range("ExportFolder_Path").Value = sItem
        Set fldr = Nothing
End Sub
