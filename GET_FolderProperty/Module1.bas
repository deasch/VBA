Attribute VB_Name = "Module1"
'Option Explicit

Sub GetFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Your Custom FileDialog Title"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Canceled"
        Else
            MsgBox .SelectedItems(1)
        End If
    End With
End Sub
