Attribute VB_Name = "modFSO"
Option Explicit

Public Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then EnsureFolder parentPath
    End If

    On Error Resume Next
    fso.CreateFolder folderPath
    On Error GoTo 0
End Sub

Public Sub EnsureBaseFolders()
    EnsureFolder BASE_PATH
    EnsureFolder INBOX_FOLDER
    EnsureFolder LOCK_FOLDER
    EnsureFolder LOG_FOLDER
    EnsureFolder ARCHIVE_FOLDER
    EnsureFolder SNAP_FOLDER
    EnsureFolder SNAP_LOCK_FOLDER
    EnsureFolder CLAIM_FOLDER
    EnsureFolder CLAIM_LOCK_FOLDER
End Sub
