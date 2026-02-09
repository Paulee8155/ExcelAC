Attribute VB_Name = "modLogging"
Option Explicit

Private g_LastRotationCheck As Date

Public Sub LogInfo(ByVal Message As String):    WriteLog "INFO", Message: End Sub
Public Sub LogWarning(ByVal Message As String): WriteLog "WARNING", Message: End Sub
Public Sub LogError(ByVal Message As String):   WriteLog "ERROR", Message: End Sub

Private Sub WriteLog(ByVal Level As String, ByVal Message As String)
    On Error Resume Next

    EnsureBaseFolders

    Dim LogPath As String
    LogPath = LOG_FOLDER & Environ$("USERNAME") & "_Log.csv"

    ' NEU: Header erstellen wenn Datei noch nicht existiert
    If Dir(LogPath) = "" Then
        Dim f0 As Integer
        f0 = FreeFile
        Open LogPath For Output As #f0
        Print #f0, "Timestamp,Level,Message"
        Close #f0
    End If

    If Now - g_LastRotationCheck > 1 Then
        RotateLogIfNeeded LogPath
        g_LastRotationCheck = Now
    End If

    Dim Timestamp As String
    Timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim EscapedMessage As String
    EscapedMessage = Replace(Message, """", """""")
    EscapedMessage = Replace(EscapedMessage, vbCrLf, " | ")
    EscapedMessage = Replace(EscapedMessage, vbCr, " | ")
    EscapedMessage = Replace(EscapedMessage, vbLf, " | ")

    Dim fNum As Integer
    fNum = FreeFile
    Open LogPath For Append As #fNum
    Print #fNum, Timestamp & "," & Level & ",""" & EscapedMessage & """"
    Close #fNum

    On Error GoTo 0
End Sub

Private Sub RotateLogIfNeeded(ByVal LogPath As String)
    On Error Resume Next
    If Dir(LogPath) = "" Then Exit Sub

    Dim FileSizeKB As Long
    FileSizeKB = FileLen(LogPath) \ 1024

    If FileSizeKB > LOG_MAX_SIZE_KB Then
        Dim ArchivePath As String
        ArchivePath = Replace(LogPath, "_Log.csv", "_Log_" & Format(Now, "yyyymmdd_hhnnss") & ".csv")
        Name LogPath As ArchivePath

        If Err.Number = 0 Then
            Dim fNum As Integer
            fNum = FreeFile
            Open LogPath For Output As #fNum
            Print #fNum, "Timestamp,Level,Message"
            Print #fNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & ",INFO,""Log rotated"""
            Close #fNum
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub


