Attribute VB_Name = "modLock"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Lock-Parameter (lokal in modLock)
Private Const LOCK_TIMEOUT_MS As Long = 15000
Private Const LOCK_RETRY_INTERVAL_MS As Long = 500
Private Const LOCK_STALE_MINUTES As Double = 10
Private Const LOCK_VERIFY_DELAY_MS As Long = 200
Private Const LOCK_VERIFY_RETRIES As Integer = 3

' Multi-lock tracking (pro Excel-Session)
Private g_LockGuids As Object ' Scripting.Dictionary: key=lockPath, value=GUID

Private Sub EnsureLockDict()
    If g_LockGuids Is Nothing Then
        Set g_LockGuids = CreateObject("Scripting.Dictionary")
    End If
End Sub

Public Function AcquireLock(ByVal lockPath As String, ByVal Action As String) As Boolean
    EnsureBaseFolders
    EnsureLockDict

    Dim tmpPath As String
    Dim RetryCount As Long
    Dim MaxRetries As Long
    Dim LockAge As Double
    Dim GUID As String
    Dim VerifyContent As String
    Dim MyLockContent As String
    Dim VerifyAttempt As Integer

    MaxRetries = LOCK_TIMEOUT_MS \ LOCK_RETRY_INTERVAL_MS

    GUID = Format(Now, "yyyymmddhhnnss") & "_" & _
           SanitizeString(Environ$("USERNAME")) & "_" & _
           SanitizeString(Environ$("COMPUTERNAME")) & "_" & _
           Format(Timer * 1000, "0")

    tmpPath = Replace(lockPath, ".lock", ".tmp." & GUID)

    MyLockContent = Environ$("USERNAME") & "|" & Environ$("COMPUTERNAME") & vbCrLf & _
                    Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                    Action & vbCrLf & _
                    GUID

    For RetryCount = 1 To MaxRetries
        On Error Resume Next
        Err.Clear

        Dim fNum As Integer
        fNum = FreeFile
        Open tmpPath For Output As #fNum
        Print #fNum, MyLockContent
        Close #fNum

        If Err.Number <> 0 Then
            LogWarning "Lock tmp creation failed: " & Err.Description
            Err.Clear
            Sleep LOCK_RETRY_INTERVAL_MS
            GoTo NextIteration
        End If

        Name tmpPath As lockPath

        If Err.Number = 0 Then
            On Error GoTo 0
            Sleep LOCK_VERIFY_DELAY_MS

            For VerifyAttempt = 1 To LOCK_VERIFY_RETRIES
                VerifyContent = ReadFileContent(lockPath)
                If VerifyContent <> "" Then Exit For
                If VerifyAttempt < LOCK_VERIFY_RETRIES Then Sleep LOCK_VERIFY_DELAY_MS
            Next VerifyAttempt

            If InStr(VerifyContent, GUID) > 0 Then
                g_LockGuids(lockPath) = GUID
                LogInfo "Lock acquired: " & lockPath & " | GUID: " & GUID
                AcquireLock = True
                Exit Function
            Else
                LogWarning "Lock race lost. My GUID: " & GUID
            End If
        Else
            Err.Clear
            On Error GoTo 0

            If Dir(lockPath) <> "" Then
                LockAge = (Now - FileDateTime(lockPath)) * 1440#
                If LockAge > LOCK_STALE_MINUTES Then
                    LogWarning "Removing stale lock (age: " & Format(LockAge, "0.0") & " min): " & lockPath
                    On Error Resume Next
                    Kill lockPath
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        End If

        On Error Resume Next
        If Dir(tmpPath) <> "" Then Kill tmpPath
        Err.Clear
        On Error GoTo 0

NextIteration:
        Sleep LOCK_RETRY_INTERVAL_MS
    Next RetryCount

    LogError "Lock timeout after " & (LOCK_TIMEOUT_MS / 1000) & "s: " & lockPath
    On Error Resume Next
    If Dir(tmpPath) <> "" Then Kill tmpPath
    On Error GoTo 0
    AcquireLock = False
End Function

Public Sub ReleaseLock(ByVal lockPath As String)
    EnsureLockDict

    Dim LockContent As String
    Dim myGuid As String
    On Error Resume Next

    If Dir(lockPath) = "" Then
        LogWarning "ReleaseLock: Lock file does not exist: " & lockPath
        Exit Sub
    End If

    myGuid = ""
    If g_LockGuids.Exists(lockPath) Then myGuid = CStr(g_LockGuids(lockPath))

    LockContent = ReadFileContent(lockPath)

    If myGuid <> "" And InStr(LockContent, myGuid) > 0 Then
        Kill lockPath
        If Err.Number = 0 Then
            LogInfo "Lock released: " & lockPath
        Else
            LogWarning "Lock release failed: " & Err.Description
        End If
    Else
        LogWarning "ReleaseLock: NOT my lock or GUID unknown: " & lockPath
    End If

    On Error Resume Next
    If g_LockGuids.Exists(lockPath) Then g_LockGuids.Remove lockPath
    On Error GoTo 0
End Sub

Private Function ReadFileContent(ByVal filePath As String) As String
    Dim fNum As Integer
    Dim content As String
    On Error Resume Next

    If Dir(filePath) = "" Then
        ReadFileContent = ""
        Exit Function
    End If

    fNum = FreeFile
    Open filePath For Input As #fNum
    If LOF(fNum) > 0 Then content = Input$(LOF(fNum), fNum) Else content = ""
    Close #fNum

    If Err.Number <> 0 Then content = "": Err.Clear
    On Error GoTo 0
    ReadFileContent = content
End Function

Private Function SanitizeString(ByVal InputStr As String) As String
    Dim Result As String, i As Long, c As String
    For i = 1 To Len(InputStr)
        c = Mid$(InputStr, i, 1)
        If c Like "[A-Za-z0-9_-]" Then Result = Result & c
    Next i
    SanitizeString = Result
End Function


