Attribute VB_Name = "modSnapshot"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const SNAP_MAX_RETRIES As Integer = 3
Private Const SNAP_RETRY_BASE_MS As Long = 500

Private g_SnapshotRunning As Boolean

Public Sub CreateSnapshotButton()
    If CreateSnapshot() Then
        MsgBox "Snapshot erstellt.", vbInformation
    Else
        MsgBox "Snapshot NICHT erstellt (siehe Log).", vbExclamation
    End If
End Sub

Public Function CreateSnapshot() As Boolean
    Dim userName As String
    Dim lockPath As String
    Dim tmpPath As String
    Dim targetPath As String
    Dim timestampPath As String

    If g_SnapshotRunning Then
        LogInfo "Snapshot already running - skipped"
        CreateSnapshot = True
        Exit Function
    End If
    g_SnapshotRunning = True

    On Error GoTo ErrorHandler

    userName = WB_USER ' Workbench-ID (Tom/Maria/Paul)
    lockPath = SNAP_LOCK_FOLDER & userName & "_SNAP.lock"
    tmpPath = SNAP_FOLDER & userName & "_SNAP_tmp.xlsx"
    targetPath = SNAP_FOLDER & userName & "_SNAP.xlsx"
    timestampPath = SNAP_FOLDER & userName & "_SNAP.timestamp"

    If Not AcquireLock(lockPath, "Snap_Create") Then
        LogWarning "Snapshot skipped: Could not acquire lock"
        CreateSnapshot = False
        GoTo Cleanup
    End If

    ' tmp löschen (best effort)
    On Error Resume Next
    If Dir(tmpPath) <> "" Then
        SetAttr tmpPath, vbNormal
        Kill tmpPath
    End If
    On Error GoTo ErrorHandler

    Dim wbSnap As Workbook
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim loSource As ListObject
    Dim rngTarget As Range

    Set wsSource = ThisWorkbook.Sheets("Aufträge")
    Set loSource = wsSource.ListObjects("tblJobs")

    Set wbSnap = Workbooks.Add(xlWBATWorksheet)
    Set wsTarget = wbSnap.Sheets(1)
    wsTarget.name = "Aufträge"

    ' Header (Werte)
    loSource.HeaderRowRange.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteValues

    ' Daten (Werte)
    If loSource.ListRows.Count > 0 Then
        loSource.DataBodyRange.Copy
        wsTarget.Range("A2").PasteSpecial xlPasteValues
    End If
    Application.CutCopyMode = False

    Set rngTarget = wsTarget.Range("A1").CurrentRegion
    wsTarget.ListObjects.Add xlSrcRange, rngTarget, , xlYes
    wsTarget.ListObjects(1).name = "tblJobs"

    Application.DisplayAlerts = False
    wbSnap.SaveAs Filename:=tmpPath, FileFormat:=xlOpenXMLWorkbook
    wbSnap.Close SaveChanges:=False
    Application.DisplayAlerts = True

    If ReplaceSnapshotSafe(tmpPath, targetPath, timestampPath) Then
        LogInfo "Snapshot created: " & targetPath
        CreateSnapshot = True
    Else
        LogError "Snapshot replace failed"
        CreateSnapshot = False
    End If

Cleanup:
    On Error Resume Next
    ReleaseLock lockPath
    Application.DisplayAlerts = True
    On Error GoTo 0
    g_SnapshotRunning = False
    Exit Function

ErrorHandler:
    LogError "Snapshot creation failed: " & Err.Description
    On Error Resume Next
    If Dir(tmpPath) <> "" Then
        SetAttr tmpPath, vbNormal
        Kill tmpPath
    End If
    ReleaseLock lockPath
    Application.DisplayAlerts = True
    On Error GoTo 0
    g_SnapshotRunning = False
    CreateSnapshot = False
End Function

Private Function ReplaceSnapshotSafe(ByVal tmpPath As String, ByVal targetPath As String, ByVal timestampPath As String) As Boolean
    Dim bakBase As String
    Dim bakPath As String
    Dim retry As Integer
    Dim hadExisting As Boolean

    bakBase = targetPath & ".bak"

    For retry = 1 To SNAP_MAX_RETRIES
        On Error Resume Next
        Err.Clear

        hadExisting = (Dir(targetPath) <> "")

        ' 1) bestehendes Target -> bak
        If hadExisting Then
            ' Target ggf. schreibbar machen (falls ReadOnly gesetzt ist)
            MakeWritable targetPath

            ' .bak löschen (falls vorhanden)
            bakPath = bakBase
            If Dir(bakPath) <> "" Then
                MakeWritable bakPath
                Kill bakPath
                Err.Clear
            End If

            ' Wenn .bak immer noch da (z.B. offen) -> eindeutigen bak-Namen nehmen
            If Dir(bakPath) <> "" Then
                bakPath = bakBase & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(retry)
            End If

            ' Rename target -> bak
            Name targetPath As bakPath
            If Err.Number <> 0 Then
                LogWarning "Snapshot rename to .bak failed (retry " & retry & "): " & Err.Description
                Err.Clear
                Sleep SNAP_RETRY_BASE_MS * retry
                GoTo NextRetry
            End If
        Else
            bakPath = "" ' kein rollback nötig
        End If

        ' 2) tmp -> target
        Name tmpPath As targetPath
        If Err.Number <> 0 Then
            LogWarning "Snapshot tmp->target failed (retry " & retry & "): " & Err.Description

            ' rollback
            If hadExisting And bakPath <> "" And Dir(bakPath) <> "" Then
                Err.Clear
                MakeWritable bakPath
                Name bakPath As targetPath
                If Err.Number = 0 Then
                    LogInfo "Snapshot rollback successful"
                Else
                    LogError "CRITICAL: Snapshot rollback failed! " & Err.Description
                End If
            End If

            Err.Clear
            Sleep SNAP_RETRY_BASE_MS * retry
            GoTo NextRetry
        End If

        On Error GoTo 0

        ' 3) timestamp schreiben (auch wenn ReadOnly)
        On Error Resume Next
        If Dir(timestampPath) <> "" Then MakeWritable timestampPath
        On Error GoTo 0

        Dim fNum As Integer
        fNum = FreeFile
        Open timestampPath For Output As #fNum
        Print #fNum, Format(Now, "yyyy-mm-dd hh:nn:ss")
        Close #fNum

        ' 4) Ziel wieder ReadOnly setzen (für PQ)
        On Error Resume Next
        SetAttr targetPath, vbReadOnly
        SetAttr timestampPath, vbNormal ' timestamp darf normal bleiben; wenn du willst: vbReadOnly
        On Error GoTo 0

        ' 5) best-effort bak cleanup
        On Error Resume Next
        If bakPath <> "" And Dir(bakPath) <> "" Then
            MakeWritable bakPath
            Kill bakPath
        End If
        On Error GoTo 0

        ReplaceSnapshotSafe = True
        Exit Function

NextRetry:
        On Error GoTo 0
    Next retry

    ' wenn alles scheitert: tmp aufräumen
    On Error Resume Next
    If Dir(tmpPath) <> "" Then
        MakeWritable tmpPath
        Kill tmpPath
    End If
    On Error GoTo 0

    ReplaceSnapshotSafe = False
End Function

Private Sub MakeWritable(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 Then
        If Dir(filePath) <> "" Then SetAttr filePath, vbNormal
    End If
    On Error GoTo 0
End Sub

