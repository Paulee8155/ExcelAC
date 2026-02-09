Attribute VB_Name = "modMasterRefresh"
Option Explicit

Public Sub RefreshDashboardSafe()
    ' 1) Alte Locks wegräumen (falls mal einer hängen blieb)
    CleanupStaleSnapLocks

    ' 2) Prüfen ob noch aktive Locks da sind
    Dim lockFile As String
    lockFile = Dir(SNAP_LOCK_FOLDER & "*.lock")

    If lockFile <> "" Then
        MsgBox "Snapshots sind gerade aktiv (" & lockFile & ")." & vbCrLf & _
               "Bitte später erneut versuchen.", vbInformation
        Exit Sub
    End If

    ' 3) Refresh ausführen
    Application.ScreenUpdating = False
    Application.StatusBar = "Dashboard wird aktualisiert..."

    ThisWorkbook.RefreshAll
    DoEvents
    Application.CalculateUntilAsyncQueriesDone

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' 4) LastRefresh setzen (optional)
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Range("LastRefresh").Value = Now
    On Error GoTo 0

    MsgBox "Dashboard aktualisiert.", vbInformation
End Sub


Private Sub CleanupStaleSnapLocks()
    Dim f As String, fp As String
    Dim ageMin As Double

    f = Dir(SNAP_LOCK_FOLDER & "*.lock")
    Do While f <> ""
        fp = SNAP_LOCK_FOLDER & f

        On Error Resume Next
        ageMin = (Now - FileDateTime(fp)) * 1440#
        On Error GoTo 0

        ' Nutzt deine LOCK_STALE_MINUTES aus modConfig, falls vorhanden.
        ' Wenn du sie im MASTER nicht hast, setz hier z.B. 10.
        If ageMin > 10 Then
            On Error Resume Next
            Kill fp
            On Error GoTo 0
        End If

        f = Dir()
    Loop
End Sub


