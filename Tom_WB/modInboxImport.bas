Attribute VB_Name = "modInboxImport"
Option Explicit

' =========================
' IMPORT PARAMETER
' =========================
Private Const INBOX_BATCH_SAVE_INTERVAL As Long = 10
Private Const INBOX_IMPORTED_FLAG_INPROGRESS As Long = 2
Private Const INBOX_IMPORTED_FLAG_DONE As Long = 1
Private Const RECOVERY_MINUTES As Double = 5

' =========================
' MAIN
' =========================
Public Sub ImportFromInbox()
    Dim userName As String
    Dim inboxPath As String
    Dim lockPath As String

    Dim wbInbox As Workbook
    Dim loInbox As ListObject
    Dim loJobs As ListObject

    Dim colFlag As Long, colAt As Long, colBy As Long
    Dim colEinsatzInbox As Long, colEinsatzJobs As Long

    Dim importCount As Long, skippedCount As Long, failedCount As Long
    Dim batchCounter As Long

    Dim dictJobs As Object

    On Error GoTo ErrorHandler

    EnsureBaseFolders

    userName = WB_USER
    inboxPath = INBOX_FOLDER & userName & "_Inbox.xlsx"
    lockPath = LOCK_FOLDER & userName & "_Inbox.lock"

    If Dir(inboxPath) = "" Then
        MsgBox "Inbox-Datei nicht gefunden: " & vbCrLf & inboxPath, vbExclamation
        Exit Sub
    End If

    If Not AcquireLock(lockPath, "Inbox_Import") Then
        MsgBox "Inbox ist gerade belegt. Bitte gleich nochmal versuchen.", vbInformation
        Exit Sub
    End If

    Set wbInbox = Workbooks.Open(inboxPath, ReadOnly:=False, UpdateLinks:=False)

    If wbInbox.ReadOnly Then
        wbInbox.Close SaveChanges:=False
        ReleaseLock lockPath
        MsgBox "Inbox ist schreibgeschützt (evtl. offen).", vbExclamation
        Exit Sub
    End If

    Set loInbox = wbInbox.Sheets(1).ListObjects("tblInbox")
    EnsureTableSchema loInbox, GetInboxColumnSchema()
    RemoveEmptyRows loInbox, "EinsatzNr"

    Set loJobs = ThisWorkbook.Sheets("Aufträge").ListObjects("tblJobs")

    ' Pflicht-Spalten Inbox
    colFlag = GetColIdx(loInbox, "ImportedFlag")
    colAt = GetColIdx(loInbox, "ImportedAt")
    colBy = GetColIdx(loInbox, "ImportedBy")

    If colFlag = 0 Or colAt = 0 Or colBy = 0 Then
        wbInbox.Close SaveChanges:=False
        ReleaseLock lockPath
        MsgBox "In tblInbox fehlen: ImportedFlag / ImportedAt / ImportedBy", vbCritical
        Exit Sub
    End If

    ' Key-Spalten
    colEinsatzInbox = GetColIdx(loInbox, "EinsatzNr")
    colEinsatzJobs = GetColIdx(loJobs, "EinsatzNr")

    If colEinsatzInbox = 0 Or colEinsatzJobs = 0 Then
        wbInbox.Close SaveChanges:=False
        ReleaseLock lockPath
        MsgBox "Spalte 'EinsatzNr' fehlt (Inbox oder Workbench).", vbCritical
        Exit Sub
    End If

    ' Crash-Recovery (Flag=2 zu alt -> zurück auf 0)
    CleanupInProgressEntries loInbox, colFlag, colAt, colBy
    wbInbox.Save

    ' Dictionary der bereits vorhandenen EinsatzNr in der Workbench (Duplikat-Schutz)
    Set dictJobs = BuildKeyDict(loJobs, colEinsatzJobs)

    ' =============================================
    ' PERFORMANCE: erst HIER aktivieren
    ' (nach allen Validierungs-Checks / Exit Sub)
    ' =============================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim i As Long
    batchCounter = 0

    For i = 1 To loInbox.ListRows.Count
        Dim rw As ListRow
        Set rw = loInbox.ListRows(i)

        Dim flagValue As Variant
        flagValue = rw.Range.Cells(1, colFlag).Value

        If IsEmpty(flagValue) Or Val(flagValue) = 0 Then
            On Error GoTo RowErrorHandler

            Dim einsatzID As String
            einsatzID = Trim$(CStr(rw.Range.Cells(1, colEinsatzInbox).Value))

            If Len(einsatzID) = 0 Then
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            ' =========================
            ' PHASE 1: IN PROGRESS
            ' =========================
            rw.Range.Cells(1, colFlag).Value = INBOX_IMPORTED_FLAG_INPROGRESS
            rw.Range.Cells(1, colAt).Value = Now
            rw.Range.Cells(1, colBy).Value = userName

            batchCounter = batchCounter + 1
            If batchCounter Mod INBOX_BATCH_SAVE_INTERVAL = 0 Then
                wbInbox.Save
            End If

            ' =========================
            ' PHASE 2: DUP CHECK (WB)
            ' =========================
            If dictJobs.Exists(einsatzID) Then
                rw.Range.Cells(1, colFlag).Value = INBOX_IMPORTED_FLAG_DONE
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            ' =========================
            ' PHASE 3: IMPORT -> tblJobs
            ' =========================
            Dim newRow As ListRow
            Set newRow = loJobs.ListRows.Add

            CopyRowByHeaders loInbox, rw, loJobs, newRow

            ' im Dict nachziehen
            dictJobs(einsatzID) = True

            ' DONE
            rw.Range.Cells(1, colFlag).Value = INBOX_IMPORTED_FLAG_DONE
            importCount = importCount + 1

            GoTo NextRow

RowErrorHandler:
            failedCount = failedCount + 1
            LogError "Import Row failed (EinsatzNr=" & einsatzID & "): " & Err.Description

            On Error Resume Next
            rw.Range.Cells(1, colFlag).Value = 0
            rw.Range.Cells(1, colAt).Value = ""
            rw.Range.Cells(1, colBy).Value = ""
            On Error GoTo 0
            Err.Clear

NextRow:
            On Error GoTo ErrorHandler
        End If
    Next i

    ' Final speichern
    wbInbox.Save

    ' Inbox-Cleanup (Archiv + Purge DONE nach deinen Regeln)
    ArchiveAndPurgeInboxDoneRows wbInbox, userName

    ' Geisterzeilen entfernen
    RemoveEmptyRows loInbox, "EinsatzNr"

    wbInbox.Save
    wbInbox.Close SaveChanges:=False

    ReleaseLock lockPath

    ' PERFORMANCE: wiederherstellen
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox importCount & " importiert" & vbCrLf & _
           skippedCount & " übersprungen" & vbCrLf & _
           failedCount & " Fehler", _
           IIf(failedCount > 0, vbExclamation, vbInformation)

    LogInfo "Inbox Import done: " & importCount & " imported, " & skippedCount & " skipped, " & failedCount & " failed"
    Exit Sub

ErrorHandler:
    LogError "ImportFromInbox failed: " & Err.Description
    On Error Resume Next
    If Not wbInbox Is Nothing Then wbInbox.Close SaveChanges:=False
    ReleaseLock lockPath
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Fehler beim Import: " & Err.Description, vbCritical
End Sub

' =========================
' RECOVERY
' =========================
Private Sub CleanupInProgressEntries(ByVal loInbox As ListObject, _
                                    ByVal colFlag As Long, ByVal colAt As Long, ByVal colBy As Long)
    Dim r As Long
    For r = 1 To loInbox.ListRows.Count
        Dim vFlag As Variant
        vFlag = loInbox.ListRows(r).Range.Cells(1, colFlag).Value

        If Val(vFlag) = INBOX_IMPORTED_FLAG_INPROGRESS Then
            Dim vAt As Variant
            vAt = loInbox.ListRows(r).Range.Cells(1, colAt).Value

            Dim ageMin As Double
            If IsDate(vAt) Then
                ageMin = (Now - CDate(vAt)) * 1440#
            Else
                ageMin = 999
            End If

            If ageMin > RECOVERY_MINUTES Then
                loInbox.ListRows(r).Range.Cells(1, colFlag).Value = 0
                loInbox.ListRows(r).Range.Cells(1, colAt).Value = ""
                loInbox.ListRows(r).Range.Cells(1, colBy).Value = ""
            End If
        End If
    Next r
End Sub

' =========================
' DICT (Duplikat-Schutz)
' =========================
Private Function BuildKeyDict(ByVal lo As ListObject, ByVal colKey As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    If lo.ListRows.Count = 0 Then
        Set BuildKeyDict = d
        Exit Function
    End If

    Dim r As Long, k As String
    For r = 1 To lo.ListRows.Count
        k = Trim$(CStr(lo.ListRows(r).Range.Cells(1, colKey).Value))
        If Len(k) > 0 Then d(k) = True
    Next r

    Set BuildKeyDict = d
End Function

