Attribute VB_Name = "modInboxCleanup"
Option Explicit

' ======= TUNING =======
Private Const KEEP_DAYS As Long = 14        ' done-Zeilen so lange in Inbox lassen
Private Const MAX_DONE_ROWS As Long = 2000  ' Notbremse: wenn mehr done drin sind -> zusätzlich purge

' Ruft man am Ende von ImportFromInbox auf (während Inbox-Workbook noch offen ist!)
Public Sub ArchiveAndPurgeInboxDoneRows(ByVal wbInbox As Workbook, ByVal userName As String)
    On Error GoTo ErrorHandler

    Dim loInbox As ListObject
    Set loInbox = wbInbox.Sheets(1).ListObjects("tblInbox")

    Dim colFlag As Long, colAt As Long
    colFlag = GetColIndexSafe(loInbox, "ImportedFlag")
    colAt = GetColIndexSafe(loInbox, "ImportedAt")

    If colFlag = 0 Then Exit Sub ' ohne Flag keine sichere Logik
    ' ImportedAt ist optional, wir können notfalls auch ohne arbeiten

    ' 1) Kandidaten bestimmen: done (=1) und (zu alt oder zu viele)
    Dim cutoff As Date
    cutoff = DateAdd("d", -KEEP_DAYS, Date)

    Dim doneCount As Long
    doneCount = CountDoneRows(loInbox, colFlag)

    Dim needPurgeByCount As Boolean
    needPurgeByCount = (doneCount > MAX_DONE_ROWS)

    ' Sammle zu archivierende Zeilen (Indexliste)
    Dim idx() As Long, n As Long
    n = 0

    Dim r As Long
    For r = 1 To loInbox.ListRows.Count
        Dim vFlag As Variant
        vFlag = loInbox.ListRows(r).Range.Cells(1, colFlag).Value

        If vFlag = 1 Then
            Dim shouldArchive As Boolean
            shouldArchive = False

            ' Alter prüfen, wenn ImportedAt vorhanden
            If colAt > 0 Then
                Dim vAt As Variant
                vAt = loInbox.ListRows(r).Range.Cells(1, colAt).Value
                If IsDate(vAt) Then
                    If DateValue(CDate(vAt)) <= cutoff Then
                        shouldArchive = True
                    End If
                Else
                    ' Kein Datum drin -> wenn Notbremse aktiv ist, darf er auch raus
                    If needPurgeByCount Then shouldArchive = True
                End If
            Else
                ' Kein ImportedAt Feld -> nur Notbremse
                If needPurgeByCount Then shouldArchive = True
            End If

            If shouldArchive Then
                n = n + 1
                ReDim Preserve idx(1 To n)
                idx(n) = r
            End If
        End If
    Next r

    If n = 0 Then Exit Sub ' nichts zu tun

    ' 2) Archivdatei (monatlich) öffnen/erstellen
    Dim archPath As String
    archPath = ARCHIVE_FOLDER & userName & "_InboxArchive_" & Format(Date, "yyyymm") & ".xlsx"

    Dim wbArch As Workbook, wsArch As Worksheet, loArch As ListObject
    Set wbArch = OpenOrCreateArchive(archPath, loInbox)

    Set wsArch = wbArch.Sheets(1)
    Set loArch = wsArch.ListObjects(1) ' einzige Tabelle

    ' 3) Zeilen ins Archiv kopieren (als Werte)
    Dim k As Long
    For k = 1 To n
        Dim srcRow As ListRow
        Set srcRow = loInbox.ListRows(idx(k))

        Dim newRow As ListRow
        Set newRow = loArch.ListRows.Add

        newRow.Range.Value = srcRow.Range.Value
    Next k

    wbArch.Close SaveChanges:=True

    ' 4) Inbox-Zeilen löschen (WICHTIG: von unten nach oben!)
    For k = n To 1 Step -1
        loInbox.ListRows(idx(k)).Delete
    Next k

    ' 5) Save (Inbox)
    wbInbox.Save

    LogInfo "Inbox cleanup: archived&purged " & n & " done rows for " & userName
    Exit Sub

ErrorHandler:
    LogError "ArchiveAndPurgeInboxDoneRows failed: " & Err.Description
End Sub

' ===== Helpers =====

Private Function GetColIndexSafe(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColIndexSafe = lo.ListColumns(colName).Index
    If Err.Number <> 0 Then GetColIndexSafe = 0
    Err.Clear
    On Error GoTo 0
End Function

Private Function CountDoneRows(ByVal lo As ListObject, ByVal colFlag As Long) As Long
    Dim r As Long, c As Long
    c = 0
    For r = 1 To lo.ListRows.Count
        If lo.ListRows(r).Range.Cells(1, colFlag).Value = 1 Then c = c + 1
    Next r
    CountDoneRows = c
End Function

Private Function OpenOrCreateArchive(ByVal archPath As String, ByVal loInbox As ListObject) As Workbook
    Dim wb As Workbook

    If Dir(archPath) <> "" Then
        Set wb = Workbooks.Open(archPath, ReadOnly:=False, UpdateLinks:=False)
        Set OpenOrCreateArchive = wb
        Exit Function
    End If

    ' Neu anlegen mit gleichen Überschriften
    Dim ws As Worksheet, lo As ListObject
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)
    ws.Name = "Archive"

    loInbox.HeaderRowRange.Copy
    ws.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = "tblArchive"

    Application.DisplayAlerts = False
    wb.SaveAs fileName:=archPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    Set OpenOrCreateArchive = wb
End Function

