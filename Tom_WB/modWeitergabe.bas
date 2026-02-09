Attribute VB_Name = "modWeitergabe"
Option Explicit

' Buttons
Public Sub WeitergebenAuswahlAnMaria(): WeitergebenAuswahlMehrfach "Maria": End Sub
Public Sub WeitergebenAuswahlAnPaul():  WeitergebenAuswahlMehrfach "Paul":  End Sub
Public Sub WeitergebenAuswahlAnTom():   WeitergebenAuswahlMehrfach "Tom":   End Sub

Public Sub WeitergebenAuswahlMehrfach(ByVal TargetUser As String)
    Dim loJobs As ListObject
    Dim rngSel As Range

    Set loJobs = ThisWorkbook.Sheets("Aufträge").ListObjects("tblJobs")
    If loJobs.DataBodyRange Is Nothing Then
        MsgBox "Keine Aufträge vorhanden.", vbInformation
        Exit Sub
    End If

    Set rngSel = Intersect(Selection, loJobs.DataBodyRange)
    If rngSel Is Nothing Then
        MsgBox "Bitte Zellen IN tblJobs markieren (mehrere Zeilen möglich).", vbInformation
        Exit Sub
    End If

    ' Zielpfade
    Dim TargetInboxPath As String, TargetLockPath As String
    TargetInboxPath = INBOX_FOLDER & TargetUser & "_Inbox.xlsx"
    TargetLockPath = LOCK_FOLDER & TargetUser & "_Inbox.lock"

    If Dir(TargetInboxPath) = "" Then
        MsgBox "Ziel-Inbox nicht gefunden:" & vbCrLf & TargetInboxPath, vbExclamation
        Exit Sub
    End If

    If Not AcquireLock(TargetLockPath, "Inbox_Write_Batch") Then
        MsgBox "Ziel-Inbox ist belegt. Bitte später erneut versuchen.", vbInformation
        Exit Sub
    End If

    On Error GoTo Fail

    ' =============================================
    ' PERFORMANCE: Screen + Calculation aus
    ' =============================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wbInbox As Workbook
    Dim wasAlreadyOpen As Boolean
    Set wbInbox = GetWorkbookIfOpen(TargetInboxPath)
    wasAlreadyOpen = Not (wbInbox Is Nothing)
    If wbInbox Is Nothing Then
        Set wbInbox = Workbooks.Open(TargetInboxPath, ReadOnly:=False, UpdateLinks:=False)
    End If

    If wbInbox.ReadOnly Then
        If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
        ReleaseLock TargetLockPath
        GoTo CleanupUI
    End If

    Dim loInbox As ListObject
    Set loInbox = wbInbox.Sheets(1).ListObjects("tblInbox")

    EnsureTableSchema loInbox, GetInboxColumnSchema()
    RemoveEmptyRows loInbox, "EinsatzNr"

    ' =============================================
    ' PERFORMANCE: Alle Spalten-Indices 1x cachen
    ' =============================================
    Dim colEinsatzJobs As Long, colEinsatzInbox As Long
    Dim colFlag As Long, colAt As Long, colBy As Long
    Dim colInfo As Long

    colEinsatzJobs = GetColIdx(loJobs, "EinsatzNr")
    colEinsatzInbox = GetColIdx(loInbox, "EinsatzNr")
    colFlag = GetColIdx(loInbox, "ImportedFlag")
    colAt = GetColIdx(loInbox, "ImportedAt")
    colBy = GetColIdx(loInbox, "ImportedBy")
    colInfo = GetColIdx(loJobs, "Info")

    If colEinsatzJobs = 0 Or colEinsatzInbox = 0 Then
        If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
        ReleaseLock TargetLockPath
        MsgBox "Spalte 'EinsatzNr' fehlt (Jobs oder Inbox).", vbCritical
        GoTo CleanupUI
    End If

    ' =============================================
    ' PERFORMANCE: Spalten-Mapping 1x vorberechnen
    ' =============================================
    Dim colMap() As Long
    Dim mapCount As Long
    BuildColumnMap loJobs, loInbox, colMap, mapCount

    ' =============================================
    ' PERFORMANCE: Dictionary statt Loop für Duplikat-Check
    ' =============================================
    Dim dictInbox As Object
    Set dictInbox = BuildEinsatzDict(loInbox, colEinsatzInbox)

    ' Einmalige Zeilen-Menge aus Selection (pro Zeile, nicht pro Zelle)
    Dim dictRows As Object: Set dictRows = CreateObject("Scripting.Dictionary")
    Dim c As Range
    For Each c In rngSel.Cells
        If Not dictRows.Exists(c.Row) Then dictRows(c.Row) = True
    Next c

    Dim okCnt As Long, dupCnt As Long, failCnt As Long, blockedCnt As Long
    Dim key As Variant

    ' Dictionary für erfolgreich weitergegebene Zeilen (für Abgegeben-Verschiebung)
    Dim dictSuccess As Object: Set dictSuccess = CreateObject("Scripting.Dictionary")

    For Each key In dictRows.Keys
        Dim idx As Long
        idx = CLng(key) - loJobs.DataBodyRange.Row + 1
        If idx < 1 Or idx > loJobs.ListRows.Count Then
            failCnt = failCnt + 1
            GoTo NextOne
        End If

        Dim srcRow As ListRow
        Set srcRow = loJobs.ListRows(idx)

        Dim einsatzID As String
        einsatzID = Trim$(CStr(srcRow.Range.Cells(1, colEinsatzJobs).Value))
        If Len(einsatzID) = 0 Then
            failCnt = failCnt + 1
            GoTo NextOne
        End If

        ' =============================================
        ' SCHRITT 1: CLAIM OWNER CHECK
        ' =============================================
        Dim owner As String
        owner = Claim_GetOwner(einsatzID)

        If owner = "" Then
            Call Claim_SetOwner(einsatzID, WB_USER, "Adopt_OnWeitergabe", WB_USER)
            owner = WB_USER
        End If

        If StrComp(owner, WB_USER, vbTextCompare) <> 0 Then
            blockedCnt = blockedCnt + 1
            If colInfo > 0 Then
                MarkiereHinweis srcRow, colInfo, "NICHT weitergegeben: Owner=" & owner
            End If
            GoTo NextOne
        End If

        ' =============================================
        ' SCHRITT 2: DUPLIKAT-CHECK
        ' =============================================
        If dictInbox.Exists(einsatzID) Then
            dupCnt = dupCnt + 1
            GoTo NextOne
        End If

        ' =============================================
        ' SCHRITT 3: CLAIM ZUERST TRANSFERIEREN
        ' (atomar sicher – nur wer Claim bekommt, darf schreiben)
        ' =============================================
        If Not Claim_Transfer(einsatzID, WB_USER, TargetUser, "Weitergabe", WB_USER) Then
            blockedCnt = blockedCnt + 1
            If colInfo > 0 Then
                MarkiereHinweis srcRow, colInfo, "Claim-Transfer fehlgeschlagen: " & einsatzID
            End If
            GoTo NextOne
        End If

        ' =============================================
        ' SCHRITT 4: NUR WENN CLAIM OK ? IN INBOX SCHREIBEN
        ' =============================================
        Dim newRow As ListRow
        Set newRow = loInbox.ListRows.Add

        CopyRowFast srcRow, newRow, colMap, mapCount

        If colFlag > 0 Then newRow.Range.Cells(1, colFlag).Value = 0
        If colAt > 0 Then newRow.Range.Cells(1, colAt).Value = ""
        If colBy > 0 Then newRow.Range.Cells(1, colBy).Value = ""

        ' Dict nachziehen
        dictInbox(einsatzID) = True

        okCnt = okCnt + 1

        ' Erfolg merken für Abgegeben-Verschiebung
        dictSuccess(key) = True

        ' Info markieren
        If colInfo > 0 Then
            MarkiereWeitergabeInInfo srcRow, colInfo, TargetUser
        End If

NextOne:
    Next key

    ' =============================================
    ' Inbox speichern + schließen (ZUERST! Daten sicher)
    ' =============================================
    wbInbox.Save
    If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
    ReleaseLock TargetLockPath

    ' =============================================
    ' DANN: Erfolgreiche Zeilen ? Blatt "Abgegeben"
    ' =============================================
    If dictSuccess.Count > 0 Then
        Dim loAbg As ListObject
        Set loAbg = EnsureAbgegebenTable(loJobs)

        ' Indices sammeln
        Dim arrIdx() As Long, nIdx As Long
        nIdx = 0
        Dim dk As Variant
        For Each dk In dictSuccess.Keys
            nIdx = nIdx + 1
            ReDim Preserve arrIdx(1 To nIdx)
            arrIdx(nIdx) = CLng(dk) - loJobs.DataBodyRange.Row + 1
        Next dk

        ' Absteigend sortieren (damit Delete die Indices nicht verschiebt)
        Dim ii As Long, jj As Long, tmp As Long
        For ii = 1 To nIdx - 1
            For jj = ii + 1 To nIdx
                If arrIdx(jj) > arrIdx(ii) Then
                    tmp = arrIdx(ii): arrIdx(ii) = arrIdx(jj): arrIdx(jj) = tmp
                End If
            Next jj
        Next ii

        ' Kopieren + Löschen (von unten nach oben)
        For ii = 1 To nIdx
            If arrIdx(ii) >= 1 And arrIdx(ii) <= loJobs.ListRows.Count Then
                Dim abgRow As ListRow
                Set abgRow = loAbg.ListRows.Add
                abgRow.Range.Value = loJobs.ListRows(arrIdx(ii)).Range.Value
                loJobs.ListRows(arrIdx(ii)).Delete
            End If
        Next ii
    End If

CleanupUI:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If okCnt + dupCnt + blockedCnt + failCnt > 0 Then
        MsgBox "Weitergabe an " & TargetUser & ":" & vbCrLf & _
               okCnt & " geschrieben + aus tblJobs entfernt" & vbCrLf & _
               dupCnt & " bereits in Ziel-Inbox" & vbCrLf & _
               blockedCnt & " blockiert (nicht Owner / Claim fehlgeschlagen)" & vbCrLf & _
               failCnt & " fehlerhaft/leer", vbInformation
    End If
    Exit Sub

Fail:
    On Error Resume Next
    If Not wbInbox Is Nothing Then
        If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
    End If
    ReleaseLock TargetLockPath
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler bei Weitergabe: " & Err.Description, vbCritical
End Sub

' =========================
' ABGEGEBEN-Tabelle sicherstellen
' =========================
Private Function EnsureAbgegebenTable(ByVal loJobs As ListObject) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Abgegeben")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "Abgegeben"
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("tblAbgegeben")
    On Error GoTo 0

    If lo Is Nothing Then
        loJobs.HeaderRowRange.Copy
        ws.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False

        Dim rng As Range
        Set rng = ws.Range("A1").Resize(1, loJobs.ListColumns.Count)
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.name = "tblAbgegeben"
    End If

    Set EnsureAbgegebenTable = lo
End Function

' =========================
' PERFORMANCE: Spalten-Mapping einmal berechnen
' =========================
Private Sub BuildColumnMap(ByVal srcTable As ListObject, ByVal destTable As ListObject, _
                           ByRef colMap() As Long, ByRef mapCount As Long)
    Dim cols As Variant
    cols = GetFullColumnSchema()

    Dim maxCols As Long
    maxCols = UBound(cols) - LBound(cols) + 1
    ReDim colMap(1 To maxCols, 0 To 1)
    mapCount = 0

    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        Dim srcCol As Long, destCol As Long
        srcCol = GetColIdx(srcTable, CStr(cols(i)))
        destCol = GetColIdx(destTable, CStr(cols(i)))

        If srcCol > 0 And destCol > 0 Then
            mapCount = mapCount + 1
            colMap(mapCount, 0) = srcCol
            colMap(mapCount, 1) = destCol
        End If
    Next i
End Sub

' =========================
' PERFORMANCE: Kopie mit gecachtem Mapping
' =========================
Private Sub CopyRowFast(ByVal srcRow As ListRow, ByVal destRow As ListRow, _
                        ByRef colMap() As Long, ByVal mapCount As Long)
    Dim i As Long
    For i = 1 To mapCount
        destRow.Range.Cells(1, colMap(i, 1)).Value = srcRow.Range.Cells(1, colMap(i, 0)).Value
    Next i
End Sub

' =========================
' PERFORMANCE: Dictionary aus Inbox bauen
' NUR pending Einträge (Flag=0) zählen als Duplikat
' Flag=1 (bereits importiert) darf erneut gesendet werden
' =========================
Private Function BuildEinsatzDict(ByVal lo As ListObject, ByVal colEinsatz As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    If lo.ListRows.Count = 0 Then
        Set BuildEinsatzDict = d
        Exit Function
    End If

    ' ImportedFlag-Spalte suchen
    Dim colFlag As Long
    colFlag = GetColIdx(lo, "ImportedFlag")

    Dim r As Long, k As String
    For r = 1 To lo.ListRows.Count
        k = Trim$(CStr(lo.ListRows(r).Range.Cells(1, colEinsatz).Value))
        If Len(k) > 0 Then
            If colFlag > 0 Then
                ' Nur Flag=0 (pending) als Duplikat werten
                Dim flagVal As Variant
                flagVal = lo.ListRows(r).Range.Cells(1, colFlag).Value
                If IsEmpty(flagVal) Or Val(flagVal) = 0 Then
                    d(k) = True
                End If
                ' Flag=1 (done) ? NICHT ins Dict ? darf erneut gesendet werden
            Else
                ' Kein Flag-Spalte ? sicherheitshalber alles als Duplikat
                d(k) = True
            End If
        End If
    Next r

    Set BuildEinsatzDict = d
End Function

' =========================
' MARK INFO
' =========================
Private Sub MarkiereWeitergabeInInfo(ByVal srcRow As ListRow, ByVal colInfo As Long, ByVal TargetUser As String)
    Dim oldInfo As String
    oldInfo = Trim$(CStr(srcRow.Range.Cells(1, colInfo).Value))

    Dim stamp As String
    stamp = "Weitergegeben an " & TargetUser & " (" & Format(Now, "dd.mm.yyyy hh:nn") & ")"

    If Len(oldInfo) = 0 Then
        srcRow.Range.Cells(1, colInfo).Value = stamp
    Else
        srcRow.Range.Cells(1, colInfo).Value = oldInfo & " | " & stamp
    End If
End Sub

Private Sub MarkiereHinweis(ByVal srcRow As ListRow, ByVal colInfo As Long, ByVal text As String)
    Dim oldInfo As String
    oldInfo = Trim$(CStr(srcRow.Range.Cells(1, colInfo).Value))
    If Len(oldInfo) = 0 Then
        srcRow.Range.Cells(1, colInfo).Value = text
    Else
        srcRow.Range.Cells(1, colInfo).Value = oldInfo & " | " & text
    End If
End Sub

' =========================
' HELPERS
' =========================
Private Function GetWorkbookIfOpen(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = LCase$(fullPath) Then
            Set GetWorkbookIfOpen = wb
            Exit Function
        End If
    Next wb
    Set GetWorkbookIfOpen = Nothing
End Function

