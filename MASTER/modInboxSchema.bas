Attribute VB_Name = "modInboxSchema"
Option Explicit

' Sorgt dafür, dass tblInbox alle benötigten Workflow-Spalten hat.
' Wichtig: Wir hängen nur hinten an, damit die ersten 15 EPOS-Spalten unangetastet bleiben.
Public Sub EnsureInboxSchema(ByVal loInbox As ListObject)
    Dim needed As Variant
    needed = Array( _
        "Info", "RNG Datum", "Status", "Klaerfall", _
        "BearbeitetVon", "BearbeitetAm", "KontrolliertVon", "KontrolliertAm" _
    )

    Dim i As Long, c As String
    For i = LBound(needed) To UBound(needed)
        c = CStr(needed(i))
        If GetColIndexSafe(loInbox, c) = 0 Then
            loInbox.ListColumns.Add
            loInbox.ListColumns(loInbox.ListColumns.Count).Name = c
        End If
    Next i
End Sub

' Entfernt leere Tabellenzeilen (z.B. nach "Inhalte löschen").
Public Sub CompactTableByKey(ByVal lo As ListObject, ByVal keyColName As String)
    Dim cKey As Long
    cKey = GetColIndexSafe(lo, keyColName)
    If cKey = 0 Then Exit Sub

    Dim r As Long
    For r = lo.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(lo.ListRows(r).Range.Cells(1, cKey).Value))) = 0 Then
            lo.ListRows(r).Delete
        End If
    Next r
End Sub

Private Function GetColIndexSafe(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColIndexSafe = lo.ListColumns(colName).Index
    If Err.Number <> 0 Then GetColIndexSafe = 0
    Err.Clear
    On Error GoTo 0
End Function


