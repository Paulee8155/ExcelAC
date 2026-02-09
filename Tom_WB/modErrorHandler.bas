Attribute VB_Name = "modErrorHandler"
Option Explicit

' Zentrale Fehlerbehandlung für alle Aktions-Subs
' Verhindert, dass der VBA-Debugger aufpoppt

Public Sub SafeRun(ByVal actionName As String)
    On Error GoTo ErrHandler

    Select Case LCase$(actionName)
        Case "import":              ImportFromInbox
        Case "klaerfall":           ToggleKlaerfall
        Case "rngfehlt":            MarkRngSubFehlt
        Case "zurkontrolle":        ZurKontrolleMarkieren
        Case "kontrolleok":         KontrolleOK_Verschieben
        Case "weitergabe_maria":    WeitergebenAuswahlAnMaria
        Case "weitergabe_paul":     WeitergebenAuswahlAnPaul
        Case "weitergabe_tom":      WeitergebenAuswahlAnTom
        Case "snapshot":            CreateSnapshotButton
        Case "filter_kontrolle":    ZurKontrolleFilternUndAuswaehlen
        Case "filter_reset":        FilterZuruecksetzen
        Case "sort":                SortTblJobsStandard
        Case Else
            MsgBox "Unbekannte Aktion: " & actionName, vbExclamation
    End Select
    Exit Sub

ErrHandler:
    Dim msg As String
    msg = FriendlyError(actionName, Err.Number, Err.Description)
    MsgBox msg, vbExclamation, "Fehler"
    LogError "SafeRun(" & actionName & ") failed: [" & Err.Number & "] " & Err.Description
End Sub

Private Function FriendlyError(ByVal action As String, ByVal errNum As Long, ByVal errDesc As String) As String
    ' Bekannte Fehler übersetzen
    Select Case errNum
        Case vbObjectError + 100, vbObjectError + 120
            FriendlyError = "Die Tabelle 'tblJobs' hat keine Daten." & vbCrLf & _
                            "Bitte zuerst Aufträge importieren."

        Case vbObjectError + 101, vbObjectError + 121
            FriendlyError = "Bitte wähle zuerst eine oder mehrere Zeilen in der Tabelle aus."

        Case vbObjectError + 200
            FriendlyError = "In der Tabelle fehlt eine benötigte Spalte." & vbCrLf & _
                            "Bitte Paul kontaktieren."

        Case Else
            FriendlyError = "Bei '" & action & "' ist ein Fehler aufgetreten." & vbCrLf & vbCrLf & _
                            "Bitte schließe die Datei und öffne sie neu." & vbCrLf & _
                            "Falls der Fehler weiterhin auftritt, melde dich bei Paul."
    End Select
End Function

' =====================================================
' BUTTON-MAKROS (eines pro Button)
' Diese werden den Form-Steuerelementen zugewiesen
' =====================================================
Public Sub Btn_Import():             SafeRun "Import":             End Sub
Public Sub Btn_Klaerfall():          SafeRun "Klaerfall":          End Sub
Public Sub Btn_RngFehlt():           SafeRun "RngFehlt":           End Sub
Public Sub Btn_ZurKontrolle():       SafeRun "ZurKontrolle":       End Sub
Public Sub Btn_KontrolleOK():        SafeRun "KontrolleOK":        End Sub
Public Sub Btn_WeitergabeMaria():    SafeRun "Weitergabe_Maria":   End Sub
Public Sub Btn_WeitergabePaul():     SafeRun "Weitergabe_Paul":    End Sub
Public Sub Btn_WeitergabeTom():      SafeRun "Weitergabe_Tom":     End Sub
Public Sub Btn_FilterKontrolle():    SafeRun "Filter_Kontrolle":   End Sub
Public Sub Btn_FilterReset():        SafeRun "Filter_Reset":       End Sub
