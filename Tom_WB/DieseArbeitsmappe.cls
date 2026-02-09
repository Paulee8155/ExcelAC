VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Snapshot automatisch erstellen beim Speichern
    ' Fehler nur loggen, Speichern NICHT blockieren
    On Error Resume Next
    CreateSnapshot
    If Err.Number <> 0 Then
        LogWarning "Auto-Snapshot bei Save fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    ' Beim Öffnen prüfen ob neue Aufträge in der Inbox liegen
    On Error Resume Next

    Dim inboxPath As String
    inboxPath = INBOX_FOLDER & WB_USER & "_Inbox.xlsx"

    If Dir(inboxPath) = "" Then Exit Sub

    ' Inbox öffnen und zählen (ohne Lock, nur lesen)
    Dim wbCheck As Workbook
    Set wbCheck = Workbooks.Open(inboxPath, ReadOnly:=True, UpdateLinks:=False)

    If wbCheck Is Nothing Then Exit Sub

    Dim loCheck As ListObject
    Set loCheck = wbCheck.Sheets(1).ListObjects("tblInbox")

    Dim colFlag As Long, pending As Long, r As Long
    colFlag = 0

    Dim lc As ListColumn
    For Each lc In loCheck.ListColumns
        If LCase$(lc.name) = "importedflag" Then colFlag = lc.Index: Exit For
    Next lc

    If colFlag > 0 Then
        For r = 1 To loCheck.ListRows.Count
            Dim v As Variant
            v = loCheck.ListRows(r).Range.Cells(1, colFlag).Value
            If IsEmpty(v) Or Val(v) = 0 Then pending = pending + 1
        Next r
    End If

    wbCheck.Close SaveChanges:=False

    On Error GoTo 0

    If pending > 0 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Du hast " & pending & " neue Aufträge in deiner Inbox." & vbCrLf & vbCrLf & _
                        "Jetzt importieren?", vbQuestion + vbYesNo, "Neue Aufträge")
        If answer = vbYes Then
            SafeRun "Import"
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' Eigene Locks aufräumen falls noch aktiv
    Dim f As String, fp As String
    f = Dir(LOCK_FOLDER & "*.lock")
    Do While f <> ""
        fp = LOCK_FOLDER & f
        ' Nur eigene Locks löschen (die unser Username enthalten)
        Dim content As String
        Dim fNum As Integer
        fNum = FreeFile
        Open fp For Input As #fNum
        If LOF(fNum) > 0 Then content = Input$(LOF(fNum), fNum) Else content = ""
        Close #fNum
        If InStr(content, Environ$("USERNAME")) > 0 And _
           InStr(content, Environ$("COMPUTERNAME")) > 0 Then
            Kill fp
        End If
        content = ""
        f = Dir()
    Loop
    On Error GoTo 0
End Sub

