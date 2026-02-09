Attribute VB_Name = "modClaim"
Option Explicit

Private Function SanitizeKey(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[A-Za-z0-9_-]" Then r = r & c
    Next i
    SanitizeKey = r
End Function

Private Function ClaimFilePath(ByVal einsatzNr As String) As String
    ClaimFilePath = CLAIM_FOLDER & SanitizeKey(einsatzNr) & ".claim"
End Function

Private Function ClaimLockPath(ByVal einsatzNr As String) As String
    ClaimLockPath = CLAIM_LOCK_FOLDER & SanitizeKey(einsatzNr) & ".lock"
End Function

Public Function Claim_GetOwner(ByVal einsatzNr As String) As String
    On Error GoTo SafeExit

    Dim fp As String, txt As String, p As Long, lineEnd As Long
    fp = ClaimFilePath(einsatzNr)
    If Dir(fp) = "" Then GoTo SafeExit

    txt = ReadAllText(fp)
    p = InStr(1, txt, "Owner=", vbTextCompare)
    If p = 0 Then GoTo SafeExit

    p = p + Len("Owner=")
    lineEnd = InStr(p, txt, vbCrLf)
    If lineEnd = 0 Then lineEnd = Len(txt) + 1

    Claim_GetOwner = Trim$(Mid$(txt, p, lineEnd - p))
    Exit Function

SafeExit:
    Claim_GetOwner = ""
End Function

' =============================================
' INTERN: Claim-Datei schreiben (OHNE Lock!)
' Darf NUR aufgerufen werden wenn Lock BEREITS gehalten wird
' =============================================
Private Function WriteClaimDirect(ByVal einsatzNr As String, ByVal newOwner As String, _
                                  ByVal reason As String, ByVal changedBy As String) As Boolean
    On Error GoTo Fail

    Dim fp As String
    fp = ClaimFilePath(einsatzNr)

    Dim content As String
    content = "Owner=" & newOwner & vbCrLf & _
              "ChangedAt=" & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "ChangedBy=" & changedBy & vbCrLf & _
              "Reason=" & reason & vbCrLf

    WriteClaimDirect = WriteTextFileSafe(fp, content)
    Exit Function

Fail:
    LogError "WriteClaimDirect failed for " & einsatzNr & ": " & Err.Description
    WriteClaimDirect = False
End Function

' =============================================
' PUBLIC: Owner setzen (holt Lock selbst)
' Für Aufrufe von AUSSEN (Workflow, Import etc.)
' =============================================
Public Function Claim_SetOwner(ByVal einsatzNr As String, ByVal newOwner As String, _
                               ByVal reason As String, Optional ByVal changedBy As String = "") As Boolean
    EnsureBaseFolders

    Dim lp As String
    lp = ClaimLockPath(einsatzNr)

    If changedBy = "" Then changedBy = WB_USER

    If Not AcquireLock(lp, "Claim_Set") Then
        LogWarning "Claim_SetOwner: could not lock claim for " & einsatzNr
        Claim_SetOwner = False
        Exit Function
    End If

    Claim_SetOwner = WriteClaimDirect(einsatzNr, newOwner, reason, changedBy)

    ReleaseLock lp
End Function

' =============================================
' PUBLIC: Owner transferieren (Lock + Prüfung + Schreiben in einem)
' Kein Deadlock mehr: nutzt WriteClaimDirect statt Claim_SetOwner
' =============================================
Public Function Claim_Transfer(ByVal einsatzNr As String, ByVal fromOwner As String, ByVal toOwner As String, _
                               ByVal reason As String, Optional ByVal changedBy As String = "") As Boolean
    EnsureBaseFolders

    Dim lp As String
    lp = ClaimLockPath(einsatzNr)

    If changedBy = "" Then changedBy = WB_USER

    If Not AcquireLock(lp, "Claim_Transfer") Then
        LogWarning "Claim_Transfer: could not lock claim for " & einsatzNr
        Claim_Transfer = False
        Exit Function
    End If

    On Error GoTo Fail

    Dim cur As String
    cur = Claim_GetOwner(einsatzNr)

    ' Migration/Legacy: kein Claim vorhanden -> adoptiere auf fromOwner
    If cur = "" Then
        WriteClaimDirect einsatzNr, fromOwner, "Adopt_NoClaim", changedBy
        cur = fromOwner
    End If

    If StrComp(cur, fromOwner, vbTextCompare) <> 0 Then
        Claim_Transfer = False
        GoTo Clean
    End If

    ' Direkt schreiben (Lock wird bereits gehalten ? kein Deadlock)
    Claim_Transfer = WriteClaimDirect(einsatzNr, toOwner, reason, changedBy)

Clean:
    ReleaseLock lp
    Exit Function

Fail:
    LogError "Claim_Transfer failed for " & einsatzNr & ": " & Err.Description
    Claim_Transfer = False
    Resume Clean
End Function

' =============================================
' Hilfs-I/O
' =============================================
Private Function ReadAllText(ByVal filePath As String) As String
    Dim f As Integer
    f = FreeFile
    Open filePath For Input As #f
    If LOF(f) > 0 Then
        ReadAllText = Input$(LOF(f), f)
    Else
        ReadAllText = ""
    End If
    Close #f
End Function

Private Function WriteTextFileSafe(ByVal targetPath As String, ByVal content As String) As Boolean
    Dim tmpPath As String, bakPath As String
    tmpPath = targetPath & ".tmp." & Format(Now, "yyyymmddhhnnss")
    bakPath = targetPath & ".bak"

    On Error GoTo Fail

    Dim f As Integer
    f = FreeFile
    Open tmpPath For Output As #f
    Print #f, content
    Close #f

    On Error Resume Next
    If Dir(bakPath) <> "" Then Kill bakPath
    On Error GoTo Fail

    If Dir(targetPath) <> "" Then
        Name targetPath As bakPath
    End If

    Name tmpPath As targetPath

    On Error Resume Next
    If Dir(bakPath) <> "" Then Kill bakPath
    On Error GoTo 0

    WriteTextFileSafe = True
    Exit Function

Fail:
    On Error Resume Next
    If Dir(tmpPath) <> "" Then Kill tmpPath
    WriteTextFileSafe = False
End Function

