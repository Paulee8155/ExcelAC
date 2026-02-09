Attribute VB_Name = "modConfig"
Option Explicit

Public Const WB_USER As String = "MASTER"
' Pfade
Public Const BASE_PATH As String = "H:\01_Workbenches\"
Public Const INBOX_FOLDER As String = BASE_PATH & "_INBOX\"
Public Const LOCK_FOLDER As String = INBOX_FOLDER & "_locks\"
Public Const LOG_FOLDER As String = BASE_PATH & "_LOG\"
Public Const ARCHIVE_FOLDER As String = BASE_PATH & "_ARCHIVE\"
Public Const SNAP_FOLDER As String = BASE_PATH & "_SNAP\"
Public Const SNAP_LOCK_FOLDER As String = SNAP_FOLDER & "_locks\"
Public Const CLAIM_FOLDER As String = BASE_PATH & "_CLAIM\"
Public Const CLAIM_LOCK_FOLDER As String = CLAIM_FOLDER & "_locks\"


' Lock-Parameter (MASTER)
Public Const LOCK_TIMEOUT_MS As Long = 15000
Public Const LOCK_RETRY_INTERVAL_MS As Long = 500
Public Const LOCK_STALE_MINUTES As Double = 10
Public Const LOCK_VERIFY_DELAY_MS As Long = 200
Public Const LOCK_VERIFY_RETRIES As Integer = 3


