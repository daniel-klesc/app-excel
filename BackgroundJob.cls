VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BackgroundJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public STR_MODULE As String

Public str_id As String

Public str_full_path As String
Public str_macro As String

Public bool_is_readonly As Boolean
Public bool_is_visible As Boolean

Public Function run()
    Dim wb As Workbook
        
    On Error GoTo ERR_LOG_OPEN
    hndl_log.open_data
    On Error GoTo 0
    
    On Error GoTo ERR_JOB_FAILED
    ' open job
    Set wb = util_file.open_wb(str_full_path, bool_is_readonly, bool_is_visible, app.str_password)
    ' run job
    Application.run "'" & wb.Name & "'!" & str_macro
    ' close job
    Windows(wb.Name).Visible = True
    wb.Close SaveChanges:=False
    ' log
    hndl_log.log db_log.TYPE_INFO, STR_MODULE, "run", "Last run>" & Now
    On Error GoTo 0
        
        
    On Error GoTo ERR_LOG_CLOSE
    hndl_log.close_data
    On Error GoTo 0
    
    Exit Function
ERR_JOB_FAILED:
    hndl_log.log db_log.TYPE_ERR, STR_MODULE, "run", "Background job has failed. " & Err.Number & ">" & Err.description
    hndl_log.close_data
    Exit Function
ERR_LOG_OPEN:
    MsgBox "Log error during opening", vbCritical, "Logging error"
    Exit Function
ERR_LOG_CLOSE:
    MsgBox "Log error during closing", vbCritical, "Logging error"
    Exit Function
End Function


Function before_run()
    Application.DisplayAlerts = False
End Function

Function after_run()
    Application.DisplayAlerts = True
End Function

Private Sub Class_Initialize()
    STR_MODULE = "background_job"
End Sub
