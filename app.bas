Attribute VB_Name = "app"
Option Explicit

Public bool_run_on_open As Boolean
Public str_name As String
Public str_password As String

Public col_background_jobs As Collection

Public Function init()
    bool_run_on_open = False
    Set col_background_jobs = New Collection
    
    ' log setup
    hndl_log.init
        
    ' settings
    hndl_local_setting.init
    ctrl_settings.init
End Function

Public Function run_jobs()
    Dim obj_background_job As BackgroundJob
    
    For Each obj_background_job In col_background_jobs
        On Error GoTo WARN_BACKGROUND_JOB_FAILED
        obj_background_job.before_run
        obj_background_job.run
        obj_background_job.after_run
        On Error GoTo 0
    Next
    Exit Function
WARN_BACKGROUND_JOB_FAILED:
    hndl_log.log db_log.TYPE_WARN, STR_MODULE, "run_jobs", "Background job: " & obj_background_job.str_id & " has failed."
    Resume Next
End Function

Public Function add_job(obj_background_job As BackgroundJob)
    col_background_jobs.Add obj_background_job, obj_background_job.str_id
End Function
