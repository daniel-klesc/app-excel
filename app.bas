Attribute VB_Name = "app"
Option Explicit

Public bool_run_on_open As Boolean
Public str_name As String
Public str_password As String

Public Function init()
    bool_run_on_open = False
    
    ' log setup
    hndl_log.init
    'hndl_log.STR_PATH = "C:\Users\czDanKle\Desktop\KLD\under-construction\wh-map\app\background_job\log\"
    'hndl_log.STR_FILE_NAME = "log.xlsx"
        
    hndl_local_setting.init
    ctrl_settings.init
End Function
