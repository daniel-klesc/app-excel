Attribute VB_Name = "ctrl_settings"
Option Explicit

Public Const STR_SETTING_FILE_DEFAULT As String = "default"
Public Const STR_SETTING_FILE_LOCAL As String = "local"

Public obj_settings As Settings

Public Function init()
    Dim str_setting_file As String
    Dim obj_background_job As BackgroundJob

    ' load settings
    Set obj_settings = New Settings
    obj_settings.init
    
    On Error GoTo ERR_LOCAL_SETTING
    str_setting_file = hndl_local_setting.get_value("setting.file")
    If Not str_setting_file = STR_SETTING_FILE_DEFAULT Then
        If str_setting_file = STR_SETTING_FILE_LOCAL Then
            obj_settings.STR_PATH = ""
        Else
            obj_settings.STR_PATH = str_setting_file
        End If
    End If
    On Error GoTo 0
    
    On Error GoTo ERR_OPEN_SETTINGS
    obj_settings.open_data
    On Error GoTo 0

    On Error GoTo ERR_INVALID_SETTING
    ' app settings
    app.bool_run_on_open = obj_settings.Item("local:app\\app.bool_run_on_open").Value
    app.str_name = obj_settings.Item("local:app\\app.str_name").Value
    app.str_password = obj_settings.Item("local:app\\app.str_password").Value
    
    ' hndl_log settings
    hndl_log.STR_PATH = obj_settings.Item("local:file\\hndl_log.str_path").Value
    hndl_log.STR_FILE_NAME = obj_settings.Item("local:file\\hndl_log.str_file_name.log").Value
    'hndl_log.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = CBool(obj_settings.Item("local:app\\hndl_log.bool_external_data_file_visibility").Value)
    
    ' background job
      ' first
    Set obj_background_job = New BackgroundJob
    obj_background_job.str_id = "1"
    obj_background_job.str_full_path = obj_settings.Item("local:file\\background_job.1.str_full_path").Value
    obj_background_job.str_macro = obj_settings.Item("local:app\\background_job.1.str_macro").Value
    app.add_job obj_background_job
    ' second
    Set obj_background_job = New BackgroundJob
    obj_background_job.str_id = "2"
    obj_background_job.str_full_path = obj_settings.Item("local:file\\background_job.2.str_full_path").Value
    obj_background_job.str_macro = obj_settings.Item("local:app\\background_job.2.str_macro").Value
    app.add_job obj_background_job
    
    On Error GoTo 0
    
    On Error GoTo ERR_CLOSE_SETTINGS
    obj_settings.close_data
    On Error GoTo 0
    
    Exit Function
ERR_LOCAL_SETTING:
    MsgBox "Error in local setting. Settings file not found.", vbCritical, "Application Initiation -> Loading settings"
    Exit Function
ERR_OPEN_SETTINGS:
    MsgBox Err.description, vbCritical, "Application Initiation -> Loading settings"
    Exit Function
ERR_INVALID_SETTING:
    MsgBox "Invalid setting", vbCritical, "Application Initiation -> Loading settings"
    Exit Function
ERR_CLOSE_SETTINGS:
    MsgBox "An error occured during closing settings file. Processing of history was terminated.", vbCritical, "Application Initiation -> Loading settings"
    Exit Function
End Function
