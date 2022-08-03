Attribute VB_Name = "mdlGlobal"
Option Explicit

Public Const THIS_PROJECT_NAME As String = "SSOP-Log"

Public log As clsCreateLog
Public Sub setUplog(ByVal is_log_on As Boolean, ByVal project As String)
    Call log.setUplog(is_log_on, project, log_level:="debug")
End Sub
