VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCreateLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const informationLevel As Integer = 5
Private Const debugLevel As Integer = 4
Private Const warningLevel As Integer = 3
Private Const errorLevel As Integer = 2
Private Const criticalLevel As Integer = 1
Private Const project_name_pattern As String = "^[a-zA-Z0-9_-]{4,255}$"

Private m_log_on As Boolean
Private m_log_file As Integer
Private m_log_path As String
Private m_file_name As String
Private m_log_level As Integer
Private m_set_up As Boolean
Private m_project As String

Public Function setupLog(ByVal logging_on As Boolean, _
                         ByVal project_name As String, _
                         Optional ByVal log_path = "c:\vba-log", _
                         Optional ByVal log_level As String = "information")
    
    Dim reg_ex As Object: Set reg_ex = CreateObject("VBScript.RegExp")
    
    With reg_ex
        .Pattern = project_name_pattern
        If Not .test(project_name) Then
            Err.Raise _
                vbObjectError + 612, _
                Err.Source, _
                "The name of your project is not valid!" & vbNewLine & _
                "You must use a letter as the first character." & vbNewLine & _
                "You can't use a space, period (.), exclamation mark (!), or the characters @, &, $, # in the name." & vbNewLine & _
                "Name can't exceed 255 characters in length."
        End If
    End With
    
    Rem free up space
    Set reg_ex = Nothing
    
    m_project = project_name

    Select Case log_level
        Case "information": m_log_level = 5
        Case "debug": m_log_level = 4
        Case "warning": m_log_level = 3
        Case "error": m_log_level = 2
        Case "critical": m_log_level = 1
        Case Else
            Err.Raise _
                vbObjectError + 610, _
                Err.Source, _
                "You have to set up log level!" & vbNewLine & _
                "setLogLevel argument can be: " & vbNewLine & _
                "information, debug, warning, error, critical" & vbNewLine & _
                log_level & " is not allowed"
    End Select
    
    m_set_up = True
    m_file_name = "\" & m_project & "-" & Format(Now(), "mmddhhmmss") & ".log"
    m_log_path = log_path
    m_log_on = logging_on

End Function

Private Sub openFile()

    Dim file_path As String
    file_path = Dir(m_log_path, vbDirectory)

    If file_path = vbNullString Then MkDir m_log_path

    file_path = m_log_path & m_file_name

    m_log_file = FreeFile
    
    Open file_path For Append As #m_log_file

End Sub

Private Sub closeFile()

    Close m_log_file
    m_log_file = vbNull

End Sub

Private Function checkIfLogable() As Boolean
    checkIfLogable = m_log_on
End Function

Private Function getLogLevel() As Integer
    getLogLevel = m_log_level
End Function

Private Sub writeLog(log_text As String)

    Dim cleaned_up As Boolean
    Dim has_error As Boolean

    Call openFile

    Write #m_log_file, Now() & ": " & log_text
    
    Call closeFile

End Sub

Public Sub logInformation(log_text As String)

    If Not m_set_up Then Exit Sub
    If getLogLevel < informationLevel Then Exit Sub

    Call writeLog("information - " & log_text)

End Sub

Public Sub logDebug(log_text As String)

    If Not m_set_up Then Exit Sub
    If getLogLevel < debugLevel Then Exit Sub

    Call writeLog("debug - " & log_text)

End Sub

Public Sub logWarning(log_text As String)

    If Not m_set_up Then Exit Sub
    If getLogLevel < warningLevel Then Exit Sub

    Call writeLog("warning - " & log_text)

End Sub

Public Sub logError(log_text As String, ByRef error As Object)

    If Not m_set_up Then Exit Sub
    If getLogLevel < errorLevel Then Exit Sub

    Call writeLog("error - " & log_text & vbNewLine & _
        "Error occured: " & vbNewLine & _
        "Description: " & error.Description & vbNewLine & _
        "Source: " & vbNewLine & _
        error.Source)

End Sub

Public Sub logCritical(log_text As String, ByRef error As Object)

    If Not m_set_up Then Exit Sub

    Call writeLog("critical - " & log_text & vbNewLine & _
        "Error occured: " & vbNewLine & _
        "Description: " & error.Description & vbNewLine & _
        "Source: " & vbNewLine & _
        error.Source)

End Sub


