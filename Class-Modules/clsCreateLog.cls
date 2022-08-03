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

Private Const DEBUG_LEVEL As Integer = 5
Private Const INFORMATION_LEVEL As Integer = 4
Private Const WARNING_LEVEL As Integer = 3
Private Const ERROR_LEVEL As Integer = 2
Private Const CRITICAL_LEVEL As Integer = 1

Private Const PROJECT_NAME_PATTERN As String = "^[a-zA-Z0-9_-]{4,255}$"

Private m_has_error As Boolean
Private m_file_path As String

Private m_log_on As Boolean
Private m_log_file As Integer
Private m_log_path As String
Private m_file_name As String
Private m_log_level As Integer
Private m_set_up As Boolean
Private m_project As String

Public Function setUplog(ByVal log_on As Boolean, _
                         ByVal project_name As String, _
                         Optional ByVal log_path = "c:\vba-log", _
                         Optional ByVal log_level As String = "information")
    
    Dim reg_ex As Object: Set reg_ex = CreateObject("VBScript.RegExp")
    
    With reg_ex
        .Pattern = PROJECT_NAME_PATTERN
        If Not .test(project_name) Then
            m_has_error = True
            Err.Raise _
                vbObjectError + 600, _
                m_project, _
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
        Case "information": m_log_level = INFORMATION_LEVEL
        Case "debug": m_log_level = DEBUG_LEVEL
        Case "warning": m_log_level = WARNING_LEVEL
        Case "error": m_log_level = ERROR_LEVEL
        Case "critical": m_log_level = CRITICAL_LEVEL
        Case Else
            m_has_error = True
            Err.Raise _
                vbObjectError + 610, _
                m_project, _
                "You have to set up log level!" & vbNewLine & _
                "log_level argument can be: " & vbNewLine & _
                "information, debug, warning, error, critical" & vbNewLine & _
                log_level & " is not allowed"
    End Select
    
    m_set_up = True
    m_file_name = Application.PathSeparator & m_project & "-" & Format(Now(), "mmddhhmmss") & ".log"
    m_log_path = IIf(Right(log_path, 1) = Application.PathSeparator, Left(log_path, Len(log_path) - 1), log_path)
    m_log_on = log_on
    
    createFilePath

End Function

Private Sub createFilePath()

    m_file_path = Dir(m_log_path, vbDirectory)
    
    If m_file_path <> vbNullString Then
        m_file_path = m_log_path & m_file_name
        Exit Sub
    End If
    
    Dim path_arr As Variant
    path_arr = Split(m_log_path, Application.PathSeparator)
    
    Dim lv As Integer
    For lv = LBound(path_arr) To UBound(path_arr)
        If Len(m_file_path) = 0 Then
            m_file_path = path_arr(lv)
        Else
            m_file_path = m_file_path & Application.PathSeparator & path_arr(lv)
        End If
        If Dir(m_file_path, vbDirectory) = vbNullString Then MkDir m_file_path
    Next lv

End Sub

Private Sub openFile()

    m_log_file = FreeFile
    Open m_file_path For Append As #m_log_file

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

Private Function isSetUp() As Boolean
    If m_set_up Then
        isSetUp = True
    ElseIf m_has_error Then
        isSetUp = False
    Else
        Err.Raise _
            vbObjectError + 620, _
            m_project, _
            "You have to use the setUpLog method of the clsCreateLog before create any log item!"
    End If
End Function

Public Sub logInformation(log_text As String)

    If Not isSetUp Then Exit Sub
    If Not m_log_on Then Exit Sub
    If getLogLevel < INFORMATION_LEVEL Then Exit Sub

    Call writeLog("information - " & log_text)

End Sub

Public Sub logDebug(log_text As String)

    If Not isSetUp Then Exit Sub
    If Not m_log_on Then Exit Sub
    If getLogLevel < DEBUG_LEVEL Then Exit Sub

    Call writeLog("debug - " & log_text)

End Sub

Public Sub logWarning(log_text As String)

    If Not isSetUp Then Exit Sub
    If Not m_log_on Then Exit Sub
    If getLogLevel < WARNING_LEVEL Then Exit Sub

    Call writeLog("warning - " & log_text)

End Sub

Public Sub logError(log_text As String, ByRef error As Object)

    If Not isSetUp Then Exit Sub
    If Not m_log_on Then Exit Sub
    If getLogLevel < ERROR_LEVEL Then Exit Sub

    Call writeLog("error - " & log_text & vbNewLine & _
        "Error occured: " & vbNewLine & _
        "Description: " & error.Description & vbNewLine & _
        "Source: " & vbNewLine & _
        error.Source)

End Sub

Public Sub logCritical(log_text As String, ByRef error As Object)

    If Not isSetUp Then Exit Sub
    If Not m_log_on Then Exit Sub

    Call writeLog("critical - " & log_text & vbNewLine & _
        "Error occured: " & vbNewLine & _
        "Description: " & error.Description & vbNewLine & _
        "Source: " & vbNewLine & _
        error.Source)

End Sub


