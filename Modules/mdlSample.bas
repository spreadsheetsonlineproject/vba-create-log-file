Attribute VB_Name = "mdlSample"
Option Explicit

Public logging As clsLogging

Private Sub setUpLogging(ByVal is_log_on As Boolean)
    Set logging = New clsLogging
    Call logging.setupLog(logging_on:=True, _
                          project_name:="SSOPLogging", _
                          log_path:="c:\vba-log", _
                          log_level:="debug")
End Sub


Public Sub main()
On Error GoTo eh
    Call setUpLogging(True)
    
logging.logDebug "The app is running."
    
    Dim number1 As Variant: number1 = 5
    Dim number2 As Variant: number2 = "asdfa"
    
    Debug.Print add(number1, number2)
    Debug.Print subtract(number1, number2)
    Debug.Print multiply(number1, number2)
    Debug.Print divide(number1, number2)
    
done:
logging.logDebug "The app has been finished."
    Exit Sub
eh:
logging.logCritical "Unexpected error!", Err
    MsgBox Err.Description & vbNewLine & Err.Source
End Sub


Private Function add(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
logging.logDebug "add function running"
    add = number1 + number2
done:
logging.logDebug "add function done"
    Exit Function
eh:
logging.logError "add function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.add", _
        Err.Description
End Function

Private Function subtract(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
logging.logDebug "subtract function running"
    subtract = number1 - number2
done:
logging.logDebug "subtract function done"
    Exit Function
eh:
logging.logError "subtract function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.subtract", _
        Err.Description
End Function

Private Function multiply(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
logging.logDebug "multiply function running"
    multiply = number1 * number2
done:
logging.logDebug "multiply function done"
    Exit Function
eh:
logging.logError "multiply function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.multiply", _
        Err.Description
End Function

Private Function divide(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
logging.logDebug "divide function running"
    divide = number1 / number2
done:
logging.logDebug "divide function done"
    Exit Function
eh:
logging.logError "divide function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.divide", _
        Err.Description
End Function
