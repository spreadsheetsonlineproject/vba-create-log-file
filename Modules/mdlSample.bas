Attribute VB_Name = "mdlSample"
Option Explicit

Public log As clsCreateLog

Private Sub setUplog(ByVal is_log_on As Boolean)
    Call log.setUplog(log_on:=is_log_on, _
                      project_name:="SSOPlog", _
                      log_path:="c:\vba-log", _
                      log_level:="debug")
End Sub


Public Sub main()
On Error GoTo eh
    Set log = New clsCreateLog
    Call setUplog(True)
    
log.logDebug "The app is running."
    
    Dim number1 As Variant: number1 = 5
    Dim number2 As Variant: number2 = "asdfa"
    
    Debug.Print add(number1, number2)
    Debug.Print subtract(number1, number2)
    Debug.Print multiply(number1, number2)
    Debug.Print divide(number1, number2)
    
done:
log.logDebug "The app has been finished."
    Exit Sub
eh:
log.logCritical "Unexpected error!", Err
    MsgBox Err.Description & vbNewLine & Err.Source
End Sub


Private Function add(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logDebug "add function running"
    add = number1 + number2
done:
log.logDebug "add function done"
    Exit Function
eh:
log.logError "add function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.add", _
        Err.Description
End Function

Private Function subtract(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logDebug "subtract function running"
    subtract = number1 - number2
done:
log.logDebug "subtract function done"
    Exit Function
eh:
log.logError "subtract function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.subtract", _
        Err.Description
End Function

Private Function multiply(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logDebug "multiply function running"
    multiply = number1 * number2
done:
log.logDebug "multiply function done"
    Exit Function
eh:
log.logError "multiply function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.multiply", _
        Err.Description
End Function

Private Function divide(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logDebug "divide function running"
    divide = number1 / number2
done:
log.logDebug "divide function done"
    Exit Function
eh:
log.logError "divide function error, input parameters: " & number1 & ", " & number2, Err
    Err.Raise _
        Err.Number, _
        Err.Source & vbNewLine & "mdlSample.divide", _
        Err.Description
End Function
