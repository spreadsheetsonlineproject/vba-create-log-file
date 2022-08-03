Attribute VB_Name = "mdlWithLogException"
Option Explicit

Public Sub mainWithLogException()
On Error GoTo eh
Set log = New clsCreateLog
Call mdlGlobal.setUplog(True, THIS_PROJECT_NAME)
    
log.logInformation "mdlWithLogException.mainWithLogException running"
    
    Dim number1 As Variant: number1 = 5
    Dim number2 As Variant: number2 = 0
    
log.logDebug "number1: " & number1 & ", number2: " & number2
    
    Debug.Print add(number1, number2)
    Debug.Print subtract(number1, number2)
    Debug.Print multiply(number1, number2)
    Debug.Print divide(number1, number2)
    
eh:
    Select Case Err.Number
        Case 0
            log.logInformation "mdlWithLogException.mainWithLogException finished"
            Debug.Print "Done!"
        Case Else
            log.logCritical "mdlWithLogException.mainWithLogException", Err
            Debug.Print "Unexpected error: " & vbNewLine & "Error description: " & Err.Description & vbNewLine & "Error source:" & vbNewLine; Err.Source
    End Select
End Sub

Private Function add(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logInformation "mdlWithLogException.add running"

    add = number1 + number2

done:
log.logInformation "mdlWithLogExceptionadd done"
    Exit Function
eh:
log.logError "mdlWithLogException.add, input parameters: " & number1 & ", " & number2, Err
    Err.Raise Err.Number, Err.Source & vbNewLine & "mdlWithLogException.add", Err.Description
End Function

Private Function subtract(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logInformation "mdlWithLogException.subtract running"

    subtract = number1 - number2

done:
log.logInformation "mdlWithLogException.subtract done"
    Exit Function
eh:
log.logError "mdlWithLogException.subtract, input parameters: " & number1 & ", " & number2, Err
    Err.Raise Err.Number, Err.Source & vbNewLine & "mdlWithLogException.subtract", Err.Description
End Function

Private Function multiply(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logInformation "mdlWithLogException.multiply running"

    multiply = number1 * number2

done:
log.logInformation "mdlWithLogException.multiply done"
    Exit Function
eh:
log.logError "mdlWithLogException.multiply, input parameters: " & number1 & ", " & number2, Err
    Err.Raise Err.Number, Err.Source & vbNewLine & "mdlWithLogException.multiply", Err.Description
End Function

Private Function divide(ByVal number1 As Variant, ByVal number2 As Variant) As Variant
On Error GoTo eh
log.logInformation "mdlWithLogException.divide running"

    divide = number1 / number2

done:
log.logInformation "mdlWithLogException.divide done"
    Exit Function
eh:
log.logError "mdlWithLogException.divide, input parameters: " & number1 & ", " & number2, Err
    Err.Raise Err.Number, Err.Source & vbNewLine & "mdlWithLogException.divide", Err.Description
End Function
