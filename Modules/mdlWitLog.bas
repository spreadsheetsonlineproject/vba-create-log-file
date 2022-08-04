Attribute VB_Name = "mdlWitLog"
Option Explicit

Public Sub mainWithLog()

Set log = New clsCreateLog
Call log.setUplog(True, mdlGlobal.THIS_PROJECT_NAME, log_level:="debug")

log.logInformation "mdlWithLog.mainWithLog running"
    
    Dim number1 As Variant: number1 = 5
    Dim number2 As Variant: number2 = 0
    
log.logDebug "number1: " & number1 & ", number2: " & number2
    
    Debug.Print add(number1, number2)
    Debug.Print subtract(number1, number2)
    Debug.Print multiply(number1, number2)
    Debug.Print divide(number1, number2)
    
log.logInformation "mdlWithLog.mainWithLog finished"
    
End Sub

Private Function add(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

log.logInformation "mdlWithLog.add running, parameters: " & number1 & ", " & number2

    add = number1 + number2
    
log.logInformation "mdlWithLog.add finished"
    
End Function

Private Function subtract(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

log.logInformation "mdlWithLog.subtract running, parameters: " & number1 & ", " & number2

    subtract = number1 - number2
    
log.logInformation "mdlWithLog.subtract finished"

End Function

Private Function multiply(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

log.logInformation "mdlWithLog.multiply running, parameters: " & number1 & ", " & number2

    multiply = number1 * number2
    
log.logInformation "mdlWithLog.multiply finished"

End Function

Private Function divide(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

log.logInformation "mdlWithLog.divide running, parameters: " & number1 & ", " & number2

    divide = number1 / number2
    
log.logInformation "mdlWithLog.divide finished"

End Function



