Attribute VB_Name = "mdlBasic"
Option Explicit

Public Sub mainWithLog()
    
    Dim number1 As Variant: number1 = 5
    Dim number2 As Variant: number2 = 0
    
    Debug.Print add(number1, number2)
    Debug.Print subtract(number1, number2)
    Debug.Print multiply(number1, number2)
    Debug.Print divide(number1, number2)
    
End Sub

Private Function add(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

    add = number1 + number2
    
End Function

Private Function subtract(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

    subtract = number1 - number2
    
End Function

Private Function multiply(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

    multiply = number1 * number2
    
End Function

Private Function divide(ByVal number1 As Variant, ByVal number2 As Variant) As Variant

    divide = number1 / number2
    
End Function


