Option Explicit

'min meaningfull value
Const minMValue As Double = 1e-6

' double comparision
Function doubleEq(a As Double, b As Double) As Boolean
	doubleEq = Abs(a - b) < 2 * minMValue
End Function

Function max(a As Double, b As Double) As Double
	If a > b Then
		max = a
	Else
		max = b
	End If
End Function
