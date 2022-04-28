Option Explicit

'Assignment 1

Function antoine(A As Double, B As Double, C As Double, t As Double) As Double

Dim result As Double
result = (10) ^ (A - (B / (t + C)))
antoine = result
End Function

Function medication(C0 As Double, k As Double, t As Double) As Double
medication = C0 * Exp(-k * t)
End Function

Function payment(P As Double, i As Double, n As Double) As Double
payment = (P * (i / 12)) / (1 - (1 + (i / 12)) ^ (-n * 12))
End Function

'Assignment 2

Sub AddNumbersA()
Dim x As Double, y As Double, z As Double
x = InputBox("Enter a number:")
y = Range("D4")
z = x + y
Range("G12") = z
End Sub

Sub AddNumbersB()
Dim x As Double, y As Double
x = InputBox("Enter a number:")
y = ActiveCell + x
ActiveCell.Offset(-3, 2) = y
'Place your code here
End Sub

Sub WherePutMe()
Dim row As Integer, col As String, res As String
row = InputBox("Enter a row number:")
col = InputBox("Enter a col number:")
res = col & row
Range(res) = Selection.Cells(2, 2)
End Sub

Sub Swap()
Dim A As Double, B As Double, Temp As Double
A = Selection.Cells(1, 1)
B = Selection.Cells(1, 2)
Temp = A
A = B
B = Temp
Selection.Cells(1, 1) = A
Selection.Cells(1, 2) = B
End Sub

'Assignment 3

Function truck(r As Double, d As Double, f As Double, w As Double, c As Boolean) As Double
Dim P As Double, TR As Double, cost As Double, s As Double, e As Double
Dim GR As Double, a As Double, b As Double, i As Double, df As Double, s1 As Double, s2 As Double, fs1 As Double, fs2 As Double
GR = (Sqr(5) - 1) / 2
a = 30
b = 100
For i = 1 To 20
    df = GR * (b - a)
    s1 = a + d
    s2 = b - df
    fs1 = -(r * s1 - f * s1) / d
    fs2 = -(r * s2 - f * s2) / d
    If fs1 < fs2 Then
        a = s2
    Else
        b = s1
    End If
Next i
s = (fs1 + fs2) / 2
truck = s
End Function

