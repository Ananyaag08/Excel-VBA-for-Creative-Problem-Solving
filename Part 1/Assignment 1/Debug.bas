[18:48, 10/20/2021] Anushree MAHE: Option Explicit
Function tank(R As Double, H As Double, d As Double) As Variant
Dim pi As Double
pi = 4 * Atn(1)


If d > H Then
        tank = Null
    ElseIf R > H Then
        tank = Null
      
    ElseIf d <= R Then
        tank = (pi * d ^ 2) / 3 * (3 * R - d)
    ElseIf R < d And d <= H - R Then
        tank = (2 * pi * R ^ 3) / 3 + pi * R ^ 2 * (d - R)
    ElseIf H - R < d And d <= H Then
        tank = (4 * pi * R ^ 3) / 3 + pi * R ^ 2 * (H - 2 * R) - (pi * (H - d) ^ 2) / 3 * (3 * R - H + d)
End If

End Function
[18:48, 10/20/2021] Anushree MAHE: 4.1
[18:48, 10/20/2021] Anushree MAHE: Function prime(n As Integer) As Boolean

Dim i As Integer

prime = True

If n = 1 Then

prime = False

ElseIf n > 2 Then

For i = 2 To n - 1

If n Mod i = 0 Then

prime = False

Exit Function

End If

Next i

End If

End Function
Function countprime(n1 As Integer, n2 As Integer)
Dim i As Integer
Dim counter As Integer

For i = n1 To n2
    If prime(i) = True Then counter = counter + 1
Next i

countprime = counter

End Function


