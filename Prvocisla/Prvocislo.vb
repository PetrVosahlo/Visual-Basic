Module Prvocislo
    Public Function Primenumbers(cislo As Integer) As Boolean
        Dim podil As Integer = Math.Ceiling(cislo / 2)
        Dim modulo As Integer

        Do
            modulo = cislo Mod podil
            podil -= 1
        Loop Until modulo = 0 Or podil = 1

        If modulo = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
End Module
