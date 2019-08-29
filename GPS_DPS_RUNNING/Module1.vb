Option Explicit On
Module Module1
    Public Function CalaWeek_(year As Integer, month As Integer, day As Integer)

        Dim Lyear%, doubleC%
        Dim y%, Ly%
        Dim m%, Lm%
        Dim d%
        Dim W%
        Lyear = Int(20 / 4)
        doubleC = 2 * 20
        y = year
        Ly = Int(y / 4)

        If month = 1 Then
            month = 13
        End If

        If month = 2 Then
            month = 14
        End If
        m = month
        Lm = Int(26 * (m + 1) / 10)
        d = day

        W = (Lyear - doubleC + y + Ly + Lm + d - 1) Mod 7

        If W < 0 Then
            W += 7
        End If

        CalaWeek_ = W


    End Function

End Module
