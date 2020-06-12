Public Class AvaliacaoFisica

    Public Function Idade(ByVal dataNascimento As Date) As Double
        Dim vAno As Double

        vAno = Now.Date.Year - dataNascimento.Date.Year

        Return vAno
    End Function


End Class
