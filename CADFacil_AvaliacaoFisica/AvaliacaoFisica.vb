Public Class AvaliacaoFisica

    Public Function Idade(ByVal dataNascimento As Date) As Integer
        Dim vAno As Integer

        vAno = Now.Date.Year - dataNascimento.Date.Year

        Return vAno
    End Function

    Public Function SaltoVertical(ByVal idade As Integer, ByVal sexo As String, ByVal saltoVert As Double) As String
        If sexo = "Masculino" Then
            If idade > 6 Then
                '=======================================================================================
                '===================================IDADE DE 7 A 10 ANOS================================
                '=======================================================================================
                If idade <= 10 Then
                    If saltoVert <= 16 Then
                        Return "Abaixo da Média"
                    ElseIf saltoVert <= 20 Then
                        If idade <= 8 Then
                            Return "Média"
                        Else
                            Return "Abaixo da Média"
                        End If
                    ElseIf saltoVert <= 27 Then
                        Return "Média"
                    ElseIf saltoVert <= 30 Then
                        If idade <= 8 Then
                            Return "Acima da Média"
                        Else
                            Return "Média"
                        End If
                    Else
                        Return "Acima da Média"
                    End If

                    '=======================================================================================
                    '===================================IDADE DE 7 A 16 ANOS================================
                    '=======================================================================================
                ElseIf idade <= 16 Then
                    If saltoVert <= 33 Then
                        Return "Fraco"
                    ElseIf saltoVert <= 36 Then
                        If idade <= 12 Then
                            Return "Regular"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert = 37 Then
                        If idade <= 12 Then
                            Return "Bom"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert <= 40 Then
                        If idade <= 12 Then
                            Return "Bom"
                        ElseIf idade <= 14 Then
                            Return "Regular"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert <= 43 Then
                        If idade <= 12 Then
                            Return "Muito Bom"
                        ElseIf idade <= 14 Then
                            Return "Regular"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert = 44 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        ElseIf idade <= 14 Then
                            Return "Bom"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert <= 49 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        ElseIf idade <= 14 Then
                            Return "Bom"
                        Else
                            Return "Regular"
                        End If
                    ElseIf saltoVert <= 54 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        ElseIf idade <= 14 Then
                            Return "Muito Bom"
                        Else
                            Return "Bom"
                        End If
                    ElseIf saltoVert = 55 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        Else
                            Return "Muito Bom"
                        End If
                    ElseIf saltoVert = 59 Then
                        If idade <= 14 Then
                            Return "Excelente"
                        Else
                            Return "Muito Bom"
                        End If
                    Else
                        Return "Excelente"
                    End If
                    '=======================================================================================
                    '===================================IDADE DE 17 A 70 ANOS================================
                    '=======================================================================================
                Else
                    If saltoVert <= 23 Then
                        Return "Abaixo da Média"
                    ElseIf saltoVert <= 30 Then
                        If idade <= 50 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert <= 32 Then
                        If idade <= 40 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert <= 36 Then
                        If idade <= 30 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert = 37 Then
                        If idade <= 20 Then
                            Return "Média"
                        ElseIf idade <= 30 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert = 38 Then
                        If idade <= 20 Then
                            Return "Média"
                        ElseIf idade <= 30 Then
                            Return "Abaixo da Média"
                        ElseIf idade <= 50 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert <= 44 Then
                        If idade <= 50 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert <= 47 Then
                        If idade <= 40 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert = 48 Then
                        If idade <= 30 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert = 52 Then
                        If idade <= 20 Then
                            Return "Acima da Média"
                        ElseIf idade <= 30 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    Else
                        Return "Acima da Média"
                    End If
                End If
            Else
                Return "Não Avaliar"
            End If
        Else

            If idade > 6 Then
                '=======================================================================================
                '===================================IDADE DE 7 A 10 ANOS================================
                '=======================================================================================
                If idade <= 10 Then
                    If saltoVert <= 17 Then
                        Return "Abaixo da Média"
                    ElseIf saltoVert <= 20 Then
                        If idade <= 8 Then
                            Return "Média"
                        Else
                            Return "Abaixo da Média"
                        End If
                    ElseIf saltoVert <= 23 Then
                        If idade = 9 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert <= 25 Then
                        Return "Média"
                    ElseIf saltoVert <= 28 Then
                        If idade = 7 Then
                            Return "Acima da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert = 29 Then
                        If idade <= 8 Then
                            Return "Acima da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert = 30 Then
                        If idade = 10 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    Else
                        Return "Acima da Média"
                    End If

                    '=======================================================================================
                    '===================================IDADE DE 7 A 16 ANOS================================
                    '=======================================================================================
                ElseIf idade <= 16 Then
                    If saltoVert <= 28 Then
                        Return "Fraco"
                    ElseIf saltoVert <= 32 Then
                        If idade <= 12 Then
                            Return "Regular"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert <= 34 Then
                        If idade <= 12 Then
                            Return "Bom"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert <= 36 Then
                        If idade <= 12 Then
                            Return "Bom"
                        ElseIf idade <= 14 Then
                            Return "Regular"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert <= 38 Then
                        If idade <= 12 Then
                            Return "Muito Bom"
                        ElseIf idade <= 14 Then
                            Return "Regular"
                        Else
                            Return "Fraco"
                        End If
                    ElseIf saltoVert = 39 Then
                        If idade <= 12 Then
                            Return "Muito Bom"
                        Else
                            Return "Regular"
                        End If
                    ElseIf saltoVert = 40 Then
                        If idade <= 12 Then
                            Return "Muito Bom"
                        ElseIf idade <= 14 Then
                            Return "Bom"
                        Else
                            Return "Regular"
                        End If
                    ElseIf saltoVert <= 42 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        ElseIf idade <= 14 Then
                            Return "Bom"
                        Else
                            Return "Regular"
                        End If
                    ElseIf saltoVert <= 44 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        Else
                            Return "Bom"
                        End If
                    ElseIf saltoVert <= 46 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        ElseIf idade <= 14 Then
                            Return "Muito Bom"
                        Else
                            Return "Bom"
                        End If
                    ElseIf saltoVert <= 49 Then
                        If idade <= 12 Then
                            Return "Excelente"
                        Else
                            Return "Muito Bom"
                        End If
                    ElseIf saltoVert = 50 Then
                        If idade <= 14 Then
                            Return "Excelente"
                        Else
                            Return "Muito Bom"
                        End If
                    Else
                        Return "Excelente"
                    End If
                    '=======================================================================================
                    '===================================IDADE DE 17 A 70 ANOS================================
                    '=======================================================================================
                Else
                    If saltoVert <= 16 Then
                        Return "Abaixo da Média"
                    ElseIf saltoVert <= 21 Then
                        If idade <= 50 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert <= 23 Then
                        If idade <= 40 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert <= 26 Then
                        If idade <= 30 Then
                            Return "Abaixo da Média"
                        Else
                            Return "Média"
                        End If
                    ElseIf saltoVert = 27 Then
                        If idade <= 20 Then
                            Return "Abaixo da Média"
                        ElseIf idade <= 50 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert <= 34 Then
                        If idade <= 40 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert <= 38 Then
                        If idade <= 30 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    ElseIf saltoVert = 39 Then
                        If idade <= 20 Then
                            Return "Média"
                        Else
                            Return "Acima da Média"
                        End If
                    Else
                        Return "Acima da Média"
                    End If
                End If
            Else
                Return "Não Avaliar"
            End If
        End If

    End Function

End Class
