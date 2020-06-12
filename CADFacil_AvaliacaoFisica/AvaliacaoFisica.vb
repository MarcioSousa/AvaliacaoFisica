Imports System.Windows.Controls

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

    Public Function Abdominais(ByVal idade As Integer, ByVal sexo As String, ByVal abdom As Double) As String
        If sexo = "Masculino" Then
            If idade < 10 Then
                Return "Deficiente"
            Else
                If abdom = 8 Then
                    Return "Deficientes"
                ElseIf abdom <= 11 Then
                    If idade <= 59 Then
                        Return "Deficientes"
                    Else
                        Return "Fraco"
                    End If
                ElseIf abdom = 12 Then
                    If idade <= 49 Then
                        Return "Deficientes"
                    Else
                        Return "Fraco"
                    End If
                ElseIf abdom <= 16 Then
                    If idade <= 49 Then
                        Return "Deficientes"
                    ElseIf idade <= 59 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf abdom <= 18 Then
                    If idade <= 39 Then
                        Return "Deficiente"
                    ElseIf idade <= 49 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf abdom <= 20 Then
                    If idade <= 39 Then
                        Return "Deficiente"
                    ElseIf idade <= 49 Then
                        Return "Fraco"
                    ElseIf idade <= 59 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf abdom = 21 Then
                    If idade <= 39 Then
                        Return "Deficiente"
                    ElseIf idade <= 59 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf abdom = 22 Then
                    If idade <= 29 Then
                        Return "Deficiente"
                    ElseIf idade <= 39 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf abdom <= 24 Then
                    If idade <= 29 Then
                        Return "Deficiente"
                    ElseIf idade <= 39 Then
                        Return "Fraco"
                    ElseIf idade <= 49 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf abdom = 25 Then
                    If idade <= 29 Then
                        Return "Deficiente"
                    ElseIf idade <= 39 Then
                        Return "Fraco"
                    ElseIf idade <= 49 Then
                        Return "Regular"
                    ElseIf idade <= 59 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom = 26 Then
                    If idade <= 29 Then
                        Return "Deficiente"
                    ElseIf idade <= 39 Then
                        Return "Fraco"
                    ElseIf idade <= 59 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 28 Then
                    If idade <= 29 Then
                        Return "Deficiente"
                    ElseIf idade <= 39 Then
                        Return "Regular"
                    ElseIf idade <= 59 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom = 29 Then
                    If idade <= 29 Then
                        Return "Deficiente"
                    ElseIf idade <= 39 Then
                        Return "Regular"
                    ElseIf idade <= 49 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 31 Then
                    If idade <= 19 Then
                        Return "Deficiente"
                    ElseIf idade <= 29 Then
                        Return "Fraco"
                    ElseIf idade <= 39 Then
                        Return "Regular"
                    ElseIf idade <= 49 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 34 Then
                    If idade <= 19 Then
                        Return "Deficiente"
                    ElseIf idade <= 29 Then
                        Return "Fraco"
                    ElseIf idade <= 39 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 36 Then
                    If idade <= 19 Then
                        Return "Deficiente"
                    ElseIf idade <= 29 Then
                        Return "Regular"
                    ElseIf idade <= 39 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom = 37 Then
                    If idade <= 19 Then
                        Return "Deficiente"
                    ElseIf idade <= 29 Then
                        Return "Regular"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 39 Then
                    If idade <= 19 Then
                        Return "Fraco"
                    ElseIf idade <= 29 Then
                        Return "Regular"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 42 Then
                    If idade <= 19 Then
                        Return "Fraco"
                    ElseIf idade <= 29 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 44 Then
                    If idade <= 19 Then
                        Return "Regular"
                    ElseIf idade <= 29 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 47 Then
                    If idade <= 19 Then
                        Return "Regular"
                    Else
                        Return "Excelente"
                    End If
                ElseIf abdom <= 52 Then
                    If idade <= 19 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                Else
                    Return "Excelente"
                End If
            End If
        Else
            If abdom <= 9 Then
                If idade <= 59 Then
                    Return "Deficiente"
                Else
                    Return "Fraco"
                End If
            ElseIf abdom = 10 Then
                If idade <= 59 Then
                    Return "Deficiente"
                Else
                    Return "Regular"
                End If
            ElseIf abdom <= 14 Then
                If idade <= 49 Then
                    Return "Deficiente"
                ElseIf idade <= 59 Then
                    Return "Fraco"
                Else
                    Return "Regular"
                End If
            ElseIf abdom = 15 Then
                If idade <= 49 Then
                    Return "Deficiente"
                ElseIf idade <= 59 Then
                    Return "Regular"
                Else
                    Return "Bom"
                End If
            ElseIf abdom <= 19 Then
                If idade <= 39 Then
                    Return "Deficiente"
                ElseIf idade <= 49 Then
                    Return "Fraco"
                ElseIf idade <= 59 Then
                    Return "Regular"
                Else
                    Return "Bom"
                End If
            ElseIf abdom = 20 Then
                If idade <= 39 Then
                    Return "Deficiente"
                ElseIf idade <= 49 Then
                    Return "Regular"
                ElseIf idade <= 59 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom <= 24 Then
                If idade <= 29 Then
                    Return "Deficiente"
                ElseIf idade <= 39 Then
                    Return "Fraco"
                ElseIf idade <= 49 Then
                    Return "Regular"
                ElseIf idade <= 59 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom = 25 Then
                If idade <= 29 Then
                    Return "Deficiente"
                ElseIf idade <= 39 Then
                    Return "Regular"
                ElseIf idade <= 49 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom <= 29 Then
                If idade <= 29 Then
                    Return "Deficiente"
                ElseIf idade <= 29 Then
                    Return "Fraco"
                ElseIf idade <= 39 Then
                    Return "Regular"
                ElseIf idade <= 49 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom = 26 Then
                If idade <= 29 Then
                    Return "Deficiente"
                ElseIf idade <= 39 Then
                    Return "Fraco"
                ElseIf idade <= 59 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom <= 34 Then
                If idade <= 19 Then
                    Return "Fraco"
                ElseIf idade <= 29 Then
                    Return "Regular"
                ElseIf idade <= 39 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom <= 39 Then
                If idade <= 19 Then
                    Return "Regular"
                ElseIf idade <= 29 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            ElseIf abdom <= 44 Then
                If idade <= 19 Then
                    Return "Bom"
                Else
                    Return "Excelente"
                End If
            Else
                Return "Excelente"
            End If
        End If
    End Function

    Public Function Flexao_Braco(ByVal idade As Integer, ByVal sexo As String, ByVal flexaoBra As Double) As String
        If sexo = "Masculino" Then
            If flexaoBra <= 4 Then
                Return "Deficiente"
            Else
                If flexaoBra <= 7 Then
                    If idade < 60 Then
                        Return "Deficiente"
                    Else
                        Return "Fraco"
                    End If
                ElseIf flexaoBra <= 9 Then
                    If idade < 50 Then
                        Return "Deficiente"
                    Else
                        Return "Fraco"
                    End If
                ElseIf flexaoBra = 10 Then
                    If idade < 50 Then
                        Return "Deficiente"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra <= 13 Then
                    If idade < 40 Then
                        Return "Deficiente"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra = 14 Then
                    If idade < 30 Then
                        Return "Deficiente"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra <= 16 Then
                    If idade < 30 Then
                        Return "Deficiente"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra = 17 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra <= 21 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra <= 23 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra = 24 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra <= 26 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Regular"
                    ElseIf idade < 60 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 29 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Regular"
                    ElseIf idade < 60 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra = 30 Then
                    If idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Regular"
                    ElseIf idade < 50 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 34 Then
                    If idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 30 Then
                        Return "Regular"
                    ElseIf idade < 50 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 39 Then
                    If idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 30 Then
                        Return "Regular"
                    ElseIf idade < 40 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 49 Then
                    If idade < 20 Then
                        Return "Regular"
                    ElseIf idade < 30 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 49 Then
                    If idade < 20 Then
                        Return "Regular"
                    ElseIf idade < 30 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 59 Then
                    If idade < 20 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                Else
                    Return "Excelente"
                End If
            End If
        Else
            'Aluno do Sexo Feminino
            If flexaoBra <= 1 Then
                Return "Deficiente"
            Else
                If flexaoBra = 2 Then
                    If idade < 60 Then
                        Return "Deficiente"
                    Else
                        Return "Fraco"
                    End If
                ElseIf flexaoBra = 3 Then
                    If idade < 50 Then
                        Return "Deficiente"
                    Else
                        Return "Fraco"
                    End If
                ElseIf flexaoBra = 4 Then
                    If idade < 40 Then
                        Return "Deficiente"
                    Else
                        Return "Fraco"
                    End If
                ElseIf flexaoBra = 5 Then
                    If idade < 30 Then
                        Return "Deficiente"
                    Else
                        Return "Fraco"
                    End If
                ElseIf flexaoBra = 6 Then
                    If idade < 30 Then
                        Return "Deficiente"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra = 7 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra <= 9 Then
                    If idade < 20 Then
                        Return "Deficiente"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra <= 12 Then
                    If idade < 40 Then
                        Return "Fraco"
                    Else
                        Return "Regular"
                    End If
                ElseIf flexaoBra <= 15 Then
                    If idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra <= 17 Then
                    If idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra <= 19 Then
                    If idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Regular"
                    Else
                        Return "Bom"
                    End If
                ElseIf flexaoBra = 20 Then
                    If idade < 50 Then
                        Return "Regular"
                    ElseIf idade < 60 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 23 Then
                    If idade < 40 Then
                        Return "Regular"
                    ElseIf idade < 60 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 26 Then
                    If idade < 30 Then
                        Return "Regular"
                    ElseIf idade < 60 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 28 Then
                    If idade < 20 Then
                        Return "Regular"
                    ElseIf idade < 60 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra = 29 Then
                    If idade < 20 Then
                        Return "Regular"
                    ElseIf idade < 50 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 31 Then
                    If idade < 50 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 34 Then
                    If idade < 40 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 37 Then
                    If idade < 30 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                ElseIf flexaoBra <= 40 Then
                    If idade < 20 Then
                        Return "Bom"
                    Else
                        Return "Excelente"
                    End If
                Else
                    Return "Excelente"
                End If
            End If
        End If
    End Function





End Class
