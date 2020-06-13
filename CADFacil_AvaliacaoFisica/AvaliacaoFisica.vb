Imports System.Runtime.InteropServices.WindowsRuntime
Imports System.Windows.Controls
Imports System.Windows.Media.Animation

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

    Public Function Cooper(ByVal idade As Integer, ByVal sexo As String, ByVal coop As Double) As String

        If coop < 860 Then
            Return "Muito Fraco"
        Else
            If sexo = "Masculino" Then
                If coop < 1260 Then
                    Return "Muito Fraco"
                ElseIf idade < 7 Then
                    Return "Não Avaliar"
                ElseIf coop <= 1340 Then
                    If idade = 7 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1360 Then
                    If idade < 9 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1370 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade = 8 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1400 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade < 10 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1420 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade < 10 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1440 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade < 11 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1500 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1540 Then
                    If idade < 9 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1600 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade = 8 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1640 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1660 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1680 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1690 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1730 Then
                    If idade < 9 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1740 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade = 8 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1780 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade < 10 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1810 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade < 11 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1820 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1870 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1880 Then
                    If idade < 9 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1900 Then
                    If idade < 10 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1930 Then
                    If idade < 10 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1950 Then
                    If idade < 10 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1970 Then
                    If idade < 11 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1990 Then
                    If idade < 12 Then
                        Return "Excelente"
                    ElseIf idade = 12 Then
                        Return "Boa"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 2000 Then
                    If idade < 12 Then
                        Return "Excelente"
                    ElseIf idade = 12 Then
                        Return "Boa"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 2090 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 2110 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 2120 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 2200 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2240 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 50 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2320 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2390 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2400 Then
                    If idade = 7 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2450 Then
                    If idade = 7 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2460 Then
                    If idade = 7 Then
                        Return "Superior"
                    ElseIf idade = 9 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2480 Then
                    If idade = 7 Then
                        Return "Superior"
                    ElseIf idade = 9 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Média"
                    ElseIf idade < 40 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 2510 Then
                    If idade = 7 Then
                        Return "Superior"
                    ElseIf idade = 9 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Média"
                    ElseIf idade < 40 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2520 Then
                    If idade = 7 Then
                        Return "Superior"
                    ElseIf idade = 9 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2540 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade = 12 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2640 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade = 12 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Boa"
                    ElseIf idade < 50 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2650 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade = 12 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 50 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2710 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade = 12 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2770 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade = 12 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2830 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade < 30 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 3000 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade < 20 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                Else
                    Return "Superior"
                End If
            Else
                'Feminino
                If idade < 7 Then
                    Return "Não Avaliar"
                ElseIf coop <= 1110 Then
                    Return "Muito Fraco"
                ElseIf coop <= 1180 Then
                    If idade = 7 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1230 Then
                    If idade < 9 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1260 Then
                    If idade < 11 Then
                        Return "Fraco"
                    Else
                        Return "Muito Fraco"
                    End If
                ElseIf coop <= 1280 Then
                    If idade < 11 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1300 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade <= 11 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1320 Then
                    If idade = 7 Then
                        Return "Média"
                    ElseIf idade <= 12 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1350 Then
                    If idade < 9 Then
                        Return "Média"
                    ElseIf idade < 12 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1380 Then
                    If idade < 10 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1390 Then
                    If idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    Else
                        Return "Fraco"
                    End If
                ElseIf coop <= 1410 Then
                    If idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1420 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1440 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1450 Then
                    If idade < 9 Then
                        Return "Boa"
                    ElseIf idade < 11 Then
                        Return "Média"
                    ElseIf idade < 13 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1500 Then
                    If idade < 9 Then
                        Return "Boa"
                    ElseIf idade < 12 Then
                        Return "Média"
                    ElseIf idade = 12 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 60 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1510 Then
                    If idade < 9 Then
                        Return "Boa"
                    ElseIf idade < 12 Then
                        Return "Média"
                    ElseIf idade = 12 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1530 Then
                    If idade < 9 Then
                        Return "Boa"
                    ElseIf idade < 12 Then
                        Return "Média"
                    ElseIf idade = 12 Then
                        Return "Fraco"
                    ElseIf idade < 30 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1540 Then
                    If idade < 10 Then
                        Return "Boa"
                    ElseIf idade < 12 Then
                        Return "Média"
                    ElseIf idade = 12 Then
                        Return "Fraco"
                    ElseIf idade < 30 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1550 Then
                    If idade < 10 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 30 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1560 Then
                    If idade < 10 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1580 Then
                    If idade < 11 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 50 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1590 Then
                    If idade < 11 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    Else
                        Return "Média"
                    End If
                ElseIf coop <= 1600 Then
                    If idade < 11 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1610 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade = 8 Then
                        Return "Excelente"
                    ElseIf idade < 11 Then
                        Return "Boa"
                    ElseIf idade < 13 Then
                        Return "Média"
                    ElseIf idade < 20 Then
                        Return "Muito Fraco"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1640 Then
                    If idade = 7 Then
                        Return "Boa"
                    ElseIf idade = 8 Then
                        Return "Excelente"
                    ElseIf idade < 12 Then
                        Return "Boa"
                    ElseIf idade = 12 Then
                        Return "Média"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1690 Then
                    If idade < 9 Then
                        Return "Excelente"
                    ElseIf idade < 12 Then
                        Return "Boa"
                    ElseIf idade = 12 Then
                        Return "Média"
                    ElseIf idade < 40 Then
                        Return "Fraco"
                    ElseIf idade < 60 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1700 Then
                    If idade < 9 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1750 Then
                    If idade < 10 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    Else
                        Return "Boa"
                    End If
                ElseIf coop <= 1780 Then
                    If idade < 10 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 1790 Then
                    If idade < 11 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Fraco"
                    ElseIf idade < 50 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 1830 Then
                    If idade < 11 Then
                        Return "Excelente"
                    ElseIf idade < 13 Then
                        Return "Boa"
                    ElseIf idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 1900 Then
                    If idade < 12 Then
                        Return "Excelente"
                    ElseIf idade = 12 Then
                        Return "Boa"
                    ElseIf idade < 20 Then
                        Return "Fraco"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 60 Then
                        Return "Boa"
                    Else
                        Return "Excelente"
                    End If
                ElseIf coop <= 1930 Then
                    If idade < 12 Then
                        Return "Excelente"
                    ElseIf idade = 12 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 1960 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 40 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 1970 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2000 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Média"
                    ElseIf idade < 50 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2080 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Média"
                    ElseIf idade < 40 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2090 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Boa"
                    ElseIf idade < 60 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2160 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 30 Then
                        Return "Boa"
                    ElseIf idade < 50 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2180 Then
                    If idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2240 Then
                    If idade = 11 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 40 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2260 Then
                    If idade = 10 Then
                        Return "Superior"
                    ElseIf idade = 11 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2300 Then
                    If idade = 8 Then
                        Return "Superior"
                    ElseIf idade = 10 Then
                        Return "Superior"
                    ElseIf idade = 11 Then
                        Return "Superior"
                    ElseIf idade < 13 Then
                        Return "Excelente"
                    ElseIf idade < 20 Then
                        Return "Boa"
                    ElseIf idade < 30 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2330 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade < 12 Then
                        Return "Superior"
                    ElseIf idade < 30 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2340 Then
                    If idade = 7 Then
                        Return "Excelente"
                    ElseIf idade < 12 Then
                        Return "Superior"
                    ElseIf idade < 20 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2370 Then
                    If idade < 12 Then
                        Return "Superior"
                    ElseIf idade < 20 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                ElseIf coop <= 2430 Then
                    If idade < 13 Then
                        Return "Superior"
                    ElseIf idade < 20 Then
                        Return "Excelente"
                    Else
                        Return "Superior"
                    End If
                Else
                    Return "Superior"
                End If

            End If
        End If

    End Function

    Public Function Punho(ByVal vPunho As Double, ByVal altura As Double) As Double
        '=========================================================================================================
        '=========================================% DA GORDURA====================================================
        '=========================================================================================================
        If vPunho < 11 Then
            Return 4
        ElseIf vPunho = 11 Then
            If altura * 100 <= 150 Then
                Return 3.5
            Else
                Return 4
            End If
        ElseIf vPunho = 12 Then
            If altura * 100 <= 150 Then
                Return 3
            ElseIf altura * 100 <= 160 Then
                Return 3.5
            Else
                Return 4
            End If
        ElseIf vPunho = 13 Then
            If altura * 100 <= 150 Then
                Return 2.5
            ElseIf altura * 100 <= 160 Then
                Return 3
            ElseIf altura * 100 <= 170 Then
                Return 3.5
            Else
                Return 4
            End If
        ElseIf vPunho = 14 Then
            If altura * 100 <= 150 Then
                Return 2
            ElseIf altura * 100 <= 160 Then
                Return 2.5
            ElseIf altura * 100 <= 170 Then
                Return 3
            ElseIf altura * 100 <= 180 Then
                Return 3.5
            Else
                Return 4
            End If
        ElseIf vPunho = 15 Then
            If altura * 100 <= 150 Then
                Return 1.5
            ElseIf altura * 100 <= 160 Then
                Return 2
            ElseIf altura * 100 <= 170 Then
                Return 2.5
            ElseIf altura * 100 <= 180 Then
                Return 3
            Else
                Return 3.5
            End If
        ElseIf vPunho = 16 Then
            If altura * 100 <= 150 Then
                Return 1
            ElseIf altura * 100 <= 160 Then
                Return 1.5
            ElseIf altura * 100 <= 170 Then
                Return 2
            ElseIf altura * 100 <= 180 Then
                Return 2.5
            Else
                Return 3
            End If
        ElseIf vPunho = 17 Then
            If altura * 100 <= 150 Then
                Return 0.5
            ElseIf altura * 100 <= 160 Then
                Return 1
            ElseIf altura * 100 <= 170 Then
                Return 1.5
            ElseIf altura * 100 <= 180 Then
                Return 2
            Else
                Return 2.5
            End If
        ElseIf vPunho = 18 Then
            If altura * 100 <= 150 Then
                Return 0
            ElseIf altura * 100 <= 160 Then
                Return 0.5
            ElseIf altura * 100 <= 170 Then
                Return 1
            ElseIf altura * 100 <= 180 Then
                Return 1.5
            Else
                Return 2
            End If
        ElseIf vPunho = 19 Then
            If altura * 100 <= 160 Then
                Return 0
            ElseIf altura * 100 <= 170 Then
                Return 0.5
            ElseIf altura * 100 <= 180 Then
                Return 1
            Else
                Return 1.5
            End If
        ElseIf vPunho = 20 Then
            If altura * 100 <= 170 Then
                Return 0
            ElseIf altura * 100 <= 180 Then
                Return 0.5
            Else
                Return 1
            End If
        ElseIf vPunho = 21 Then
            If altura * 100 <= 180 Then
                Return 0
            Else
                Return 0.5
            End If
        Else
            Return 0
        End If
    End Function

    Public Function Imc(ByVal sexo As String, ByVal valorimc As Double) As String
        If sexo = "Masculino" Then
            If valorimc = 22 Then
                Return "IDEAL"
            ElseIf valorimc < 20 Then
                Return "ABAIXO DO NORMAL"
            ElseIf valorimc <= 24.9 Then
                Return "NORMAL"
            ElseIf valorimc = 29.9 Then
                Return "OBESO -- 1"
            ElseIf valorimc <= 40 Then
                Return "OBESO -- 2"
            ElseIf valorimc > 40 Then
                Return "OBESO -- 3"
            Else
                Return "Não encontrado"
            End If
        Else
            If valorimc = 21 Then
                Return "IDEAL"
            ElseIf valorimc < 19 Then
                Return "ABAIXO DO NORMAL"
            ElseIf valorimc <= 23.9 Then
                Return "NORMAL"
            ElseIf valorimc = 28.9 Then
                Return "OBESA -- 1"
            ElseIf valorimc <= 39 Then
                Return "OBESA -- 2"
            ElseIf valorimc > 39 Then
                Return "OBESA -- 3"
            Else
                Return "Não encontrado"
            End If
        End If

    End Function

    Public Function Calculo_Cintura_Quadril(ByVal sexo As String, ByVal idade As Integer, ByVal cintura As Double, ByVal quadril As Double) As String
        Dim vResultado As Double

        vResultado = String.Format("{0:N2}", cintura / quadril)

        If sexo = "Masculino" Then
            If vResultado < 0.83 Then
                Return "Baixo"
            ElseIf vResultado = 0.83 Then
                If idade < 30 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado < 0.88 Then
                If idade < 40 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado = 0.88 Then
                If idade < 50 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado = 0.89 Then
                If idade < 30 Then
                    Return "Elevado"
                ElseIf idade < 50 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado = 0.9 Then
                If idade < 30 Then
                    Return "Elevado"
                ElseIf idade < 60 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado = 0.91 Then
                If idade < 30 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado < 0.95 Then
                If idade < 40 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado = 0.95 Then
                If idade < 30 Then
                    Return "Muito Elevado"
                ElseIf idade < 40 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado = 0.96 Then
                If idade < 30 Then
                    Return "Muito Elevado"
                ElseIf idade < 50 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado < 0.99 Then
                If idade < 40 Then
                    Return "Muito Elevado"
                ElseIf idade < 60 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado < 1.01 Then
                If idade < 40 Then
                    Return "Muito Elevado"
                Else
                    Return "Elevado"
                End If
            ElseIf vResultado < 1.03 Then
                If idade < 50 Then
                    Return "Muito Elevado"
                Else
                    Return "Elevado"
                End If
            ElseIf vResultado = 1.03 Then
                If idade < 60 Then
                    Return "Muito Elevado"
                Else
                    Return "Elevado"
                End If
            Else
                Return "Muito Elevado"
            End If
        Else
            'Feminino
            If vResultado < 0.71 Then
                Return "Baixo"
            ElseIf vResultado = 0.71 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 30 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado = 0.72 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 40 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado = 0.73 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 50 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado < 0.76 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 60 Then
                    Return "Moderado"
                Else
                    Return "Baixo"
                End If
            ElseIf vResultado < 0.78 Then
                If idade < 20 Then
                    Return "Baixo"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado = 0.8 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 40 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado < 0.82 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 50 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado = 0.82 Then
                If idade < 20 Then
                    Return "Baixo"
                ElseIf idade < 60 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado = 0.83 Then
                If idade < 20 Then
                    Return "Moderado"
                ElseIf idade < 30 Then
                    Return "Muito Elevado"
                ElseIf idade < 60 Then
                    Return "Elevado"
                Else
                    Return "Moderado"
                End If
            ElseIf vResultado = 0.84 Then
                If idade < 20 Then
                    Return "Moderado"
                ElseIf idade < 30 Then
                    Return "Muito Elevado"
                Else
                    Return "Elevado"
                End If
            ElseIf vResultado < 0.88 Then
                If idade < 20 Then
                    Return "Moderado"
                ElseIf idade < 40 Then
                    Return "Muito Elevado"
                Else
                    Return "Elevado"
                End If
            ElseIf vResultado < 0.91 Then
                If idade < 20 Then
                    Return "Elevado"
                ElseIf idade < 60 Then
                    Return "Muito Elevado"
                Else
                    Return "Elevado"
                End If
            ElseIf vResultado < 0.95 Then
                If idade < 20 Then
                    Return "Elevado"
                Else
                    Return "Muito Elevado"
                End If
            Else
                Return "Muito Elevado"
            End If

        End If

    End Function

End Class
