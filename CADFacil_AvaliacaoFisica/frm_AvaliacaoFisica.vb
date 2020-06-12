Imports System.Drawing.Printing
Imports System.IO

Public Class frm_AvaliacaoFisica
    Dim vPerseteInic As Integer = 0
    Dim vProcessar As Boolean = True
    Dim vPunho As Double = 0
    Dim sLeitura As String = ""
    Dim avaliacaoFisica As New AvaliacaoFisica

    Private StringToPrint As String
    Private PrintFont As New Font("Consolas", 12)
    Private PrintFont2 As New Font("Consolas", 7.5)
    Private PrintPageSettings As New PageSettings


    ''' <summary>
    ''' Idade do Aluno
    ''' </summary>
    ''' <remarks>Calcula a Idade do Aluno</remarks>
    Private Sub DtpDataNasc_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DtpDataNasc.ValueChanged
        TxtIdade.Text = avaliacaoFisica.Idade(DtpDataNasc.Value)
    End Sub
    Private Sub Cooper()
        If CDec(TxtCooper.Text) < 860 Then
            TxtClaCoo.Text = "Muito Fraco"
        Else
            If CbxSexo.Text = "Masculino" Then
                If CDec(TxtCooper.Text) < 1260 Then
                    TxtClaCoo.Text = "Muito Fraco"
                ElseIf CDec(TxtIdade.Text) < 7 Then
                    TxtClaCoo.Text = "Não Avaliar"
                ElseIf CDec(TxtCooper.Text) <= 1340 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1360 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1370 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) = 8 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1400 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1420 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1440 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1500 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1540 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1600 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) = 8 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1640 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1660 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1680 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1690 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1730 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1740 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) = 8 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1780 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1810 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1820 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1870 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1880 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1900 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1930 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1950 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1970 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1990 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2000 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2090 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2110 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2120 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2200 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2240 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2320 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2390 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2400 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2450 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2460 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 9 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2480 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 9 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2510 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 9 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2520 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 9 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2540 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2640 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2650 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2710 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2770 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2830 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 3000 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                Else
                    TxtClaCoo.Text = "Superior"
                End If
            Else
                'Feminino
                If CDec(TxtIdade.Text) < 7 Then
                    TxtClaCoo.Text = "Não Avaliar"
                ElseIf CDec(TxtCooper.Text) <= 1110 Then
                    TxtClaCoo.Text = "Muito Fraco"
                ElseIf CDec(TxtCooper.Text) <= 1180 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1230 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1260 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1280 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1300 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) <= 11 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1320 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) <= 12 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1350 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1380 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1390 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    Else
                        TxtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1410 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1420 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1440 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1450 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1500 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1510 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1530 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1540 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1550 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1560 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1580 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1590 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    Else
                        TxtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1600 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1610 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) = 8 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1640 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) = 8 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1690 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1700 Then
                    If CDec(TxtIdade.Text) < 9 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1750 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    Else
                        TxtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1780 Then
                    If CDec(TxtIdade.Text) < 10 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1790 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1830 Then
                    If CDec(TxtIdade.Text) < 11 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1900 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Fraco"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Boa"
                    Else
                        TxtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1930 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) = 12 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1960 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 1970 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2000 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2080 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Média"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2090 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 60 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2160 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 50 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2180 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2240 Then
                    If CDec(TxtIdade.Text) = 11 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 40 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2260 Then
                    If CDec(TxtIdade.Text) = 10 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 11 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2300 Then
                    If CDec(TxtIdade.Text) = 8 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 10 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) = 11 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Boa"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2330 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 30 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2340 Then
                    If CDec(TxtIdade.Text) = 7 Then
                        TxtClaCoo.Text = "Excelente"
                    ElseIf CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2370 Then
                    If CDec(TxtIdade.Text) < 12 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(TxtCooper.Text) <= 2430 Then
                    If CDec(TxtIdade.Text) < 13 Then
                        TxtClaCoo.Text = "Superior"
                    ElseIf CDec(TxtIdade.Text) < 20 Then
                        TxtClaCoo.Text = "Excelente"
                    Else
                        TxtClaCoo.Text = "Superior"
                    End If
                Else
                    TxtClaCoo.Text = "Superior"
                End If

            End If
        End If

    End Sub
    Private Sub VO2Max()
        TxtVO2.Text = String.Format("{0:N2}", (CDec(TxtCooper.Text) - 504) / 45)
    End Sub


    ''' <summary>
    ''' Composição Corporal (% de Gordura, Pesos (Gordura, Muscular, Residual, Ósseo), M. C. M)
    ''' </summary>
    ''' <remarks>Composição Corporal</remarks>
    Private Sub Composicao_Corporal()
        '=========================================================================================================
        '=========================================% DA GORDURA====================================================
        '=========================================================================================================
        Dim vPorGordura As Double

        If CDec(TxtPunho.Text) < 11 Then
            vPunho = 4
        ElseIf CDec(TxtPunho.Text) = 11 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(TxtPunho.Text) = 12 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 3
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(TxtPunho.Text) = 13 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 2.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 3
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(TxtPunho.Text) = 14 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 2
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 2.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 3
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(TxtPunho.Text) = 15 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 1.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 2
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 2.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 3
            Else
                vPunho = 3.5
            End If
        ElseIf CDec(TxtPunho.Text) = 16 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 1
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 1.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 2
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 2.5
            Else
                vPunho = 3
            End If
        ElseIf CDec(TxtPunho.Text) = 17 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 0.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 1
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 1.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 2
            Else
                vPunho = 2.5
            End If
        ElseIf CDec(TxtPunho.Text) = 18 Then
            If CDec(TxtAltura.Text) * 100 <= 150 Then
                vPunho = 0
            ElseIf CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 0.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 1
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 1.5
            Else
                vPunho = 2
            End If
        ElseIf CDec(TxtPunho.Text) = 19 Then
            If CDec(TxtAltura.Text) * 100 <= 160 Then
                vPunho = 0
            ElseIf CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 0.5
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 1
            Else
                vPunho = 1.5
            End If
        ElseIf CDec(TxtPunho.Text) = 20 Then
            If CDec(TxtAltura.Text) * 100 <= 170 Then
                vPunho = 0
            ElseIf CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 0.5
            Else
                vPunho = 1
            End If
        ElseIf CDec(TxtPunho.Text) = 21 Then
            If CDec(TxtAltura.Text) * 100 <= 180 Then
                vPunho = 0
            Else
                vPunho = 0.5
            End If
        Else
            vPunho = 0
        End If

        If CbxSexo.Text = "Masculino" Then
            vPorGordura = ((vPunho + 115) + CDec(TxtPeso.Text)) - (CDec(TxtAltura.Text) * 100)

        ElseIf CbxSexo.Text = "Feminino" Then
            vPorGordura = ((vPunho + 125) + CDec(TxtPeso.Text)) - (CDec(TxtAltura.Text) * 100)
        Else
            vPorGordura = 6
        End If

        If vPorGordura < 6 Then
            vPorGordura = 6
        End If

        TxtPorcGordura.Text = vPorGordura
        '=========================================================================================================
        '========================================PESO DA GORDURA==================================================
        '=========================================================================================================
        TxtPesoGordura.Text = String.Format("{0:N3}", vPorGordura / 100) * CDec(TxtPeso.Text)
        '=========================================================================================================
        '=========================================PESO RESIDUAL===================================================
        '=========================================================================================================
        If CbxSexo.Text = "Masculino" Then
            TxtPesoResid.Text = String.Format("{0:N3}", CDec(TxtPeso.Text) * 0.241)
        ElseIf CbxSexo.Text = "Feminino" Then
            TxtPesoResid.Text = String.Format("{0:N3}", CDec(TxtPeso.Text) * 0.209)
        Else
            TxtPesoResid.Text = 0
        End If
        '=========================================================================================================
        '==========================================PESO ÓSSEO=====================================================
        '=========================================================================================================
        TxtPesoOsseo.Text = String.Format("{0:N3}", 3.02 * ((CDec(TxtAltura.Text) * CDec(TxtAltura.Text)) * (CDec(TxtFemur.Text) * 0.001) * (CDec(TxtRadio.Text) * 0.001) * 400) ^ 0.712)
        '=========================================================================================================
        '=========================================PESO MUSCULAR===================================================
        '=========================================================================================================
        TxtPesoMusc.Text = CDec(TxtPeso.Text) - (CDec(TxtPesoGordura.Text) + CDec(TxtPesoResid.Text) + CDec(TxtPesoOsseo.Text))
        '=========================================================================================================
        '=============================================M.C.M=======================================================
        '=========================================================================================================
        TxtMcm.Text = String.Format("{0:N1}", CDec(TxtPeso.Text) - CDec(TxtPesoGordura.Text))
    End Sub


    ''' <summary>
    ''' Massa Provável das Várias Partes do Corpo
    ''' </summary>
    ''' <remarks>Massa Partes do Corpo</remarks>
    Private Sub MassaPartesCorpo()
        'Cabeça
        TxtCabeca.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 7.3) / 100)
        'Tronto
        TxtTronco.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 50.7) / 100)
        'Braço
        TxtBraco.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 2.6) / 100)
        'Antebraço
        TxtAntebraco.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 1.6) / 100)
        'Mão
        TxtMao.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 0.7) / 100)
        'Coxa
        TxtCoxa.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 10.3) / 100)
        'Perna
        TxtPerna.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 4.3) / 100)
        'Pé
        TxtPe.Text = String.Format("{0:N3}", (CDec(TxtPeso.Text) * 1.5) / 100)
    End Sub


    ''' <summary>
    ''' Calcula o Peso Ideal ao Aluno
    ''' </summary>
    ''' <remarks>Peso Ideal</remarks>
    Private Sub PesosIdeais()
        Dim vImc As Double
        Dim vOssatura As Double
        Dim vComplFisico As Double

        'Normal
        TxtPesoNormal.Text = String.Format("{0:N2}", (((CDec(TxtAltura.Text) * 100) - 100) + 4 * CDec(TxtPunho.Text)) / 2)

        'Mínimo
        TxtPesoMinimo.Text = String.Format("{0:N2}", CDec(TxtPesoNormal.Text) - ((CDec(TxtPesoNormal.Text) / 100) * 10))

        'Máximo
        TxtPesoMaximo.Text = String.Format("{0:N2}", ((CDec(TxtPesoNormal.Text) / 100) * 10) + CDec(TxtPesoNormal.Text))

        'IMC
        TxtIMC.Text = String.Format("{0:N2}", CDec(TxtPeso.Text) / (CDec(TxtAltura.Text) * CDec(TxtAltura.Text)))
        vImc = String.Format("{0:N2}", CDec(TxtPeso.Text) / (CDec(TxtAltura.Text) * CDec(TxtAltura.Text)))

        If CbxSexo.Text = "Masculino" Then
            If CDec(TxtIMC.Text) = 22 Then
                TxtResIMC.Text = "IDEAL"
            ElseIf CDec(TxtIMC.Text) < 20 Then
                TxtResIMC.Text = "ABAIXO DO NORMAL"
            ElseIf CDec(TxtIMC.Text) <= 24.9 Then
                TxtResIMC.Text = "NORMAL"
            ElseIf CDec(TxtIMC.Text) = 29.9 Then
                TxtResIMC.Text = "OBESO -- 1"
            ElseIf CDec(TxtIMC.Text) <= 40 Then
                TxtResIMC.Text = "OBESO -- 2"
            ElseIf CDec(TxtIMC.Text) > 40 Then
                TxtResIMC.Text = "OBESO -- 3"
            End If
        Else
            If CDec(TxtIMC.Text) = 21 Then
                TxtResIMC.Text = "IDEAL"
            ElseIf CDec(TxtIMC.Text) < 19 Then
                TxtResIMC.Text = "ABAIXO DO NORMAL"
            ElseIf CDec(TxtIMC.Text) <= 23.9 Then
                TxtResIMC.Text = "NORMAL"
            ElseIf CDec(TxtIMC.Text) = 28.9 Then
                TxtResIMC.Text = "OBESA -- 1"
            ElseIf CDec(TxtIMC.Text) <= 39 Then
                TxtResIMC.Text = "OBESA -- 2"
            ElseIf CDec(TxtIMC.Text) > 39 Then
                TxtResIMC.Text = "OBESA -- 3"
            End If
        End If

        'Ossatura
        TxtOssatura.Text = String.Format("{0:N1}", (CDec(txtTornozelo.Text) + CDec(TxtJoelho.Text) + CDec(TxtPunho.Text)) / CDec(TxtAltura.Text))
        vOssatura = String.Format("{0:N2}", (CDec(txtTornozelo.Text) + CDec(TxtJoelho.Text) + CDec(TxtPunho.Text)) / CDec(TxtAltura.Text))

        If vOssatura < 43 Then
            TxtResOssatura.Text = "FRACA"
        ElseIf vOssatura < 46 Then
            TxtResOssatura.Text = "MÉDIA"
        Else
            TxtResOssatura.Text = "FORTE"
        End If

        'Complemento Físico
        TxtCFisico.Text = String.Format("{0:N2}", (CDec(TxtAltura.Text) * 100) / CDec(TxtPunho.Text))
        vComplFisico = String.Format("{0:N2}", (CDec(TxtAltura.Text) * 100) / CDec(TxtPunho.Text))

        If CbxSexo.Text = "Masculino" Then
            If vComplFisico > 10.4 Then
                TxtResCFisico.Text = "PEQUENA"
            ElseIf vComplFisico < 9.6 Then
                TxtResCFisico.Text = "GRANDE"
            Else
                TxtResCFisico.Text = "MÉDIA"
            End If
        Else
            If vComplFisico > 11 Then
                TxtResCFisico.Text = "PEQUENA"
            ElseIf vComplFisico < 10.1 Then
                TxtResCFisico.Text = "GRANDE"
            Else
                TxtResCFisico.Text = "MÉDIA"
            End If
        End If

    End Sub


    ''' <summary>
    ''' Calcula a Relação da Cintura com o Quadril
    ''' </summary>
    ''' <remarks>Cintura com o Quadril</remarks>
    Private Sub Calculo_Cintura_Quadril()
        Dim vResultado As Double
        'txtRelCinQua.Text = String.Format("{0:N2}", CDec(txtCintura.Text) / CDec(txtQuadril.Text))
        vResultado = String.Format("{0:N2}", CDec(TxtCintura.Text) / CDec(TxtQuadril.Text))

        If CbxSexo.Text = "Masculino" Then
            If vResultado < 0.83 Then
                LblRelCinQua.Text = "Baixo"
            ElseIf vResultado = 0.83 Then
                If CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado < 0.88 Then
                If CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.88 Then
                If CDec(TxtIdade.Text) < 50 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.89 Then
                If CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Elevado"
                ElseIf CDec(TxtIdade.Text) < 50 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.9 Then
                If CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Elevado"
                ElseIf CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.91 Then
                If CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 0.95 Then
                If CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.95 Then
                If CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.96 Then
                If CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(TxtIdade.Text) < 50 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 0.99 Then
                If CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 1.01 Then
                If CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Muito Elevado"
                Else
                    LblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 1.03 Then
                If CDec(TxtIdade.Text) < 50 Then
                    LblRelCinQua.Text = "Muito Elevado"
                Else
                    LblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado = 1.03 Then
                If CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Muito Elevado"
                Else
                    LblRelCinQua.Text = "Elevado"
                End If
            Else
                LblRelCinQua.Text = "Muito Elevado"
            End If
        Else
            'Feminino
            If vResultado < 0.71 Then
                LblRelCinQua.Text = "Baixo"
            ElseIf vResultado = 0.71 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.72 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.73 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 50 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado < 0.76 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Moderado"
                Else
                    LblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado < 0.78 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.8 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 0.82 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 50 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.82 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Baixo"
                ElseIf CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.83 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Moderado"
                ElseIf CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.84 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Moderado"
                ElseIf CDec(TxtIdade.Text) < 30 Then
                    LblRelCinQua.Text = "Muito Elevado"
                Else
                    LblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 0.88 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Moderado"
                ElseIf CDec(TxtIdade.Text) < 40 Then
                    LblRelCinQua.Text = "Muito Elevado"
                Else
                    LblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 0.91 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Elevado"
                ElseIf CDec(TxtIdade.Text) < 60 Then
                    LblRelCinQua.Text = "Muito Elevado"
                Else
                    LblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 0.95 Then
                If CDec(TxtIdade.Text) < 20 Then
                    LblRelCinQua.Text = "Elevado"
                Else
                    LblRelCinQua.Text = "Muito Elevado"
                End If
            Else
                LblRelCinQua.Text = "Muito Elevado"
            End If

        End If



        TxtRelCinQua.Text = vResultado
    End Sub


    ''' <summary>
    ''' Calcula toda a Prescrição do Treinamento
    ''' </summary>
    ''' <remarks>Precrição do Treinamento</remarks>
    Private Sub Prescricao_Treinamento()
        'Cálculo do Rítmo de Atividade
        Dim vRitmoAtiv As Double
        Dim vDistAtiv As Integer

        vRitmoAtiv = String.Format("{0:N2}", ((CDec(TxtCooper.Text) - 504) / 45) * CDec(TxtVo2Trab.Text) / 100)

        If vRitmoAtiv < 7 Then
            TxtRitmoAtiv.Text = "Sem Valor"
        ElseIf vRitmoAtiv < 7.76 Then
            TxtRitmoAtiv.Text = 50
        ElseIf vRitmoAtiv < 8.5 Then
            TxtRitmoAtiv.Text = 55.5
        ElseIf vRitmoAtiv < 9 Then
            TxtRitmoAtiv.Text = 60
        ElseIf vRitmoAtiv < 10.4 Then
            TxtRitmoAtiv.Text = 62.5
        ElseIf vRitmoAtiv < 10.7 Then
            TxtRitmoAtiv.Text = 70
        ElseIf vRitmoAtiv < 12.5 Then
            TxtRitmoAtiv.Text = 71.4
        ElseIf vRitmoAtiv < 13.3 Then
            TxtRitmoAtiv.Text = 80
        ElseIf vRitmoAtiv < 15 Then
            TxtRitmoAtiv.Text = 83.3
        ElseIf vRitmoAtiv < 15.14 Then
            TxtRitmoAtiv.Text = 95
        ElseIf vRitmoAtiv < 16.2 Then
            TxtRitmoAtiv.Text = 90.9
        ElseIf vRitmoAtiv < 17.6 Then
            TxtRitmoAtiv.Text = 95
        ElseIf vRitmoAtiv < 20.5 Then
            TxtRitmoAtiv.Text = 100
        ElseIf vRitmoAtiv < 20.88 Then
            TxtRitmoAtiv.Text = 110
        ElseIf vRitmoAtiv < 22.1 Then
            TxtRitmoAtiv.Text = 111.11
        ElseIf vRitmoAtiv < 23 Then
            TxtRitmoAtiv.Text = 115
        ElseIf vRitmoAtiv < 23.8 Then
            TxtRitmoAtiv.Text = 117.6
        ElseIf vRitmoAtiv < 25.5 Then
            TxtRitmoAtiv.Text = 120
        ElseIf vRitmoAtiv < 27.3 Then
            TxtRitmoAtiv.Text = 125
        ElseIf vRitmoAtiv < 30.5 Then
            TxtRitmoAtiv.Text = 130
        ElseIf vRitmoAtiv < 31.5 Then
            TxtRitmoAtiv.Text = 135
        ElseIf vRitmoAtiv < 32.07 Then
            TxtRitmoAtiv.Text = 140
        ElseIf vRitmoAtiv < 33.5 Then
            TxtRitmoAtiv.Text = 142.85
        ElseIf vRitmoAtiv < 34.26 Then
            TxtRitmoAtiv.Text = 150
        ElseIf vRitmoAtiv < 35.4 Then
            TxtRitmoAtiv.Text = 153.84
        ElseIf vRitmoAtiv < 36.83 Then
            TxtRitmoAtiv.Text = 160
        ElseIf vRitmoAtiv < 37.5 Then
            TxtRitmoAtiv.Text = 166.66
        ElseIf vRitmoAtiv < 39.5 Then
            TxtRitmoAtiv.Text = 170
        ElseIf vRitmoAtiv < 41.5 Then
            TxtRitmoAtiv.Text = 180
        ElseIf vRitmoAtiv < 43.5 Then
            TxtRitmoAtiv.Text = 190
        ElseIf vRitmoAtiv < 45.5 Then
            TxtRitmoAtiv.Text = 200
        ElseIf vRitmoAtiv < 47.5 Then
            TxtRitmoAtiv.Text = 210
        ElseIf vRitmoAtiv < 49.5 Then
            TxtRitmoAtiv.Text = 220
        ElseIf vRitmoAtiv < 51.5 Then
            TxtRitmoAtiv.Text = 230
        ElseIf vRitmoAtiv < 53.5 Then
            TxtRitmoAtiv.Text = 240
        ElseIf vRitmoAtiv < 55.5 Then
            TxtRitmoAtiv.Text = 250
        ElseIf vRitmoAtiv < 60.64 Then
            TxtRitmoAtiv.Text = 260
        ElseIf vRitmoAtiv < 69.5 Then
            TxtRitmoAtiv.Text = 285.7
        ElseIf vRitmoAtiv < 40.16 Then
            TxtRitmoAtiv.Text = 330
        ElseIf vRitmoAtiv < 71.05 Then
            TxtRitmoAtiv.Text = 333.3
        Else
            TxtRitmoAtiv.Text = 340
        End If

        vDistAtiv = CDec(TxtRitmoAtiv.Text) * CDec(TxtTempoAtiv.Text)
        TxtDistanciaAtiv.Text = vDistAtiv

    End Sub


    ''' <summary>
    ''' Faz o Cálculo do Grupo de Gasto Calórico
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Gasto_Calorico()
        '========== Calculando a Atividade ==============
        Dim vCalculadoInicio As Double
        Dim vCalculoFinal As Integer
        Dim vCalculoPorcentagem As Double

        vCalculadoInicio = ((((CDec(TxtCooper.Text) - 504) / 45) / 3.5) * 1.25)
        vCalculoFinal = ((vCalculadoInicio * CDec(TxtPeso.Text)) / 60) * CDec(TxtTempoAtiv.Text)
        vCalculoPorcentagem = (CDec(TxtVo2Trab.Text)) / 100
        TxtAtividade.Text = CInt(vCalculoFinal * vCalculoPorcentagem)

        '========== Calculandoo Repouso ==================
        ''=================== MASCULINO ==================
        If CbxSexo.Text = "Masculino" Then
            If CDec(TxtIdade.Text) < 19 Then
                TxtRepouso.Text = CInt(16.6 * CDec(TxtPeso.Text) + (77 * CDec(TxtAltura.Text)) + 572)
            ElseIf CDec(TxtIdade.Text) < 31 Then
                TxtRepouso.Text = CInt(15.4 * CDec(TxtPeso.Text) + (27 * CDec(TxtAltura.Text)) + 717)
            ElseIf CDec(TxtIdade.Text) < 60 Then
                TxtRepouso.Text = CInt(11.3 * CDec(TxtPeso.Text) + (16 * CDec(TxtAltura.Text)) + 901)
            Else
                TxtRepouso.Text = CInt(8.8 * CDec(TxtPeso.Text) + (11 * CDec(TxtAltura.Text)) + 1071)
            End If
        Else
            If CDec(TxtIdade.Text) < 19 Then
                TxtRepouso.Text = CInt(7.4 * CDec(TxtPeso.Text) + (482 * CDec(TxtAltura.Text)) + 217)
            ElseIf CDec(TxtIdade.Text) < 31 Then
                TxtRepouso.Text = CInt(13.3 * CDec(TxtPeso.Text) + (334 * CDec(TxtAltura.Text)) + 35)
            ElseIf CDec(TxtIdade.Text) < 60 Then
                TxtRepouso.Text = CInt(8.7 * CDec(TxtPeso.Text) - (25 * CDec(TxtAltura.Text)) + 862)
            Else
                TxtRepouso.Text = CInt(9.2 * CDec(TxtPeso.Text) + (637 * CDec(TxtAltura.Text)) + 302)
            End If
        End If
        '========== Calculando a Diário/Ativ =============
        If CbxSexo.Text = "Masculino" Then
            '=SE(F10="leve";B46*1,55;SE(F10="moderada";B46*1,78;B46*2,1))
            If CbxAtivFisDia.Text = "Leve" Then
                'A75 * 1,55
                TxtDiario.Text = CInt((TxtRepouso.Text) * 1.55)
            ElseIf CbxAtivFisDia.Text = "Moderada" Then
                'A75 * 1,78
                TxtDiario.Text = CInt((TxtRepouso.Text) * 1.78)
            Else
                'A75 * 2,1
                TxtDiario.Text = CInt((TxtRepouso.Text) * 2.1)
            End If
        Else
            '=SE(F10="LEVE";B46*1,56;SE(F10="MODERADA";B46*1,64;B46*1,82))
            If CbxAtivFisDia.Text = "Leve" Then
                'A76 * 1,56
                TxtDiario.Text = CInt((TxtRepouso.Text) * 1.56)
            ElseIf CbxAtivFisDia.Text = "Moderada" Then
                'A76 * 1,64
                TxtDiario.Text = CInt((TxtRepouso.Text) * 1.64)
            Else
                'A76 * 1,82
                TxtDiario.Text = CInt((TxtRepouso.Text) * 1.82)
            End If
        End If
        '===========Calculando o Total ===================
        TxtTotal.Text = CInt(TxtAtividade.Text) + CInt(TxtDiario.Text)

    End Sub


    ''' <summary>
    ''' Calcula a Redução de Peso e a Qtde de Peso Perdido
    ''' </summary>
    ''' <remarks>Redução de Peso/Peso Perdido</remarks>
    Private Sub ReducaoPeso_PesoPerdido()
        Dim vConversao As Double

        vConversao = CDec(TxtPeso.Text) - CDec(TxtPesoNormal.Text)
        TxtPerdaDesGord.Text = String.Format("{0:f3}", vConversao)

        vConversao = (CDec(TxtAtividade.Text) * 0.453) / 3500
        TxtPesoDia.Text = String.Format("{0:f3}", vConversao)

        vConversao = CInt(CDec(TxtPerdaDesGord.Text) / CDec(TxtPesoDia.Text))
        TxtDiasNeces.Text = vConversao

        vConversao = CDec(TxtDiasNeces.Text) / 30
        TxtMesesNecess.Text = String.Format("{0:f1}", vConversao)

        vConversao = CDec(TxtPesoDia.Text) * CDec(TxtQuantSeman.Text)
        TxtPesoSemana.Text = String.Format("{0:f3}", vConversao)

        vConversao = CDec(TxtPesoSemana.Text) * 4
        TxtPesoMes.Text = String.Format("{0:f3}", vConversao)
    End Sub
    Private Sub Carrega_campos()
        TxtProf.Text = "Antenor José Augusto"
        TxtNome.Text = "Marcos José Antônio da Silva"
        TxtLocal.Text = "Academia Bastos de Musculação Ltda 123456"
        DtpDataNasc.Text = "26/03/1967"
        CbxSexo.Text = "Masculino"
        TxtPeso.Text = 84
        TxtAltura.Text = 1.77
        TxtFCrepouso.Text = 88
        TxtFCmaxima.Text = 172
        TxtBraDireito.Text = 13
        TxtBraEsquerdo.Text = 15
        TxtAnteDireito.Text = 20
        TxtAnteEsquerdo.Text = 30
        TxtTorax.Text = 30
        TxtAbdomen.Text = 20
        TxtCoxaDireita.Text = 10
        TxtCoxaEsquerda.Text = 50
        TxtPernaDireita.Text = 40
        TxtPernaEsquerda.Text = 10
        txtTornozelo.Text = 20
        TxtJoelho.Text = 35
        TxtPunho.Text = 18
        TxtCintura.Text = 120
        TxtQuadril.Text = 85
        TxtFemur.Text = 42
        TxtRadio.Text = 35
        TxtSaltoVert.Text = 45
        TxtAbdominais.Text = 48
        TxtFlexaoBraco.Text = 33
        TxtCooper.Text = 2000
        TxtVo2Trab.Text = 70
        TxtTempoAtiv.Text = 90
        TxtQuantSeman.Text = 5
    End Sub
    Private Sub BtnProcessar_Click(sender As System.Object, e As System.EventArgs) Handles BtnProcessar.Click
        Try
            If CbxSexo.Text <> "Masculino" And CbxSexo.Text <> "Feminino" Then
                MessageBox.Show("Escolha entre o sexo Masculino ou o sexo Feminino!", "Sexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Error)
                CbxSexo.Focus()
                Exit Sub
            End If

            If TxtPunho.Text = "" Then
                MessageBox.Show("Campo Punho Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                TxtPunho.Focus()
            ElseIf TxtAltura.Text = "" Then
                MessageBox.Show("Campo Altura Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                TxtAltura.Focus()
            ElseIf TxtIdade.Text = "" Then
                MessageBox.Show("Campo Idade Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                TxtIdade.Focus()
            ElseIf TxtSaltoVert.Text = "" Then
                MessageBox.Show("Campo Salto Vertical Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                TxtSaltoVert.Focus()
            ElseIf CbxSexo.Text = "" Then
                MessageBox.Show("Campo Sexo Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                CbxSexo.Focus()
            End If

            If vProcessar = True Then
                TxtClasv.Text = avaliacaoFisica.SaltoVertical(TxtIdade.Text, CbxSexo.Text, TxtSaltoVert.Text)
                TxtClaAbd.Text = avaliacaoFisica.Abdominais(TxtIdade.Text, CbxSexo.Text, TxtAbdominais.Text)
                TxtClaFle.Text = avaliacaoFisica.Flexao_Braco(TxtIdade.Text, CbxSexo.Text, TxtFlexaoBraco.Text)

                Cooper()
                VO2Max()
                Composicao_Corporal()
                Calculo_Cintura_Quadril()
                MassaPartesCorpo()
                PesosIdeais()
                Prescricao_Treinamento()
                Gasto_Calorico()
                ReducaoPeso_PesoPerdido()
            End If

            vProcessar = True
        Catch ex As Exception
            MessageBox.Show("Foi encontrado um erro em um dos campos, verifique os campos novamente!", "Valor de campo inválido!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub
    Private Sub Frm_Principal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Carrega_campos()
    End Sub


    ''' <summary>
    ''' Usando Keypress de todos os campos
    ''' </summary>
    ''' <remarks>Nessa parte é pressionado um botão e sempre aceitará número ou vírgula</remarks>
    Private Sub TxtPeso_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPeso.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtAltura_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAltura.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtFCrepouso_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtFCrepouso.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtFCmaxima_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtFCmaxima.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtBraDireito_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBraDireito.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtBraEsquerdo_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBraEsquerdo.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtAnteDireito_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAnteDireito.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtAnteEsquerdo_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAnteEsquerdo.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTorax_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtTorax.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtAbdomen_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAbdomen.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtCoxaDireita_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCoxaDireita.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtCoxaEsquerda_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCoxaEsquerda.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtPernaDireita_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPernaDireita.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtPernaEsquerda_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPernaEsquerda.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTornozelo_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtTornozelo.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtJoelho_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtJoelho.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtPunho_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPunho.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtFemur_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtFemur.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRadio_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtRadio.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtCintura_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCintura.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtQuadril_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtQuadril.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtSaltoVert_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtSaltoVert.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtAbdominais_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAbdominais.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtFlexaoBraco_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtFlexaoBraco.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtCooper_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCooper.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtVo2Trab_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtVo2Trab.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTempoAtiv_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtTempoAtiv.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtQuantSeman_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtQuantSeman.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtIdade_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtIdade.KeyPress
        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub


    ''' <summary>
    ''' Impressão do documento de AVALIAÇÃO FÍSICA
    ''' </summary>
    ''' <remarks>Aqui está anotado todo o código para impressão do documento preenchido no formulário</remarks>
    Private Sub PrintDocument1_PrintPage(sender As System.Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim numChars As Integer
        Dim numLines As Integer
        Dim stringForPage As String
        Dim strFormat As New StringFormat

        'Com base na configuração de página, define o retângulo desenhável na página
        Dim rectDraw As New Rectangle(e.MarginBounds.Left, e.MarginBounds.Top, e.MarginBounds.Width, e.MarginBounds.Height)

        If rectDraw.Width = 969 Then
            'Define a área para determinar quanto texto cabe em uma página
            'Diminui a altura por uma linha para assegurar que o texto não seja cortado
            Dim sizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - PrintFont.GetHeight(e.Graphics))

            'Ao desenhar strings longas, quebra entre palavras
            strFormat.Trimming = StringTrimming.Word


            'Calcula quantos caracteres e linhas podem caber com base em sizeMeasure
            e.Graphics.MeasureString(StringToPrint, PrintFont, sizeMeasure, strFormat, numChars, numLines)

            'Calcula a string que caberá em uma página
            stringForPage = StringToPrint.Substring(0, numChars)

            'Imprime a string na página atual
            e.Graphics.DrawString(stringForPage, PrintFont, Brushes.Black, rectDraw, strFormat)

            'Se houver mais texto, indica que há mais páginas
            If numChars < StringToPrint.Length Then
                'Subtrai texto da string que foi impressa
                StringToPrint = StringToPrint.Substring(numChars)
                e.HasMorePages = True
            Else
                e.HasMorePages = False
                'Todo o texto foi impresso, então restaura string
                StringToPrint = sLeitura
            End If
        Else
            'Define a área para determinar quanto texto cabe em uma página
            'Diminui a altura por uma linha para assegurar que o texto não seja cortado
            Dim sizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - PrintFont2.GetHeight(e.Graphics))

            'Ao desenhar strings longas, quebra entre palavras
            strFormat.Trimming = StringTrimming.Word


            'Calcula quantos caracteres e linhas podem caber com base em sizeMeasure
            e.Graphics.MeasureString(StringToPrint, PrintFont2, sizeMeasure, strFormat, numChars, numLines)

            'Calcula a string que caberá em uma página
            stringForPage = StringToPrint.Substring(0, numChars)

            'Imprime a string na página atual
            e.Graphics.DrawString(stringForPage, PrintFont2, Brushes.Black, rectDraw, strFormat)

            'Se houver mais texto, indica que há mais páginas
            If numChars < StringToPrint.Length Then
                'Subtrai texto da string que foi impressa
                StringToPrint = StringToPrint.Substring(numChars)
                e.HasMorePages = True
            Else
                e.HasMorePages = False
                'Todo o texto foi impresso, então restaura string
                StringToPrint = sLeitura
            End If

        End If
    End Sub
    Private Sub FitaMétricaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FitaMétricaToolStripMenuItem.Click
        Dim result As DialogResult = PrintDialog1.ShowDialog()

        Try
            '====================================================================================================
            '===============================================SLEITURA=============================================
            sLeitura = "                                                                                   " & Date.Now & vbCrLf
            sLeitura = sLeitura & "                              AVALIAÇÃO FÍSICA - FITA MÉTRICA" & vbCrLf
            '=============================================102 ESPAÇOS============================================
            sLeitura = sLeitura & "Professor: " & (TxtProf.Text.Trim & Space(40)).Substring(0, 40) & Space(10) & (TxtLocal.Text & Space(40)).Substring(0, 40) & vbCrLf
            sLeitura = sLeitura & "======================================================================================================" & vbCrLf
            sLeitura = sLeitura & " NOME:" & (TxtNome.Text.Trim & Space(30)).Substring(0, 30) & Space(2) & "IDADE:" & TxtIdade.Text & Space(2) & "SEXO:" & CbxSexo.Text & Space(2) & "PESO:" & TxtPeso.Text & Space(2) & "ALTURA:" & TxtAltura.Text & Space(2) & "DN:" & DtpDataNasc.Value & vbCrLf
            sLeitura = sLeitura & "------------------------------------------------------------------------------------------------------" & vbCrLf
            sLeitura = sLeitura & "CIRCUNFERÊNCIAS CORPORAIS                  DIÂMETROS ÓSSEOS(mm)" & vbCrLf
            sLeitura = sLeitura & "Tornozelo: " & (txtTornozelo.Text.Trim & Space(6)).Substring(0, 6) & Space(26) & "Fêmur: " & TxtFemur.Text & Space(10) & vbCrLf
            sLeitura = sLeitura & "Joelho:    " & (TxtJoelho.Text.Trim & Space(6)).Substring(0, 6) & Space(26) & "Rádio: " & TxtRadio.Text & Space(10) & vbCrLf
            sLeitura = sLeitura & "Punho:     " & (TxtPunho.Text.Trim & Space(6)).Substring(0, 6) & Space(26) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "RISCOS DE DOENÇAS CRÔNICAS (CINTURA X QUADRIL)" & vbCrLf
            sLeitura = sLeitura & "Cintura: " & (TxtCintura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Relação cintura ao quadril: " & TxtRelCinQua.Text & vbCrLf
            sLeitura = sLeitura & "Quadril: " & (TxtQuadril.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Risco: " & LblRelCinQua.Text & vbCrLf & vbCrLf

            sLeitura = sLeitura & "COMPOSIÇÃO CORPORAL" & vbCrLf
            sLeitura = sLeitura & "% de gordura:  " & (TxtPorcGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "MASSA PROVÁVEL DAS VÁRIAS PARTES DO CORPO (Kg)" & vbCrLf
            sLeitura = sLeitura & "Peso Gordura:  " & (TxtPesoGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Cabeça:    " & (TxtCabeca.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Mão:   " & TxtMao.Text & vbCrLf
            sLeitura = sLeitura & "Peso Muscular: " & (TxtPesoMusc.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Tronco:    " & (TxtTronco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Coxa:  " & TxtCoxa.Text & vbCrLf
            sLeitura = sLeitura & "Peso Residual: " & (TxtPesoResid.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Braço:     " & (TxtBraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Perna: " & TxtPerna.Text & vbCrLf
            sLeitura = sLeitura & "Peso Ósseo:    " & (TxtPesoOsseo.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Antebraço: " & (TxtAntebraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Pé:    " & TxtPe.Text & vbCrLf
            sLeitura = sLeitura & "M. C. M:       " & (TxtMcm.Text.Trim & Space(6)).Substring(0, 6) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "PESO IDEAL" & vbCrLf
            sLeitura = sLeitura & "Normal: " & (TxtPesoNormal.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Imc:           " & (TxtIMC.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResIMC.Text & vbCrLf
            sLeitura = sLeitura & "Mínimo: " & (TxtPesoMinimo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Ossatura:      " & (TxtOssatura.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResOssatura.Text & vbCrLf
            sLeitura = sLeitura & "Máximo: " & (TxtPesoMaximo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Compl. Físico: " & (TxtCFisico.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResCFisico.Text & vbCrLf & vbCrLf & vbCrLf

            sLeitura = sLeitura & "                                     ____________________________________________________________" & vbCrLf
            sLeitura = sLeitura & "                                         " & "Ass: " & (TxtProf.Text.Trim & Space(40)).Substring(0, 40)
            PrintDocument1.DefaultPageSettings = PrintPageSettings

            'Especifica o documento para a caixa de diálogo de impressão e exibe 
            StringToPrint = sLeitura
            PrintDialog1.Document = PrintDocument1

            'Se clicar em OK, imprime o documento para impressora
            If result = Windows.Forms.DialogResult.OK Then
                PrintDocument1.Print()
            End If
        Catch ex As Exception
            'Exibe mensagem de erro
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub AdipômetroToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdipômetroToolStripMenuItem.Click
        Dim result As DialogResult = PrintDialog1.ShowDialog()

        Try
            '====================================================================================================
            '===============================================SLEITURA=============================================
            sLeitura = "                                                                                   " & Date.Now & vbCrLf
            sLeitura = sLeitura & "                              AVALIAÇÃO FÍSICA - ADIPÔMETRO" & vbCrLf
            '=============================================102 ESPAÇOS============================================
            sLeitura = sLeitura & "Professor: " & (TxtProf.Text.Trim & Space(40)).Substring(0, 40) & Space(10) & (TxtLocal.Text & Space(40)).Substring(0, 40) & vbCrLf
            sLeitura = sLeitura & "======================================================================================================" & vbCrLf
            sLeitura = sLeitura & " NOME:" & (TxtNome.Text.Trim & Space(30)).Substring(0, 30) & Space(2) & "IDADE:" & TxtIdade.Text & Space(2) & "SEXO:" & CbxSexo.Text & Space(2) & "PESO:" & TxtPeso.Text & Space(2) & "ALTURA:" & TxtAltura.Text & Space(2) & "DN:" & DtpDataNasc.Value & vbCrLf
            sLeitura = sLeitura & "------------------------------------------------------------------------------------------------------" & vbCrLf
            sLeitura = sLeitura & "CIRCUNFERÊNCIAS CORPORAIS            DIÂMETROS ÓSSEOS(mm)        Dobras" & vbCrLf
            sLeitura = sLeitura & "Tornozelo: " & (txtTornozelo.Text.Trim & Space(6)).Substring(0, 6) & Space(20) & "Fêmur: " & (TxtFemur.Text.Trim & Space(21)).Substring(0, 21) & "Suprailíaca:  18   Abdomen:    19" & vbCrLf
            sLeitura = sLeitura & "Joelho:    " & (TxtJoelho.Text.Trim & Space(6)).Substring(0, 6) & Space(20) & "Rádio: " & (TxtRadio.Text.Trim & Space(21)).Substring(0, 21) & "Subescapular: 13   Tricipital: 20" & vbCrLf
            sLeitura = sLeitura & "Punho:     " & (TxtPunho.Text.Trim & Space(6)).Substring(0, 6) & Space(2) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "RISCOS DE DOENÇAS CRÔNICAS (CINTURA X QUADRIL)" & vbCrLf
            sLeitura = sLeitura & "Cintura: " & (TxtCintura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Relação cintura ao quadril: " & TxtRelCinQua.Text & vbCrLf
            sLeitura = sLeitura & "Quadril: " & (TxtQuadril.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Risco: " & LblRelCinQua.Text & vbCrLf & vbCrLf

            sLeitura = sLeitura & "COMPOSIÇÃO CORPORAL" & vbCrLf
            sLeitura = sLeitura & "% de gordura:  " & (TxtPorcGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "MASSA PROVÁVEL DAS VÁRIAS PARTES DO CORPO (Kg)" & vbCrLf
            sLeitura = sLeitura & "Peso Gordura:  " & (TxtPesoGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Cabeça:    " & (TxtCabeca.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Mão:   " & TxtMao.Text & vbCrLf
            sLeitura = sLeitura & "Peso Muscular: " & (TxtPesoMusc.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Tronco:    " & (TxtTronco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Coxa:  " & TxtCoxa.Text & vbCrLf
            sLeitura = sLeitura & "Peso Residual: " & (TxtPesoResid.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Braço:     " & (TxtBraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Perna: " & TxtPerna.Text & vbCrLf
            sLeitura = sLeitura & "Peso Ósseo:    " & (TxtPesoOsseo.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Antebraço: " & (TxtAntebraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Pé:    " & TxtPe.Text & vbCrLf
            sLeitura = sLeitura & "M. C. M:       " & (TxtMcm.Text.Trim & Space(6)).Substring(0, 6) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "PESO IDEAL" & vbCrLf
            sLeitura = sLeitura & "Normal: " & (TxtPesoNormal.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Imc:           " & (TxtIMC.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResIMC.Text & vbCrLf
            sLeitura = sLeitura & "Mínimo: " & (TxtPesoMinimo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Ossatura:      " & (TxtOssatura.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResOssatura.Text & vbCrLf
            sLeitura = sLeitura & "Máximo: " & (TxtPesoMaximo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Compl. Físico: " & (TxtCFisico.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResCFisico.Text & vbCrLf & vbCrLf & vbCrLf

            sLeitura = sLeitura & "                                     ____________________________________________________________" & vbCrLf
            sLeitura = sLeitura & "                                         " & "Ass: " & (TxtProf.Text.Trim & Space(40)).Substring(0, 40)
            PrintDocument1.DefaultPageSettings = PrintPageSettings

            'Especifica o documento para a caixa de diálogo de impressão e exibe 
            StringToPrint = sLeitura
            PrintDialog1.Document = PrintDocument1

            'Se clicar em OK, imprime o documento para impressora
            If result = Windows.Forms.DialogResult.OK Then
                PrintDocument1.Print()
            End If
        Catch ex As Exception
            'Exibe mensagem de erro
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub ImpedânciaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImpedânciaToolStripMenuItem.Click
        Dim result As DialogResult = PrintDialog1.ShowDialog()

        Try
            '====================================================================================================
            '===============================================SLEITURA=============================================
            sLeitura = "                                                                                   " & Date.Now & vbCrLf
            sLeitura = sLeitura & "                              AVALIAÇÃO FÍSICA - ADIPÔMETRO" & vbCrLf
            '=============================================102 ESPAÇOS============================================
            sLeitura = sLeitura & "Professor: " & (TxtProf.Text.Trim & Space(40)).Substring(0, 40) & Space(10) & (TxtLocal.Text & Space(40)).Substring(0, 40) & vbCrLf
            sLeitura = sLeitura & "======================================================================================================" & vbCrLf
            sLeitura = sLeitura & " NOME:" & (TxtNome.Text.Trim & Space(30)).Substring(0, 30) & Space(2) & "IDADE:" & TxtIdade.Text & Space(2) & "SEXO:" & CbxSexo.Text & Space(2) & "PESO:" & TxtPeso.Text & Space(2) & "ALTURA:" & TxtAltura.Text & Space(2) & "DN:" & DtpDataNasc.Value & vbCrLf
            sLeitura = sLeitura & "------------------------------------------------------------------------------------------------------" & vbCrLf
            sLeitura = sLeitura & "IMPEDÂNCIA BIOELÈTRICA                  DIÂMETROS ÓSSEOS(mm)" & vbCrLf
            sLeitura = sLeitura & "% Gordura: 55" & Space(27) & "Fêmur: " & (TxtFemur.Text.Trim & Space(21)).Substring(0, 21) & vbCrLf
            sLeitura = sLeitura & Space(40) & "Rádio: " & (TxtRadio.Text.Trim & Space(21)).Substring(0, 21) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "RISCOS DE DOENÇAS CRÔNICAS (CINTURA X QUADRIL)" & vbCrLf
            sLeitura = sLeitura & "Cintura: " & (TxtCintura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Relação cintura ao quadril: " & TxtRelCinQua.Text & vbCrLf
            sLeitura = sLeitura & "Quadril: " & (TxtQuadril.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Risco: " & LblRelCinQua.Text & vbCrLf & vbCrLf

            sLeitura = sLeitura & "COMPOSIÇÃO CORPORAL" & vbCrLf
            sLeitura = sLeitura & "% de gordura:  " & (TxtPorcGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "MASSA PROVÁVEL DAS VÁRIAS PARTES DO CORPO (Kg)" & vbCrLf
            sLeitura = sLeitura & "Peso Gordura:  " & (TxtPesoGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Cabeça:    " & (TxtCabeca.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Mão:   " & TxtMao.Text & vbCrLf
            sLeitura = sLeitura & "Peso Muscular: " & (TxtPesoMusc.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Tronco:    " & (TxtTronco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Coxa:  " & TxtCoxa.Text & vbCrLf
            sLeitura = sLeitura & "Peso Residual: " & (TxtPesoResid.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Braço:     " & (TxtBraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Perna: " & TxtPerna.Text & vbCrLf
            sLeitura = sLeitura & "Peso Ósseo:    " & (TxtPesoOsseo.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Antebraço: " & (TxtAntebraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Pé:    " & TxtPe.Text & vbCrLf
            sLeitura = sLeitura & "M. C. M:       " & (TxtMcm.Text.Trim & Space(6)).Substring(0, 6) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "PESO IDEAL" & vbCrLf
            sLeitura = sLeitura & "Normal: " & (TxtPesoNormal.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Imc:           " & (TxtIMC.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResIMC.Text & vbCrLf
            sLeitura = sLeitura & "Mínimo: " & (TxtPesoMinimo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Ossatura:      " & (TxtOssatura.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResOssatura.Text & vbCrLf
            sLeitura = sLeitura & "Máximo: " & (TxtPesoMaximo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Compl. Físico: " & (TxtCFisico.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & TxtResCFisico.Text & vbCrLf & vbCrLf & vbCrLf

            sLeitura = sLeitura & "                                     ____________________________________________________________" & vbCrLf
            sLeitura = sLeitura & "                                         " & "Ass: " & (TxtProf.Text.Trim & Space(40)).Substring(0, 40)
            PrintDocument1.DefaultPageSettings = PrintPageSettings

            'Especifica o documento para a caixa de diálogo de impressão e exibe 
            StringToPrint = sLeitura
            PrintDialog1.Document = PrintDocument1

            'Se clicar em OK, imprime o documento para impressora
            If result = Windows.Forms.DialogResult.OK Then
                PrintDocument1.Print()
            End If
        Catch ex As Exception
            'Exibe mensagem de erro
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class