Imports System.Drawing.Printing
Imports System.IO
Public Class frm_AvaliacaoFisica
    Dim vPerseteInic As Integer = 0
    Dim vProcessar As Boolean = True
    Dim vPunho As Double = 0
    Dim sLeitura As String = ""

    Private StringToPrint As String
    Private PrintFont As New Font("Consolas", 12)
    Private PrintFont2 As New Font("Consolas", 7.5)
    Private PrintPageSettings As New PageSettings


    ''' <summary>
    ''' Idade do Aluno
    ''' </summary>
    ''' <remarks>Calcula a Idade do Aluno</remarks>
    Private Sub dtpDataNasc_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpDataNasc.ValueChanged
        Dim vAno As Double
        Dim vDataNasc As Date = dtpDataNasc.Text

        vAno = Now.Date.Year - vDataNasc.Date.Year
        txtIdade.Text = vAno
    End Sub


    ''' <summary>
    ''' Testes Físicos (Salto Vertical, Abdominais, Flexão de Braço, Cooper e VO2Max)
    ''' </summary>
    ''' <remarks>Testes Físicos</remarks>
    Private Sub Salto_Vertical()

        If cbxSexo.Text = "Masculino" Then
            If CDec(CDec(txtIdade.Text)) > 6 Then
                '=======================================================================================
                '===================================IDADE DE 7 A 10 ANOS================================
                '=======================================================================================
                If CDec(CDec(txtIdade.Text)) <= 10 Then
                    If CDec(txtSaltoVert.Text) <= 16 Then
                        txtClasv.Text = "Abaixo da Média"
                    ElseIf CDec(txtSaltoVert.Text) <= 20 Then
                        If CDec(CDec(txtIdade.Text)) <= 8 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Abaixo da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 27 Then
                        txtClasv.Text = "Média"
                    ElseIf CDec(txtSaltoVert.Text) <= 30 Then
                        If CDec(CDec(txtIdade.Text)) <= 8 Then
                            txtClasv.Text = "Acima da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    Else
                        txtClasv.Text = "Acima da Média"
                    End If

                    '=======================================================================================
                    '===================================IDADE DE 7 A 16 ANOS================================
                    '=======================================================================================
                ElseIf CDec(CDec(txtIdade.Text)) <= 16 Then
                    If CDec(txtSaltoVert.Text) <= 33 Then
                        txtClasv.Text = "Fraco"
                    ElseIf CDec(txtSaltoVert.Text) <= 36 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Regular"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 37 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Bom"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 40 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Bom"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Regular"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 43 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Muito Bom"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Regular"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 44 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Bom"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 49 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Bom"
                        Else
                            txtClasv.Text = "Regular"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 54 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Muito Bom"
                        Else
                            txtClasv.Text = "Bom"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 55 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        Else
                            txtClasv.Text = "Muito Bom"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 59 Then
                        If CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Excelente"
                        Else
                            txtClasv.Text = "Muito Bom"
                        End If
                    Else
                        txtClasv.Text = "Excelente"
                    End If
                    '=======================================================================================
                    '===================================IDADE DE 17 A 70 ANOS================================
                    '=======================================================================================
                Else
                    If CDec(txtSaltoVert.Text) <= 23 Then
                        txtClasv.Text = "Abaixo da Média"
                    ElseIf CDec(txtSaltoVert.Text) <= 30 Then
                        If CDec(CDec(txtIdade.Text)) <= 50 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 32 Then
                        If CDec(CDec(txtIdade.Text)) <= 40 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 36 Then
                        If CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 37 Then
                        If CDec(CDec(txtIdade.Text)) <= 20 Then
                            txtClasv.Text = "Média"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 38 Then
                        If CDec(CDec(txtIdade.Text)) <= 20 Then
                            txtClasv.Text = "Média"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Abaixo da Média"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 50 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 44 Then
                        If CDec(CDec(txtIdade.Text)) <= 50 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 47 Then
                        If CDec(CDec(txtIdade.Text)) <= 40 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 48 Then
                        If CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 52 Then
                        If CDec(CDec(txtIdade.Text)) <= 20 Then
                            txtClasv.Text = "Acima da Média"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    Else
                        txtClasv.Text = "Acima da Média"
                    End If
                End If
            Else
                txtClasv.Text = "Não Avaliar"
            End If
        Else

            If CDec(CDec(txtIdade.Text)) > 6 Then
                '=======================================================================================
                '===================================IDADE DE 7 A 10 ANOS================================
                '=======================================================================================
                If CDec(CDec(txtIdade.Text)) <= 10 Then
                    If CDec(txtSaltoVert.Text) <= 17 Then
                        txtClasv.Text = "Abaixo da Média"
                    ElseIf CDec(txtSaltoVert.Text) <= 20 Then
                        If CDec(CDec(txtIdade.Text)) <= 8 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Abaixo da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 23 Then
                        If CDec(CDec(txtIdade.Text)) = 9 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 25 Then
                        txtClasv.Text = "Média"
                    ElseIf CDec(txtSaltoVert.Text) <= 28 Then
                        If CDec(CDec(txtIdade.Text)) = 7 Then
                            txtClasv.Text = "Acima da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 29 Then
                        If CDec(CDec(txtIdade.Text)) <= 8 Then
                            txtClasv.Text = "Acima da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 30 Then
                        If CDec(CDec(txtIdade.Text)) = 10 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    Else
                        txtClasv.Text = "Acima da Média"
                    End If

                    '=======================================================================================
                    '===================================IDADE DE 7 A 16 ANOS================================
                    '=======================================================================================
                ElseIf CDec(CDec(txtIdade.Text)) <= 16 Then
                    If CDec(txtSaltoVert.Text) <= 28 Then
                        txtClasv.Text = "Fraco"
                    ElseIf CDec(txtSaltoVert.Text) <= 32 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Regular"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 34 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Bom"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 36 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Bom"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Regular"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 38 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Muito Bom"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Regular"
                        Else
                            txtClasv.Text = "Fraco"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 39 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Muito Bom"
                        Else
                            txtClasv.Text = "Regular"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 40 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Muito Bom"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Bom"
                        Else
                            txtClasv.Text = "Regular"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 42 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Bom"
                        Else
                            txtClasv.Text = "Regular"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 44 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        Else
                            txtClasv.Text = "Bom"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 46 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Muito Bom"
                        Else
                            txtClasv.Text = "Bom"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 49 Then
                        If CDec(CDec(txtIdade.Text)) <= 12 Then
                            txtClasv.Text = "Excelente"
                        Else
                            txtClasv.Text = "Muito Bom"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 50 Then
                        If CDec(CDec(txtIdade.Text)) <= 14 Then
                            txtClasv.Text = "Excelente"
                        Else
                            txtClasv.Text = "Muito Bom"
                        End If
                    Else
                        txtClasv.Text = "Excelente"
                    End If
                    '=======================================================================================
                    '===================================IDADE DE 17 A 70 ANOS================================
                    '=======================================================================================
                Else
                    If CDec(txtSaltoVert.Text) <= 16 Then
                        txtClasv.Text = "Abaixo da Média"
                    ElseIf CDec(txtSaltoVert.Text) <= 21 Then
                        If CDec(CDec(txtIdade.Text)) <= 50 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 23 Then
                        If CDec(CDec(txtIdade.Text)) <= 40 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 26 Then
                        If CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Abaixo da Média"
                        Else
                            txtClasv.Text = "Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 27 Then
                        If CDec(CDec(txtIdade.Text)) <= 20 Then
                            txtClasv.Text = "Abaixo da Média"
                        ElseIf CDec(CDec(txtIdade.Text)) <= 50 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 34 Then
                        If CDec(CDec(txtIdade.Text)) <= 40 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) <= 38 Then
                        If CDec(CDec(txtIdade.Text)) <= 30 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    ElseIf CDec(txtSaltoVert.Text) = 39 Then
                        If CDec(CDec(txtIdade.Text)) <= 20 Then
                            txtClasv.Text = "Média"
                        Else
                            txtClasv.Text = "Acima da Média"
                        End If
                    Else
                        txtClasv.Text = "Acima da Média"
                    End If
                End If
            Else
                txtClasv.Text = "Não Avaliar"
            End If
        End If
    End Sub
    Private Sub Abdominais()
        If cbxSexo.Text = "Masculino" Then
            If CDec(CDec(txtIdade.Text)) < 10 Then
                txtClaAbd.Text = "Deficiente"
            Else
                If CDec(txtAbdominais.Text) = 8 Then
                    txtClaAbd.Text = "Deficientes"
                ElseIf CDec(txtAbdominais.Text) <= 11 Then
                    If CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Deficientes"
                    Else
                        txtClaAbd.Text = "Fraco"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 12 Then
                    If CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Deficientes"
                    Else
                        txtClaAbd.Text = "Fraco"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 16 Then
                    If CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Deficientes"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Fraco"
                    Else
                        txtClaAbd.Text = "Regular"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 18 Then
                    If CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Fraco"
                    Else
                        txtClaAbd.Text = "Regular"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 20 Then
                    If CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Regular"
                    Else
                        txtClaAbd.Text = "Bom"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 21 Then
                    If CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Regular"
                    Else
                        txtClaAbd.Text = "Bom"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 22 Then
                    If CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Fraco"
                    Else
                        txtClaAbd.Text = "Regular"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 24 Then
                    If CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Regular"
                    Else
                        txtClaAbd.Text = "Bom"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 25 Then
                    If CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Regular"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 26 Then
                    If CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 28 Then
                    If CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Regular"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 29 Then
                    If CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Regular"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 31 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Regular"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 34 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 36 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Regular"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) = 37 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Deficiente"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Regular"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 39 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Regular"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 42 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Fraco"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 44 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Regular"
                    ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 47 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Regular"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                ElseIf CDec(txtAbdominais.Text) <= 52 Then
                    If CDec(CDec(txtIdade.Text)) <= 19 Then
                        txtClaAbd.Text = "Bom"
                    Else
                        txtClaAbd.Text = "Excelente"
                    End If
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            End If
        Else
            If CDec(txtAbdominais.Text) <= 9 Then
                If CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Deficiente"
                Else
                    txtClaAbd.Text = "Fraco"
                End If
            ElseIf CDec(txtAbdominais.Text) = 10 Then
                If CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Deficiente"
                Else
                    txtClaAbd.Text = "Regular"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 14 Then
                If CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Fraco"
                Else
                    txtClaAbd.Text = "Regular"
                End If
            ElseIf CDec(txtAbdominais.Text) = 15 Then
                If CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Regular"
                Else
                    txtClaAbd.Text = "Bom"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 19 Then
                If CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Fraco"
                ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Regular"
                Else
                    txtClaAbd.Text = "Bom"
                End If
            ElseIf CDec(txtAbdominais.Text) = 20 Then
                If CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Regular"
                ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 24 Then
                If CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Fraco"
                ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Regular"
                ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) = 25 Then
                If CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Regular"
                ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 29 Then
                If CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Fraco"
                ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Regular"
                ElseIf CDec(CDec(txtIdade.Text)) <= 49 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) = 26 Then
                If CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Deficiente"
                ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Fraco"
                ElseIf CDec(CDec(txtIdade.Text)) <= 59 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 34 Then
                If CDec(CDec(txtIdade.Text)) <= 19 Then
                    txtClaAbd.Text = "Fraco"
                ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Regular"
                ElseIf CDec(CDec(txtIdade.Text)) <= 39 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 39 Then
                If CDec(CDec(txtIdade.Text)) <= 19 Then
                    txtClaAbd.Text = "Regular"
                ElseIf CDec(CDec(txtIdade.Text)) <= 29 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            ElseIf CDec(txtAbdominais.Text) <= 44 Then
                If CDec(CDec(txtIdade.Text)) <= 19 Then
                    txtClaAbd.Text = "Bom"
                Else
                    txtClaAbd.Text = "Excelente"
                End If
            Else
                txtClaAbd.Text = "Excelente"
            End If
        End If
    End Sub
    Private Sub Flexao_Braco()

        'Aluno do Sexo Masculino
        If cbxSexo.Text = "Masculino" Then
            If CDec(txtFlexaoBraco.Text) <= 4 Then
                txtClaFle.Text = "Deficiente"
            Else
                If CDec(txtFlexaoBraco.Text) <= 7 Then
                    If CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Deficiente"
                    Else
                        txtClaFle.Text = "Fraco"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 9 Then
                    If CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Deficiente"
                    Else
                        txtClaFle.Text = "Fraco"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 10 Then
                    If CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 13 Then
                    If CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 14 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 16 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 17 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 21 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 23 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 24 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 26 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 29 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 30 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 34 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 39 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 49 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 49 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 59 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                Else
                    txtClaFle.Text = "Excelente"
                End If
            End If
        Else
            'Aluno do Sexo Feminino
            If CDec(txtFlexaoBraco.Text) <= 1 Then
                txtClaFle.Text = "Deficiente"
            Else
                If CDec(txtFlexaoBraco.Text) = 2 Then
                    If CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Deficiente"
                    Else
                        txtClaFle.Text = "Fraco"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 3 Then
                    If CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Deficiente"
                    Else
                        txtClaFle.Text = "Fraco"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 4 Then
                    If CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Deficiente"
                    Else
                        txtClaFle.Text = "Fraco"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 5 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Deficiente"
                    Else
                        txtClaFle.Text = "Fraco"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 6 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 7 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 9 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Deficiente"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 12 Then
                    If CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Fraco"
                    Else
                        txtClaFle.Text = "Regular"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 15 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 17 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 19 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Regular"
                    Else
                        txtClaFle.Text = "Bom"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 20 Then
                    If CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 23 Then
                    If CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 26 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 28 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) = 29 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Regular"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 31 Then
                    If CDec(txtIdade.Text) < 50 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 34 Then
                    If CDec(txtIdade.Text) < 40 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 37 Then
                    If CDec(txtIdade.Text) < 30 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                ElseIf CDec(txtFlexaoBraco.Text) <= 40 Then
                    If CDec(txtIdade.Text) < 20 Then
                        txtClaFle.Text = "Bom"
                    Else
                        txtClaFle.Text = "Excelente"
                    End If
                Else
                    txtClaFle.Text = "Excelente"
                End If
            End If
        End If
    End Sub
    Private Sub Cooper()
        If CDec(txtCooper.Text) < 860 Then
            txtClaCoo.Text = "Muito Fraco"
        Else
            If cbxSexo.Text = "Masculino" Then
                If CDec(txtCooper.Text) < 1260 Then
                    txtClaCoo.Text = "Muito Fraco"
                ElseIf CDec(txtIdade.Text) < 7 Then
                    txtClaCoo.Text = "Não Avaliar"
                ElseIf CDec(txtCooper.Text) <= 1340 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1360 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1370 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) = 8 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1400 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1420 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1440 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1500 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1540 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1600 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) = 8 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1640 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1660 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1680 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1690 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1730 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1740 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) = 8 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1780 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1810 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1820 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1870 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1880 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1900 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1930 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1950 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1970 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1990 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2000 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2090 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2110 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2120 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2200 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2240 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2320 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2390 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2400 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2450 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2460 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 9 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2480 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 9 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2510 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 9 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2520 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 9 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2540 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2640 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2650 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2710 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2770 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2830 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 3000 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                Else
                    txtClaCoo.Text = "Superior"
                End If
            Else
                'Feminino
                If CDec(txtIdade.Text) < 7 Then
                    txtClaCoo.Text = "Não Avaliar"
                ElseIf CDec(txtCooper.Text) <= 1110 Then
                    txtClaCoo.Text = "Muito Fraco"
                ElseIf CDec(txtCooper.Text) <= 1180 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1230 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1260 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Muito Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1280 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1300 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) <= 11 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1320 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) <= 12 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1350 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1380 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1390 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    Else
                        txtClaCoo.Text = "Fraco"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1410 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1420 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1440 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1450 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1500 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1510 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1530 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1540 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1550 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1560 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1580 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1590 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    Else
                        txtClaCoo.Text = "Média"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1600 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1610 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) = 8 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Muito Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1640 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) = 8 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1690 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1700 Then
                    If CDec(txtIdade.Text) < 9 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1750 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    Else
                        txtClaCoo.Text = "Boa"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1780 Then
                    If CDec(txtIdade.Text) < 10 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1790 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1830 Then
                    If CDec(txtIdade.Text) < 11 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1900 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Fraco"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Boa"
                    Else
                        txtClaCoo.Text = "Excelente"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1930 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) = 12 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1960 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 1970 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2000 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2080 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Média"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2090 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 60 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2160 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 50 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2180 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2240 Then
                    If CDec(txtIdade.Text) = 11 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 40 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2260 Then
                    If CDec(txtIdade.Text) = 10 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 11 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2300 Then
                    If CDec(txtIdade.Text) = 8 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 10 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) = 11 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Boa"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2330 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 30 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2340 Then
                    If CDec(txtIdade.Text) = 7 Then
                        txtClaCoo.Text = "Excelente"
                    ElseIf CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2370 Then
                    If CDec(txtIdade.Text) < 12 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                ElseIf CDec(txtCooper.Text) <= 2430 Then
                    If CDec(txtIdade.Text) < 13 Then
                        txtClaCoo.Text = "Superior"
                    ElseIf CDec(txtIdade.Text) < 20 Then
                        txtClaCoo.Text = "Excelente"
                    Else
                        txtClaCoo.Text = "Superior"
                    End If
                Else
                    txtClaCoo.Text = "Superior"
                End If

            End If
        End If

    End Sub
    Private Sub VO2Max()
        txtVO2.Text = String.Format("{0:N2}", (CDec(txtCooper.Text) - 504) / 45)
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

        If CDec(txtPunho.Text) < 11 Then
            vPunho = 4
        ElseIf CDec(txtPunho.Text) = 11 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(txtPunho.Text) = 12 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 3
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(txtPunho.Text) = 13 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 2.5
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 3
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(txtPunho.Text) = 14 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 2
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 2.5
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 3
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 3.5
            Else
                vPunho = 4
            End If
        ElseIf CDec(txtPunho.Text) = 15 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 1.5
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 2
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 2.5
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 3
            Else
                vPunho = 3.5
            End If
        ElseIf CDec(txtPunho.Text) = 16 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 1
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 1.5
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 2
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 2.5
            Else
                vPunho = 3
            End If
        ElseIf CDec(txtPunho.Text) = 17 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 0.5
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 1
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 1.5
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 2
            Else
                vPunho = 2.5
            End If
        ElseIf CDec(txtPunho.Text) = 18 Then
            If CDec(txtAltura.Text) * 100 <= 150 Then
                vPunho = 0
            ElseIf CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 0.5
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 1
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 1.5
            Else
                vPunho = 2
            End If
        ElseIf CDec(txtPunho.Text) = 19 Then
            If CDec(txtAltura.Text) * 100 <= 160 Then
                vPunho = 0
            ElseIf CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 0.5
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 1
            Else
                vPunho = 1.5
            End If
        ElseIf CDec(txtPunho.Text) = 20 Then
            If CDec(txtAltura.Text) * 100 <= 170 Then
                vPunho = 0
            ElseIf CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 0.5
            Else
                vPunho = 1
            End If
        ElseIf CDec(txtPunho.Text) = 21 Then
            If CDec(txtAltura.Text) * 100 <= 180 Then
                vPunho = 0
            Else
                vPunho = 0.5
            End If
        Else
            vPunho = 0
        End If

        If cbxSexo.Text = "Masculino" Then
            vPorGordura = ((vPunho + 115) + CDec(txtPeso.Text)) - (CDec(txtAltura.Text) * 100)

        ElseIf cbxSexo.Text = "Feminino" Then
            vPorGordura = ((vPunho + 125) + CDec(txtPeso.Text)) - (CDec(txtAltura.Text) * 100)
        Else
            vPorGordura = 6
        End If

        If vPorGordura < 6 Then
            vPorGordura = 6
        End If

        txtPorcGordura.Text = vPorGordura
        '=========================================================================================================
        '========================================PESO DA GORDURA==================================================
        '=========================================================================================================
        txtPesoGordura.Text = String.Format("{0:N3}", vPorGordura / 100) * CDec(txtPeso.Text)
        '=========================================================================================================
        '=========================================PESO RESIDUAL===================================================
        '=========================================================================================================
        If cbxSexo.Text = "Masculino" Then
            txtPesoResid.Text = String.Format("{0:N3}", CDec(txtPeso.Text) * 0.241)
        ElseIf cbxSexo.Text = "Feminino" Then
            txtPesoResid.Text = String.Format("{0:N3}", CDec(txtPeso.Text) * 0.209)
        Else
            txtPesoResid.Text = 0
        End If
        '=========================================================================================================
        '==========================================PESO ÓSSEO=====================================================
        '=========================================================================================================
        txtPesoOsseo.Text = String.Format("{0:N3}", 3.02 * ((CDec(txtAltura.Text) * CDec(txtAltura.Text)) * (CDec(txtFemur.Text) * 0.001) * (CDec(txtRadio.Text) * 0.001) * 400) ^ 0.712)
        '=========================================================================================================
        '=========================================PESO MUSCULAR===================================================
        '=========================================================================================================
        txtPesoMusc.Text = CDec(txtPeso.Text) - (CDec(txtPesoGordura.Text) + CDec(txtPesoResid.Text) + CDec(txtPesoOsseo.Text))
        '=========================================================================================================
        '=============================================M.C.M=======================================================
        '=========================================================================================================
        txtMcm.Text = String.Format("{0:N1}", CDec(txtPeso.Text) - CDec(txtPesoGordura.Text))
    End Sub


    ''' <summary>
    ''' Massa Provável das Várias Partes do Corpo
    ''' </summary>
    ''' <remarks>Massa Partes do Corpo</remarks>
    Private Sub MassaPartesCorpo()
        'Cabeça
        txtCabeca.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 7.3) / 100)
        'Tronto
        txtTronco.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 50.7) / 100)
        'Braço
        txtBraco.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 2.6) / 100)
        'Antebraço
        txtAntebraco.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 1.6) / 100)
        'Mão
        txtMao.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 0.7) / 100)
        'Coxa
        txtCoxa.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 10.3) / 100)
        'Perna
        txtPerna.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 4.3) / 100)
        'Pé
        txtPe.Text = String.Format("{0:N3}", (CDec(txtPeso.Text) * 1.5) / 100)
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
        txtPesoNormal.Text = String.Format("{0:N2}", (((CDec(txtAltura.Text) * 100) - 100) + 4 * CDec(txtPunho.Text)) / 2)

        'Mínimo
        txtPesoMinimo.Text = String.Format("{0:N2}", CDec(txtPesoNormal.Text) - ((CDec(txtPesoNormal.Text) / 100) * 10))

        'Máximo
        txtPesoMaximo.Text = String.Format("{0:N2}", ((CDec(txtPesoNormal.Text) / 100) * 10) + CDec(txtPesoNormal.Text))

        'IMC
        txtIMC.Text = String.Format("{0:N2}", CDec(txtPeso.Text) / (CDec(txtAltura.Text) * CDec(txtAltura.Text)))
        vImc = String.Format("{0:N2}", CDec(txtPeso.Text) / (CDec(txtAltura.Text) * CDec(txtAltura.Text)))

        If cbxSexo.Text = "Masculino" Then
            If CDec(txtIMC.Text) = 22 Then
                txtResIMC.Text = "IDEAL"
            ElseIf CDec(txtIMC.Text) < 20 Then
                txtResIMC.Text = "ABAIXO DO NORMAL"
            ElseIf CDec(txtIMC.Text) <= 24.9 Then
                txtResIMC.Text = "NORMAL"
            ElseIf CDec(txtIMC.Text) = 29.9 Then
                txtResIMC.Text = "OBESO -- 1"
            ElseIf CDec(txtIMC.Text) <= 40 Then
                txtResIMC.Text = "OBESO -- 2"
            ElseIf CDec(txtIMC.Text) > 40 Then
                txtResIMC.Text = "OBESO -- 3"
            End If
        Else
            If CDec(txtIMC.Text) = 21 Then
                txtResIMC.Text = "IDEAL"
            ElseIf CDec(txtIMC.Text) < 19 Then
                txtResIMC.Text = "ABAIXO DO NORMAL"
            ElseIf CDec(txtIMC.Text) <= 23.9 Then
                txtResIMC.Text = "NORMAL"
            ElseIf CDec(txtIMC.Text) = 28.9 Then
                txtResIMC.Text = "OBESA -- 1"
            ElseIf CDec(txtIMC.Text) <= 39 Then
                txtResIMC.Text = "OBESA -- 2"
            ElseIf CDec(txtIMC.Text) > 39 Then
                txtResIMC.Text = "OBESA -- 3"
            End If
        End If

        'Ossatura
        txtOssatura.Text = String.Format("{0:N1}", (CDec(txtTornozelo.Text) + CDec(txtJoelho.Text) + CDec(txtPunho.Text)) / CDec(txtAltura.Text))
        vOssatura = String.Format("{0:N2}", (CDec(txtTornozelo.Text) + CDec(txtJoelho.Text) + CDec(txtPunho.Text)) / CDec(txtAltura.Text))

        If vOssatura < 43 Then
            txtResOssatura.Text = "FRACA"
        ElseIf vOssatura < 46 Then
            txtResOssatura.Text = "MÉDIA"
        Else
            txtResOssatura.Text = "FORTE"
        End If

        'Complemento Físico
        txtCFisico.Text = String.Format("{0:N2}", (CDec(txtAltura.Text) * 100) / CDec(txtPunho.Text))
        vComplFisico = String.Format("{0:N2}", (CDec(txtAltura.Text) * 100) / CDec(txtPunho.Text))

        If cbxSexo.Text = "Masculino" Then
            If vComplFisico > 10.4 Then
                txtResCFisico.Text = "PEQUENA"
            ElseIf vComplFisico < 9.6 Then
                txtResCFisico.Text = "GRANDE"
            Else
                txtResCFisico.Text = "MÉDIA"
            End If
        Else
            If vComplFisico > 11 Then
                txtResCFisico.Text = "PEQUENA"
            ElseIf vComplFisico < 10.1 Then
                txtResCFisico.Text = "GRANDE"
            Else
                txtResCFisico.Text = "MÉDIA"
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
        vResultado = String.Format("{0:N2}", CDec(txtCintura.Text) / CDec(txtQuadril.Text))

        If cbxSexo.Text = "Masculino" Then
            If vResultado < 0.83 Then
                lblRelCinQua.Text = "Baixo"
            ElseIf vResultado = 0.83 Then
                If CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado < 0.88 Then
                If CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.88 Then
                If CDec(txtIdade.Text) < 50 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.89 Then
                If CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Elevado"
                ElseIf CDec(txtIdade.Text) < 50 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.9 Then
                If CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Elevado"
                ElseIf CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.91 Then
                If CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 0.95 Then
                If CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.95 Then
                If CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.96 Then
                If CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(txtIdade.Text) < 50 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 0.99 Then
                If CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 1.01 Then
                If CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Muito Elevado"
                Else
                    lblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 1.03 Then
                If CDec(txtIdade.Text) < 50 Then
                    lblRelCinQua.Text = "Muito Elevado"
                Else
                    lblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado = 1.03 Then
                If CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Muito Elevado"
                Else
                    lblRelCinQua.Text = "Elevado"
                End If
            Else
                lblRelCinQua.Text = "Muito Elevado"
            End If
        Else
            'Feminino
            If vResultado < 0.71 Then
                lblRelCinQua.Text = "Baixo"
            ElseIf vResultado = 0.71 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.72 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado = 0.73 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 50 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado < 0.76 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Moderado"
                Else
                    lblRelCinQua.Text = "Baixo"
                End If
            ElseIf vResultado < 0.78 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.8 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado < 0.82 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 50 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.82 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Baixo"
                ElseIf CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.83 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Moderado"
                ElseIf CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Muito Elevado"
                ElseIf CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Moderado"
                End If
            ElseIf vResultado = 0.84 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Moderado"
                ElseIf CDec(txtIdade.Text) < 30 Then
                    lblRelCinQua.Text = "Muito Elevado"
                Else
                    lblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 0.88 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Moderado"
                ElseIf CDec(txtIdade.Text) < 40 Then
                    lblRelCinQua.Text = "Muito Elevado"
                Else
                    lblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 0.91 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Elevado"
                ElseIf CDec(txtIdade.Text) < 60 Then
                    lblRelCinQua.Text = "Muito Elevado"
                Else
                    lblRelCinQua.Text = "Elevado"
                End If
            ElseIf vResultado < 0.95 Then
                If CDec(txtIdade.Text) < 20 Then
                    lblRelCinQua.Text = "Elevado"
                Else
                    lblRelCinQua.Text = "Muito Elevado"
                End If
            Else
                lblRelCinQua.Text = "Muito Elevado"
            End If

        End If



        txtRelCinQua.Text = vResultado
    End Sub


    ''' <summary>
    ''' Calcula toda a Prescrição do Treinamento
    ''' </summary>
    ''' <remarks>Precrição do Treinamento</remarks>
    Private Sub Prescricao_Treinamento()
        'Cálculo do Rítmo de Atividade
        Dim vRitmoAtiv As Double
        Dim vDistAtiv As Integer

        vRitmoAtiv = String.Format("{0:N2}", ((CDec(txtCooper.Text) - 504) / 45) * CDec(txtVo2Trab.Text) / 100)

        If vRitmoAtiv < 7 Then
            txtRitmoAtiv.Text = "Sem Valor"
        ElseIf vRitmoAtiv < 7.76 Then
            txtRitmoAtiv.Text = 50
        ElseIf vRitmoAtiv < 8.5 Then
            txtRitmoAtiv.Text = 55.5
        ElseIf vRitmoAtiv < 9 Then
            txtRitmoAtiv.Text = 60
        ElseIf vRitmoAtiv < 10.4 Then
            txtRitmoAtiv.Text = 62.5
        ElseIf vRitmoAtiv < 10.7 Then
            txtRitmoAtiv.Text = 70
        ElseIf vRitmoAtiv < 12.5 Then
            txtRitmoAtiv.Text = 71.4
        ElseIf vRitmoAtiv < 13.3 Then
            txtRitmoAtiv.Text = 80
        ElseIf vRitmoAtiv < 15 Then
            txtRitmoAtiv.Text = 83.3
        ElseIf vRitmoAtiv < 15.14 Then
            txtRitmoAtiv.Text = 95
        ElseIf vRitmoAtiv < 16.2 Then
            txtRitmoAtiv.Text = 90.9
        ElseIf vRitmoAtiv < 17.6 Then
            txtRitmoAtiv.Text = 95
        ElseIf vRitmoAtiv < 20.5 Then
            txtRitmoAtiv.Text = 100
        ElseIf vRitmoAtiv < 20.88 Then
            txtRitmoAtiv.Text = 110
        ElseIf vRitmoAtiv < 22.1 Then
            txtRitmoAtiv.Text = 111.11
        ElseIf vRitmoAtiv < 23 Then
            txtRitmoAtiv.Text = 115
        ElseIf vRitmoAtiv < 23.8 Then
            txtRitmoAtiv.Text = 117.6
        ElseIf vRitmoAtiv < 25.5 Then
            txtRitmoAtiv.Text = 120
        ElseIf vRitmoAtiv < 27.3 Then
            txtRitmoAtiv.Text = 125
        ElseIf vRitmoAtiv < 30.5 Then
            txtRitmoAtiv.Text = 130
        ElseIf vRitmoAtiv < 31.5 Then
            txtRitmoAtiv.Text = 135
        ElseIf vRitmoAtiv < 32.07 Then
            txtRitmoAtiv.Text = 140
        ElseIf vRitmoAtiv < 33.5 Then
            txtRitmoAtiv.Text = 142.85
        ElseIf vRitmoAtiv < 34.26 Then
            txtRitmoAtiv.Text = 150
        ElseIf vRitmoAtiv < 35.4 Then
            txtRitmoAtiv.Text = 153.84
        ElseIf vRitmoAtiv < 36.83 Then
            txtRitmoAtiv.Text = 160
        ElseIf vRitmoAtiv < 37.5 Then
            txtRitmoAtiv.Text = 166.66
        ElseIf vRitmoAtiv < 39.5 Then
            txtRitmoAtiv.Text = 170
        ElseIf vRitmoAtiv < 41.5 Then
            txtRitmoAtiv.Text = 180
        ElseIf vRitmoAtiv < 43.5 Then
            txtRitmoAtiv.Text = 190
        ElseIf vRitmoAtiv < 45.5 Then
            txtRitmoAtiv.Text = 200
        ElseIf vRitmoAtiv < 47.5 Then
            txtRitmoAtiv.Text = 210
        ElseIf vRitmoAtiv < 49.5 Then
            txtRitmoAtiv.Text = 220
        ElseIf vRitmoAtiv < 51.5 Then
            txtRitmoAtiv.Text = 230
        ElseIf vRitmoAtiv < 53.5 Then
            txtRitmoAtiv.Text = 240
        ElseIf vRitmoAtiv < 55.5 Then
            txtRitmoAtiv.Text = 250
        ElseIf vRitmoAtiv < 60.64 Then
            txtRitmoAtiv.Text = 260
        ElseIf vRitmoAtiv < 69.5 Then
            txtRitmoAtiv.Text = 285.7
        ElseIf vRitmoAtiv < 40.16 Then
            txtRitmoAtiv.Text = 330
        ElseIf vRitmoAtiv < 71.05 Then
            txtRitmoAtiv.Text = 333.3
        Else
            txtRitmoAtiv.Text = 340
        End If

        vDistAtiv = CDec(txtRitmoAtiv.Text) * CDec(txtTempoAtiv.Text)
        txtDistanciaAtiv.Text = vDistAtiv

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

        vCalculadoInicio = ((((CDec(txtCooper.Text) - 504) / 45) / 3.5) * 1.25)
        vCalculoFinal = ((vCalculadoInicio * CDec(txtPeso.Text)) / 60) * CDec(txtTempoAtiv.Text)
        vCalculoPorcentagem = (CDec(txtVo2Trab.Text)) / 100
        txtAtividade.Text = CInt(vCalculoFinal * vCalculoPorcentagem)

        '========== Calculandoo Repouso ==================
        ''=================== MASCULINO ==================
        If cbxSexo.Text = "Masculino" Then
            If CDec(txtIdade.Text) < 19 Then
                txtRepouso.Text = CInt(16.6 * CDec(txtPeso.Text) + (77 * CDec(txtAltura.Text)) + 572)
            ElseIf CDec(txtIdade.Text) < 31 Then
                txtRepouso.Text = CInt(15.4 * CDec(txtPeso.Text) + (27 * CDec(txtAltura.Text)) + 717)
            ElseIf CDec(txtIdade.Text) < 60 Then
                txtRepouso.Text = CInt(11.3 * CDec(txtPeso.Text) + (16 * CDec(txtAltura.Text)) + 901)
            Else
                txtRepouso.Text = CInt(8.8 * CDec(txtPeso.Text) + (11 * CDec(txtAltura.Text)) + 1071)
            End If
        Else
            If CDec(txtIdade.Text) < 19 Then
                txtRepouso.Text = CInt(7.4 * CDec(txtPeso.Text) + (482 * CDec(txtAltura.Text)) + 217)
            ElseIf CDec(txtIdade.Text) < 31 Then
                txtRepouso.Text = CInt(13.3 * CDec(txtPeso.Text) + (334 * CDec(txtAltura.Text)) + 35)
            ElseIf CDec(txtIdade.Text) < 60 Then
                txtRepouso.Text = CInt(8.7 * CDec(txtPeso.Text) - (25 * CDec(txtAltura.Text)) + 862)
            Else
                txtRepouso.Text = CInt(9.2 * CDec(txtPeso.Text) + (637 * CDec(txtAltura.Text)) + 302)
            End If
        End If
        '========== Calculando a Diário/Ativ =============
        If cbxSexo.Text = "Masculino" Then
            '=SE(F10="leve";B46*1,55;SE(F10="moderada";B46*1,78;B46*2,1))
            If cbxAtivFisDia.Text = "Leve" Then
                'A75 * 1,55
                txtDiario.Text = CInt((txtRepouso.Text) * 1.55)
            ElseIf cbxAtivFisDia.Text = "Moderada" Then
                'A75 * 1,78
                txtDiario.Text = CInt((txtRepouso.Text) * 1.78)
            Else
                'A75 * 2,1
                txtDiario.Text = CInt((txtRepouso.Text) * 2.1)
            End If
        Else
            '=SE(F10="LEVE";B46*1,56;SE(F10="MODERADA";B46*1,64;B46*1,82))
            If cbxAtivFisDia.Text = "Leve" Then
                'A76 * 1,56
                txtDiario.Text = CInt((txtRepouso.Text) * 1.56)
            ElseIf cbxAtivFisDia.Text = "Moderada" Then
                'A76 * 1,64
                txtDiario.Text = CInt((txtRepouso.Text) * 1.64)
            Else
                'A76 * 1,82
                txtDiario.Text = CInt((txtRepouso.Text) * 1.82)
            End If
        End If
        '===========Calculando o Total ===================
        txtTotal.Text = CInt(txtAtividade.Text) + CInt(txtDiario.Text)

    End Sub


    ''' <summary>
    ''' Calcula a Redução de Peso e a Qtde de Peso Perdido
    ''' </summary>
    ''' <remarks>Redução de Peso/Peso Perdido</remarks>
    Private Sub ReducaoPeso_PesoPerdido()
        Dim vConversao As Double

        vConversao = CDec(txtPeso.Text) - CDec(txtPesoNormal.Text)
        txtPerdaDesGord.Text = String.Format("{0:f3}", vConversao)

        vConversao = (CDec(txtAtividade.Text) * 0.453) / 3500
        txtPesoDia.Text = String.Format("{0:f3}", vConversao)

        vConversao = CInt(CDec(txtPerdaDesGord.Text) / CDec(txtPesoDia.Text))
        txtDiasNeces.Text = vConversao

        vConversao = CDec(txtDiasNeces.Text) / 30
        txtMesesNecess.Text = String.Format("{0:f1}", vConversao)

        vConversao = CDec(txtPesoDia.Text) * CDec(txtQuantSeman.Text)
        txtPesoSemana.Text = String.Format("{0:f3}", vConversao)

        vConversao = CDec(txtPesoSemana.Text) * 4
        txtPesoMes.Text = String.Format("{0:f3}", vConversao)
    End Sub
    Private Sub carrega_campos()
        txtProf.Text = "Antenor José Augusto"
        txtNome.Text = "Marcos José Antônio da Silva"
        txtLocal.Text = "Academia Bastos de Musculação Ltda 123456"
        dtpDataNasc.Text = "26/03/1967"
        cbxSexo.Text = "Masculino"
        txtPeso.Text = 84
        txtAltura.Text = 1.77
        txtFCrepouso.Text = 88
        txtFCmaxima.Text = 172
        txtBraDireito.Text = 13
        txtBraEsquerdo.Text = 15
        txtAnteDireito.Text = 20
        txtAnteEsquerdo.Text = 30
        txtTorax.Text = 30
        txtAbdomen.Text = 20
        txtCoxaDireita.Text = 10
        txtCoxaEsquerda.Text = 50
        txtPernaDireita.Text = 40
        txtPernaEsquerda.Text = 10
        txtTornozelo.Text = 20
        txtJoelho.Text = 35
        txtPunho.Text = 18
        txtCintura.Text = 120
        txtQuadril.Text = 85
        txtFemur.Text = 42
        txtRadio.Text = 35
        txtSaltoVert.Text = 45
        txtAbdominais.Text = 48
        txtFlexaoBraco.Text = 33
        txtCooper.Text = 2000
        txtVo2Trab.Text = 70
        txtTempoAtiv.Text = 90
        txtQuantSeman.Text = 5
    End Sub
    Private Sub btnProcessar_Click(sender As System.Object, e As System.EventArgs) Handles btnProcessar.Click
        Try
            If cbxSexo.Text <> "Masculino" And cbxSexo.Text <> "Feminino" Then
                MessageBox.Show("Escolha entre o sexo Masculino ou o sexo Feminino!", "Sexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cbxSexo.Focus()
                Exit Sub
            End If

            If txtPunho.Text = "" Then
                MessageBox.Show("Campo Punho Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                txtPunho.Focus()
            ElseIf txtAltura.Text = "" Then
                MessageBox.Show("Campo Altura Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                txtAltura.Focus()
            ElseIf txtIdade.Text = "" Then
                MessageBox.Show("Campo Idade Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                txtIdade.Focus()
            ElseIf txtSaltoVert.Text = "" Then
                MessageBox.Show("Campo Salto Vertical Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                txtSaltoVert.Focus()
            ElseIf cbxSexo.Text = "" Then
                MessageBox.Show("Campo Sexo Vazio", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vProcessar = False
                cbxSexo.Focus()
            End If

            If vProcessar = True Then
                Flexao_Braco()
                Salto_Vertical()
                Abdominais()
                Flexao_Braco()
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
    Private Sub frm_Principal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        carrega_campos()
    End Sub
    Private Sub SerialToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SerialToolStripMenuItem.Click

        Try
            Dim contador As Integer = 0
            Dim leitor As StreamReader

            leitor = My.Computer.FileSystem.OpenTextFileReader("C:\AvaliacaoFisica\avalfisica.txt")
            Do Until leitor.EndOfStream  'lê linha do arquivo
                If contador = 8 Then
                    MessageBox.Show("SUA " & leitor.ReadLine, "SENHA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                Else
                    leitor.ReadLine()
                End If
                contador += 1
            Loop
            leitor.Close()
        Catch ex As Exception
            MessageBox.Show("Erro de leitura !! " & ex.Message, "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

    End Sub


    ''' <summary>
    ''' Usando Keypress de todos os campos
    ''' </summary>
    ''' <remarks>Nessa parte é pressionado um botão e sempre aceitará número ou vírgula</remarks>
    Private Sub txtPeso_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPeso.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAltura_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAltura.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtFCrepouso_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFCrepouso.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtFCmaxima_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFCmaxima.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtBraDireito_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtBraDireito.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtBraEsquerdo_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtBraEsquerdo.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAnteDireito_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAnteDireito.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAnteEsquerdo_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAnteEsquerdo.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtTorax_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtTorax.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAbdomen_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAbdomen.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtCoxaDireita_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCoxaDireita.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtCoxaEsquerda_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCoxaEsquerda.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtPernaDireita_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPernaDireita.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtPernaEsquerda_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPernaEsquerda.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtTornozelo_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtTornozelo.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtJoelho_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtJoelho.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtPunho_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPunho.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtFemur_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFemur.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtRadio_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtRadio.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtCintura_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCintura.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtQuadril_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtQuadril.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtSaltoVert_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSaltoVert.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAbdominais_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAbdominais.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtFlexaoBraco_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFlexaoBraco.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtCooper_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCooper.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtVo2Trab_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtVo2Trab.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtTempoAtiv_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtTempoAtiv.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtQuantSeman_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtQuantSeman.KeyPress
        If e.KeyChar = "," Then
            e.Handled = False
            Exit Sub
        End If

        If Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtIdade_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtIdade.KeyPress
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
            sLeitura = sLeitura & "Professor: " & (txtProf.Text.Trim & Space(40)).Substring(0, 40) & Space(10) & (txtLocal.Text & Space(40)).Substring(0, 40) & vbCrLf
            sLeitura = sLeitura & "======================================================================================================" & vbCrLf
            sLeitura = sLeitura & " NOME:" & (txtNome.Text.Trim & Space(30)).Substring(0, 30) & Space(2) & "IDADE:" & txtIdade.Text & Space(2) & "SEXO:" & cbxSexo.Text & Space(2) & "PESO:" & txtPeso.Text & Space(2) & "ALTURA:" & txtAltura.Text & Space(2) & "DN:" & dtpDataNasc.Value & vbCrLf
            sLeitura = sLeitura & "------------------------------------------------------------------------------------------------------" & vbCrLf
            sLeitura = sLeitura & "CIRCUNFERÊNCIAS CORPORAIS                  DIÂMETROS ÓSSEOS(mm)" & vbCrLf
            sLeitura = sLeitura & "Tornozelo: " & (txtTornozelo.Text.Trim & Space(6)).Substring(0, 6) & Space(26) & "Fêmur: " & txtFemur.Text & Space(10) & vbCrLf
            sLeitura = sLeitura & "Joelho:    " & (txtJoelho.Text.Trim & Space(6)).Substring(0, 6) & Space(26) & "Rádio: " & txtRadio.Text & Space(10) & vbCrLf
            sLeitura = sLeitura & "Punho:     " & (txtPunho.Text.Trim & Space(6)).Substring(0, 6) & Space(26) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "RISCOS DE DOENÇAS CRÔNICAS (CINTURA X QUADRIL)" & vbCrLf
            sLeitura = sLeitura & "Cintura: " & (txtCintura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Relação cintura ao quadril: " & txtRelCinQua.Text & vbCrLf
            sLeitura = sLeitura & "Quadril: " & (txtQuadril.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Risco: " & lblRelCinQua.Text & vbCrLf & vbCrLf

            sLeitura = sLeitura & "COMPOSIÇÃO CORPORAL" & vbCrLf
            sLeitura = sLeitura & "% de gordura:  " & (txtPorcGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "MASSA PROVÁVEL DAS VÁRIAS PARTES DO CORPO (Kg)" & vbCrLf
            sLeitura = sLeitura & "Peso Gordura:  " & (txtPesoGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Cabeça:    " & (txtCabeca.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Mão:   " & txtMao.Text & vbCrLf
            sLeitura = sLeitura & "Peso Muscular: " & (txtPesoMusc.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Tronco:    " & (txtTronco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Coxa:  " & txtCoxa.Text & vbCrLf
            sLeitura = sLeitura & "Peso Residual: " & (txtPesoResid.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Braço:     " & (txtBraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Perna: " & txtPerna.Text & vbCrLf
            sLeitura = sLeitura & "Peso Ósseo:    " & (txtPesoOsseo.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Antebraço: " & (txtAntebraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Pé:    " & txtPe.Text & vbCrLf
            sLeitura = sLeitura & "M. C. M:       " & (txtMcm.Text.Trim & Space(6)).Substring(0, 6) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "PESO IDEAL" & vbCrLf
            sLeitura = sLeitura & "Normal: " & (txtPesoNormal.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Imc:           " & (txtIMC.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResIMC.Text & vbCrLf
            sLeitura = sLeitura & "Mínimo: " & (txtPesoMinimo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Ossatura:      " & (txtOssatura.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResOssatura.Text & vbCrLf
            sLeitura = sLeitura & "Máximo: " & (txtPesoMaximo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Compl. Físico: " & (txtCFisico.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResCFisico.Text & vbCrLf & vbCrLf & vbCrLf

            sLeitura = sLeitura & "                                     ____________________________________________________________" & vbCrLf
            sLeitura = sLeitura & "                                         " & "Ass: " & (txtProf.Text.Trim & Space(40)).Substring(0, 40)
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
            sLeitura = sLeitura & "Professor: " & (txtProf.Text.Trim & Space(40)).Substring(0, 40) & Space(10) & (txtLocal.Text & Space(40)).Substring(0, 40) & vbCrLf
            sLeitura = sLeitura & "======================================================================================================" & vbCrLf
            sLeitura = sLeitura & " NOME:" & (txtNome.Text.Trim & Space(30)).Substring(0, 30) & Space(2) & "IDADE:" & txtIdade.Text & Space(2) & "SEXO:" & cbxSexo.Text & Space(2) & "PESO:" & txtPeso.Text & Space(2) & "ALTURA:" & txtAltura.Text & Space(2) & "DN:" & dtpDataNasc.Value & vbCrLf
            sLeitura = sLeitura & "------------------------------------------------------------------------------------------------------" & vbCrLf
            sLeitura = sLeitura & "CIRCUNFERÊNCIAS CORPORAIS            DIÂMETROS ÓSSEOS(mm)        Dobras" & vbCrLf
            sLeitura = sLeitura & "Tornozelo: " & (txtTornozelo.Text.Trim & Space(6)).Substring(0, 6) & Space(20) & "Fêmur: " & (txtFemur.Text.Trim & Space(21)).Substring(0, 21) & "Suprailíaca:  18   Abdomen:    19" & vbCrLf
            sLeitura = sLeitura & "Joelho:    " & (txtJoelho.Text.Trim & Space(6)).Substring(0, 6) & Space(20) & "Rádio: " & (txtRadio.Text.Trim & Space(21)).Substring(0, 21) & "Subescapular: 13   Tricipital: 20" & vbCrLf
            sLeitura = sLeitura & "Punho:     " & (txtPunho.Text.Trim & Space(6)).Substring(0, 6) & Space(2) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "RISCOS DE DOENÇAS CRÔNICAS (CINTURA X QUADRIL)" & vbCrLf
            sLeitura = sLeitura & "Cintura: " & (txtCintura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Relação cintura ao quadril: " & txtRelCinQua.Text & vbCrLf
            sLeitura = sLeitura & "Quadril: " & (txtQuadril.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Risco: " & lblRelCinQua.Text & vbCrLf & vbCrLf

            sLeitura = sLeitura & "COMPOSIÇÃO CORPORAL" & vbCrLf
            sLeitura = sLeitura & "% de gordura:  " & (txtPorcGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "MASSA PROVÁVEL DAS VÁRIAS PARTES DO CORPO (Kg)" & vbCrLf
            sLeitura = sLeitura & "Peso Gordura:  " & (txtPesoGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Cabeça:    " & (txtCabeca.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Mão:   " & txtMao.Text & vbCrLf
            sLeitura = sLeitura & "Peso Muscular: " & (txtPesoMusc.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Tronco:    " & (txtTronco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Coxa:  " & txtCoxa.Text & vbCrLf
            sLeitura = sLeitura & "Peso Residual: " & (txtPesoResid.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Braço:     " & (txtBraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Perna: " & txtPerna.Text & vbCrLf
            sLeitura = sLeitura & "Peso Ósseo:    " & (txtPesoOsseo.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Antebraço: " & (txtAntebraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Pé:    " & txtPe.Text & vbCrLf
            sLeitura = sLeitura & "M. C. M:       " & (txtMcm.Text.Trim & Space(6)).Substring(0, 6) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "PESO IDEAL" & vbCrLf
            sLeitura = sLeitura & "Normal: " & (txtPesoNormal.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Imc:           " & (txtIMC.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResIMC.Text & vbCrLf
            sLeitura = sLeitura & "Mínimo: " & (txtPesoMinimo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Ossatura:      " & (txtOssatura.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResOssatura.Text & vbCrLf
            sLeitura = sLeitura & "Máximo: " & (txtPesoMaximo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Compl. Físico: " & (txtCFisico.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResCFisico.Text & vbCrLf & vbCrLf & vbCrLf

            sLeitura = sLeitura & "                                     ____________________________________________________________" & vbCrLf
            sLeitura = sLeitura & "                                         " & "Ass: " & (txtProf.Text.Trim & Space(40)).Substring(0, 40)
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
            sLeitura = sLeitura & "Professor: " & (txtProf.Text.Trim & Space(40)).Substring(0, 40) & Space(10) & (txtLocal.Text & Space(40)).Substring(0, 40) & vbCrLf
            sLeitura = sLeitura & "======================================================================================================" & vbCrLf
            sLeitura = sLeitura & " NOME:" & (txtNome.Text.Trim & Space(30)).Substring(0, 30) & Space(2) & "IDADE:" & txtIdade.Text & Space(2) & "SEXO:" & cbxSexo.Text & Space(2) & "PESO:" & txtPeso.Text & Space(2) & "ALTURA:" & txtAltura.Text & Space(2) & "DN:" & dtpDataNasc.Value & vbCrLf
            sLeitura = sLeitura & "------------------------------------------------------------------------------------------------------" & vbCrLf
            sLeitura = sLeitura & "IMPEDÂNCIA BIOELÈTRICA                  DIÂMETROS ÓSSEOS(mm)" & vbCrLf
            sLeitura = sLeitura & "% Gordura: 55" & Space(27) & "Fêmur: " & (txtFemur.Text.Trim & Space(21)).Substring(0, 21) & vbCrLf
            sLeitura = sLeitura & Space(40) & "Rádio: " & (txtRadio.Text.Trim & Space(21)).Substring(0, 21) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "RISCOS DE DOENÇAS CRÔNICAS (CINTURA X QUADRIL)" & vbCrLf
            sLeitura = sLeitura & "Cintura: " & (txtCintura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Relação cintura ao quadril: " & txtRelCinQua.Text & vbCrLf
            sLeitura = sLeitura & "Quadril: " & (txtQuadril.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Risco: " & lblRelCinQua.Text & vbCrLf & vbCrLf

            sLeitura = sLeitura & "COMPOSIÇÃO CORPORAL" & vbCrLf
            sLeitura = sLeitura & "% de gordura:  " & (txtPorcGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "MASSA PROVÁVEL DAS VÁRIAS PARTES DO CORPO (Kg)" & vbCrLf
            sLeitura = sLeitura & "Peso Gordura:  " & (txtPesoGordura.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Cabeça:    " & (txtCabeca.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Mão:   " & txtMao.Text & vbCrLf
            sLeitura = sLeitura & "Peso Muscular: " & (txtPesoMusc.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Tronco:    " & (txtTronco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Coxa:  " & txtCoxa.Text & vbCrLf
            sLeitura = sLeitura & "Peso Residual: " & (txtPesoResid.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Braço:     " & (txtBraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Perna: " & txtPerna.Text & vbCrLf
            sLeitura = sLeitura & "Peso Ósseo:    " & (txtPesoOsseo.Text.Trim & Space(6)).Substring(0, 6) & Space(10) & "Antebraço: " & (txtAntebraco.Text.Trim & Space(6)).Substring(0, 6) & Space(3) & "Pé:    " & txtPe.Text & vbCrLf
            sLeitura = sLeitura & "M. C. M:       " & (txtMcm.Text.Trim & Space(6)).Substring(0, 6) & vbCrLf & vbCrLf

            sLeitura = sLeitura & "PESO IDEAL" & vbCrLf
            sLeitura = sLeitura & "Normal: " & (txtPesoNormal.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Imc:           " & (txtIMC.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResIMC.Text & vbCrLf
            sLeitura = sLeitura & "Mínimo: " & (txtPesoMinimo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Ossatura:      " & (txtOssatura.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResOssatura.Text & vbCrLf
            sLeitura = sLeitura & "Máximo: " & (txtPesoMaximo.Text.Trim & Space(6)).Substring(0, 6) & " Kg" & Space(4) & "Compl. Físico: " & (txtCFisico.Text.Trim & Space(6)).Substring(0, 6) & Space(1) & txtResCFisico.Text & vbCrLf & vbCrLf & vbCrLf

            sLeitura = sLeitura & "                                     ____________________________________________________________" & vbCrLf
            sLeitura = sLeitura & "                                         " & "Ass: " & (txtProf.Text.Trim & Space(40)).Substring(0, 40)
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