Public Class gif

    

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If ProgressBar1.Value < 100 Then

            ProgressBar1.Value = ProgressBar1.Value + 4.6

            If ProgressBar1.Value = 10 Then
                lbl_carregar.Text = "Carregando... 10%"

            End If
            If ProgressBar1.Value = 20 Then
                lbl_carregar.Text = "Carregando... 20%"

            End If
            If ProgressBar1.Value = 30 Then
                lbl_carregar.Text = "Carregando... 30%"

            End If
            If ProgressBar1.Value = 40 Then
                lbl_carregar.Text = "Carregando... 40%"

            End If
            If ProgressBar1.Value = 50 Then
                lbl_carregar.Text = "Carregando... 50%"

            End If
            If ProgressBar1.Value = 60 Then
                lbl_carregar.Text = "Carregando... 60%"

            End If
            If ProgressBar1.Value = 70 Then
                lbl_carregar.Text = "Carregando... 70%"

            End If
            If ProgressBar1.Value = 80 Then
                lbl_carregar.Text = "Carregando... 80%"

            End If
            If ProgressBar1.Value = 90 Then
                lbl_carregar.Text = "Carregando... 90%"

            End If
            If ProgressBar1.Value = 100 Then
                lbl_carregar.Text = "Carregando... 100%"


            End If

        Else
            Timer1.Enabled = False
            Login.Show()
            Me.Hide()

        End If


    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub
End Class