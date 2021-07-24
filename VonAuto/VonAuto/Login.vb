Public Class Login

    Private Sub btn_entrar_Click(sender As Object, e As EventArgs) Handles btn_entrar.Click
        conectar_banco()
        sql = "select * from tb_funcionarios where email_func='" & txt_email_login.Text & "'"
        rs = db.Execute(UCase(sql))
        If rs.EOF = False Then
            sql = "select * from tb_funcionarios where senha_func='" & txt_senha_login.Text & "'"

            rs = db.Execute(UCase(sql))
            If rs.EOF = False Then

                sql = "select * from tb_funcionarios where email_func='" & txt_email_login.Text & "'"
                rs = db.Execute(UCase(sql))


                lbl_nome_func.Text = rs.Fields(2).Value
                vonauto.lbl_nome_func.Text = rs.Fields(2).Value
                vonauto.lbl_vaga_func.Text = rs.Fields(5).Value
                vonauto.img_login.Load(rs.Fields(13).Value)
                vonauto.lbl_rg_func.Text = rs.Fields(1).Value
                MsgBox("BEM VINDO " + lbl_nome_func.Text + "!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")

                ' função fechar form

                Me.Close()

                'Abrir Outro form
                vonauto.Show()


            End If
        ElseIf txt_email_login.Text = "admin" And txt_senha_login.Text = "admin" Then
            txt_email_login.Text = ""
            txt_senha_login.Text = ""
            MsgBox("BEM VINDO! CONTA DE ADMINISTRADOR CONECTADA!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")

            ' função fechar form
            Me.Close()


            conectar_banco()

            'Abrir Outro form
            vonauto.Show()
           

        Else
            MsgBox("EMAIL OU SENHA NÃO CADASTRADOS!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        End If





    End Sub
End Class