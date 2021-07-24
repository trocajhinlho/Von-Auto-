Public Class vonauto

    Private Sub btn_fipe_Click(sender As Object, e As EventArgs) Handles btn_fipe.Click

        btn_deslogar.Visible = False
        btn_tabfipe()

    End Sub

    Private Sub btn_veiculos_Click(sender As Object, e As EventArgs) Handles btn_veiculos.Click

        btn_deslogar.Visible = False
        btn_demanda()

    End Sub

    Private Sub btn_cadastro_veiculos_Click(sender As Object, e As EventArgs) Handles btn_cadastro_veiculos.Click

        btn_deslogar.Visible = False
        cad_veiculo()

    End Sub

    
    Private Sub btn_cadastro_clientes_Click(sender As Object, e As EventArgs) Handles btn_cadastro_clientes.Click

        btn_deslogar.Visible = False
        cad_clientes()

    End Sub

    Private Sub btn_cadastro_compras_Click(sender As Object, e As EventArgs) Handles btn_cadastro_compras.Click

        btn_deslogar.Visible = False
        cad_compras()

    End Sub

    Private Sub btn_cadastro_func_Click(sender As Object, e As EventArgs) Handles btn_cadastro_func.Click

        
        If lbl_vaga_func.Text = "VENDEDOR" Then
            MsgBox("VENDEDOR NÃO PODE CADASTRAR FUNCIONARIOS! FAÇA LOGIN COM UMA CONTA DE ADM OU GERENTE", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
        Else
            btn_deslogar.Visible = False
            cad_func()

        End If
        
    End Sub

    Private Sub btn_sair_MouseHover(sender As Object, e As EventArgs) Handles btn_sair.MouseHover

        btn_sair.Font = New Font(btn_sair.Font, FontStyle.Bold)

    End Sub

    Private Sub btn_sair_MouseLeave(sender As Object, e As EventArgs) Handles btn_sair.MouseLeave

        btn_sair.Font = New Font(btn_sair.Font, FontStyle.Bold = False)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btn_float.Click
        If btn_deslogar.Visible = False Then
            btn_deslogar.Visible = True
        Else
            btn_deslogar.Visible = False
        End If


    End Sub

    Private Sub vonauto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.FormBorderStyle = FormBorderStyle.None

    End Sub


    Private Sub img_carro_Click(sender As Object, e As EventArgs) Handles img_carro.Click
        Try
            With OpenFileDialog1
                .Title = "SELECIONE UMA IMAGEM DO VEICULO"
                .InitialDirectory = Application.StartupPath & "\FOTOS_CARROS"
                .ShowDialog()
                diretorio = .FileName
                img_carro.Load(diretorio)

            End With
        Catch ex As Exception
            Exit Sub 'parar a ação, evitar exibir mensagem de erro (para user final)! 
        End Try

    End Sub

    Private Sub btn_cadastro_veic_Click(sender As Object, e As EventArgs) Handles btn_cadastro_veic.Click
        Try
            sql = "select * from tb_veiculos where placa_veiculo='" & txt_placa_veiculo.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then

                sql = "update tb_veiculos set valor_veiculo='" & txt_valor_carro.Text & "',ano_fabric_veiculo='" & txt_ano_veiculo.Text & "',marca_veiculo='" & cb_marca.Text & "',modelo_veiculo='" & txt_modelo_carro.Text & "',renavam_veiculo='" & txt_renavam_veiculo.Text & "',cor_veiculo='" & txt_cor_veiculo.Text & "',foto_veiculo='" & diretorio & "' where placa_veiculo='" & txt_placa_veiculo.Text & "'"

                rs = db.Execute(UCase(sql))

                MsgBox("Dados do Veiculo atualizados com sucesso!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
                carregar_dados()

            Else
                sql = "insert into tb_veiculos values ('" & txt_placa_veiculo.Text & "','" & txt_valor_carro.Text & "','" & txt_ano_veiculo.Text & "','" & cb_marca.Text & "','" & txt_modelo_carro.Text & "','" & txt_renavam_veiculo.Text & "','" & txt_cor_veiculo.Text & "','" & diretorio & "')"
                ' ***UTILIZAR PARA REFERENCIAR NO CONSULTA AVANÇADA DE VEICULOS***   sql = "select * from tb_veiculos set placa_veiculo='" & txt_placa_veiculo.Text & "',valor_veiculo='" & txt_valor_carro.Text & "',ano_fabric_veiculo='" & txt_ano_veiculo.Text & "',marca_veiculo='" & cb_marca.Text & "',modelo_veiculo='" & txt_modelo_carro.Text & "',renavam='" & txt_renavam_veiculo.Text & "'cor_veiculo='" & txt_cor_veiculo.Text & "',foto_veiculo=" & diretorio & "'"
                rs = db.Execute(UCase(sql))
                MsgBox("Veiculo Cadastrado com Sucesso!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
                carregar_dados()

            End If

            limpar_cadastro_carros()

        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try
    End Sub

    Private Sub txt_placa_veiculo_DoubleClick(sender As Object, e As EventArgs) Handles txt_placa_veiculo.DoubleClick
        limpar_cadastro_carros()

    End Sub

    Private Sub txt_placa_veiculo_LostFocus(sender As Object, e As EventArgs) Handles txt_placa_veiculo.LostFocus
        Try
            sql = "select * from tb_veiculos where placa_veiculo='" & txt_placa_veiculo.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then
                MsgBox("Veiculo já cadastrado!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")

                txt_valor_carro.Text = rs.Fields(1).Value
                txt_ano_veiculo.Text = rs.Fields(2).Value
                cb_marca.Text = rs.Fields(3).Value
                txt_modelo_carro.Text = rs.Fields(4).Value
                txt_renavam_veiculo.Text = rs.Fields(5).Value
                txt_cor_veiculo.Text = rs.Fields(6).Value
                img_carro.Load(rs.Fields(7).Value)
                diretorio = (rs.Fields(7).Value)

            Else
                txt_valor_carro.Focus()

            End If
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try
    End Sub

    Private Sub btn_gravar_clientes_Click(sender As Object, e As EventArgs) Handles btn_gravar_clientes.Click

        Try
            sql = "select * from tb_clientes where cpf_cliente='" & txt_cpf_cliente.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then

                sql = "update tb_clientes set nome='" & txt_nome_cliente.Text & "',data_nasc_cliente='" & dtp_cliente.Text & "',email_cliente='" & txt_email_cliente.Text & "',cep_cliente='" & txt_cep_cliente.Text & "',endereco_cliente='" & txt_rua_cliente.Text & "',cidade_cliente='" & txt_cidade_cliente.Text & "',bairro_cliente='" & txt_bairro_cliente.Text & "',uf_cliente='" & txt_uf_cliente.Text & "',num_func='" & txt_numero_cliente.Text & "',telefone='" & txt_telefone_cliente.Text & "' where cpf_cliente='" & txt_cpf_cliente.Text & "'"

                rs = db.Execute(UCase(sql))

                MsgBox(diretorio + " Dados do Cliente atualizados com sucesso!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
            Else
                sql = "insert into tb_clientes values ('" & txt_cpf_cliente.Text & "','" & txt_nome_cliente.Text & "','" & dtp_cliente.Text & "','" & txt_email_cliente.Text & "','" & txt_cep_cliente.Text & "','" & txt_rua_cliente.Text & "','" & txt_cidade_cliente.Text & "','" & txt_bairro_cliente.Text & "','" & txt_uf_cliente.Text & "','" & txt_numero_cliente.Text & "','" & txt_telefone_cliente.Text & "')"
                ' ***UTILIZAR PARA REFERENCIAR NO CONSULTA AVANÇADA***   sql = "select * from tb_veiculos set placa_veiculo='" & txt_placa_veiculo.Text & "',valor_veiculo='" & txt_valor_carro.Text & "',ano_fabric_veiculo='" & txt_ano_veiculo.Text & "',marca_veiculo='" & cb_marca.Text & "',modelo_veiculo='" & txt_modelo_carro.Text & "',renavam='" & txt_renavam_veiculo.Text & "'cor_veiculo='" & txt_cor_veiculo.Text & "',foto_veiculo=" & diretorio & "'"
                rs = db.Execute(UCase(sql))
                MsgBox("Cliente Cadastrado com Sucesso!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")

            End If

            limpar_cadastro_clientes()

        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try


    End Sub

    Private Sub txt_cpf_cliente_DoubleClick(sender As Object, e As EventArgs) Handles txt_cpf_cliente.DoubleClick
        limpar_cadastro_clientes()

    End Sub

    Private Sub txt_cpf_cliente_LostFocus(sender As Object, e As EventArgs) Handles txt_cpf_cliente.LostFocus

        Try
            sql = "select * from tb_clientes where cpf_cliente='" & txt_cpf_cliente.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then
                MsgBox("Cliente já cadastrado!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")

                txt_nome_cliente.Text = rs.Fields(1).Value
                dtp_cliente.Text = rs.Fields(2).Value
                txt_email_cliente.Text = rs.Fields(3).Value
                txt_cep_cliente.Text = rs.Fields(4).Value
                txt_rua_cliente.Text = rs.Fields(5).Value
                txt_cidade_cliente.Text = rs.Fields(6).Value
                txt_bairro_cliente.Text = rs.Fields(7).Value
                txt_uf_cliente.Text = rs.Fields(8).Value
                txt_numero_cliente.Text = rs.Fields(9).Value
                txt_telefone_cliente.Text = rs.Fields(10).Value
               

            Else

                txt_nome_cliente.Focus()

            End If
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try

    End Sub

    
    Private Sub txt_cep_cliente_LostFocus(sender As Object, e As EventArgs) Handles txt_cep_cliente.LostFocus
        Try
            sql = "select * from tb_cep where cep='" & txt_cep_cliente.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then

                txt_rua_cliente.Text = rs.Fields(1).Value
                txt_cidade_cliente.Text = rs.Fields(2).Value
                txt_bairro_cliente.Text = rs.Fields(3).Value
                txt_uf_cliente.Text = rs.Fields(4).Value
                txt_numero_cliente.Focus()

            Else

                txt_rua_cliente.Focus()

            End If



        Catch ex As Exception
            Exit Sub
            'MsgBox("CEP NÃO ENCONTRADO NO BANCO DE DADOS!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        End Try

    End Sub

   
    Private Sub btn_cadastro_funcionarios_Click(sender As Object, e As EventArgs) Handles btn_cadastro_funcionarios.Click
        Try
            sql = "select * from tb_funcionarios where cpf_funcionario='" & txt_cpf_func.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then

                sql = "update tb_funcionarios set rg_func='" & txt_rg_func.Text & "',nome_func='" & txt_nome_func.Text & "',data_nasc_func='" & dtp_data_func.Text & "',email_func='" & txt_email_func.Text & "',vaga_func='" & cbl_vaga_func.Text & "',senha_func='" & txt_senha_func.Text & "',cep_func='" & txt_cep_func.Text & "',endereco_func='" & txt_endereco_func.Text & "',cidade_func='" & txt_cidade_func.Text & "',bairro_func='" & txt_bairro_func.Text & "',uf_func='" & txt_uf_func.Text & "',num_func='" & txt_numero_func.Text & "',foto_func='" & diretorio_func & "' where cpf_funcionario='" & txt_cpf_func.Text & "'"

                rs = db.Execute(UCase(sql))

                MsgBox(" Dados do Funcionario " + txt_nome_func.Text + " atualizados com sucesso!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
            Else
                sql = "insert into tb_funcionarios values ('" & txt_cpf_func.Text & "','" & txt_rg_func.Text & "','" & txt_nome_func.Text & "','" & dtp_data_func.Text & "','" & txt_email_func.Text & "','" & cbl_vaga_func.Text & "','" & txt_senha_func.Text & "','" & txt_cep_func.Text & "','" & txt_endereco_func.Text & "','" & txt_cidade_func.Text & "','" & txt_bairro_func.Text & "','" & txt_uf_func.Text & "','" & txt_numero_func.Text & "','" & diretorio_func & "')"
                ' ***UTILIZAR PARA REFERENCIAR NO CONSULTA AVANÇADA***   sql = "select * from tb_veiculos set placa_veiculo='" & txt_placa_veiculo.Text & "',valor_veiculo='" & txt_valor_carro.Text & "',ano_fabric_veiculo='" & txt_ano_veiculo.Text & "',marca_veiculo='" & cb_marca.Text & "',modelo_veiculo='" & txt_modelo_carro.Text & "',renavam='" & txt_renavam_veiculo.Text & "'cor_veiculo='" & txt_cor_veiculo.Text & "',foto_veiculo=" & diretorio & "'"
                rs = db.Execute(UCase(sql))
                MsgBox("Funcionário " + txt_nome_func.Text + " Cadastrado com Sucesso!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")

            End If

            limpar_cadastro_funcionario()

        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try
    End Sub

    Private Sub img_func_Click(sender As Object, e As EventArgs) Handles img_func.Click
        Try
            With OpenFileDialog2
                .Title = "SELECIONE UMA IMAGEM DO FUNCIONÁRIO"
                .InitialDirectory = Application.StartupPath & "\FOTOS_FUNCIONARIOS"
                .ShowDialog()
                diretorio_func = .FileName
                img_func.Load(diretorio_func)

            End With
        Catch ex As Exception
            Exit Sub 'parar a ação, evitar exibir mensagem de erro (para user final)! 
        End Try
    End Sub

    Private Sub txt_cpf_func_LostFocus(sender As Object, e As EventArgs) Handles txt_cpf_func.LostFocus
        Try
            sql = "select * from tb_funcionarios where cpf_funcionario='" & txt_cpf_func.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then
                MsgBox("Funcionário já cadastrado!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
                txt_rg_func.Text = rs.Fields(1).Value
                txt_nome_func.Text = rs.Fields(2).Value
                dtp_data_func.Text = rs.Fields(3).Value
                txt_email_func.Text = rs.Fields(4).Value
                cbl_vaga_func.Text = rs.Fields(5).Value
                txt_senha_func.Text = rs.Fields(6).Value
                txt_cep_func.Text = rs.Fields(7).Value
                txt_endereco_func.Text = rs.Fields(8).Value
                txt_cidade_func.Text = rs.Fields(9).Value
                txt_bairro_func.Text = rs.Fields(10).Value
                txt_uf_func.Text = rs.Fields(11).Value
                txt_numero_func.Text = rs.Fields(12).Value
                img_func.Load(rs.Fields(13).Value)
                diretorio_func = rs.Fields(13).Value

            Else

                txt_nome_cliente.Focus()

            End If
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try
    End Sub

    Private Sub txt_cep_func_LostFocus(sender As Object, e As EventArgs) Handles txt_cep_func.LostFocus
        Try
            sql = "select * from tb_cep where cep='" & txt_cep_func.Text & "'"
            rs = db.Execute(sql)
            If rs.EOF = False Then

                txt_endereco_func.Text = rs.Fields(1).Value
                txt_cidade_func.Text = rs.Fields(2).Value
                txt_bairro_func.Text = rs.Fields(3).Value
                txt_uf_func.Text = rs.Fields(4).Value
                txt_numero_func.Focus()

            Else

                txt_endereco_func.Focus()

            End If



        Catch ex As Exception
            Exit Sub
            'MsgBox("CEP NÃO ENCONTRADO NO BANCO DE DADOS!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub

    Private Sub btn_sair_Click(sender As Object, e As EventArgs) Handles btn_sair.Click
        resp = MsgBox("Deseja Sair do Sistema?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATENÇÃO")

        If resp = MsgBoxResult.Yes Then
            MsgBox("Até uma outra hora " & lbl_nome_func.Text & "!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
            Me.Close()
            Login.Close()
        Else
            MsgBox("Obrigado por continuar conosco " & lbl_nome_func.Text & "!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")

        End If


    End Sub

    Private Sub btn_deslogar_Click(sender As Object, e As EventArgs) Handles btn_deslogar.Click
        resp = MsgBox("Deseja Deslogar-se do Sistema?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATENÇÃO")

        If resp = MsgBoxResult.Yes Then
            MsgBox("Até uma outra hora " & lbl_nome_func.Text & "!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
            Me.Close()
            Login.Show()

        Else
            MsgBox("Obrigado por continuar conosco " & lbl_nome_func.Text & "!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")

        End If
    End Sub

   
    Private Sub btn_compras_Click(sender As Object, e As EventArgs) Handles btn_compras.Click


        Try
            sql = "select * from tb_compras where placa_veiculo='" & txt_placa_compra.Text & "'"
            rs = db.Execute(UCase(sql))
            If rs.EOF = False Then
                
                MsgBox("Veiculo já vendido ou não cadastrado!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
            Else
                sql = "insert into tb_compras values ('" & txt_cpf_cliente_compra.Text & "','" & lbl_rg_func.Text & "','" & txt_placa_compra.Text & "','" & txt_ano_compra.Text & "','" & lbl_valor_veic.Text & "','" & cbl_forma_pag.Text & "','" & lbl_valor_total.Text & "','" & diretorio_compra & "')"
                rs = db.Execute(UCase(sql))
                sql = "Delete from tb_veiculos where placa_veiculo='" & txt_placa_compra.Text & "'"
                ' ***UTILIZAR PARA REFERENCIAR NO CONSULTA AVANÇADA***   sql = "select * from tb_veiculos set placa_veiculo='" & txt_placa_veiculo.Text & "',valor_veiculo='" & txt_valor_carro.Text & "',ano_fabric_veiculo='" & txt_ano_veiculo.Text & "',marca_veiculo='" & cb_marca.Text & "',modelo_veiculo='" & txt_modelo_carro.Text & "',renavam='" & txt_renavam_veiculo.Text & "'cor_veiculo='" & txt_cor_veiculo.Text & "',foto_veiculo=" & diretorio & "'"
                rs = db.Execute(UCase(sql))
                MsgBox("Compra Realizada com Sucesso!!! Parabéns!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")



            End If

            limpar_compras()

        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try









    End Sub

    Private Sub txt_cpf_cliente_compra_DoubleClick(sender As Object, e As EventArgs) Handles txt_cpf_cliente_compra.DoubleClick
        limpar_compras()

    End Sub

    Private Sub txt_cpf_cliente_compra_LostFocus(sender As Object, e As EventArgs) Handles txt_cpf_cliente_compra.LostFocus
        conectar_banco()
        sql = "select * from tb_clientes where cpf_cliente='" & txt_cpf_cliente_compra.Text & "'"
        rs = db.Execute(sql)
        If rs.EOF = False Then
            lbl_Nome_cliente.Text = ("Nome do Cliente: " + rs.Fields(1).Value)
            txt_placa_compra.Focus()


        Else
            MsgBox("Para Cadastrar uma compra o Cliente tem que ser cadastrado!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
            txt_cpf_cliente_compra.Clear()



        End If
    End Sub

    Private Sub txt_placa_compra_LostFocus(sender As Object, e As EventArgs) Handles txt_placa_compra.LostFocus
        conectar_banco()
        sql = "select * from tb_veiculos where placa_veiculo='" & txt_placa_compra.Text & "'"
        rs = db.Execute(sql)
        If rs.EOF = False Then
            lbl_valor_veic.Text = ("R$ " + rs.Fields(1).Value)
            lbl_valor_total.Text = ("Total a Pagar: " + lbl_valor_veic.Text)
            lbl_modelo_veiculo.Text = ("Modelo do Veiculo: " + rs.Fields(4).Value)
            lbl_marca_veiculo.Text = ("Marca do Veiculo: " + rs.Fields(3).Value)

            txt_ano_compra.Text = rs.Fields(2).Value
            img_compra.Load(rs.Fields(7).Value)
            diretorio_compra = (rs.Fields(7).Value)
        Else
            MsgBox("Placa não Localizada! Para cadastrar uma Compra o Carro deve ser Cadastrado!!!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
            txt_placa_compra.Clear()


        End If
    End Sub

    
   
    
    Private Sub txt_nome_func_DoubleClick(sender As Object, e As EventArgs) Handles txt_nome_func.DoubleClick
        limpar_cadastro_funcionario()

    End Sub

    Private Sub btn_consulta_veiculos_Click(sender As Object, e As EventArgs) Handles btn_consulta_veiculos.Click

        btn_consulta_veiculo()
        btn_deslogar.Visible = False
        carregar_dados()
        carregar_tipo_veiculos()

    End Sub


    Private Sub txt_consulta_cliente_TextChanged(sender As Object, e As EventArgs) Handles txt_consulta_cliente.TextChanged
        Try
            With dgv_consulta_cliente
                cont = 1
                sql = "select * from tb_clientes where " & cmb_cunsulta_cliente.Text & " like '" & txt_consulta_cliente.Text & "%'"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False
                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, rs.Fields(7).Value, rs.Fields(8).Value, rs.Fields(9).Value, rs.Fields(10).Value, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            Exit Sub

        End Try
    End Sub








    Private Sub txt_busca_veiculos_TextChanged(sender As Object, e As EventArgs) Handles txt_busca_veiculos.TextChanged
        Try
            With dgv_consulta_veiculos
                cont = 1
                sql = "select * from tb_veiculos where " & cmb_tipo_veiculo.Text & " like '" & txt_busca_veiculos.Text & "%'"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False
                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, Nothing, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            Exit Sub

        End Try

    End Sub

    Private Sub dgv_consulta_veiculos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_consulta_veiculos.CellContentClick
        Try
            With dgv_consulta_veiculos
                If .CurrentRow.Cells(8).Selected = True Then
                    aux_carros = .CurrentRow.Cells(1).Value

                    resp = MsgBox("Deseja Realmente excluir" + vbNewLine & _
                                  "PLACA: " & aux_carros & "?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "ATENÇÃO")
                    If resp = MsgBoxResult.Yes Then
                        sql = "delete from tb_veiculos where placa_veiculo='" & aux_carros & "'"
                        rs = db.Execute(sql)
                        carregar_dados()

                    End If
                Else
                    Exit Sub

                End If
            End With
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try
    End Sub

    Private Sub btn_consulta_cliente_Click(sender As Object, e As EventArgs) Handles btn_consulta_cliente.Click
        btn_consulta_clientes()
        btn_deslogar.Visible = False
        carregar_dados_clientes()
        carregar_tipo_clientes()
    End Sub

    Private Sub dgv_consulta_cliente_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_consulta_cliente.CellContentClick
        Try
            With dgv_consulta_cliente
                If .CurrentRow.Cells(12).Selected = True Then
                    aux_clientes = .CurrentRow.Cells(1).Value

                    resp = MsgBox("Deseja Realmente excluir" + vbNewLine & _
                                  "CPF: " & aux_clientes & "?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "ATENÇÃO")
                    If resp = MsgBoxResult.Yes Then
                        sql = "delete from tb_clientes where cpf_cliente='" & aux_clientes & "'"
                        rs = db.Execute(sql)
                        carregar_dados_clientes()

                    End If
                Else
                    Exit Sub

                End If
            End With
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try
    End Sub

    Private Sub btn_historico_compras_Click(sender As Object, e As EventArgs) Handles btn_historico_compras.Click
        carregar_dados_compras()
        carregar_tipo_compras()
        btn_consulta_compras()
        btn_deslogar.Visible = False


    End Sub

    Private Sub btn_consulta_funcionarios_Click(sender As Object, e As EventArgs) Handles btn_consulta_funcionarios.Click

        If lbl_vaga_func.Text = "VENDEDOR" Then
            MsgBox("VENDEDOR NÃO PODE CONSULTAR FUNCIONARIOS! FAÇA LOGIN COM UMA CONTA DE ADM OU GERENTE", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
        Else
            carregar_dados_funcionario()
            carregar_tipo_funcionarios()
            btn_consulta_funcionario()
            btn_deslogar.Visible = False
        End If


        
    End Sub



    Private Sub txt_busca_compras_TextChanged(sender As Object, e As EventArgs) Handles txt_busca_compras.TextChanged

        Try
            With dgv_consulta_compras
                cont = 1
                sql = "select * from tb_compras where " & cmb_tipo_compras.Text & " like '" & txt_busca_compras.Text & "%'"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False
                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, rs.Fields(7).Value, Nothing, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            Exit Sub

        End Try


    End Sub

    Private Sub txt_busca_funcionarios_TextChanged(sender As Object, e As EventArgs) Handles txt_busca_funcionarios.TextChanged

        Try
            With dgv_consulta_funcionarios
                cont = 1
                sql = "select * from tb_funcionarios where " & cmb_tipo_funcionarios.Text & " like '" & txt_busca_funcionarios.Text & "%'"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False
                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, rs.Fields(7).Value, rs.Fields(8).Value, rs.Fields(9).Value, rs.Fields(10).Value, rs.Fields(11).Value, rs.Fields(12).Value, Nothing, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            Exit Sub

        End Try


    End Sub

    Private Sub dgv_consulta_compras_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_consulta_compras.CellContentClick
        Try
            With dgv_consulta_compras
                If .CurrentRow.Cells(9).Selected = True Then
                    aux_compras = .CurrentRow.Cells(1).Value

                    resp = MsgBox("Deseja Realmente excluir" + vbNewLine & _
                                  "COMPRA: " & aux_compras & "?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "ATENÇÃO")
                    If resp = MsgBoxResult.Yes Then
                        sql = "delete from tb_compras where ID_compra='" & aux_compras & "'"
                        rs = db.Execute(sql)
                        carregar_dados_compras()

                    End If
                Else
                    Exit Sub

                End If
            End With
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try



    End Sub

    Private Sub dgv_consulta_funcionarios_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_consulta_funcionarios.CellContentClick

        Try
            With dgv_consulta_funcionarios
                If .CurrentRow.Cells(14).Selected = True Then
                    aux_funcionario = .CurrentRow.Cells(1).Value

                    resp = MsgBox("Deseja Realmente excluir" + vbNewLine & _
                                  "CPF: " & aux_funcionario & "?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "ATENÇÃO")
                    If resp = MsgBoxResult.Yes Then
                        sql = "delete from tb_funcionarios where cpf_funcionario='" & aux_funcionario & "'"
                        rs = db.Execute(sql)
                        carregar_dados_funcionario()

                    End If
                Else
                    Exit Sub

                End If
            End With
        Catch ex As Exception
            MsgBox("Erro ao processar!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End Try



    End Sub
End Class