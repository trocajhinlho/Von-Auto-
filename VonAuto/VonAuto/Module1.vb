Module Module1

    Public db As New ADODB.Connection 'Variavel conexão com database (nn esquecer de habilitar a referencia)
    Public rs As New ADODB.Recordset 'Variavel da tabela do database
    Public sql As String 'Trabalhar as querys do CRUID de leitura e escrita (utiliza pra usar comandos de sql no codigo do programa)
    Public resp, diretorio, diretorio_func, diretorio_compra, cont, aux_carros, aux_clientes, aux_funcionario, aux_compras As String

    Sub carregar_tipo_veiculos()
        Try
            With vonauto.cmb_tipo_veiculo.Items
                .Add("marca_veiculo")
                .Add("modelo_veiculo")
                .Add("placa_veiculo")

            End With
            vonauto.cmb_tipo_veiculo.SelectedIndex = 0
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub
    Sub carregar_dados()
        Try
            With vonauto.dgv_consulta_veiculos
                cont = 1
                sql = "select * from tb_veiculos order by marca_veiculo asc"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False

                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub


    Sub carregar_tipo_clientes()
        Try
            With vonauto.cmb_cunsulta_cliente.Items
                .Add("nome")
                .Add("cpf_cliente")
                .Add("email_cliente")

            End With
            vonauto.cmb_cunsulta_cliente.SelectedIndex = 0
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub



    Sub carregar_dados_clientes()

        Try
            With vonauto.dgv_consulta_cliente
                cont = 1
                sql = "select * from tb_clientes order by nome asc"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False

                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, rs.Fields(7).Value, rs.Fields(8).Value, rs.Fields(9).Value, rs.Fields(10).Value, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try

    End Sub








    Sub carregar_dados_funcionario()

        Try
            With vonauto.dgv_consulta_funcionarios
                cont = 1
                sql = "select * from tb_funcionarios order by nome_func asc"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False

                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, rs.Fields(7).Value, rs.Fields(8).Value, rs.Fields(9).Value, rs.Fields(10).Value, rs.Fields(11).Value, rs.Fields(12).Value, Nothing, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try


    End Sub


    Sub carregar_tipo_funcionarios()

        Try
            With vonauto.cmb_tipo_funcionarios.Items
                .Add("nome_func")
                .Add("cpf_funcionario")
                .Add("rg_func")

            End With
            vonauto.cmb_tipo_funcionarios.SelectedIndex = 0
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try


    End Sub



    Sub carregar_dados_compras()


        Try
            With vonauto.dgv_consulta_compras
                cont = 1
                sql = "select * from tb_compras order by ID_compra asc"
                rs = db.Execute(sql)
                .Rows.Clear()
                Do While rs.EOF = False

                    .Rows.Add(cont, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value, rs.Fields(6).Value, rs.Fields(7).Value, Nothing, Nothing)
                    cont = cont + 1
                    rs.MoveNext()

                Loop
            End With
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try

    End Sub

    Sub carregar_tipo_compras()

        Try
            With vonauto.cmb_tipo_compras.Items
                .Add("ID_compra")
                .Add("cpf_cliente")
                .Add("rg_funcionario")
                .Add("placa_veiculo")

            End With
            vonauto.cmb_tipo_compras.SelectedIndex = 0
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try


    End Sub







    Sub conectar_banco()
        'tratamento de erro (melhor visibilidade para cliente final)
        Try
            db = CreateObject("ADODB.Connection")
            db.Open("Provider=SQLOLEDB;Data Source=DESKTOP-99DG4UV;Initial Catalog=von_auto;trusted_connection=yes;")
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao conectar com o Banco de dados do Sistema! Contate agora mesmo o Desenvolvedor responsável!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub






































    Sub btn_consulta_compras()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o da tabela fipe
            .painel_consulta_compras.Visible = True
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False
        End With
    End Sub


    Sub btn_consulta_funcionario()

        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o da tabela fipe

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = True
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False
        End With
    End Sub


    Sub btn_consulta_clientes()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o da tabela fipe

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = True
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False
        End With
    End Sub






    Sub btn_consulta_veiculo()

        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o da tabela fipe

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = True
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False
        End With

    End Sub
    Sub btn_tabfipe()
        With vonauto
            .painel_cor_fipe.BackColor = Color.FromArgb(255, 227, 50)
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o da tabela fipe

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = True
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False
        End With

    End Sub
    Sub btn_demanda()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.FromArgb(255, 227, 50)
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o da demanda de veiculos

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = True
            .painel_cor_demanda.BringToFront()
        End With
    End Sub
    Sub cad_veiculo()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.FromArgb(255, 227, 50)
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o do cadastro de veiculos

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = True
            .painel_demanda.Visible = False
        End With
    End Sub
    Sub cad_clientes()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.FromArgb(255, 227, 50)
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o do cadastro de clientes

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = True
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False

        End With
    End Sub
    Sub cad_compras()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.FromArgb(255, 227, 50)
            .painel_cor_funcinarios.BackColor = Color.Transparent

            'deixar oculto todos os paineis exceto o do cadastro de compras

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = True
            .painel_cadastro_funcionarios.Visible = False
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False

        End With
    End Sub
    Sub cad_func()
        With vonauto
            .painel_cor_fipe.BackColor = Color.Transparent
            .painel_cor_demanda.BackColor = Color.Transparent
            .painel_cor_veiculos.BackColor = Color.Transparent
            .painel_cor_clientes.BackColor = Color.Transparent
            .Painel_cor_compras.BackColor = Color.Transparent
            .painel_cor_funcinarios.BackColor = Color.FromArgb(255, 227, 50)

            'deixar oculto todos os paineis exceto o do cadastro de funcionarios

            .painel_consulta_compras.Visible = False
            .painel_consulta_funcionarios.Visible = False
            .painel_consulta_clientes.Visible = False
            .painel_consulta_veiculo.Visible = False
            .web_fipe.Visible = False
            .painel_cadastro_clientes.Visible = False
            .painel_cadastro_compras.Visible = False
            .painel_cadastro_funcionarios.Visible = True
            .painel_cadastro_veiculos.Visible = False
            .painel_demanda.Visible = False

        End With
    End Sub



































    Sub limpar_cadastro_carros()
        With vonauto
            .txt_placa_veiculo.Clear()
            .txt_valor_carro.Clear()
            .txt_ano_veiculo.Clear()
            .cb_marca.SelectedIndex = -1
            .txt_modelo_carro.Clear()
            .txt_renavam_veiculo.Clear()
            .txt_cor_veiculo.Clear()
            .img_carro.Load(Application.StartupPath & "\ICONES\2389240701596026966-512.png")
        End With
    End Sub

    Sub limpar_cadastro_clientes()
        With vonauto
            .txt_nome_cliente.Clear()
            .txt_cpf_cliente.Clear()
            .txt_email_cliente.Clear()
            .txt_cep_cliente.Clear()
            .txt_rua_cliente.Clear()
            .txt_bairro_cliente.Clear()
            .txt_cidade_cliente.Clear()
            .txt_uf_cliente.Clear()
            .txt_numero_cliente.Clear()
            .txt_telefone_cliente.Clear()
            .dtp_cliente.Value = "31/12/2000"
        End With
    End Sub


    Sub limpar_cadastro_funcionario()
        With vonauto
            .txt_cpf_func.Clear()
            .txt_rg_func.Clear()
            .txt_nome_func.Clear()
            .dtp_data_func.Value = "31/12/2000"
            .txt_email_func.Clear()
            .cbl_vaga_func.SelectedIndex = -1
            .txt_senha_func.Clear()
            .txt_cep_func.Clear()
            .txt_endereco_func.Clear()
            .txt_cidade_func.Clear()
            .txt_bairro_func.Clear()
            .txt_uf_func.Clear()
            .txt_numero_func.Clear()
            .img_func.Load(Application.StartupPath & "\FOTOS_FUNCIONARIOS\foto_branco.png")

        End With
    End Sub

    Sub limpar_login()
        With Login
            .txt_email_login.Clear()
            .txt_senha_login.Clear()
            .txt_email_login.Focus()

        End With
    End Sub


    Sub limpar_compras()
        With vonauto
            .txt_cpf_cliente_compra.Clear()
            .txt_placa_compra.Clear()
            .txt_ano_compra.Clear()
            .lbl_valor_veic.Text = "R$ "
            .cbl_forma_pag.SelectedIndex = -1
            .lbl_Nome_cliente.Text = "Nome do Cliente: "
            .lbl_valor_total.Text = "Total a Pagar: "
            .lbl_marca_veiculo.Text = "Marca do Veiculo: "
            .lbl_modelo_veiculo.Text = "Modelo do Veiculo: "
            .img_compra.Load(Application.StartupPath & "\ICONES\Car-2-icon.png")
        End With
    End Sub
End Module
