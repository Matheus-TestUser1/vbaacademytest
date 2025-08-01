VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMadeireira 
   Caption         =   "Sistema Madeireira Maria Luiza - v1.0"
   ClientHeight    =   11040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16680
   OleObjectBlob   =   "frmMadeireira.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMadeireira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' UserForm Principal - Sistema Madeireira Maria Luiza
' Versão: 1.0
' Data: 01/08/2025
' Autor: Matheus-TestUser1

Option Explicit

' Variáveis do formulário
Private ProdutoSelecionado As Produto
Private ResultadosBusca() As Produto

' Evento de inicialização do formulário
Private Sub UserForm_Initialize()
    ' Inicializar sistema
    Call InicializarSistema
    
    ' Configurar interface
    Call ConfigurarInterface
    
    ' Carregar seções no ComboBox
    Call CarregarSecoes
    
    ' Limpar formulário
    Call LimparFormulario
End Sub

' Configurar interface
Private Sub ConfigurarInterface()
    ' Configurar UserForm
    Me.Width = 600
    Me.Height = 500
    Me.Caption = "Sistema Madeireira Maria Luiza - v1.0"
    
    ' Adicionar controles programaticamente
    Call CriarControles
End Sub

' Criar controles do formulário
Private Sub CriarControles()
    Dim lblTitulo As MSForms.Label
    Dim lblCliente As MSForms.Label
    Dim lblEndereco As MSForms.Label
    Dim lblCPF As MSForms.Label
    Dim lblUF As MSForms.Label
    Dim lblCEP As MSForms.Label
    Dim lblBusca As MSForms.Label
    Dim lblSecao As MSForms.Label
    Dim lblProdutos As MSForms.Label
    Dim lblQuantidade As MSForms.Label
    Dim lblObservacao As MSForms.Label
    Dim lblItens As MSForms.Label
    Dim lblTotal As MSForms.Label
    
    ' Título
    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo")
    With lblTitulo
        .Caption = "SISTEMA MADEIREIRA MARIA LUIZA"
        .Font.Bold = True
        .Font.Size = 14
        .ForeColor = RGB(0, 0, 139)
        .Left = 180
        .Top = 12
        .Width = 240
        .Height = 20
    End With
    
    ' Dados do Cliente
    Set lblCliente = Me.Controls.Add("Forms.Label.1", "lblCliente")
    With lblCliente
        .Caption = "Nome do Cliente:"
        .Left = 12
        .Top = 48
        .Width = 84
        .Height = 15
    End With
    
    Set txtCliente = Me.Controls.Add("Forms.TextBox.1", "txtCliente")
    With txtCliente
        .Left = 108
        .Top = 45
        .Width = 200
        .Height = 18
    End With
    
    Set lblEndereco = Me.Controls.Add("Forms.Label.1", "lblEndereco")
    With lblEndereco
        .Caption = "Endereço:"
        .Left = 12
        .Top = 72
        .Width = 60
        .Height = 15
    End With
    
    Set txtEndereco = Me.Controls.Add("Forms.TextBox.1", "txtEndereco")
    With txtEndereco
        .Left = 108
        .Top = 69
        .Width = 200
        .Height = 18
    End With
    
    Set lblCPF = Me.Controls.Add("Forms.Label.1", "lblCPF")
    With lblCPF
        .Caption = "CPF/CNPJ:"
        .Left = 12
        .Top = 96
        .Width = 60
        .Height = 15
    End With
    
    Set txtCPF = Me.Controls.Add("Forms.TextBox.1", "txtCPF")
    With txtCPF
        .Left = 108
        .Top = 93
        .Width = 120
        .Height = 18
    End With
    
    Set lblUF = Me.Controls.Add("Forms.Label.1", "lblUF")
    With lblUF
        .Caption = "UF:"
        .Left = 240
        .Top = 96
        .Width = 24
        .Height = 15
    End With
    
    Set txtUF = Me.Controls.Add("Forms.TextBox.1", "txtUF")
    With txtUF
        .Left = 264
        .Top = 93
        .Width = 30
        .Height = 18
        .MaxLength = 2
    End With
    
    Set lblCEP = Me.Controls.Add("Forms.Label.1", "lblCEP")
    With lblCEP
        .Caption = "CEP:"
        .Left = 306
        .Top = 96
        .Width = 30
        .Height = 15
    End With
    
    Set txtCEP = Me.Controls.Add("Forms.TextBox.1", "txtCEP")
    With txtCEP
        .Left = 336
        .Top = 93
        .Width = 80
        .Height = 18
    End With
    
    ' Busca de Produtos
    Set lblBusca = Me.Controls.Add("Forms.Label.1", "lblBusca")
    With lblBusca
        .Caption = "Buscar Produto (Código/Nome):"
        .Left = 12
        .Top = 132
        .Width = 150
        .Height = 15
    End With
    
    Set txtBusca = Me.Controls.Add("Forms.TextBox.1", "txtBusca")
    With txtBusca
        .Left = 12
        .Top = 150
        .Width = 200
        .Height = 18
    End With
    
    Set btnBuscar = Me.Controls.Add("Forms.CommandButton.1", "btnBuscar")
    With btnBuscar
        .Caption = "Buscar"
        .Left = 225
        .Top = 147
        .Width = 60
        .Height = 24
        .BackColor = RGB(0, 123, 255)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Set lblSecao = Me.Controls.Add("Forms.Label.1", "lblSecao")
    With lblSecao
        .Caption = "Filtrar por Seção:"
        .Left = 300
        .Top = 132
        .Width = 100
        .Height = 15
    End With
    
    Set cmbSecao = Me.Controls.Add("Forms.ComboBox.1", "cmbSecao")
    With cmbSecao
        .Left = 300
        .Top = 150
        .Width = 120
        .Height = 18
        .Style = fmStyleDropDownList
    End With
    
    ' Lista de Produtos
    Set lblProdutos = Me.Controls.Add("Forms.Label.1", "lblProdutos")
    With lblProdutos
        .Caption = "Produtos Encontrados:"
        .Left = 12
        .Top = 180
        .Width = 120
        .Height = 15
    End With
    
    Set lstProdutos = Me.Controls.Add("Forms.ListBox.1", "lstProdutos")
    With lstProdutos
        .Left = 12
        .Top = 198
        .Width = 408
        .Height = 90
    End With
    
    ' Quantidade e Observação
    Set lblQuantidade = Me.Controls.Add("Forms.Label.1", "lblQuantidade")
    With lblQuantidade
        .Caption = "Quantidade:"
        .Left = 12
        .Top = 300
        .Width = 72
        .Height = 15
    End With
    
    Set txtQuantidade = Me.Controls.Add("Forms.TextBox.1", "txtQuantidade")
    With txtQuantidade
        .Left = 12
        .Top = 318
        .Width = 60
        .Height = 18
        .Value = "1"
    End With
    
    Set lblObservacao = Me.Controls.Add("Forms.Label.1", "lblObservacao")
    With lblObservacao
        .Caption = "Observação:"
        .Left = 84
        .Top = 300
        .Width = 72
        .Height = 15
    End With
    
    Set txtObservacao = Me.Controls.Add("Forms.TextBox.1", "txtObservacao")
    With txtObservacao
        .Left = 84
        .Top = 318
        .Width = 120
        .Height = 18
    End With
    
    Set btnAdicionar = Me.Controls.Add("Forms.CommandButton.1", "btnAdicionar")
    With btnAdicionar
        .Caption = "Adicionar Item"
        .Left = 216
        .Top = 315
        .Width = 80
        .Height = 24
        .BackColor = RGB(40, 167, 69)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Lista de Itens da Venda
    Set lblItens = Me.Controls.Add("Forms.Label.1", "lblItens")
    With lblItens
        .Caption = "Itens da Venda:"
        .Left = 12
        .Top = 354
        .Width = 100
        .Height = 15
    End With
    
    Set lstItens = Me.Controls.Add("Forms.ListBox.1", "lstItens")
    With lstItens
        .Left = 12
        .Top = 372
        .Width = 408
        .Height = 90
    End With
    
    Set btnRemover = Me.Controls.Add("Forms.CommandButton.1", "btnRemover")
    With btnRemover
        .Caption = "Remover Item"
        .Left = 432
        .Top = 372
        .Width = 80
        .Height = 24
        .BackColor = RGB(220, 53, 69)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Total
    Set lblTotal = Me.Controls.Add("Forms.Label.1", "lblTotal")
    With lblTotal
        .Caption = "Total da Venda: R$ 0,00"
        .Font.Bold = True
        .Font.Size = 12
        .ForeColor = RGB(40, 167, 69)
        .Left = 12
        .Top = 474
        .Width = 200
        .Height = 20
    End With
    
    ' Botões principais
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With btnOK
        .Caption = "Finalizar Venda"
        .Left = 240
        .Top = 474
        .Width = 90
        .Height = 30
        .BackColor = RGB(40, 167, 69)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    Set btnLimpar = Me.Controls.Add("Forms.CommandButton.1", "btnLimpar")
    With btnLimpar
        .Caption = "Limpar"
        .Left = 342
        .Top = 474
        .Width = 60
        .Height = 30
        .BackColor = RGB(108, 117, 125)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar")
    With btnCancelar
        .Caption = "Cancelar"
        .Left = 414
        .Top = 474
        .Width = 60
        .Height = 30
        .BackColor = RGB(220, 53, 69)
        .ForeColor = RGB(255, 255, 255)
    End With
End Sub

' Carregar seções no ComboBox
Private Sub CarregarSecoes()
    Dim secoes() As String
    Dim i As Integer
    
    secoes = ObterSecoes()
    cmbSecao.AddItem "TODAS"
    
    For i = 1 To UBound(secoes)
        cmbSecao.AddItem secoes(i)
    Next i
    
    cmbSecao.Value = "TODAS"
End Sub

' Evento de busca de produtos
Private Sub btnBuscar_Click()
    Call BuscarProdutos
End Sub

' Evento de mudança na seção
Private Sub cmbSecao_Change()
    Call BuscarProdutos
End Sub

' Buscar produtos
Private Sub BuscarProdutos()
    Dim textoBusca As String
    Dim secaoSelecionada As String
    Dim produtosPorNome() As Produto
    Dim produtosPorSecao() As Produto
    Dim produtosFiltrados() As Produto
    Dim i As Integer, j As Integer, contador As Integer
    
    textoBusca = Trim(txtBusca.Value)
    secaoSelecionada = cmbSecao.Value
    
    lstProdutos.Clear
    
    ' Se não há busca e seção é "TODAS", mostrar todos os produtos
    If textoBusca = "" And secaoSelecionada = "TODAS" Then
        ReDim ResultadosBusca(1 To 67)
        For i = 1 To 67
            ResultadosBusca(i) = Produtos(i)
        Next i
    ElseIf textoBusca = "" And secaoSelecionada <> "TODAS" Then
        ' Apenas filtro por seção
        produtosPorSecao = BuscarProdutosPorSecao(secaoSelecionada)
        ResultadosBusca = produtosPorSecao
    ElseIf textoBusca <> "" And secaoSelecionada = "TODAS" Then
        ' Busca por código primeiro
        Dim produtoPorCodigo As Produto
        produtoPorCodigo = BuscarProdutoPorCodigo(textoBusca)
        
        If produtoPorCodigo.Codigo <> "" Then
            ReDim ResultadosBusca(1 To 1)
            ResultadosBusca(1) = produtoPorCodigo
        Else
            ' Busca por nome
            produtosPorNome = BuscarProdutosPorNome(textoBusca)
            ResultadosBusca = produtosPorNome
        End If
    Else
        ' Busca com filtro de seção
        produtosPorNome = BuscarProdutosPorNome(textoBusca)
        
        ' Filtrar por seção
        contador = 0
        For i = 1 To UBound(produtosPorNome)
            If UCase(produtosPorNome(i).Secao) = UCase(secaoSelecionada) Then
                contador = contador + 1
            End If
        Next i
        
        If contador > 0 Then
            ReDim produtosFiltrados(1 To contador)
            j = 0
            For i = 1 To UBound(produtosPorNome)
                If UCase(produtosPorNome(i).Secao) = UCase(secaoSelecionada) Then
                    j = j + 1
                    produtosFiltrados(j) = produtosPorNome(i)
                End If
            Next i
            ResultadosBusca = produtosFiltrados
        End If
    End If
    
    ' Preencher lista
    If UBound(ResultadosBusca) > 0 Then
        For i = 1 To UBound(ResultadosBusca)
            lstProdutos.AddItem ResultadosBusca(i).Codigo & " - " & _
                              ResultadosBusca(i).Nome & " - " & _
                              ResultadosBusca(i).Secao & " - " & _
                              FormatarMoeda(ResultadosBusca(i).Valor) & "/" & _
                              ResultadosBusca(i).Unidade
        Next i
    Else
        lstProdutos.AddItem "Nenhum produto encontrado"
    End If
End Sub

' Evento de seleção de produto
Private Sub lstProdutos_Click()
    Dim indice As Integer
    indice = lstProdutos.ListIndex + 1
    
    If indice > 0 And indice <= UBound(ResultadosBusca) Then
        ProdutoSelecionado = ResultadosBusca(indice)
    End If
End Sub

' Evento de adicionar item
Private Sub btnAdicionar_Click()
    Dim quantidade As Double
    Dim observacao As String
    
    ' Validar produto selecionado
    If ProdutoSelecionado.Codigo = "" Then
        MsgBox "Selecione um produto da lista!", vbExclamation
        Exit Sub
    End If
    
    ' Validar quantidade
    If Not EhNumero(txtQuantidade.Value) Then
        MsgBox "Quantidade deve ser um número válido!", vbExclamation
        txtQuantidade.SetFocus
        Exit Sub
    End If
    
    quantidade = CDbl(txtQuantidade.Value)
    If quantidade <= 0 Then
        MsgBox "Quantidade deve ser maior que zero!", vbExclamation
        txtQuantidade.SetFocus
        Exit Sub
    End If
    
    observacao = Trim(txtObservacao.Value)
    
    ' Adicionar item
    Call AdicionarItemVenda(ProdutoSelecionado, quantidade, observacao)
    
    ' Atualizar interface
    Call AtualizarListaItens
    Call AtualizarTotal
    
    ' Limpar campos
    txtQuantidade.Value = "1"
    txtObservacao.Value = ""
    ProdutoSelecionado = CriarProduto("", "", "", "", 0)
    lstProdutos.ListIndex = -1
    
    MsgBox "Item adicionado com sucesso!", vbInformation
End Sub

' Evento de remover item
Private Sub btnRemover_Click()
    Dim indice As Integer
    indice = lstItens.ListIndex + 1
    
    If indice > 0 And indice <= ContadorItens Then
        Call RemoverItemVenda(indice)
        Call AtualizarListaItens
        Call AtualizarTotal
        MsgBox "Item removido com sucesso!", vbInformation
    Else
        MsgBox "Selecione um item para remover!", vbExclamation
    End If
End Sub

' Atualizar lista de itens
Private Sub AtualizarListaItens()
    Dim lista() As String
    Dim i As Integer
    
    lstItens.Clear
    lista = ObterListaItens()
    
    For i = 1 To UBound(lista)
        lstItens.AddItem lista(i)
    Next i
End Sub

' Atualizar total
Private Sub AtualizarTotal()
    Dim total As Double
    total = CalcularTotalVenda()
    lblTotal.Caption = "Total da Venda: " & FormatarMoeda(total)
End Sub

' Evento OK - Finalizar venda
Private Sub btnOK_Click()
    ' Definir dados do cliente
    Call DefinirDadosCliente(txtCliente.Value, txtEndereco.Value, txtCPF.Value, txtUF.Value, txtCEP.Value)
    
    ' Transferir para talão
    Call TransferirParaTalao
    
    ' Limpar formulário
    Call LimparFormulario
End Sub

' Evento Limpar
Private Sub btnLimpar_Click()
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja limpar todos os dados? Esta ação não pode ser desfeita.", vbQuestion + vbYesNo)
    
    If resposta = vbYes Then
        Call LimparFormulario
    End If
End Sub

' Evento Cancelar
Private Sub btnCancelar_Click()
    Dim resposta As VbMsgBoxResult
    
    If ContadorItens > 0 Then
        resposta = MsgBox("Existem itens na venda. Deseja realmente cancelar?", vbQuestion + vbYesNo)
        If resposta = vbNo Then Exit Sub
    End If
    
    Unload Me
End Sub

' Limpar formulário
Private Sub LimparFormulario()
    ' Limpar dados do cliente
    txtCliente.Value = ""
    txtEndereco.Value = ""
    txtCPF.Value = ""
    txtUF.Value = ""
    txtCEP.Value = ""
    
    ' Limpar busca
    txtBusca.Value = ""
    cmbSecao.Value = "TODAS"
    lstProdutos.Clear
    
    ' Limpar produto e quantidade
    txtQuantidade.Value = "1"
    txtObservacao.Value = ""
    
    ' Limpar venda
    Call LimparVenda
    
    ' Atualizar interface
    Call AtualizarListaItens
    Call AtualizarTotal
    
    ' Reset variáveis
    ProdutoSelecionado = CriarProduto("", "", "", "", 0)
    ReDim ResultadosBusca(1 To 1)
End Sub

' Declaração de controles (serão criados dinamicamente)
Private txtCliente As MSForms.TextBox
Private txtEndereco As MSForms.TextBox
Private txtCPF As MSForms.TextBox
Private txtUF As MSForms.TextBox
Private txtCEP As MSForms.TextBox
Private txtBusca As MSForms.TextBox
Private btnBuscar As MSForms.CommandButton
Private cmbSecao As MSForms.ComboBox
Private lstProdutos As MSForms.ListBox
Private txtQuantidade As MSForms.TextBox
Private txtObservacao As MSForms.TextBox
Private btnAdicionar As MSForms.CommandButton
Private lstItens As MSForms.ListBox
Private btnRemover As MSForms.CommandButton
Private lblTotal As MSForms.Label
Private btnOK As MSForms.CommandButton
Private btnLimpar As MSForms.CommandButton
Private btnCancelar As MSForms.CommandButton

' Função auxiliar para criar produto vazio
Private Function CriarProduto(codigo As String, nome As String, secao As String, unidade As String, valor As Double) As Produto
    CriarProduto.Codigo = codigo
    CriarProduto.Nome = nome
    CriarProduto.Secao = secao
    CriarProduto.Unidade = unidade
    CriarProduto.Valor = valor
End Function