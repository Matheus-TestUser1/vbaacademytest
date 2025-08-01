Attribute VB_Name = "modSistema"
' Módulo do Sistema - Sistema Madeireira Maria Luiza
' Versão: 1.0
' Data: 01/08/2025
' Autor: Matheus-TestUser1

Option Explicit

' Estrutura para item de venda
Public Type ItemVenda
    Produto As Produto
    Quantidade As Double
    Observacao As String
    Total As Double
End Type

' Estrutura para dados do cliente
Public Type DadosCliente
    Nome As String
    Endereco As String
    CPF As String
    UF As String
    CEP As String
End Type

' Constantes do sistema
Public Const VENDEDOR_PADRAO As String = "Matheus-TestUser1"
Public Const DATA_PADRAO As String = "01/08/2025"
Public Const PLANILHA_DESTINO As String = "marialuiza(1)"

' Variáveis globais
Public ItensVenda() As ItemVenda
Public ContadorItens As Integer
Public ClienteAtual As DadosCliente

' Inicializar sistema
Public Sub InicializarSistema()
    ' Inicializar produtos
    Call InicializarProdutos
    
    ' Limpar itens de venda
    ContadorItens = 0
    ReDim ItensVenda(1 To 100) ' Máximo 100 itens por venda
    
    ' Limpar dados do cliente
    ClienteAtual.Nome = ""
    ClienteAtual.Endereco = ""
    ClienteAtual.CPF = ""
    ClienteAtual.UF = ""
    ClienteAtual.CEP = ""
End Sub

' Adicionar item à venda
Public Sub AdicionarItemVenda(produto As Produto, quantidade As Double, observacao As String)
    If ContadorItens >= 100 Then
        MsgBox "Máximo de 100 itens por venda atingido!", vbExclamation
        Exit Sub
    End If
    
    ContadorItens = ContadorItens + 1
    ItensVenda(ContadorItens).Produto = produto
    ItensVenda(ContadorItens).Quantidade = quantidade
    ItensVenda(ContadorItens).Observacao = observacao
    ItensVenda(ContadorItens).Total = produto.Valor * quantidade
End Sub

' Remover item da venda
Public Sub RemoverItemVenda(indice As Integer)
    Dim i As Integer
    
    If indice < 1 Or indice > ContadorItens Then
        Exit Sub
    End If
    
    ' Mover itens para cima
    For i = indice To ContadorItens - 1
        ItensVenda(i) = ItensVenda(i + 1)
    Next i
    
    ContadorItens = ContadorItens - 1
End Sub

' Calcular total da venda
Public Function CalcularTotalVenda() As Double
    Dim total As Double
    Dim i As Integer
    
    total = 0
    For i = 1 To ContadorItens
        total = total + ItensVenda(i).Total
    Next i
    
    CalcularTotalVenda = total
End Function

' Validar dados do cliente
Public Function ValidarDadosCliente(cliente As DadosCliente) As Boolean
    If Trim(cliente.Nome) = "" Then
        MsgBox "Nome do cliente é obrigatório!", vbExclamation
        ValidarDadosCliente = False
        Exit Function
    End If
    
    If Trim(cliente.CPF) = "" Then
        MsgBox "CPF do cliente é obrigatório!", vbExclamation
        ValidarDadosCliente = False
        Exit Function
    End If
    
    ValidarDadosCliente = True
End Function

' Transferir venda para talão
Public Sub TransferirParaTalao()
    Dim ws As Worksheet
    Dim i As Integer
    Dim linhaEsquerda As Integer
    Dim linhaDireita As Integer
    
    ' Validar dados
    If Not ValidarDadosCliente(ClienteAtual) Then
        Exit Sub
    End If
    
    If ContadorItens = 0 Then
        MsgBox "Nenhum item adicionado à venda!", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se a planilha existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PLANILHA_DESTINO)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Planilha '" & PLANILHA_DESTINO & "' não encontrada!", vbCritical
        Exit Sub
    End If
    
    ' Preencher dados do cliente - Lado Esquerdo
    ws.Range("B7").Value = ClienteAtual.Nome ' Nome do Cliente
    ws.Range("B8").Value = ClienteAtual.Endereco ' Endereço
    ws.Range("B9").Value = ClienteAtual.CPF ' CPF/CNPJ
    ws.Range("F9").Value = ClienteAtual.UF ' UF
    ws.Range("H9").Value = ClienteAtual.CEP ' CEP
    
    ' Preencher dados do cliente - Lado Direito
    ws.Range("O7").Value = ClienteAtual.Nome ' Nome do Cliente
    ws.Range("O8").Value = ClienteAtual.Endereco ' Endereço
    ws.Range("O9").Value = ClienteAtual.CPF ' CPF/CNPJ
    ws.Range("S9").Value = ClienteAtual.UF ' UF
    ws.Range("U9").Value = ClienteAtual.CEP ' CEP
    
    ' Preencher produtos (máximo 11 linhas de produtos: 11-21)
    linhaEsquerda = 11
    linhaDireita = 11
    
    For i = 1 To ContadorItens
        If i <= 11 Then ' Lado esquerdo (A-F)
            ws.Cells(linhaEsquerda, 1).Value = ItensVenda(i).Produto.Nome ' Descrição
            ws.Cells(linhaEsquerda, 2).Value = ItensVenda(i).Produto.Unidade ' UN
            ws.Cells(linhaEsquerda, 3).Value = ItensVenda(i).Produto.Valor ' Valor
            ws.Cells(linhaEsquerda, 4).Value = ItensVenda(i).Quantidade ' Qtd
            ws.Cells(linhaEsquerda, 5).Value = ItensVenda(i).Observacao ' Obs
            ws.Cells(linhaEsquerda, 6).Value = ItensVenda(i).Total ' Total
            linhaEsquerda = linhaEsquerda + 1
        ElseIf i <= 22 Then ' Lado direito (L-R)
            ws.Cells(linhaDireita, 12).Value = ItensVenda(i).Produto.Codigo ' Código
            ws.Cells(linhaDireita, 13).Value = ItensVenda(i).Produto.Nome ' Descrição
            ws.Cells(linhaDireita, 14).Value = ItensVenda(i).Produto.Unidade ' UN
            ws.Cells(linhaDireita, 15).Value = ItensVenda(i).Produto.Valor ' Valor
            ws.Cells(linhaDireita, 16).Value = ItensVenda(i).Quantidade ' Qtd
            ws.Cells(linhaDireita, 17).Value = ItensVenda(i).Observacao ' Obs
            ws.Cells(linhaDireita, 18).Value = ItensVenda(i).Total ' Total
            linhaDireita = linhaDireita + 1
        End If
    Next i
    
    ' Preencher totais
    Dim totalVenda As Double
    totalVenda = CalcularTotalVenda()
    ws.Range("I25").Value = totalVenda ' Total esquerdo
    ws.Range("U25").Value = totalVenda ' Total direito
    
    ' Preencher data e vendedor
    ws.Range("B5").Value = DATA_PADRAO ' Data esquerdo
    ws.Range("O5").Value = DATA_PADRAO ' Data direito
    ws.Range("B6").Value = VENDEDOR_PADRAO ' Vendedor esquerdo
    ws.Range("O6").Value = VENDEDOR_PADRAO ' Vendedor direito
    
    MsgBox "Venda transferida com sucesso para o talão!" & vbCrLf & _
           "Total: R$ " & Format(totalVenda, "#,##0.00") & vbCrLf & _
           "Itens: " & ContadorItens, vbInformation
End Sub

' Limpar venda atual
Public Sub LimparVenda()
    ContadorItens = 0
    ReDim ItensVenda(1 To 100)
    
    ClienteAtual.Nome = ""
    ClienteAtual.Endereco = ""
    ClienteAtual.CPF = ""
    ClienteAtual.UF = ""
    ClienteAtual.CEP = ""
End Sub

' Obter resumo da venda
Public Function ObterResumoVenda() As String
    Dim resumo As String
    Dim i As Integer
    
    resumo = "RESUMO DA VENDA" & vbCrLf & String(30, "-") & vbCrLf
    resumo = resumo & "Cliente: " & ClienteAtual.Nome & vbCrLf
    resumo = resumo & "Itens: " & ContadorItens & vbCrLf
    resumo = resumo & "Total: R$ " & Format(CalcularTotalVenda(), "#,##0.00") & vbCrLf
    resumo = resumo & String(30, "-") & vbCrLf
    
    For i = 1 To ContadorItens
        resumo = resumo & i & ". " & ItensVenda(i).Produto.Nome & vbCrLf
        resumo = resumo & "   Qtd: " & ItensVenda(i).Quantidade & " " & ItensVenda(i).Produto.Unidade
        resumo = resumo & " - R$ " & Format(ItensVenda(i).Total, "#,##0.00") & vbCrLf
        If ItensVenda(i).Observacao <> "" Then
            resumo = resumo & "   Obs: " & ItensVenda(i).Observacao & vbCrLf
        End If
    Next i
    
    ObterResumoVenda = resumo
End Function

' Formatar valor monetário
Public Function FormatarMoeda(valor As Double) As String
    FormatarMoeda = "R$ " & Format(valor, "#,##0.00")
End Function

' Validar número
Public Function EhNumero(texto As String) As Boolean
    EhNumero = IsNumeric(texto) And Len(Trim(texto)) > 0
End Function

' Definir dados do cliente
Public Sub DefinirDadosCliente(nome As String, endereco As String, cpf As String, uf As String, cep As String)
    ClienteAtual.Nome = Trim(nome)
    ClienteAtual.Endereco = Trim(endereco)
    ClienteAtual.CPF = Trim(cpf)
    ClienteAtual.UF = Trim(UCase(uf))
    ClienteAtual.CEP = Trim(cep)
End Sub

' Obter lista de itens para exibição
Public Function ObterListaItens() As String()
    Dim lista() As String
    Dim i As Integer
    
    If ContadorItens = 0 Then
        ReDim lista(1 To 1)
        lista(1) = "Nenhum item adicionado"
        ObterListaItens = lista
        Exit Function
    End If
    
    ReDim lista(1 To ContadorItens)
    
    For i = 1 To ContadorItens
        lista(i) = ItensVenda(i).Produto.Codigo & " - " & _
                  ItensVenda(i).Produto.Nome & " - " & _
                  "Qtd: " & ItensVenda(i).Quantidade & " - " & _
                  FormatarMoeda(ItensVenda(i).Total)
    Next i
    
    ObterListaItens = lista
End Function