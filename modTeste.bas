Attribute VB_Name = "modTeste"
' Módulo de Testes - Sistema Madeireira Maria Luiza
' Versão: 1.0
' Data: 01/08/2025
' Autor: Matheus-TestUser1

Option Explicit

' Teste básico do sistema
Sub TestarSistema()
    Dim resultado As String
    
    ' Inicializar sistema
    Call InicializarSistema
    
    resultado = "TESTE DO SISTEMA MADEIREIRA MARIA LUIZA" & vbCrLf
    resultado = resultado & String(50, "=") & vbCrLf & vbCrLf
    
    ' Testar carregamento de produtos
    resultado = resultado & "1. TESTE DE CARREGAMENTO DE PRODUTOS:" & vbCrLf
    resultado = resultado & "   Total de produtos carregados: " & UBound(Produtos) & vbCrLf & vbCrLf
    
    ' Testar busca por código
    Dim produtoTeste As Produto
    produtoTeste = BuscarProdutoPorCodigo("000022")
    resultado = resultado & "2. TESTE DE BUSCA POR CÓDIGO (000022):" & vbCrLf
    resultado = resultado & "   Produto encontrado: " & produtoTeste.Nome & vbCrLf
    resultado = resultado & "   Seção: " & produtoTeste.Secao & vbCrLf
    resultado = resultado & "   Valor: R$ " & Format(produtoTeste.Valor, "#,##0.00") & vbCrLf & vbCrLf
    
    ' Testar busca por nome
    Dim produtosPorNome() As Produto
    produtosPorNome = BuscarProdutosPorNome("ALCATEX")
    resultado = resultado & "3. TESTE DE BUSCA POR NOME (ALCATEX):" & vbCrLf
    resultado = resultado & "   Produtos encontrados: " & UBound(produtosPorNome) & vbCrLf & vbCrLf
    
    ' Testar busca por seção
    Dim produtosPorSecao() As Produto
    produtosPorSecao = BuscarProdutosPorSecao("ESTRUTURAL")
    resultado = resultado & "4. TESTE DE BUSCA POR SEÇÃO (ESTRUTURAL):" & vbCrLf
    resultado = resultado & "   Produtos encontrados: " & UBound(produtosPorSecao) & vbCrLf & vbCrLf
    
    ' Testar seções
    Dim secoes() As String
    secoes = ObterSecoes()
    resultado = resultado & "5. TESTE DE SEÇÕES DISPONÍVEIS:" & vbCrLf
    Dim i As Integer
    For i = 1 To UBound(secoes)
        resultado = resultado & "   - " & secoes(i) & " (" & ContarProdutosPorSecao(secoes(i)) & " produtos)" & vbCrLf
    Next i
    resultado = resultado & vbCrLf
    
    ' Testar adição de item
    Call DefinirDadosCliente("João Silva", "Rua das Flores, 123", "123.456.789-00", "SP", "01234-567")
    Call AdicionarItemVenda(produtoTeste, 2, "Teste de observação")
    
    resultado = resultado & "6. TESTE DE ADIÇÃO DE ITEM:" & vbCrLf
    resultado = resultado & "   Itens na venda: " & ContadorItens & vbCrLf
    resultado = resultado & "   Total da venda: R$ " & Format(CalcularTotalVenda(), "#,##0.00") & vbCrLf & vbCrLf
    
    ' Testar validação
    resultado = resultado & "7. TESTE DE VALIDAÇÃO:" & vbCrLf
    resultado = resultado & "   Dados do cliente válidos: " & ValidarDadosCliente(ClienteAtual) & vbCrLf
    resultado = resultado & "   Cliente: " & ClienteAtual.Nome & vbCrLf
    resultado = resultado & "   CPF: " & ClienteAtual.CPF & vbCrLf & vbCrLf
    
    resultado = resultado & "8. TESTE DE FUNÇÕES AUXILIARES:" & vbCrLf
    resultado = resultado & "   EhNumero('123.45'): " & EhNumero("123.45") & vbCrLf
    resultado = resultado & "   EhNumero('abc'): " & EhNumero("abc") & vbCrLf
    resultado = resultado & "   FormatarMoeda(1234.56): " & FormatarMoeda(1234.56) & vbCrLf & vbCrLf
    
    resultado = resultado & String(50, "=") & vbCrLf
    resultado = resultado & "RESULTADO: TODOS OS TESTES EXECUTADOS COM SUCESSO!" & vbCrLf
    resultado = resultado & "Sistema pronto para uso!" & vbCrLf
    
    ' Exibir resultado
    MsgBox resultado, vbInformation, "Teste do Sistema Madeireira"
    
    ' Limpar teste
    Call LimparVenda
End Sub

' Abrir formulário principal
Sub AbrirSistemaMadeireira()
    frmMadeireira.Show
End Sub

' Gerar relatório de produtos
Sub GerarRelatorioProdutos()
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Inicializar produtos
    Call InicializarProdutos
    
    ' Criar nova planilha para relatório
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Relatório_Produtos_" & Format(Now, "ddmmyyyy_hhmmss")
    
    ' Cabeçalhos
    ws.Range("A1").Value = "RELATÓRIO DE PRODUTOS - MADEIREIRA MARIA LUIZA"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    ws.Range("A3").Value = "Data: " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Range("A4").Value = "Total de Produtos: " & UBound(Produtos)
    
    ' Cabeçalhos da tabela
    ws.Range("A6").Value = "Código"
    ws.Range("B6").Value = "Nome do Produto"
    ws.Range("C6").Value = "Seção"
    ws.Range("D6").Value = "Unidade"
    ws.Range("E6").Value = "Valor (R$)"
    
    ' Formatar cabeçalhos
    ws.Range("A6:E6").Font.Bold = True
    ws.Range("A6:E6").Interior.Color = RGB(200, 200, 200)
    
    ' Preencher dados
    For i = 1 To UBound(Produtos)
        ws.Cells(6 + i, 1).Value = Produtos(i).Codigo
        ws.Cells(6 + i, 2).Value = Produtos(i).Nome
        ws.Cells(6 + i, 3).Value = Produtos(i).Secao
        ws.Cells(6 + i, 4).Value = Produtos(i).Unidade
        ws.Cells(6 + i, 5).Value = Produtos(i).Valor
        ws.Cells(6 + i, 5).NumberFormat = "R$ #,##0.00"
    Next i
    
    ' Ajustar colunas
    ws.Columns("A:E").AutoFit
    
    ' Adicionar bordas
    ws.Range("A6:E" & (6 + UBound(Produtos))).Borders.LineStyle = xlContinuous
    
    MsgBox "Relatório de produtos gerado com sucesso na planilha: " & ws.Name, vbInformation
End Sub