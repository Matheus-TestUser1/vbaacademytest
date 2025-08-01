Attribute VB_Name = "modProdutos"
' Módulo de Produtos - Sistema Madeireira Maria Luiza
' Versão: 1.0
' Data: 01/08/2025
' Autor: Matheus-TestUser1

Option Explicit

' Estrutura para produtos
Public Type Produto
    Codigo As String
    Nome As String
    Secao As String
    Unidade As String
    Valor As Double
End Type

' Array global de produtos
Public Produtos(1 To 67) As Produto

' Inicializar base de dados de produtos
Public Sub InicializarProdutos()
    ' PORTA
    Produtos(1) = CriarProduto("000001", "PORTA MADEIRA MACICA 210X80", "PORTA", "UN", 180.00)
    Produtos(2) = CriarProduto("000002", "PORTA MADEIRA MACICA 210X70", "PORTA", "UN", 165.00)
    Produtos(3) = CriarProduto("000003", "PORTA MADEIRA MACICA 210X60", "PORTA", "UN", 150.00)
    Produtos(4) = CriarProduto("000004", "PORTA COMPENSADO 210X80", "PORTA", "UN", 120.00)
    Produtos(5) = CriarProduto("000005", "PORTA COMPENSADO 210X70", "PORTA", "UN", 110.00)
    
    ' MADEIRA
    Produtos(6) = CriarProduto("000006", "MADEIRA PEROBA 4X4 3M", "MADEIRA", "UN", 45.00)
    Produtos(7) = CriarProduto("000007", "MADEIRA PEROBA 6X6 3M", "MADEIRA", "UN", 85.00)
    Produtos(8) = CriarProduto("000008", "MADEIRA ANGELIM 4X4 3M", "MADEIRA", "UN", 38.00)
    Produtos(9) = CriarProduto("000009", "MADEIRA ANGELIM 6X6 3M", "MADEIRA", "UN", 75.00)
    Produtos(10) = CriarProduto("000010", "MADEIRA EUCALIPTO 4X4 3M", "MADEIRA", "UN", 25.00)
    
    ' ESTRUTURAL
    Produtos(11) = CriarProduto("000011", "BARROTE MASSARANDUBA 6X12 3M", "ESTRUTURAL", "UN", 65.00)
    Produtos(12) = CriarProduto("000012", "BARROTE MASSARANDUBA 6X12 4M", "ESTRUTURAL", "UN", 85.00)
    Produtos(13) = CriarProduto("000013", "BARROTE MASSARANDUBA 6X12 5M", "ESTRUTURAL", "UN", 105.00)
    Produtos(14) = CriarProduto("000014", "BARROTE MISTO 6X12 3M", "ESTRUTURAL", "UN", 45.00)
    Produtos(15) = CriarProduto("000015", "BARROTE MISTO 6X12 4M", "ESTRUTURAL", "UN", 60.00)
    Produtos(16) = CriarProduto("000016", "BARROTE MISTO 6X12 5M", "ESTRUTURAL", "UN", 75.00)
    Produtos(17) = CriarProduto("000017", "CAIBRO MASSARANDUBA 6X8 3M", "ESTRUTURAL", "UN", 42.00)
    Produtos(18) = CriarProduto("000018", "CAIBRO MASSARANDUBA 6X8 4M", "ESTRUTURAL", "UN", 56.00)
    Produtos(19) = CriarProduto("000019", "CAIBRO MISTO 6X8 3M", "ESTRUTURAL", "UN", 28.00)
    Produtos(20) = CriarProduto("000020", "CAIBRO MISTO 6X8 4M", "ESTRUTURAL", "UN", 38.00)
    Produtos(21) = CriarProduto("000021", "ESTRONCAS EUCALIPTO 3M", "ESTRUTURAL", "UN", 15.00)
    
    ' ACABAMENTO
    Produtos(22) = CriarProduto("000022", "ALCATEX 60MM", "ACABAMENTO", "M", 12.50)
    Produtos(23) = CriarProduto("000023", "ALCATEX 70MM", "ACABAMENTO", "M", 14.50)
    Produtos(24) = CriarProduto("000024", "ALCATEX 80MM", "ACABAMENTO", "M", 16.50)
    Produtos(25) = CriarProduto("000025", "ALIZARES MASSARANDUBA 5CM", "ACABAMENTO", "M", 18.00)
    Produtos(26) = CriarProduto("000026", "ALIZARES MASSARANDUBA 7CM", "ACABAMENTO", "M", 22.00)
    Produtos(27) = CriarProduto("000027", "ALIZARES MISTA 5CM", "ACABAMENTO", "M", 12.00)
    Produtos(28) = CriarProduto("000028", "ALIZARES MISTA 7CM", "ACABAMENTO", "M", 16.00)
    Produtos(29) = CriarProduto("000029", "FORRA MASSARANDUBA 1X10 3M", "ACABAMENTO", "UN", 28.00)
    Produtos(30) = CriarProduto("000030", "FORRA MASSARANDUBA 1X10 4M", "ACABAMENTO", "UN", 38.00)
    Produtos(31) = CriarProduto("000031", "FORRA MISTA 1X10 3M", "ACABAMENTO", "UN", 18.00)
    Produtos(32) = CriarProduto("000032", "FORRA MISTA 1X10 4M", "ACABAMENTO", "UN", 24.00)
    Produtos(33) = CriarProduto("000033", "LINHA MASSARANDUBA 2X2", "ACABAMENTO", "M", 8.50)
    Produtos(34) = CriarProduto("000034", "LINHA MASSARANDUBA 3X3", "ACABAMENTO", "M", 12.00)
    Produtos(35) = CriarProduto("000035", "LINHA MISTA 2X2", "ACABAMENTO", "M", 5.50)
    Produtos(36) = CriarProduto("000036", "LINHA MISTA 3X3", "ACABAMENTO", "M", 8.00)
    Produtos(37) = CriarProduto("000037", "RIPA MASSARANDUBA 1X5 3M", "ACABAMENTO", "UN", 12.00)
    Produtos(38) = CriarProduto("000038", "RIPA MASSARANDUBA 1X5 4M", "ACABAMENTO", "UN", 16.00)
    Produtos(39) = CriarProduto("000039", "RIPA MISTA 1X5 3M", "ACABAMENTO", "UN", 8.00)
    Produtos(40) = CriarProduto("000040", "RIPA MISTA 1X5 4M", "ACABAMENTO", "UN", 11.00)
    
    ' COMPENSADO
    Produtos(41) = CriarProduto("000041", "MADEIRITE COMPENSADO 15MM", "COMPENSADO", "CHP", 75.00)
    Produtos(42) = CriarProduto("000042", "MADEIRITE COMPENSADO 18MM", "COMPENSADO", "CHP", 85.00)
    Produtos(43) = CriarProduto("000043", "MADEIRITE COMPENSADO 20MM", "COMPENSADO", "CHP", 95.00)
    Produtos(44) = CriarProduto("000044", "COMPENSADO NAVAL 15MM", "COMPENSADO", "CHP", 120.00)
    Produtos(45) = CriarProduto("000045", "COMPENSADO NAVAL 18MM", "COMPENSADO", "CHP", 135.00)
    Produtos(46) = CriarProduto("000046", "COMPENSADO NAVAL 20MM", "COMPENSADO", "CHP", 150.00)
    Produtos(47) = CriarProduto("000047", "COMPENSADO PLASTIFICADO 15MM", "COMPENSADO", "CHP", 165.00)
    Produtos(48) = CriarProduto("000048", "COMPENSADO PLASTIFICADO 18MM", "COMPENSADO", "CHP", 185.00)
    
    ' TABUA
    Produtos(49) = CriarProduto("000049", "TABUA MASSARANDUBA 2.5X20 3M", "TABUA", "UN", 45.00)
    Produtos(50) = CriarProduto("000050", "TABUA MASSARANDUBA 2.5X20 4M", "TABUA", "UN", 60.00)
    Produtos(51) = CriarProduto("000051", "TABUA MASSARANDUBA 2.5X20 5M", "TABUA", "UN", 75.00)
    Produtos(52) = CriarProduto("000052", "TABUA MASSARANDUBA 2.5X25 3M", "TABUA", "UN", 55.00)
    Produtos(53) = CriarProduto("000053", "TABUA MASSARANDUBA 2.5X25 4M", "TABUA", "UN", 75.00)
    Produtos(54) = CriarProduto("000054", "TABUA MASSARANDUBA 2.5X25 5M", "TABUA", "UN", 95.00)
    Produtos(55) = CriarProduto("000055", "TABUA MASSARANDUBA 2.5X30 3M", "TABUA", "UN", 65.00)
    Produtos(56) = CriarProduto("000056", "TABUA MASSARANDUBA 2.5X30 4M", "TABUA", "UN", 85.00)
    Produtos(57) = CriarProduto("000057", "TABUA MASSARANDUBA 2.5X30 5M", "TABUA", "UN", 105.00)
    Produtos(58) = CriarProduto("000058", "TABUA MISTA 2.5X20 3M", "TABUA", "UN", 28.00)
    Produtos(59) = CriarProduto("000059", "TABUA MISTA 2.5X20 4M", "TABUA", "UN", 38.00)
    Produtos(60) = CriarProduto("000060", "TABUA MISTA 2.5X20 5M", "TABUA", "UN", 48.00)
    Produtos(61) = CriarProduto("000061", "TABUA MISTA 2.5X25 3M", "TABUA", "UN", 35.00)
    Produtos(62) = CriarProduto("000062", "TABUA MISTA 2.5X25 4M", "TABUA", "UN", 46.00)
    Produtos(63) = CriarProduto("000063", "TABUA MISTA 2.5X25 5M", "TABUA", "UN", 58.00)
    Produtos(64) = CriarProduto("000064", "TABUA MISTA 2.5X30 3M", "TABUA", "UN", 42.00)
    Produtos(65) = CriarProduto("000065", "TABUA MISTA 2.5X30 4M", "TABUA", "UN", 56.00)
    Produtos(66) = CriarProduto("000066", "TABUA MISTA 2.5X30 5M", "TABUA", "UN", 70.00)
    Produtos(67) = CriarProduto("000067", "TABUA PINUS 2.5X20 3M", "TABUA", "UN", 22.00)
End Sub

' Função auxiliar para criar produto
Private Function CriarProduto(codigo As String, nome As String, secao As String, unidade As String, valor As Double) As Produto
    CriarProduto.Codigo = codigo
    CriarProduto.Nome = nome
    CriarProduto.Secao = secao
    CriarProduto.Unidade = unidade
    CriarProduto.Valor = valor
End Function

' Buscar produto por código
Public Function BuscarProdutoPorCodigo(codigo As String) As Produto
    Dim i As Integer
    For i = 1 To 67
        If UCase(Produtos(i).Codigo) = UCase(codigo) Then
            BuscarProdutoPorCodigo = Produtos(i)
            Exit Function
        End If
    Next i
    ' Retorna produto vazio se não encontrado
    BuscarProdutoPorCodigo = CriarProduto("", "", "", "", 0)
End Function

' Buscar produtos por nome (busca parcial)
Public Function BuscarProdutosPorNome(nome As String) As Produto()
    Dim resultados() As Produto
    Dim contador As Integer
    Dim i As Integer
    
    contador = 0
    
    ' Primeira passagem: contar resultados
    For i = 1 To 67
        If InStr(1, UCase(Produtos(i).Nome), UCase(nome)) > 0 Then
            contador = contador + 1
        End If
    Next i
    
    If contador = 0 Then
        BuscarProdutosPorNome = resultados
        Exit Function
    End If
    
    ' Segunda passagem: preencher resultados
    ReDim resultados(1 To contador)
    contador = 0
    
    For i = 1 To 67
        If InStr(1, UCase(Produtos(i).Nome), UCase(nome)) > 0 Then
            contador = contador + 1
            resultados(contador) = Produtos(i)
        End If
    Next i
    
    BuscarProdutosPorNome = resultados
End Function

' Buscar produtos por seção
Public Function BuscarProdutosPorSecao(secao As String) As Produto()
    Dim resultados() As Produto
    Dim contador As Integer
    Dim i As Integer
    
    contador = 0
    
    ' Primeira passagem: contar resultados
    For i = 1 To 67
        If UCase(Produtos(i).Secao) = UCase(secao) Then
            contador = contador + 1
        End If
    Next i
    
    If contador = 0 Then
        BuscarProdutosPorSecao = resultados
        Exit Function
    End If
    
    ' Segunda passagem: preencher resultados
    ReDim resultados(1 To contador)
    contador = 0
    
    For i = 1 To 67
        If UCase(Produtos(i).Secao) = UCase(secao) Then
            contador = contador + 1
            resultados(contador) = Produtos(i)
        End If
    Next i
    
    BuscarProdutosPorSecao = resultados
End Function

' Obter todas as seções disponíveis
Public Function ObterSecoes() As String()
    Dim secoes() As String
    ReDim secoes(1 To 6)
    
    secoes(1) = "PORTA"
    secoes(2) = "MADEIRA"
    secoes(3) = "ESTRUTURAL"
    secoes(4) = "ACABAMENTO"
    secoes(5) = "COMPENSADO"
    secoes(6) = "TABUA"
    
    ObterSecoes = secoes
End Function

' Contar produtos por seção
Public Function ContarProdutosPorSecao(secao As String) As Integer
    Dim contador As Integer
    Dim i As Integer
    
    contador = 0
    For i = 1 To 67
        If UCase(Produtos(i).Secao) = UCase(secao) Then
            contador = contador + 1
        End If
    Next i
    
    ContarProdutosPorSecao = contador
End Function

' Obter todos os produtos
Public Function ObterTodosProdutos() As Produto()
    ObterTodosProdutos = Produtos
End Function