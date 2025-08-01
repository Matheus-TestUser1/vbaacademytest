# Sistema Madeireira Maria Luiza - VBA Completo

## 📋 Descrição
Sistema completo em VBA para gerenciamento de vendas da Madeireira Maria Luiza, incluindo interface gráfica, base de dados de produtos e integração com planilha de talão.

## 🗂️ Arquivos do Sistema

### 1. **modProdutos.bas**
Módulo contendo a base de dados completa com 67 produtos da madeireira organizados por categorias:
- **PORTA** (5 produtos): Portas de madeira maciça e compensado
- **MADEIRA** (5 produtos): Madeiras Peroba, Angelim e Eucalipto
- **ESTRUTURAL** (11 produtos): Barrotes, Caibros e Estroncas
- **ACABAMENTO** (19 produtos): Alcatex, Alizares, Forras, Linhas e Ripas
- **COMPENSADO** (8 produtos): Madeirite e Compensados diversos
- **TABUA** (19 produtos): Tábuas Massaranduba, Mista e Pinus

### 2. **modSistema.bas**
Módulo do sistema com funções para:
- Gerenciamento de vendas e itens
- Cálculo automático de totais
- Validação de dados
- Transferência para planilha "marialuiza(1)"
- Formatação e utilitários

### 3. **frmMadeireira.frm**
UserForm principal com interface completa para:
- Entrada de dados do cliente (Nome, Endereço, CPF, UF, CEP)
- Busca inteligente de produtos por código ou nome
- Filtro por seção de produtos
- Adição/remoção de itens com quantidade e observações
- Visualização do total da venda
- Botões com cores profissionais

### 4. **frmMadeireira.frx**
Arquivo binário do UserForm (gerado automaticamente pelo VBA)

### 5. **modTeste.bas**
Módulo de testes e utilitários com:
- Teste completo do sistema (`TestarSistema`)
- Abertura do formulário principal (`AbrirSistemaMadeireira`)  
- Geração de relatório de produtos (`GerarRelatorioProdutos`)

## 🎯 Funcionalidades Principais

### Busca de Produtos
- **Por Código**: Digite o código do produto (ex: 000022) para localização direta
- **Por Nome**: Digite parte do nome (ex: ALCATEX) para busca parcial
- **Por Seção**: Filtre produtos por categoria usando o ComboBox

### Gestão de Vendas
- Adicione múltiplos produtos com quantidades específicas
- Inclua observações para cada item
- Visualize o total calculado automaticamente
- Remova itens da venda se necessário

### Transferência para Talão
Os dados são transferidos automaticamente para a planilha "marialuiza(1)" nas seguintes posições:

**Dados do Cliente:**
- B7/O7: Nome do Cliente
- B8/O8: Endereço
- B9/O9: CPF/CNPJ
- F9/S9: UF
- H9/U9: CEP

**Produtos (Linhas 11-21):**
- **Lado Esquerdo (A-F)**: Descrição, UN, Valor, Qtd, Obs, Total
- **Lado Direito (L-R)**: Código, Descrição, UN, Valor, Qtd, Obs, Total

**Totais:**
- I25/U25: Valor Total Final

**Sistema:**
- B5/O5: Data (01/08/2025)
- B6/O6: Vendedor (Matheus-TestUser1)

## 🎨 Interface

### Cores dos Botões
- **Finalizar Venda**: Verde (#28a745)
- **Buscar/Adicionar**: Azul (#007bff)
- **Cancelar/Remover**: Vermelho (#dc3545)
- **Limpar**: Cinza (#6c757d)

### Layout
- Formulário organizado em seções lógicas
- Controles intuitivos e bem posicionados
- Labels informativos para cada campo
- ListBoxes para seleção de produtos e itens

## 🚀 Como Usar

1. **Importar Arquivos**: 
   - Importe os módulos .bas no VBA Editor
   - Importe o UserForm .frm/.frx

2. **Executar Sistema**:
   ```vba
   ' Abrir interface principal
   Sub AbrirSistemaMadeireira()
       frmMadeireira.Show
   End Sub
   
   ' Testar funcionalidades
   Sub TestarSistema()
       ' Executa bateria completa de testes
   End Sub
   
   ' Gerar relatório de produtos
   Sub GerarRelatorioProdutos()
       ' Cria planilha com todos os produtos
   End Sub
   ```

3. **Fluxo de Venda**:
   - Preencha os dados do cliente
   - Busque produtos por código, nome ou seção
   - Selecione produto e defina quantidade
   - Adicione à venda
   - Repita para mais produtos
   - Finalize a venda para transferir ao talão

## 🔧 Configurações

### Constantes do Sistema (modSistema.bas)
```vba
Public Const VENDEDOR_PADRAO As String = "Matheus-TestUser1"
Public Const DATA_PADRAO As String = "01/08/2025"
Public Const PLANILHA_DESTINO As String = "marialuiza(1)"
```

### Capacidades
- Máximo 100 itens por venda
- 67 produtos cadastrados
- 6 seções de produtos
- Talão com espaço para 22 itens (11 de cada lado)

## 📚 Estruturas de Dados

### Produto
```vba
Type Produto
    Codigo As String
    Nome As String
    Secao As String
    Unidade As String
    Valor As Double
End Type
```

### ItemVenda
```vba
Type ItemVenda
    Produto As Produto
    Quantidade As Double
    Observacao As String
    Total As Double
End Type
```

### DadosCliente
```vba
Type DadosCliente
    Nome As String
    Endereco As String
    CPF As String
    UF As String
    CEP As String
End Type
```

## 📄 Versão
- **Versão**: 1.0
- **Data**: 01/08/2025
- **Autor**: Matheus-TestUser1

## 🛠️ Requisitos
- Microsoft Excel com VBA habilitado
- Planilha "marialuiza(1)" existente na pasta de trabalho
- Macros habilitadas

---
*Sistema desenvolvido para Madeireira Maria Luiza - Gestão completa de vendas em VBA*