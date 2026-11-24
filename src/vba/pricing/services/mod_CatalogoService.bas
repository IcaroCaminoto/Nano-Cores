Attribute VB_Name = "mod_CatalogoService"
' ==============================================================================
' Módulo: mod_CatalogoService
' Camada: Application Service (PURA - pronta para DLL)
' Finalidade: Regras de inclusão de produtos na precificação, sem tocar na
'             planilha nem em formulários:
'   - cruzamento Produto x Impostos (DTOs)
'   - filtro de produtos já presentes na tabela
'   - montagem da matriz (linhas x colunas da CfPcfic) pronta para injeção
' Quem chama (mod_PricingCatalogo) lê/escreve a planilha; este módulo só
' recebe e devolve dados. Na DLL, vira a Action "catalogo.montarLinhas".
' ==============================================================================

Option Explicit

' ==============================================================================
' Função: Catalogo_ColunasCadastro
' Finalidade: Colunas da CfPcfic preenchidas a partir da base de produtos.
'             (Código de Barras é opcional: só é escrita se existir na tabela.)
' ==============================================================================
Public Function Catalogo_ColunasCadastro() As Variant
    Catalogo_ColunasCadastro = Array(Col_Descricao(), Col_QtdPadrao(), Col_Grupo(), _
                                     Col_PesoKg(), Col_CodigoBarras())
End Function

' ==============================================================================
' Função: Catalogo_ColunasLote
' Finalidade: Colunas gravadas na inclusão em lote: Código, cadastro e a
'             Operação padrão (OPERACAO_PADRAO_LOTE).
' ==============================================================================
Public Function Catalogo_ColunasLote() As Variant
    Catalogo_ColunasLote = Array(Col_Codigo(), Col_Descricao(), Col_QtdPadrao(), Col_Grupo(), _
                                 Col_PesoKg(), Col_CodigoBarras(), Col_Operacao())
End Function

' ==============================================================================
' Procedimento: Catalogo_CruzarImpostos
' Finalidade: Ligar cada produto às suas alíquotas da UF (ou Nothing).
' Parâmetros:
'   - produtos: Collection de clsProdutoDTO
'   - impostosUF: Dictionary chave -> clsAliquotasDTO (pode ser Nothing)
' ==============================================================================
Public Sub Catalogo_CruzarImpostos(ByVal produtos As Collection, ByVal impostosUF As Object)
    Dim dto As clsProdutoDTO
    For Each dto In produtos
        Set dto.Impostos = Nothing
        If Not impostosUF Is Nothing Then
            If impostosUF.Exists(dto.Chave) Then Set dto.Impostos = impostosUF(dto.Chave)
        End If
    Next dto
End Sub

' ==============================================================================
' Função: Catalogo_FiltrarNovos
' Finalidade: Separar os produtos que ainda não estão na tabela.
' Parâmetros:
'   - chavesExistentes: Dictionary com as chaves já presentes na CfPcfic
'   - outRepetidos (saída): quantos produtos foram ignorados por já existirem
' ==============================================================================
Public Function Catalogo_FiltrarNovos(ByVal produtos As Collection, ByVal chavesExistentes As Object, _
                                      Optional ByRef outRepetidos As Long) As Collection
    Dim dto As clsProdutoDTO
    Dim novos As New Collection

    outRepetidos = 0
    For Each dto In produtos
        If chavesExistentes.Exists(dto.Chave) Then
            outRepetidos = outRepetidos + 1
        Else
            novos.Add dto
        End If
    Next dto
    Set Catalogo_FiltrarNovos = novos
End Function

' ==============================================================================
' Função: Catalogo_CodigosSemImpostos
' Finalidade: Lista dos códigos sem alíquota na UF (para avisar o usuário).
' ==============================================================================
Public Function Catalogo_CodigosSemImpostos(ByVal produtos As Collection) As Collection
    Dim dto As clsProdutoDTO
    Dim lista As New Collection
    For Each dto In produtos
        If Not dto.TemImpostos Then lista.Add dto.Chave
    Next dto
    Set Catalogo_CodigosSemImpostos = lista
End Function

' ==============================================================================
' Função: Catalogo_ValorDaColuna
' Finalidade: Valor que um produto grava numa coluna da CfPcfic.
' Retorno: False se a coluna não for cadastral (não deve ser tocada).
' ==============================================================================
Public Function Catalogo_ValorDaColuna(ByVal dto As clsProdutoDTO, ByVal nomeColuna As String, _
                                       ByRef outValor As Variant) As Boolean
    Catalogo_ValorDaColuna = True
    Select Case nomeColuna
        Case Col_Codigo():        outValor = dto.Codigo
        Case Col_Descricao():     outValor = dto.Descricao
        Case Col_QtdPadrao():     outValor = dto.QtdPadrao
        Case Col_Grupo():         outValor = dto.Grupo
        Case Col_PesoKg():        outValor = dto.PesoKg
        Case Col_CodigoBarras():  outValor = dto.CodigoBarras
        Case Col_Operacao():      outValor = OPERACAO_PADRAO_LOTE   ' só usada no lote
        Case Else
            outValor = Empty
            Catalogo_ValorDaColuna = False
    End Select
End Function

' ==============================================================================
' Função: Catalogo_MontarMatriz
' Finalidade: Matriz (1..produtos, 1..colunas) com os valores de cada produto
'             na ordem das colunas pedidas. Colunas não cadastrais ficam vazias.
' ==============================================================================
Public Function Catalogo_MontarMatriz(ByVal produtos As Collection, ByVal nomesColunas As Variant) As Variant
    Dim matriz() As Variant
    Dim dto As clsProdutoDTO
    Dim valor As Variant
    Dim nCols As Long
    Dim i As Long, j As Long

    If produtos.Count = 0 Then Exit Function
    nCols = UBound(nomesColunas) - LBound(nomesColunas) + 1
    ReDim matriz(1 To produtos.Count, 1 To nCols)

    For Each dto In produtos
        i = i + 1
        For j = 1 To nCols
            If Catalogo_ValorDaColuna(dto, CStr(nomesColunas(LBound(nomesColunas) + j - 1)), valor) Then
                matriz(i, j) = valor
            End If
        Next j
    Next dto

    Catalogo_MontarMatriz = matriz
End Function
