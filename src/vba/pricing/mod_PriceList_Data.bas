Attribute VB_Name = "mod_PriceList_Data"
' ==============================================================================
' Módulo: mod_PriceList_Data
' Finalidade: Extração de dados (ETL), indexação hash em memória (Scripting.Dictionary),
'             auto-provisionamento de tabelas de parâmetros e motor taxonômico
'             em 3 camadas (Seção -> Associação Exata -> Fallback Heurístico).
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' ==============================================================================
Option Explicit

' Constantes de nomenclatura de planilhas e tabelas
Public Const SHEET_PRODUTOS As String = "Produtos"
Public Const TABLE_PRODUTOS As String = "dProdutos"

Public Const SHEET_CUSTOS As String = "TabelaCustos"
Public Const TABLE_PRECOS As String = "CfPcfic"

Public Const SHEET_PARAM_GRUPOS As String = "Param_GruposPreco"
Public Const TABLE_PARAM_GRUPOS As String = "tb_Param_GruposPreco"

' ==============================================================================
' Procedimento: EnsureParameterSheetExists
' Finalidade: Verifica a existência da planilha Param_GruposPreco e da tabela
'             tb_Param_GruposPreco. Se ausentes, cria dinamicamente com cabeçalhos
'             e registro padrão, garantindo resiliência sem interrupção.
' ==============================================================================
Public Sub EnsureParameterSheetExists(Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim wsParam As Worksheet
    Dim loParam As ListObject
    Dim tblRange As Range
    
    On Error Resume Next
    Set wsParam = wb.Worksheets(SHEET_PARAM_GRUPOS)
    On Error GoTo 0
    
    If wsParam Is Nothing Then
        Set wsParam = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsParam.Name = SHEET_PARAM_GRUPOS
        wsParam.Tab.Color = RGB(20, 66, 114) ' Cerulean
    End If
    
    On Error Resume Next
    Set loParam = wsParam.ListObjects(TABLE_PARAM_GRUPOS)
    On Error GoTo 0
    
    If loParam Is Nothing Then
        ' Injeta cabeçalhos padrão e 1 registro de inicialização
        wsParam.Cells(1, 1).Value = "Cod_Grupo_Produtos"
        wsParam.Cells(1, 2).Value = "Nome_Grupo_Tabela"
        wsParam.Cells(2, 1).Value = "0"
        wsParam.Cells(2, 2).Value = "GERAL"
        
        Set tblRange = wsParam.Range("A1:B2")
        Set loParam = wsParam.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        loParam.Name = TABLE_PARAM_GRUPOS
        loParam.TableStyle = "TableStyleLight1"
        
        wsParam.Columns("A:B").AutoFit
    End If
End Sub

' ==============================================================================
' Função: LoadPriceIndex
' Finalidade: Carrega a matriz de preços da tabela CfPcfic (planilha TabelaCustos)
'             em um Scripting.Dictionary em memória (Chave = Código, Item = P.Base).
'             Atua como filtro rigoroso de inclusão.
' ==============================================================================
Public Function LoadPriceIndex(Optional ByVal wb As Workbook = Nothing) As Object
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim wsCustos As Worksheet
    Dim loPrecos As ListObject
    Dim dictPreco As Object
    Dim rawPrices As Variant
    Dim colCodIdx As Long
    Dim colPrecoIdx As Long
    Dim i As Long
    Dim totalRows As Long
    Dim skuKey As String
    Dim precoVal As Double
    
    On Error Resume Next
    Set wsCustos = wb.Worksheets(SHEET_CUSTOS)
    On Error GoTo 0
    
    If wsCustos Is Nothing Then
        Err.Raise vbObjectError + 1001, "mod_PriceList_Data.LoadPriceIndex", _
            "A planilha obrigatória '" & SHEET_CUSTOS & "' não foi encontrada na pasta de trabalho."
    End If
    
    On Error Resume Next
    Set loPrecos = wsCustos.ListObjects(TABLE_PRECOS)
    On Error GoTo 0
    
    If loPrecos Is Nothing Then
        Err.Raise vbObjectError + 1002, "mod_PriceList_Data.LoadPriceIndex", _
            "A tabela de preços '" & TABLE_PRECOS & "' não foi encontrada na planilha '" & SHEET_CUSTOS & "'."
    End If
    
    If loPrecos.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1003, "mod_PriceList_Data.LoadPriceIndex", _
            "A tabela de preços '" & TABLE_PRECOS & "' não contém registros de dados."
    End If
    
    ' Identifica dinamicamente as colunas necessárias
    colCodIdx = FindColumnIndex(loPrecos, "Código", "Codigo", "Codigo_Produto")
    colPrecoIdx = FindColumnIndex(loPrecos, "P.Base - Ex IPI/ST", "P. Base - Ex IPI/ST", "P.Base", "Preco Base")
    
    If colCodIdx = 0 Or colPrecoIdx = 0 Then
        Err.Raise vbObjectError + 1004, "mod_PriceList_Data.LoadPriceIndex", _
            "Colunas obrigatórias ('Código' ou 'P.Base - Ex IPI/ST') não foram identificadas na tabela '" & TABLE_PRECOS & "'."
    End If
    
    ' Leitura massiva em matriz de memória (Zero cell looping)
    rawPrices = loPrecos.DataBodyRange.Value2
    totalRows = UBound(rawPrices, 1)
    
    Set dictPreco = CreateObject("Scripting.Dictionary")
    dictPreco.CompareMode = 1 ' vbTextCompare para evitar inconsistência de caixa
    
    For i = 1 To totalRows
        skuKey = Trim(CStr(rawPrices(i, colCodIdx)))
        If Len(skuKey) > 0 Then
            If IsNumeric(rawPrices(i, colPrecoIdx)) Then
                precoVal = CDbl(rawPrices(i, colPrecoIdx))
            Else
                precoVal = 0#
            End If
            ' Adiciona ou atualiza a chave com o preço base
            dictPreco(skuKey) = precoVal
        End If
    Next i
    
    Set LoadPriceIndex = dictPreco
End Function

' ==============================================================================
' Função: LoadAssociationIndex
' Finalidade: Carrega a tabela de parâmetros tb_Param_GruposPreco em um
'             Scripting.Dictionary (Chave = Cod_Grupo_Produtos, Item = Nome_Grupo_Tabela).
' ==============================================================================
Public Function LoadAssociationIndex(Optional ByVal wb As Workbook = Nothing) As Object
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Call EnsureParameterSheetExists(wb)
    
    Dim wsParam As Worksheet
    Dim loParam As ListObject
    Dim dictAssoc As Object
    Dim rawAssoc As Variant
    Dim colCodGrupo As Long
    Dim colNomeGrupo As Long
    Dim i As Long
    Dim totalRows As Long
    Dim groupKey As String
    Dim groupName As String
    
    Set wsParam = wb.Worksheets(SHEET_PARAM_GRUPOS)
    Set loParam = wsParam.ListObjects(TABLE_PARAM_GRUPOS)
    
    Set dictAssoc = CreateObject("Scripting.Dictionary")
    dictAssoc.CompareMode = 1 ' vbTextCompare
    
    If Not loParam.DataBodyRange Is Nothing Then
        colCodGrupo = FindColumnIndex(loParam, "Cod_Grupo_Produtos", "Codigo_Grupo", "CodGrupo")
        colNomeGrupo = FindColumnIndex(loParam, "Nome_Grupo_Tabela", "Nome_Grupo", "Descricao_Grupo")
        
        If colCodGrupo > 0 And colNomeGrupo > 0 Then
            rawAssoc = loParam.DataBodyRange.Value2
            totalRows = UBound(rawAssoc, 1)
            
            For i = 1 To totalRows
                groupKey = Trim(CStr(rawAssoc(i, colCodGrupo)))
                groupName = Trim(CStr(rawAssoc(i, colNomeGrupo)))
                If Len(groupKey) > 0 And Len(groupName) > 0 Then
                    dictAssoc(groupKey) = groupName
                End If
            Next i
        End If
    End If
    
    Set LoadAssociationIndex = dictAssoc
End Function

' ==============================================================================
' Função: ResolveGroupName
' Finalidade: Executa o pipeline taxonômico das Camadas 2 e 3:
'             - Layer 2: Correspondência exata em tb_Param_GruposPreco via Cod_Grupo_Produtos.
'             - Layer 3: Correspondência heurística por palavras-chave na descrição
'                        ('CERA' -> 'CERAS', 'FREEZER' -> 'FREEZERS',
'                         'COPO' -> 'COPOS E DESCARTÁVEIS', default -> Desc_Grupo_Produtos).
' ==============================================================================
Public Function ResolveGroupName(ByVal codGrupo As Variant, ByVal descGrupo As String, ByRef dictAssoc As Object) As String
    Dim cleanCod As String
    Dim cleanDesc As String
    
    cleanCod = Trim(CStr(codGrupo))
    cleanDesc = UCase(Trim(descGrupo))
    
    ' --------------------------------------------------------------------------
    ' Layer 2: Associação Exata por Código de Grupo (PK)
    ' --------------------------------------------------------------------------
    If Len(cleanCod) > 0 And Not dictAssoc Is Nothing Then
        If dictAssoc.Exists(cleanCod) Then
            Dim matchedName As String
            matchedName = Trim(CStr(dictAssoc(cleanCod)))
            If Len(matchedName) > 0 Then
                ResolveGroupName = matchedName
                Exit Function
            End If
        End If
    End If
    
    ' --------------------------------------------------------------------------
    ' Layer 3: Fallback Heurístico por Palavra-Chave na Descrição do Grupo
    ' --------------------------------------------------------------------------
    If InStr(1, cleanDesc, "CERA", vbBinaryCompare) > 0 Then
        ResolveGroupName = "CERAS"
    ElseIf InStr(1, cleanDesc, "FREEZER", vbBinaryCompare) > 0 Then
        ResolveGroupName = "FREEZERS"
    ElseIf InStr(1, cleanDesc, "COPO", vbBinaryCompare) > 0 Then
        ResolveGroupName = "COPOS E DESCARTÁVEIS"
    ElseIf Len(cleanDesc) > 0 Then
        ResolveGroupName = cleanDesc
    Else
        ResolveGroupName = "DIVERSOS"
    End If
End Function

' ==============================================================================
' Função: ExtractAndProcessPriceData
' Finalidade: Lê o cadastro mestre dProdutos, aplica o filtro de inclusão por SKU
'             a partir de CfPcfic, resolve a taxonomia em cascata e dimensiona
'             a matriz de registros qualificados:
'             (1: Section_ID, 2: Resolved_Group_Name, 3: Codigo, 4: Descricao,
'              5: Qtd_Padrao, 6: Preco_Venda).
' ==============================================================================
Public Function ExtractAndProcessPriceData(ByRef outputRecords As Variant, ByRef outCount As Long, Optional ByVal wb As Workbook = Nothing) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim wsProd As Worksheet
    Dim loProd As ListObject
    Dim dictPreco As Object
    Dim dictAssoc As Object
    Dim rawProducts As Variant
    Dim tempRecords() As Variant
    
    Dim colCodigoProd As Long
    Dim colDescProd As Long
    Dim colCodGrupo As Long
    Dim colDescGrupo As Long
    Dim colQtdPadrao As Long
    
    Dim totalRows As Long
    Dim matchCount As Long
    Dim i As Long
    Dim j As Long
    Dim sku As String
    Dim sectionId As String
    Dim resolvedGroup As String
    Dim qtdVal As Variant
    
    ExtractAndProcessPriceData = False
    outCount = 0
    
    ' 1. Validação precoce de tabelas obrigatórias
    On Error Resume Next
    Set wsProd = wb.Worksheets(SHEET_PRODUTOS)
    On Error GoTo 0
    
    If wsProd Is Nothing Then
        Err.Raise vbObjectError + 1005, "mod_PriceList_Data.ExtractAndProcessPriceData", _
            "A planilha mestre '" & SHEET_PRODUTOS & "' não foi encontrada."
    End If
    
    On Error Resume Next
    Set loProd = wsProd.ListObjects(TABLE_PRODUTOS)
    On Error GoTo 0
    
    If loProd Is Nothing Then
        Err.Raise vbObjectError + 1006, "mod_PriceList_Data.ExtractAndProcessPriceData", _
            "A tabela mestre '" & TABLE_PRODUTOS & "' não foi encontrada na planilha '" & SHEET_PRODUTOS & "'."
    End If
    
    If loProd.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1007, "mod_PriceList_Data.ExtractAndProcessPriceData", _
            "A tabela mestre '" & TABLE_PRODUTOS & "' não possui linhas de dados."
    End If
    
    ' 2. Carregamento dos índices em memória
    Set dictPreco = LoadPriceIndex(wb)
    Set dictAssoc = LoadAssociationIndex(wb)
    
    ' 3. Mapeamento de colunas da tabela dProdutos
    colCodigoProd = FindColumnIndex(loProd, "Codigo_Produto", "Código", "Codigo")
    colDescProd = FindColumnIndex(loProd, "Descricao_Produto", "Descrição", "Descricao")
    colCodGrupo = FindColumnIndex(loProd, "Cod_Grupo_Produtos", "Cod_Grupo", "Grupo")
    colDescGrupo = FindColumnIndex(loProd, "Desc_Grupo_Produtos", "Desc_Grupo", "Descricao_Grupo")
    colQtdPadrao = FindColumnIndex(loProd, "Qtd_Padrao", "Qtd Padrão", "Qtd_Embalagem", "Qtd")
    
    If colCodigoProd = 0 Or colDescProd = 0 Then
        Err.Raise vbObjectError + 1008, "mod_PriceList_Data.ExtractAndProcessPriceData", _
            "As colunas mínimas 'Codigo_Produto' e 'Descricao_Produto' não foram localizadas em '" & TABLE_PRODUTOS & "'."
    End If
    
    ' 4. Extração em memória
    rawProducts = loProd.DataBodyRange.Value2
    totalRows = UBound(rawProducts, 1)
    
    ' Alocação temporária dinâmica (capacidade máxima = total de produtos cadastrados)
    ReDim tempRecords(1 To totalRows, 1 To 6)
    matchCount = 0
    
    ' 5. Iteração e processamento algorítmico em memória
    For i = 1 To totalRows
        sku = Trim(CStr(rawProducts(i, colCodigoProd)))
        
        If Len(sku) > 0 Then
            ' Filtro de Inclusão: Se o SKU não consta na matriz de preços CfPcfic, descarta
            If dictPreco.Exists(sku) Then
                matchCount = matchCount + 1
                
                ' Camada 1: Resolução de Seção (2 primeiros caracteres do código)
                sectionId = Left$(sku, 2)
                
                ' Camadas 2 e 3: Resolução de Grupo (Associação Exata -> Heurística)
                Dim vCodG As Variant: vCodG = ""
                Dim vDescG As String: vDescG = ""
                
                If colCodGrupo > 0 Then vCodG = rawProducts(i, colCodGrupo)
                If colDescGrupo > 0 Then vDescG = CStr(rawProducts(i, colDescGrupo))
                
                resolvedGroup = ResolveGroupName(vCodG, vDescG, dictAssoc)
                
                ' Quantidade padrão / embalagem
                If colQtdPadrao > 0 Then
                    qtdVal = rawProducts(i, colQtdPadrao)
                    If IsNumeric(qtdVal) And Not IsEmpty(qtdVal) Then
                        qtdVal = CLng(qtdVal)
                    Else
                        qtdVal = 1
                    End If
                Else
                    qtdVal = 1
                End If
                
                tempRecords(matchCount, 1) = sectionId
                tempRecords(matchCount, 2) = resolvedGroup
                tempRecords(matchCount, 3) = sku
                tempRecords(matchCount, 4) = Trim(CStr(rawProducts(i, colDescProd)))
                tempRecords(matchCount, 5) = qtdVal
                tempRecords(matchCount, 6) = CDbl(dictPreco(sku))
            End If
        End If
    Next i
    
    ' 6. Compactação da matriz final para o total efetivo de SKUs qualificados
    If matchCount > 0 Then
        ReDim outputRecords(1 To matchCount, 1 To 6)
        For i = 1 To matchCount
            For j = 1 To 6
                outputRecords(i, j) = tempRecords(i, j)
            Next j
        Next i
        
        outCount = matchCount
        ExtractAndProcessPriceData = True
    Else
        outputRecords = Empty
        outCount = 0
        ExtractAndProcessPriceData = False
    End If
End Function

' ==============================================================================
' Função Auxiliar: FindColumnIndex
' Finalidade: Localiza com segurança o índice 1-based de uma coluna em um ListObject
'             testando múltiplos nomes possíveis (aliases).
' ==============================================================================
Private Function FindColumnIndex(ByVal lo As ListObject, ParamArray possibleNames() As Variant) As Long
    Dim i As Long
    Dim targetName As Variant
    Dim colName As String
    
    FindColumnIndex = 0
    If lo Is Nothing Then Exit Function
    
    For Each targetName In possibleNames
        colName = Trim(CStr(targetName))
        On Error Resume Next
        FindColumnIndex = lo.ListColumns(colName).Index
        On Error GoTo 0
        If FindColumnIndex > 0 Then Exit Function
    Next targetName
    
    ' Busca adicional case-insensitive se não encontrou correspondência exata
    Dim lc As ListColumn
    For Each targetName In possibleNames
        colName = UCase(Trim(CStr(targetName)))
        For Each lc In lo.ListColumns
            If UCase(Trim(lc.Name)) = colName Then
                FindColumnIndex = lc.Index
                Exit Function
            End If
        Next lc
    Next targetName
End Function
