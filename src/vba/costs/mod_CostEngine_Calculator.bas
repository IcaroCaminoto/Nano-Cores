Attribute VB_Name = "mod_CostEngine_Calculator"
' ==============================================================================
' Módulo: mod_CostEngine_Calculator
' Camada: Pure In-Memory Math Engine
' Finalidade: Executar agregações numéricas, filtro de janelas temporais móveis,
'             snapshot de última atualização, resolução de clusters familiares
'             e montagem da matriz tabular de saída orientada ao catálogo mestre.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' Conformidade: SPEC-003 (Seções 1.3, 3.1, 3.2, 3.4, 3.5 e 4)
' Proibições: Zero chamadas a objetos de tela (Worksheets, Range, Cells, Application).
'             100% isolado em memória RAM via estruturas nativas (Variant e Scripting.Dictionary).
' ==============================================================================

Option Explicit

' ==============================================================================
' Função: ComputeSKUMovingAverages
' Finalidade: Calcular a média móvel individual por SKU para janelas de dias
'             determinadas (30, 60, 90 dias) a partir da data de corte (Date - windowDays).
' Equação Determinística (Critério 1):
'   Custo_Medio_SKU = (1 / c) * Sum(C_i)
'   Caso c = 0 (sem lançamentos), atribui 0.00 sem divisão por zero.
' Parâmetros:
'   - arrERPRaw: Matriz 2D Variant contendo os dados brutos a partir da linha 4.
'   - windowDays (Long): Quantidade de dias da janela móvel (ex: 30, 60, 90).
'   - idxCodigo (Long): Índice 1-based da coluna de código no arrERPRaw.
'   - idxData (Long): Índice 1-based da coluna de data de atualização no arrERPRaw.
'   - idxCusto (Long): Índice 1-based da coluna de custo gerencial no arrERPRaw.
'   - outTxCount (Object ByRef, Opcional): Dicionário retornado com SKU -> Qtd_Transacoes.
' Retorno: Scripting.Dictionary contendo SKU -> Custo_Medio_SKU (Double, 2 decimais).
' ==============================================================================
Public Function ComputeSKUMovingAverages(ByRef arrERPRaw As Variant, _
                                         ByVal windowDays As Long, _
                                         ByVal idxCodigo As Long, _
                                         ByVal idxData As Long, _
                                         ByVal idxCusto As Long, _
                                         Optional ByRef outTxCount As Object = Nothing) As Object
    
    Dim dictSKUAvg As Object
    Dim dictTxCount As Object
    Dim dictSum As Object
    
    Dim r As Long
    Dim totalRows As Long
    Dim rawCode As Variant
    Dim rawDate As Variant
    Dim rawCost As Variant
    Dim skuKey As String
    Dim txDate As Date
    Dim txCost As Double
    Dim cutoffDate As Date
    Dim isDateValid As Boolean
    Dim isCostValid As Boolean
    
    Set dictSKUAvg = CreateObject("Scripting.Dictionary")
    dictSKUAvg.CompareMode = vbTextCompare
    
    Set dictTxCount = CreateObject("Scripting.Dictionary")
    dictTxCount.CompareMode = vbTextCompare
    
    ' Se a janela for <= 0, redireciona para o snapshot de última atualização
    If windowDays <= 0 Then
        Set ComputeSKUMovingAverages = ComputeSKULatestUpdate(arrERPRaw, idxCodigo, idxData, idxCusto, outTxCount)
        Exit Function
    End If
    
    ' Defesa para conjunto vazio (Empty Set Defense)
    If IsEmpty(arrERPRaw) Then
        If Not outTxCount Is Nothing Then
            outTxCount.RemoveAll
        End If
        On Error Resume Next
        Set outTxCount = dictTxCount
        On Error GoTo 0
        Set ComputeSKUMovingAverages = dictSKUAvg
        Exit Function
    End If
    
    totalRows = UBound(arrERPRaw, 1)
    cutoffDate = Date - windowDays
    
    Set dictSum = CreateObject("Scripting.Dictionary")
    dictSum.CompareMode = vbTextCompare
    
    For r = 1 To totalRows
        rawCode = arrERPRaw(r, idxCodigo)
        rawDate = arrERPRaw(r, idxData)
        rawCost = arrERPRaw(r, idxCusto)
        
        If Not IsNull(rawCode) Then
            skuKey = Trim$(CStr(rawCode))
            
            If Len(skuKey) > 0 Then
                isDateValid = TryParseDate(rawDate, txDate)
                isCostValid = TryParseCost(rawCost, txCost)
                
                If isDateValid And isCostValid Then
                    ' Filtro determinístico da janela móvel
                    If txDate >= cutoffDate Then
                        If dictSum.Exists(skuKey) Then
                            dictSum(skuKey) = dictSum(skuKey) + txCost
                            dictTxCount(skuKey) = dictTxCount(skuKey) + 1&
                        Else
                            dictSum(skuKey) = txCost
                            dictTxCount(skuKey) = 1&
                        End If
                    End If
                End If
            End If
        End If
    Next r
    
    ' Apuração das médias aritméticas individuais (Critério 1: Tratamento de indeterminação)
    For Each rawCode In dictSum.Keys
        skuKey = CStr(rawCode)
        If dictTxCount(skuKey) > 0& Then
            dictSKUAvg(skuKey) = Round(dictSum(skuKey) / CDbl(dictTxCount(skuKey)), 2)
        Else
            dictSKUAvg(skuKey) = 0#
        End If
    Next rawCode
    
    ' Popula diretamente o dicionário pré-instanciado pelo chamador e assegura a referência
    If Not outTxCount Is Nothing Then
        outTxCount.RemoveAll
        Dim k As Variant
        For Each k In dictTxCount.Keys
            outTxCount.Item(k) = dictTxCount.Item(k)
        Next k
    End If
    On Error Resume Next
    Set outTxCount = dictTxCount
    On Error GoTo 0
    
    Set ComputeSKUMovingAverages = dictSKUAvg
End Function

' ==============================================================================
' Função: ComputeSKULatestUpdate
' Finalidade: Isolar o registro com a maior Data Atualização por SKU (Snapshot
'             "Última Atualização") e atribuir como custo de referência.
' Diagrama: SPEC-003, Seção 3.5
' Parâmetros:
'   - arrERPRaw: Matriz 2D Variant contendo os dados brutos a partir da linha 4.
'   - idxCodigo (Long): Índice 1-based da coluna de código no arrERPRaw.
'   - idxData (Long): Índice 1-based da coluna de data de atualização no arrERPRaw.
'   - idxCusto (Long): Índice 1-based da coluna de custo gerencial no arrERPRaw.
'   - outTxCount (Object ByRef, Opcional): Dicionário retornado com SKU -> Qtd_Transacoes.
' Retorno: Scripting.Dictionary contendo SKU -> Custo_Medio_SKU (Double, 2 decimais).
' ==============================================================================
Public Function ComputeSKULatestUpdate(ByRef arrERPRaw As Variant, _
                                       ByVal idxCodigo As Long, _
                                       ByVal idxData As Long, _
                                       ByVal idxCusto As Long, _
                                       Optional ByRef outTxCount As Object = Nothing) As Object
    
    Dim dictSKUAvg As Object
    Dim dictTxCount As Object
    Dim dictLatestDate As Object
    Dim dictLatestCost As Object
    
    Dim r As Long
    Dim totalRows As Long
    Dim rawCode As Variant
    Dim rawDate As Variant
    Dim rawCost As Variant
    Dim skuKey As String
    Dim txDate As Date
    Dim txCost As Double
    Dim isDateValid As Boolean
    Dim isCostValid As Boolean
    
    Set dictSKUAvg = CreateObject("Scripting.Dictionary")
    dictSKUAvg.CompareMode = vbTextCompare
    
    Set dictTxCount = CreateObject("Scripting.Dictionary")
    dictTxCount.CompareMode = vbTextCompare
    
    ' Defesa para conjunto vazio (Empty Set Defense)
    If IsEmpty(arrERPRaw) Then
        If Not outTxCount Is Nothing Then
            outTxCount.RemoveAll
        End If
        On Error Resume Next
        Set outTxCount = dictTxCount
        On Error GoTo 0
        Set ComputeSKULatestUpdate = dictSKUAvg
        Exit Function
    End If
    
    totalRows = UBound(arrERPRaw, 1)
    
    Set dictLatestDate = CreateObject("Scripting.Dictionary")
    dictLatestDate.CompareMode = vbTextCompare
    
    Set dictLatestCost = CreateObject("Scripting.Dictionary")
    dictLatestCost.CompareMode = vbTextCompare
    
    For r = 1 To totalRows
        rawCode = arrERPRaw(r, idxCodigo)
        rawDate = arrERPRaw(r, idxData)
        rawCost = arrERPRaw(r, idxCusto)
        
        If Not IsNull(rawCode) Then
            skuKey = Trim$(CStr(rawCode))
            
            If Len(skuKey) > 0 Then
                isDateValid = TryParseDate(rawDate, txDate)
                isCostValid = TryParseCost(rawCost, txCost)
                
                If isDateValid And isCostValid Then
                    If Not dictLatestDate.Exists(skuKey) Then
                        dictLatestDate(skuKey) = txDate
                        dictLatestCost(skuKey) = txCost
                    Else
                        If txDate > dictLatestDate(skuKey) Then
                            dictLatestDate(skuKey) = txDate
                            dictLatestCost(skuKey) = txCost
                        End If
                    End If
                End If
            End If
        End If
    Next r
    
    ' Consolidação do snapshot de última atualização
    For Each rawCode In dictLatestCost.Keys
        skuKey = CStr(rawCode)
        dictSKUAvg(skuKey) = Round(CDbl(dictLatestCost(skuKey)), 2)
        dictTxCount(skuKey) = 1&
    Next rawCode
    
    ' Popula diretamente o dicionário pré-instanciado pelo chamador e assegura a referência
    If Not outTxCount Is Nothing Then
        outTxCount.RemoveAll
        Dim k As Variant
        For Each k In dictTxCount.Keys
            outTxCount.Item(k) = dictTxCount.Item(k)
        Next k
    End If
    On Error Resume Next
    Set outTxCount = dictTxCount
    On Error GoTo 0
    
    Set ComputeSKULatestUpdate = dictSKUAvg
End Function

' ==============================================================================
' Função: ComputeFamilyAverages
' Finalidade: Resolver famílias e calcular médias agregadas dos clusters taxonômicos.
' Contrato de Algoritmo:
'   1. Resolução de Órfãos: SKUs do Cadastro Mestre sem cluster recebem cluster unitário
'      (Cod_Familia = Cod_Produto, Custo_Medio_Familia = Custo_Medio_SKU).
'   2. Consolidação Reflexiva: Média ponderada com o SKU Pai e seus respectivos Filhos.
' Parâmetros:
'   - dictClusters (Object): Scripting.Dictionary Cod_Pai -> SubDict com Membros.
'   - dictSKUAvg (Object): Scripting.Dictionary SKU -> Custo_Medio_SKU.
'   - dictProductMaster (Object): Scripting.Dictionary SKU -> Descrição.
'   - dictFamilyMap (Object): Scripting.Dictionary SKU -> Cod_Pai.
' Retorno: Scripting.Dictionary Cod_Pai -> Custo_Medio_Familia.
' ==============================================================================
Public Function ComputeFamilyAverages(ByRef dictClusters As Object, _
                                      ByRef dictSKUAvg As Object, _
                                      ByRef dictProductMaster As Object, _
                                      ByRef dictFamilyMap As Object) As Object
    Dim dictFamilyAvg As Object
    Dim skuItem As Variant
    Dim skuCode As String
    Dim parentKey As Variant
    Dim memberItem As Variant
    Dim memberCode As String
    Dim clusterSubDict As Object
    Dim clusterSum As Double
    Dim clusterCount As Long
    Dim memberCost As Double
    
    Set dictFamilyAvg = CreateObject("Scripting.Dictionary")
    dictFamilyAvg.CompareMode = vbTextCompare
    
    ' --------------------------------------------------------------------------
    ' Etapa 1: Resolução de Órfãos (Cadastro Mestre vs. Cluster)
    ' Critério 9: Se o produto não possui vínculo, Cod_Familia = Cod_Produto
    ' --------------------------------------------------------------------------
    If Not dictProductMaster Is Nothing Then
        For Each skuItem In dictProductMaster.Keys
            skuCode = CStr(skuItem)
            
            If Not dictFamilyMap.Exists(skuCode) Then
                dictFamilyMap(skuCode) = skuCode
                
                If Not dictClusters.Exists(skuCode) Then
                    Set clusterSubDict = CreateObject("Scripting.Dictionary")
                    clusterSubDict.CompareMode = vbTextCompare
                    clusterSubDict(skuCode) = True
                    Set dictClusters(skuCode) = clusterSubDict
                End If
            End If
        Next skuItem
    End If
    
    ' --------------------------------------------------------------------------
    ' Etapa 2: Consolidação da Média do Cluster (Pai + Filhos)
    ' --------------------------------------------------------------------------
    For Each parentKey In dictClusters.Keys
        clusterSum = 0#
        clusterCount = 0&
        Set clusterSubDict = dictClusters(parentKey)
        
        For Each memberItem In clusterSubDict.Keys
            memberCode = CStr(memberItem)
            
            If dictSKUAvg.Exists(memberCode) Then
                memberCost = CDbl(dictSKUAvg(memberCode))
                If memberCost > 0# Then
                    clusterSum = clusterSum + memberCost
                    clusterCount = clusterCount + 1&
                End If
            End If
        Next memberItem
        
        If clusterCount > 0& Then
            dictFamilyAvg(parentKey) = Round(clusterSum / CDbl(clusterCount), 2)
        Else
            dictFamilyAvg(parentKey) = 0#
        End If
    Next parentKey
    
    Set ComputeFamilyAverages = dictFamilyAvg
End Function

' ==============================================================================
' Função: BuildFinalOutputMatrix
' Finalidade: Executar o loop diretor orientado ao Cadastro Mestre (Left Join Semântico)
'             e montar a matriz 2D Variant pronta para persistência em bloco único.
' Dimensão: (1 To dictMaster.Count, 1 To 7)
' Colunas:
'   1. Cod_Produto
'   2. Descricao_Produto
'   3. Custo_Medio_SKU (0.00 se sem histórico)
'   4. Custo_Medio_Familia
'   5. Qtd_Transacoes_SKU
'   6. Janela_Dias_Base ("ULT", "30D", "60D", "90D")
'   7. Data_Processamento
' ==============================================================================
Public Function BuildFinalOutputMatrix(ByRef dictMaster As Object, _
                                       ByRef dictSKUAvg As Object, _
                                       ByRef dictFamilyAvg As Object, _
                                       ByRef dictFamilyMap As Object, _
                                       ByVal windowDays As Long, _
                                       Optional ByRef dictTxCount As Object = Nothing) As Variant
    Dim outArray() As Variant
    Dim masterCount As Long
    Dim i As Long
    Dim skuKey As Variant
    Dim skuCode As String
    Dim parentCode As String
    Dim skuCost As Double
    Dim familyCost As Double
    Dim txQty As Long
    Dim windowLabel As String
    Dim procTimestamp As Date
    
    If dictMaster Is Nothing Or dictMaster.Count = 0 Then
        BuildFinalOutputMatrix = Empty
        Exit Function
    End If
    
    masterCount = dictMaster.Count
    ReDim outArray(1 To masterCount, 1 To 7)
    
    ' Definição do rótulo da janela
    If windowDays <= 0 Then
        windowLabel = "ULT"
    Else
        windowLabel = CStr(windowDays) & "D"
    End If
    
    procTimestamp = Now
    i = 0
    
    ' Loop diretor pelo Cadastro Mestre (Preservação Integral do Catálogo - Critérios 6 e 7)
    For Each skuKey In dictMaster.Keys
        i = i + 1
        skuCode = CStr(skuKey)
        
        ' Resolução taxonômica do Pai
        If dictFamilyMap.Exists(skuCode) Then
            parentCode = CStr(dictFamilyMap(skuCode))
        Else
            parentCode = skuCode
        End If
        
        ' Resolução de custo individual (Critério 1 e 7: sentinela 0.00 se sem histórico)
        If dictSKUAvg.Exists(skuCode) Then
            skuCost = CDbl(dictSKUAvg(skuCode))
        Else
            skuCost = 0#
        End If
        
        ' Resolução de custo da família (Critério 9: fallback se cluster vazio)
        If dictFamilyAvg.Exists(parentCode) Then
            familyCost = CDbl(dictFamilyAvg(parentCode))
        Else
            familyCost = skuCost
        End If
        
        ' Resolução de transações
        If Not dictTxCount Is Nothing Then
            If dictTxCount.Exists(skuCode) Then
                txQty = CLng(dictTxCount(skuCode))
            Else
                txQty = 0&
            End If
        Else
            txQty = 0&
        End If
        
        ' Atribuição pura em memória
        outArray(i, 1) = skuCode
        outArray(i, 2) = CStr(dictMaster(skuCode))
        outArray(i, 3) = skuCost
        outArray(i, 4) = familyCost
        outArray(i, 5) = txQty
        outArray(i, 6) = windowLabel
        outArray(i, 7) = procTimestamp
    Next skuKey
    
    BuildFinalOutputMatrix = outArray
End Function

' ==============================================================================
' Funções Auxiliares Privadas de Conversão Segura
' ==============================================================================
Private Function TryParseDate(ByVal valInput As Variant, ByRef outDate As Date) As Boolean
    On Error Resume Next
    If IsDate(valInput) Then
        outDate = CDate(valInput)
        TryParseDate = True
    Else
        TryParseDate = False
    End If
    On Error GoTo 0
End Function

Private Function TryParseCost(ByVal valInput As Variant, ByRef outCost As Double) As Boolean
    Dim strVal As String
    On Error Resume Next
    
    If IsNumeric(valInput) Then
        outCost = CDbl(valInput)
        TryParseCost = True
        Exit Function
    End If
    
    If VarType(valInput) = vbString Then
        strVal = Trim$(CStr(valInput))
        ' Tratamento de separador decimal regional (substitui vírgula por ponto)
        strVal = Replace(strVal, "R$", "")
        strVal = Replace(strVal, " ", "")
        strVal = Replace(strVal, ",", ".")
        If IsNumeric(strVal) Then
            outCost = Val(strVal)
            TryParseCost = True
            Exit Function
        End If
    End If
    
    TryParseCost = False
    On Error GoTo 0
End Function

' ==============================================================================
' Procedimento SafeZone: Test_CostEngine_Calculator_SafeZone
' Finalidade: Auditoria asseriva do motor matemático contra fixtures sintéticas.
' ==============================================================================
Public Sub Test_CostEngine_Calculator_SafeZone()
    Dim mockMaster As Object
    Dim mockFamilyMap As Object
    Dim mockClusters As Object
    Dim mockERP(1 To 6, 1 To 3) As Variant
    Dim dictSKUAvg As Object
    Dim dictTxCount As Object
    Dim dictFamilyAvg As Object
    Dim outMatrix As Variant
    
    ' 1. Cria Cadastro Mestre com 3 SKUs
    Set mockMaster = CreateObject("Scripting.Dictionary")
    mockMaster("40001") = "PRODUTO BASE PAI"
    mockMaster("40002") = "PRODUTO DERIVADO FILHO"
    mockMaster("40099") = "PRODUTO ORFAO SEM HISTORICO"
    
    ' 2. Cria Associação: 40001 é Pai de 40002
    Set mockFamilyMap = CreateObject("Scripting.Dictionary")
    mockFamilyMap("40001") = "40001"
    mockFamilyMap("40002") = "40001"
    
    Set mockClusters = CreateObject("Scripting.Dictionary")
    Set mockClusters("40001") = CreateObject("Scripting.Dictionary")
    mockClusters("40001")("40001") = True
    mockClusters("40001")("40002") = True
    
    ' 3. Fixture Sintética do ERP
    ' Col 1: Codigo | Col 2: Data | Col 3: Custo
    mockERP(1, 1) = "40001": mockERP(1, 2) = Date - 10: mockERP(1, 3) = 100#
    mockERP(2, 1) = "40001": mockERP(2, 2) = Date - 20: mockERP(2, 3) = 200#
    mockERP(3, 1) = "40002": mockERP(3, 2) = Date - 5:  mockERP(3, 3) = 150#
    mockERP(4, 1) = "40001": mockERP(4, 2) = Date - 100: mockERP(4, 3) = 500# ' Fora da janela de 90D
    mockERP(5, 1) = "99999": mockERP(5, 2) = Date - 1:  mockERP(5, 3) = 300# ' SKU fora do mestre
    mockERP(6, 1) = "40002": mockERP(6, 2) = Date - 2:  mockERP(6, 3) = 170# ' Lançamento mais recente para snapshot
    
    ' 4. Executa cálculo de 90 dias
    Set dictSKUAvg = ComputeSKUMovingAverages(mockERP, 90, 1, 2, 3, dictTxCount)
    
    If dictSKUAvg Is Nothing Or dictTxCount Is Nothing Then
        Err.Raise vbObjectError + 3000, "Test_CostEngine_Calculator_SafeZone", "Falha: dictSKUAvg ou dictTxCount retornou Nothing."
    End If
    
    ' Asserções SKU (90D):
    ' 40001: (100 + 200) / 2 = 150.00 | Qtd = 2
    If dictSKUAvg("40001") <> 150# Or dictTxCount("40001") <> 2 Then
        Err.Raise vbObjectError + 3001, "Test_CostEngine_Calculator_SafeZone", "Falha: Média de 40001 incorreta."
    End If
    ' 40002: (150 + 170) / 2 = 160.00 | Qtd = 2
    If dictSKUAvg("40002") <> 160# Or dictTxCount("40002") <> 2 Then
        Err.Raise vbObjectError + 3002, "Test_CostEngine_Calculator_SafeZone", "Falha: Média de 40002 incorreta."
    End If
    
    ' 5. Executa teste de Snapshot de Última Atualização
    Dim dictSnapAvg As Object
    Dim dictSnapCount As Object
    Set dictSnapAvg = ComputeSKULatestUpdate(mockERP, 1, 2, 3, dictSnapCount)
    
    If dictSnapAvg Is Nothing Or dictSnapCount Is Nothing Then
        Err.Raise vbObjectError + 3009, "Test_CostEngine_Calculator_SafeZone", "Falha: dictSnapAvg ou dictSnapCount retornou Nothing."
    End If
    
    ' 40002 mais recente foi em Date - 2 com valor 170
    If dictSnapAvg("40002") <> 170# Or dictSnapCount("40002") <> 1 Then
        Err.Raise vbObjectError + 3003, "Test_CostEngine_Calculator_SafeZone", "Falha: Snapshot de 40002 incorreto."
    End If
    
    ' 6. Executa consolidação de famílias
    Set dictFamilyAvg = ComputeFamilyAverages(mockClusters, dictSKUAvg, mockMaster, mockFamilyMap)
    
    ' Família 40001: Média de (150 + 160) / 2 = 155.00
    If dictFamilyAvg("40001") <> 155# Then
        Err.Raise vbObjectError + 3004, "Test_CostEngine_Calculator_SafeZone", "Falha: Média da Família 40001 incorreta."
    End If
    
    ' Órfão 40099: deve ter virado família de si mesmo e ter média 0.00
    If Not mockFamilyMap.Exists("40099") Then
        Err.Raise vbObjectError + 3005, "Test_CostEngine_Calculator_SafeZone", "Falha: Órfão não registrado no mapa de famílias."
    End If
    If dictFamilyAvg("40099") <> 0# Then
        Err.Raise vbObjectError + 3006, "Test_CostEngine_Calculator_SafeZone", "Falha: Média da Família Órfã deve ser 0.00."
    End If
    
    ' 7. Montagem da matriz final
    outMatrix = BuildFinalOutputMatrix(mockMaster, dictSKUAvg, dictFamilyAvg, mockFamilyMap, 90, dictTxCount)
    
    If UBound(outMatrix, 1) <> 3 Or UBound(outMatrix, 2) <> 7 Then
        Err.Raise vbObjectError + 3007, "Test_CostEngine_Calculator_SafeZone", "Falha: Dimensão da matriz de saída inválida."
    End If
    
    ' Asserção do SKU sem histórico na matriz final: Custo_Medio_SKU = 0.00, Qtd = 0
    If outMatrix(3, 3) <> 0# Or outMatrix(3, 5) <> 0 Then
        Err.Raise vbObjectError + 3008, "Test_CostEngine_Calculator_SafeZone", "Falha: SKU sem histórico não recebeu 0.00."
    End If
    
    Debug.Print "[PASS] Test_CostEngine_Calculator_SafeZone concluído com 100% de sucesso nas asserções."
End Sub
