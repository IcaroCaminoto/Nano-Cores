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
                                        ByVal cutoffDays As Long, _
                                        ByVal idxCod As Long, _
                                        ByVal idxDate As Long, _
                                        ByVal idxCost As Long, _
                                        ByRef outTxCount As Object) As Object
    On Error GoTo CalcErrorHandler

    Dim dictSums As Object
    Dim dictAvg As Object
    Set dictSums = CreateObject("Scripting.Dictionary")
    Set dictAvg = CreateObject("Scripting.Dictionary")
    
    Set outTxCount = CreateObject("Scripting.Dictionary")
    dictSums.CompareMode = vbTextCompare
    dictAvg.CompareMode = vbTextCompare
    outTxCount.CompareMode = vbTextCompare

    If IsEmpty(arrERPRaw) Then
        Set ComputeSKUMovingAverages = dictAvg
        Exit Function
    End If

    Dim cutoffDate As Date
    cutoffDate = Date - cutoffDays

    Dim r As Long
    Dim rStart As Long, rEnd As Long
    rStart = LBound(arrERPRaw, 1)
    rEnd = UBound(arrERPRaw, 1)

    Dim rawSKU As String
    Dim rawCost As Double
    Dim rawDate As Date
    Dim valCell As Variant

    For r = rStart To rEnd
        valCell = arrERPRaw(r, idxCod)
        If Not IsError(valCell) And Not IsEmpty(valCell) Then
            rawSKU = Trim$(CStr(valCell))
            
            If Len(rawSKU) > 0 Then
                ' Lê a célula de data convertendo inclusive seriais Double vindos de .Value2
                If TryParseERPDate(arrERPRaw(r, idxDate), rawDate) Then
                    
                    ' Filtro da janela móvel (ex: >= 22/07/2026)
                    If rawDate >= cutoffDate Then
                        valCell = arrERPRaw(r, idxCost)
                        
                        If Not IsError(valCell) And Not IsEmpty(valCell) Then
                            If IsNumeric(valCell) Then
                                rawCost = CDbl(valCell)
                                
                                ' Acúmulo de soma
                                If dictSums.Exists(rawSKU) Then
                                    dictSums(rawSKU) = CDbl(dictSums(rawSKU)) + rawCost
                                Else
                                    dictSums.Add rawSKU, rawCost
                                End If
                                
                                ' Acúmulo de contagem
                                If outTxCount.Exists(rawSKU) Then
                                    outTxCount(rawSKU) = CLng(outTxCount(rawSKU)) + 1&
                                Else
                                    outTxCount.Add rawSKU, 1&
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    ' Consolidação das médias (Soma / Contagem)
    Dim skuKey As Variant
    Dim txTotal As Long
    Dim sumTotal As Double

    For Each skuKey In dictSums.Keys
        txTotal = CLng(outTxCount(skuKey))
        sumTotal = CDbl(dictSums(skuKey))
        
        If txTotal > 0 Then
            dictAvg(skuKey) = Round(sumTotal / txTotal, 2)
        Else
            dictAvg(skuKey) = 0#
        End If
    Next skuKey

    Set ComputeSKUMovingAverages = dictAvg
    Exit Function

CalcErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Err.Raise errNum, "mod_CostEngine_Calculator.ComputeSKUMovingAverages", _
              "Falha no cálculo de médias móveis: " & errDesc
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
                                      ByVal idxCod As Long, _
                                      ByVal idxDate As Long, _
                                      ByVal idxCost As Long, _
                                      ByRef outTxCount As Object) As Object
    On Error GoTo LatestErrorHandler

    Dim dictLatestDate As Object
    Dim dictAvg As Object
    Set dictLatestDate = CreateObject("Scripting.Dictionary")
    Set dictAvg = CreateObject("Scripting.Dictionary")
    
    Set outTxCount = CreateObject("Scripting.Dictionary")
    dictLatestDate.CompareMode = vbTextCompare
    dictAvg.CompareMode = vbTextCompare
    outTxCount.CompareMode = vbTextCompare

    If IsEmpty(arrERPRaw) Then
        Set ComputeSKULatestUpdate = dictAvg
        Exit Function
    End If

    Dim r As Long
    Dim rStart As Long, rEnd As Long
    rStart = LBound(arrERPRaw, 1)
    rEnd = UBound(arrERPRaw, 1)

    Dim rawSKU As String
    Dim rawCost As Double
    Dim rawDate As Date
    Dim valCell As Variant

    For r = rStart To rEnd
        valCell = arrERPRaw(r, idxCod)
        If Not IsError(valCell) And Not IsEmpty(valCell) Then
            rawSKU = Trim$(CStr(valCell))
            
            If Len(rawSKU) > 0 Then
                ' Lê a célula de data convertendo inclusive seriais Double vindos de .Value2
                If TryParseERPDate(arrERPRaw(r, idxDate), rawDate) Then
                    valCell = arrERPRaw(r, idxCost)
                    
                    If Not IsError(valCell) And Not IsEmpty(valCell) Then
                        If IsNumeric(valCell) Then
                            rawCost = CDbl(valCell)
                            
                            If Not dictLatestDate.Exists(rawSKU) Then
                                dictLatestDate.Add rawSKU, rawDate
                                dictAvg.Add rawSKU, Round(rawCost, 2)
                                outTxCount.Add rawSKU, 1&
                            ElseIf rawDate > CDate(dictLatestDate(rawSKU)) Then
                                dictLatestDate(rawSKU) = rawDate
                                dictAvg(rawSKU) = Round(rawCost, 2)
                                outTxCount(rawSKU) = 1&
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    Set ComputeSKULatestUpdate = dictAvg
    Exit Function

LatestErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Err.Raise errNum, "mod_CostEngine_Calculator.ComputeSKULatestUpdate", _
              "Falha no snapshot de última atualização: " & errDesc
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
    On Error GoTo FamErrorHandler

    Dim dictFamAvg As Object
    Set dictFamAvg = CreateObject("Scripting.Dictionary")
    dictFamAvg.CompareMode = vbTextCompare

    If dictFamilyMap Is Nothing Then
        Set dictFamilyMap = CreateObject("Scripting.Dictionary")
        dictFamilyMap.CompareMode = vbTextCompare
    End If

    If dictClusters Is Nothing Then
        Set dictClusters = CreateObject("Scripting.Dictionary")
        dictClusters.CompareMode = vbTextCompare
    End If

    ' Etapa 1: Resolução de Órfãos (Cadastro Mestre vs. Cluster)
    Dim skuKey As Variant
    Dim skuCode As String
    Dim orphanCluster As Object

    If Not dictProductMaster Is Nothing Then
        For Each skuKey In dictProductMaster.Keys
            skuCode = Trim$(CStr(skuKey))
            If Not dictFamilyMap.Exists(skuCode) Then
                dictFamilyMap.Add skuCode, skuCode
                If Not dictClusters.Exists(skuCode) Then
                    Set orphanCluster = CreateObject("Scripting.Dictionary")
                    orphanCluster.CompareMode = vbTextCompare
                    orphanCluster.Add skuCode, True
                    dictClusters.Add skuCode, orphanCluster
                End If
            End If
        Next skuKey
    End If

    ' Etapa 2: Consolidação da Média do Cluster (Pai + Filhos)
    Dim parentKey As Variant
    Dim memberKey As Variant
    Dim clusterMembers As Object
    Dim clusterSum As Double
    Dim clusterCount As Long
    Dim memberCost As Double

    For Each parentKey In dictClusters.Keys
        clusterSum = 0#
        clusterCount = 0&
        Set clusterMembers = dictClusters(parentKey)

        For Each memberKey In clusterMembers.Keys
            skuCode = Trim$(CStr(memberKey))
            If Not dictSKUAvg Is Nothing Then
                If dictSKUAvg.Exists(skuCode) Then
                    memberCost = CDbl(dictSKUAvg(skuCode))
                    If memberCost > 0# Then
                        clusterSum = clusterSum + memberCost
                        clusterCount = clusterCount + 1&
                    End If
                End If
            End If
        Next memberKey

        If clusterCount > 0 Then
            dictFamAvg(Trim$(CStr(parentKey))) = Round(clusterSum / clusterCount, 2)
        Else
            dictFamAvg(Trim$(CStr(parentKey))) = 0#
        End If
    Next parentKey

    Set ComputeFamilyAverages = dictFamAvg
    Exit Function

FamErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Err.Raise errNum, "mod_CostEngine_Calculator.ComputeFamilyAverages", _
              "Falha na consolidação das médias de família: " & errDesc
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
                                      ByRef dictTxCount As Object) As Variant
    On Error GoTo BuildMatrixErrorHandler

    If dictMaster Is Nothing Or dictMaster.Count = 0 Then
        BuildFinalOutputMatrix = Empty
        Exit Function
    End If

    Dim totalRows As Long
    totalRows = dictMaster.Count

    Dim outArray() As Variant
    ReDim outArray(1 To totalRows, 1 To 7)

    Dim windowLabel As String
    If windowDays <= 0 Then
        windowLabel = "ULT"
    Else
        windowLabel = CStr(windowDays) & "D"
    End If

    Dim procTimestamp As Date
    procTimestamp = Now

    Dim i As Long: i = 1
    Dim skuKey As Variant
    Dim skuCode As String
    Dim skuDesc As String
    Dim parentCode As String
    Dim skuCost As Double
    Dim famCost As Double
    Dim txCount As Long

    ' Loop diretor estritamente orientado por dictMaster (Left Join Semântico)
    For Each skuKey In dictMaster.Keys
        skuCode = Trim$(CStr(skuKey))
        skuDesc = Trim$(CStr(dictMaster(skuKey)))

        ' 1. Resolve código do Pai representante
        If Not dictFamilyMap Is Nothing Then
            If dictFamilyMap.Exists(skuCode) Then
                parentCode = Trim$(CStr(dictFamilyMap(skuCode)))
            Else
                parentCode = skuCode
            End If
        Else
            parentCode = skuCode
        End If

        ' 2. Resolve Custo Médio Individual (Col 3)
        skuCost = 0#
        If Not dictSKUAvg Is Nothing Then
            If dictSKUAvg.Exists(skuCode) Then
                skuCost = CDbl(dictSKUAvg(skuCode))
            End If
        End If

        ' 3. Resolve Custo Médio da Família (Col 4)
        famCost = 0#
        If Not dictFamilyAvg Is Nothing Then
            If dictFamilyAvg.Exists(parentCode) Then
                famCost = CDbl(dictFamilyAvg(parentCode))
            Else
                famCost = skuCost
            End If
        End If

        ' 4. Resolve Quantidade de Transações do SKU (Col 5)
        txCount = 0&
        If Not dictTxCount Is Nothing Then
            If dictTxCount.Exists(skuCode) Then
                txCount = CLng(dictTxCount(skuCode))
            End If
        End If

        ' 5. Gravação nas 7 colunas contratuais (Output Contract 01)
        outArray(i, 1) = skuCode                                         ' Col 1: Cod_Produto
        outArray(i, 2) = IIf(Len(skuDesc) > 0, skuDesc, "Sem Descrição") ' Col 2: Descricao_Produto
        outArray(i, 3) = skuCost                                         ' Col 3: Custo_Medio_SKU
        outArray(i, 4) = famCost                                         ' Col 4: Custo_Medio_Familia
        outArray(i, 5) = txCount                                         ' Col 5: Qtd_Transacoes_SKU
        outArray(i, 6) = windowLabel                                     ' Col 6: Janela_Dias_Base
        outArray(i, 7) = procTimestamp                                   ' Col 7: Data_Processamento

        i = i + 1
    Next skuKey

    BuildFinalOutputMatrix = outArray
    Exit Function

BuildMatrixErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Err.Raise errNum, "mod_CostEngine_Calculator.BuildFinalOutputMatrix", _
              "Falha na montagem da matriz de saída: " & errDesc
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
' Finalidade: Auditoria assertiva do motor matemático contra fixtures sintéticas.
' ==============================================================================
Private Sub Test_CostEngine_Calculator_SafeZone()
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
    mockERP(1, 1) = "40001": mockERP(1, 2) = Date - 10:  mockERP(1, 3) = 100#
    mockERP(2, 1) = "40001": mockERP(2, 2) = Date - 20:  mockERP(2, 3) = 200#
    mockERP(3, 1) = "40002": mockERP(3, 2) = Date - 5:   mockERP(3, 3) = 150#
    mockERP(4, 1) = "40001": mockERP(4, 2) = Date - 100: mockERP(4, 3) = 500# ' Fora da janela de 90D
    mockERP(5, 1) = "99999": mockERP(5, 2) = Date - 1:   mockERP(5, 3) = 300# ' SKU fora do mestre
    mockERP(6, 1) = "40002": mockERP(6, 2) = Date - 2:   mockERP(6, 3) = 170# ' Lançamento mais recente para snapshot
    
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

' ==============================================================================
' Função Privada Auxiliar: TryParseERPDate
' Finalidade: Converte valores de data vindos de .Value2 (Double serial Excel),
'             Strings ("30/09/2026") ou Date nativo sem falhar no IsDate().
' ==============================================================================
Private Function TryParseERPDate(ByVal valCell As Variant, ByRef outDate As Date) As Boolean
    On Error GoTo ParseFail
    
    If IsError(valCell) Or IsEmpty(valCell) Then
        TryParseERPDate = False
        Exit Function
    End If
    
    ' Caso 1: Já é Date nativo ou String de data reconhecida pelo VBA
    If IsDate(valCell) Then
        outDate = CDate(valCell)
        TryParseERPDate = True
        Exit Function
    End If
    
    ' Caso 2: É um número serial Double retornado por .Value2 (ex: 46295 para 30/09/2026)
    If IsNumeric(valCell) Then
        Dim serialVal As Double
        serialVal = CDbl(valCell)
        ' Intervalo válido de datas no Excel (entre 01/01/1980 e 31/12/2199)
        If serialVal >= 29221# And serialVal <= 109574# Then
            outDate = CDate(serialVal)
            TryParseERPDate = True
            Exit Function
        End If
    End If

ParseFail:
    TryParseERPDate = False
End Function

Private Sub Run_GuardRails_Verification()
    Debug.Print "=== INICIANDO AUDITORIA DE CONTRATOS E CONTAGEM ==="

    ' Fixture Sintética: 4 registros (2 válidos dentro da janela, 1 fora, 1 outro SKU)
    Dim mockERP(1 To 4, 1 To 3) As Variant
    Dim today As Date: today = Date
    
    ' SKU-01: 2 compras recentes (100.00 e 200.00) -> Média Esperada: 150.00 | Qtd: 2
    mockERP(1, 1) = "40001": mockERP(1, 2) = today - 10: mockERP(1, 3) = 100#
    mockERP(2, 1) = 40001:   mockERP(2, 2) = today - 30: mockERP(2, 3) = 200# ' Teste de chave numérica
    
    ' SKU-01: 1 compra antiga (fora da janela de 90 dias) -> Deve ser descartada
    mockERP(3, 1) = "40001": mockERP(3, 2) = today - 95: mockERP(3, 3) = 999#
    
    ' SKU-02: 1 compra recente (50.00) -> Média Esperada: 50.00 | Qtd: 1
    mockERP(4, 1) = "40002": mockERP(4, 2) = today - 5:  mockERP(4, 3) = 50#

    Dim dictCounts As Object
    Dim dictAvgs As Object

    ' Execução do Motor de Cálculo
    Set dictAvgs = mod_CostEngine_Calculator.ComputeSKUMovingAverages( _
                       mockERP, 90, 1, 2, 3, dictCounts)

    ' Guard-Rail 1: Instanciação do Ponteiro ByRef
    If dictCounts Is Nothing Then
        Debug.Print "[ERRO] FALHA G-1: dictCounts retornou Nothing!"
    Else
        Debug.Print "[OK] SUCESSO G-1: dictCounts instanciado corretamente."
    End If

    ' Guard-Rail 2: Acurácia da Contagem com Chaves Heterogêneas (String vs Double)
    If dictCounts("40001") = 2 Then
        Debug.Print "[OK] SUCESSO G-2: SKU '40001' acumulou 2 transações (filtro temporal e tipagem OK)."
    Else
        Debug.Print "[ERRO] FALHA G-2: Esperado 2, obtido: " & dictCounts("40001")
    End If

    ' Guard-Rail 3: Acurácia da Média Ponderada por Contagem
    If dictAvgs("40001") = 150# Then
        Debug.Print "[OK] SUCESSO G-3: Custo médio de '40001' apurado exatamente em R$ 150,00."
    Else
        Debug.Print "[ERRO] FALHA G-3: Esperado 150.00, obtido: " & dictAvgs("40001")
    End If

    ' Guard-Rail 4: Isolamento de SKU Adjacente
    If dictCounts("40002") = 1 And dictAvgs("40002") = 50# Then
        Debug.Print "[OK] SUCESSO G-4: SKU '40002' computado com Qtd: 1 e Custo: R$ 50,00."
    Else
        Debug.Print "[ERRO] FALHA G-4: Inconsistência no SKU '40002'."
    End If

    Debug.Print "=== AUDITORIA CONCLUÍDA COM 0 FALHAS ==="
End Sub

Private Sub GuardRail_Test_BuildMatrix()
    Dim dMaster As Object: Set dMaster = CreateObject("Scripting.Dictionary")
    Dim dSKUAvg As Object: Set dSKUAvg = CreateObject("Scripting.Dictionary")
    Dim dFamAvg As Object: Set dFamAvg = CreateObject("Scripting.Dictionary")
    Dim dFamMap As Object: Set dFamMap = CreateObject("Scripting.Dictionary")
    Dim dCount As Object:  Set dCount = CreateObject("Scripting.Dictionary")

    ' Fixture de 2 produtos: 1 com dados e 1 sem dados (órfão)
    dMaster.Add "40001", "PRODUTO TESTE COM MOVIMENTO"
    dMaster.Add "40002", "PRODUTO TESTE SEM MOVIMENTO"

    dFamMap.Add "40001", "40001"
    
    dSKUAvg.Add "40001", 125.5
    dFamAvg.Add "40001", 125.5
    dCount.Add "40001", 5&

    Dim mOut As Variant
    mOut = mod_CostEngine_Calculator.BuildFinalOutputMatrix(dMaster, dSKUAvg, dFamAvg, dFamMap, 90, dCount)

    ' Asserções:
    Debug.Print "--- TESTE CONSTRUTOR DE MATRIZ ---"
    Debug.Print "Linha 1 - SKU: " & mOut(1, 1) & " | Qtd: " & mOut(1, 5) & " (Esperado: 5) | Custo: " & mOut(1, 3)
    Debug.Print "Linha 2 - SKU: " & mOut(2, 1) & " | Qtd: " & mOut(2, 5) & " (Esperado: 0) | Custo: " & mOut(2, 3)
    
    If mOut(1, 5) = 5 And mOut(2, 5) = 0 And mOut(1, 3) = 125.5 And mOut(2, 3) = 0# Then
        Debug.Print "[OK] SUCESSO: Construtor alinhado e defensivo!"
    Else
        Debug.Print "[ERRO] FALHA: Divergência na matriz montada."
    End If
End Sub