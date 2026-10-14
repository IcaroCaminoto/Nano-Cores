Attribute VB_Name = "mod_CostEngine_Calculator"
' ==============================================================================
' M�dulo: mod_CostEngine_Calculator
' Camada: Pure In-Memory Math Engine
' Finalidade: Executar agrega��es num�ricas, filtro de janelas temporais m�veis,
'             snapshot de �ltima atualiza��o, resolu��o de clusters familiares
'             e montagem da matriz tabular de sa�da orientada ao cat�logo mestre.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compat�vel)
' Conformidade: SPEC-003 (Se��es 1.3, 3.1, 3.2, 3.4, 3.5 e 4)
' Proibi��es: Zero chamadas a objetos de tela (Worksheets, Range, Cells, Application).
'             100% isolado em mem�ria RAM via estruturas nativas (Variant e Scripting.Dictionary).
' ==============================================================================

Option Explicit

' ==============================================================================
' Fun��o: ComputeSKUMovingAverages
' Finalidade: Calcular a m�dia m�vel individual por SKU para janelas de dias
'             determinadas (30, 60, 90 dias) a partir da data de corte (Date - windowDays).
' Equa��o Determin�stica (Crit�rio 1):
'   Custo_Medio_SKU = (1 / c) * Sum(C_i)
'   Caso c = 0 (sem lan�amentos), atribui 0.00 sem divis�o por zero.
' Par�metros:
'   - arrERPRaw: Matriz 2D Variant contendo os dados brutos a partir da linha 4.
'   - windowDays (Long): Quantidade de dias da janela m�vel (ex: 30, 60, 90).
'   - idxCodigo (Long): �ndice 1-based da coluna de c�digo no arrERPRaw.
'   - idxData (Long): �ndice 1-based da coluna de data de atualiza��o no arrERPRaw.
'   - idxCusto (Long): �ndice 1-based da coluna de custo gerencial no arrERPRaw.
'   - outTxCount (Object ByRef, Opcional): Dicion�rio retornado com SKU -> Qtd_Transacoes.
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
    
    ' 1. Instanciação e vinculação do objeto de saída ByRef
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

    ' 2. Varredura do extrato com normalização estrita de tipos
    For r = rStart To rEnd
        valCell = arrERPRaw(r, idxCod)
        If Not IsError(valCell) And Not IsEmpty(valCell) Then
            rawSKU = Trim$(CStr(valCell))
            
            If Len(rawSKU) > 0 Then
                valCell = arrERPRaw(r, idxDate)
                
                If IsDate(valCell) Then
                    rawDate = CDate(valCell)
                    
                    ' Filtro da janela móvel
                    If rawDate >= cutoffDate Then
                        valCell = arrERPRaw(r, idxCost)
                        
                        If IsNumeric(valCell) Then
                            rawCost = CDbl(valCell)
                            
                            ' Acúmulo de soma
                            If dictSums.Exists(rawSKU) Then
                                dictSums(rawSKU) = dictSums(rawSKU) + rawCost
                            Else
                                dictSums.Add rawSKU, rawCost
                            End If
                            
                            ' Acúmulo estrutural de contagem
                            If outTxCount.Exists(rawSKU) Then
                                outTxCount(rawSKU) = CLng(outTxCount(rawSKU)) + 1
                            Else
                                outTxCount.Add rawSKU, 1&
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    ' 3. Consolidação das médias (Soma / Contagem)
    Dim skuKey As Variant
    Dim txTotal As Long
    Dim sumTotal As Double

    For Each skuKey In dictSums.Keys
        txTotal = outTxCount(skuKey)
        sumTotal = dictSums(skuKey)
        
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
' Fun��o: ComputeSKULatestUpdate
' Finalidade: Isolar o registro com a maior Data Atualiza��o por SKU (Snapshot
'             "�ltima Atualiza��o") e atribuir como custo de refer�ncia.
' Diagrama: SPEC-003, Se��o 3.5
' Par�metros:
'   - arrERPRaw: Matriz 2D Variant contendo os dados brutos a partir da linha 4.
'   - idxCodigo (Long): �ndice 1-based da coluna de c�digo no arrERPRaw.
'   - idxData (Long): �ndice 1-based da coluna de data de atualiza��o no arrERPRaw.
'   - idxCusto (Long): �ndice 1-based da coluna de custo gerencial no arrERPRaw.
'   - outTxCount (Object ByRef, Opcional): Dicion�rio retornado com SKU -> Qtd_Transacoes.
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
                valCell = arrERPRaw(r, idxDate)
                
                If IsDate(valCell) Then
                    rawDate = CDate(valCell)
                    valCell = arrERPRaw(r, idxCost)
                    
                    If IsNumeric(valCell) Then
                        rawCost = CDbl(valCell)
                        
                        If Not dictLatestDate.Exists(rawSKU) Then
                            dictLatestDate.Add rawSKU, rawDate
                            dictAvg.Add rawSKU, Round(rawCost, 2)
                            outTxCount.Add rawSKU, 1&
                        ElseIf rawDate > dictLatestDate(rawSKU) Then
                            dictLatestDate(rawSKU) = rawDate
                            dictAvg(rawSKU) = Round(rawCost, 2)
                            outTxCount(rawSKU) = 1&
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
' Fun��o: ComputeFamilyAverages
' Finalidade: Resolver fam�lias e calcular m�dias agregadas dos clusters taxon�micos.
' Contrato de Algoritmo:
'   1. Resolu��o de �rf�os: SKUs do Cadastro Mestre sem cluster recebem cluster unit�rio
'      (Cod_Familia = Cod_Produto, Custo_Medio_Familia = Custo_Medio_SKU).
'   2. Consolida��o Reflexiva: M�dia ponderada com o SKU Pai e seus respectivos Filhos.
' Par�metros:
'   - dictClusters (Object): Scripting.Dictionary Cod_Pai -> SubDict com Membros.
'   - dictSKUAvg (Object): Scripting.Dictionary SKU -> Custo_Medio_SKU.
'   - dictProductMaster (Object): Scripting.Dictionary SKU -> Descri��o.
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
    ' Etapa 1: Resolu��o de �rf�os (Cadastro Mestre vs. Cluster)
    ' Crit�rio 9: Se o produto n�o possui v�nculo, Cod_Familia = Cod_Produto
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
    ' Etapa 2: Consolida��o da M�dia do Cluster (Pai + Filhos)
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
' Fun��o: BuildFinalOutputMatrix
' Finalidade: Executar o loop diretor orientado ao Cadastro Mestre (Left Join Sem�ntico)
'             e montar a matriz 2D Variant pronta para persist�ncia em bloco �nico.
' Dimens�o: (1 To dictMaster.Count, 1 To 7)
' Colunas:
'   1. Cod_Produto
'   2. Descricao_Produto
'   3. Custo_Medio_SKU (0.00 se sem hist�rico)
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

    ' 1. Validação defensiva do condutor principal (Spine)
    If dictMaster Is Nothing Then
        Err.Raise vbObjectError + 5101, "mod_CostEngine_Calculator.BuildFinalOutputMatrix", _
                  "O catálogo mestre (dictMaster) não foi inicializado."
    End If

    Dim totalRows As Long
    totalRows = dictMaster.Count

    If totalRows = 0 Then
        BuildFinalOutputMatrix = Empty
        Exit Function
    End If

    ' 2. Dimensionamento da matriz em memória: (1 To N, 1 To 7 Colunas Contratuais)
    Dim outArray() As Variant
    ReDim outArray(1 To totalRows, 1 To 7)

    ' 3. Resolução da etiqueta da janela temporal (Janela_Dias_Base)
    Dim windowLabel As String
    If windowDays <= 0 Then
        windowLabel = "ULT"
    Else
        windowLabel = CStr(windowDays) & "D"
    End If

    Dim procTimestamp As Date
    procTimestamp = Now

    ' 4. Variáveis de iteração e leitura
    Dim i As Long: i = 1
    Dim skuKey As Variant
    Dim skuCode As String
    Dim skuDesc As String
    Dim parentCode As String
    Dim skuCost As Double
    Dim famCost As Double
    Dim txCount As Long

    ' 5. Loop diretor estritamente orientado pelas chaves do Cadastro Mestre
    For Each skuKey In dictMaster.Keys
        skuCode = Trim$(CStr(skuKey))
        skuDesc = Trim$(CStr(dictMaster(skuKey)))

        ' --- A. Resolução taxonômica da família (Pai / Órfão) ---
        If Not dictFamilyMap Is Nothing Then
            If dictFamilyMap.Exists(skuCode) Then
                parentCode = Trim$(CStr(dictFamilyMap(skuCode)))
            Else
                parentCode = skuCode ' Órfão é pai de si mesmo
            End If
        Else
            parentCode = skuCode
        End If

        ' --- B. Resolução defensiva do custo individual do SKU ---
        skuCost = 0#
        If Not dictSKUAvg Is Nothing Then
            If dictSKUAvg.Exists(skuCode) Then
                skuCost = CDbl(dictSKUAvg(skuCode))
            End If
        End If

        ' --- C. Resolução defensiva do custo da família ---
        famCost = 0#
        If Not dictFamilyAvg Is Nothing Then
            If dictFamilyAvg.Exists(parentCode) Then
                famCost = CDbl(dictFamilyAvg(parentCode))
            Else
                famCost = skuCost ' Fallback para órfão sem cálculo de cluster
            End If
        End If

        ' --- D. Resolução defensiva da contagem de transações (dictTxCount) ---
        txCount = 0&
        If Not dictTxCount Is Nothing Then
            If dictTxCount.Exists(skuCode) Then
                txCount = CLng(dictTxCount(skuCode))
            End If
        End If

        ' --- E. Povoamento das 7 colunas contratuais (SPEC-003) ---
        outArray(i, 1) = skuCode                                    ' Col 1: Cod_Produto
        outArray(i, 2) = IIf(Len(skuDesc) > 0, skuDesc, "Sem Descrição") ' Col 2: Descricao_Produto
        outArray(i, 3) = skuCost                                    ' Col 3: Custo_Medio_SKU
        outArray(i, 4) = famCost                                    ' Col 4: Custo_Medio_Familia
        outArray(i, 5) = txCount                                    ' Col 5: Qtd_Transacoes_SKU
        outArray(i, 6) = windowLabel                                ' Col 6: Janela_Dias_Base
        outArray(i, 7) = procTimestamp                              ' Col 7: Data_Processamento

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
' Fun��es Auxiliares Privadas de Convers�o Segura
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
        ' Tratamento de separador decimal regional (substitui v�rgula por ponto)
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
' Finalidade: Auditoria asseriva do motor matem�tico contra fixtures sint�ticas.
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
    
    ' 2. Cria Associa��o: 40001 � Pai de 40002
    Set mockFamilyMap = CreateObject("Scripting.Dictionary")
    mockFamilyMap("40001") = "40001"
    mockFamilyMap("40002") = "40001"
    
    Set mockClusters = CreateObject("Scripting.Dictionary")
    Set mockClusters("40001") = CreateObject("Scripting.Dictionary")
    mockClusters("40001")("40001") = True
    mockClusters("40001")("40002") = True
    
    ' 3. Fixture Sint�tica do ERP
    ' Col 1: Codigo | Col 2: Data | Col 3: Custo
    mockERP(1, 1) = "40001": mockERP(1, 2) = Date - 10: mockERP(1, 3) = 100#
    mockERP(2, 1) = "40001": mockERP(2, 2) = Date - 20: mockERP(2, 3) = 200#
    mockERP(3, 1) = "40002": mockERP(3, 2) = Date - 5:  mockERP(3, 3) = 150#
    mockERP(4, 1) = "40001": mockERP(4, 2) = Date - 100: mockERP(4, 3) = 500# ' Fora da janela de 90D
    mockERP(5, 1) = "99999": mockERP(5, 2) = Date - 1:  mockERP(5, 3) = 300# ' SKU fora do mestre
    mockERP(6, 1) = "40002": mockERP(6, 2) = Date - 2:  mockERP(6, 3) = 170# ' Lan�amento mais recente para snapshot
    
    ' 4. Executa c�lculo de 90 dias
    Set dictSKUAvg = ComputeSKUMovingAverages(mockERP, 90, 1, 2, 3, dictTxCount)
    
    If dictSKUAvg Is Nothing Or dictTxCount Is Nothing Then
        Err.Raise vbObjectError + 3000, "Test_CostEngine_Calculator_SafeZone", "Falha: dictSKUAvg ou dictTxCount retornou Nothing."
    End If
    
    ' Asser��es SKU (90D):
    ' 40001: (100 + 200) / 2 = 150.00 | Qtd = 2
    If dictSKUAvg("40001") <> 150# Or dictTxCount("40001") <> 2 Then
        Err.Raise vbObjectError + 3001, "Test_CostEngine_Calculator_SafeZone", "Falha: M�dia de 40001 incorreta."
    End If
    ' 40002: (150 + 170) / 2 = 160.00 | Qtd = 2
    If dictSKUAvg("40002") <> 160# Or dictTxCount("40002") <> 2 Then
        Err.Raise vbObjectError + 3002, "Test_CostEngine_Calculator_SafeZone", "Falha: M�dia de 40002 incorreta."
    End If
    
    ' 5. Executa teste de Snapshot de �ltima Atualiza��o
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
    
    ' 6. Executa consolida��o de fam�lias
    Set dictFamilyAvg = ComputeFamilyAverages(mockClusters, dictSKUAvg, mockMaster, mockFamilyMap)
    
    ' Fam�lia 40001: M�dia de (150 + 160) / 2 = 155.00
    If dictFamilyAvg("40001") <> 155# Then
        Err.Raise vbObjectError + 3004, "Test_CostEngine_Calculator_SafeZone", "Falha: M�dia da Fam�lia 40001 incorreta."
    End If
    
    ' �rf�o 40099: deve ter virado fam�lia de si mesmo e ter m�dia 0.00
    If Not mockFamilyMap.Exists("40099") Then
        Err.Raise vbObjectError + 3005, "Test_CostEngine_Calculator_SafeZone", "Falha: �rf�o n�o registrado no mapa de fam�lias."
    End If
    If dictFamilyAvg("40099") <> 0# Then
        Err.Raise vbObjectError + 3006, "Test_CostEngine_Calculator_SafeZone", "Falha: M�dia da Fam�lia �rf� deve ser 0.00."
    End If
    
    ' 7. Montagem da matriz final
    outMatrix = BuildFinalOutputMatrix(mockMaster, dictSKUAvg, dictFamilyAvg, mockFamilyMap, 90, dictTxCount)
    
    If UBound(outMatrix, 1) <> 3 Or UBound(outMatrix, 2) <> 7 Then
        Err.Raise vbObjectError + 3007, "Test_CostEngine_Calculator_SafeZone", "Falha: Dimens�o da matriz de sa�da inv�lida."
    End If
    
    ' Asser��o do SKU sem hist�rico na matriz final: Custo_Medio_SKU = 0.00, Qtd = 0
    If outMatrix(3, 3) <> 0# Or outMatrix(3, 5) <> 0 Then
        Err.Raise vbObjectError + 3008, "Test_CostEngine_Calculator_SafeZone", "Falha: SKU sem hist�rico n�o recebeu 0.00."
    End If
    
    Debug.Print "[PASS] Test_CostEngine_Calculator_SafeZone conclu�do com 100% de sucesso nas asser��es."
End Sub

Private Sub Run_GuardRails_Verification()
    Debug.Print "=== INICIANDO AUDITORIA DE CONTATOS E CONTAGEM ==="

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
        Debug.Print "❌ FALHA G-1: dictCounts retornou Nothing!"
    Else
        Debug.Print "✅ SUCESSO G-1: dictCounts instanciado corretamente."
    End If

    ' Guard-Rail 2: Acurácia da Contagem com Chaves Heterogêneas (String vs Double)
    If dictCounts("40001") = 2 Then
        Debug.Print "✅ SUCESSO G-2: SKU '40001' acumulou 2 transações (filtro temporal e tipagem OK)."
    Else
        Debug.Print "❌ FALHA G-2: Esperado 2, obtido: " & dictCounts("40001")
    End If

    ' Guard-Rail 3: Acurácia da Média Ponderada por Contagem
    If dictAvgs("40001") = 150# Then
        Debug.Print "✅ SUCESSO G-3: Custo médio de '40001' apurado exatamente em R$ 150,00."
    Else
        Debug.Print "❌ FALHA G-3: Esperado 150.00, obtido: " & dictAvgs("40001")
    End If

    ' Guard-Rail 4: Isolamento de SKU Adjacente
    If dictCounts("40002") = 1 And dictAvgs("40002") = 50# Then
        Debug.Print "✅ SUCESSO G-4: SKU '40002' computado com Qtd: 1 e Custo: R$ 50,00."
    Else
        Debug.Print "❌ FALHA G-4: Inconsistência no SKU '40002'."
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
    
    dSKUAvg.Add "40001", 125.50
    dFamAvg.Add "40001", 125.50
    dCount.Add "40001", 5&

    Dim mOut As Variant
    mOut = mod_CostEngine_Calculator.BuildFinalOutputMatrix(dMaster, dSKUAvg, dFamAvg, dFamMap, 90, dCount)

    ' Asserções:
    Debug.Print "--- TESTE CONSTRUTOR DE MATRIZ ---"
    Debug.Print "Linha 1 - SKU: " & mOut(1, 1) & " | Qtd: " & mOut(1, 5) & " (Esperado: 5) | Custo: " & mOut(1, 3)
    Debug.Print "Linha 2 - SKU: " & mOut(2, 1) & " | Qtd: " & mOut(2, 5) & " (Esperado: 0) | Custo: " & mOut(2, 3)
    
    If mOut(1, 5) = 5 And mOut(2, 5) = 0 And mOut(1, 3) = 125.50 And mOut(2, 3) = 0# Then
        Debug.Print "✅ SUCESSO: Construtor alinhado e defensivo!"
    Else
        Debug.Print "❌ FALHA: Divergência na matriz montada."
    End If
End Sub