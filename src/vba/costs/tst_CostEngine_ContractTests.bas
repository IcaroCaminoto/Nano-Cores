Attribute VB_Name = "tst_CostEngine_ContractTests"
' ==============================================================================
' Módulo de Testes: tst_CostEngine_ContractTests
' Finalidade: Suíte de testes automatizados e asserções em memória pura para
'             validação dos contratos da SPEC-003 conforme diretrizes da skill
'             code-spec-validator.
' Conformidade: 100% não-destrutivo, executado em estruturas sintéticas isoladas.
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Contadores globais de teste (Assertion Helpers)
' ------------------------------------------------------------------------------
Public TotalTests As Long
Public PassedTests As Long
Public FailedTests As Long

Public Sub ResetTestCounters()
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
End Sub

Public Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal testName As String)
    TotalTests = TotalTests + 1
    If actual = expected Then
        PassedTests = PassedTests + 1
        Debug.Print " [PASS] " & testName
    Else
        FailedTests = FailedTests + 1
        Debug.Print "![FAIL] " & testName & " | Esperado: [" & expected & "] mas obteve: [" & actual & "]"
    End If
End Sub

Public Sub AssertDoubleApprox(ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double, ByVal testName As String)
    TotalTests = TotalTests + 1
    If Abs(actual - expected) <= tolerance Then
        PassedTests = PassedTests + 1
        Debug.Print " [PASS] " & testName
    Else
        FailedTests = FailedTests + 1
        Debug.Print "![FAIL] " & testName & " | Esperado ~[" & expected & "] +/- " & tolerance & " mas obteve: [" & actual & "]"
    End If
End Sub

Public Sub PrintTestSummary()
    Debug.Print "=========================================================="
    Debug.Print "EXECUÇÃO DE TESTES CONCLUÍDA: " & TotalTests & " Asserções Executadas."
    Debug.Print " Aprovados: " & PassedTests & " | Falhas: " & FailedTests
    Debug.Print "=========================================================="
    If FailedTests > 0 Then
        Err.Raise vbObjectError + 999, "tst_CostEngine_ContractTests", _
                  "Uma ou mais asserções contratuais falharam durante a auditoria da SPEC-003!"
    End If
End Sub

' ==============================================================================
' Ponto de Entrada: RunAllCostEngineContractTests
' ==============================================================================
Public Sub RunAllCostEngineContractTests()
    ResetTestCounters
    Debug.Print "--- INICIANDO PROTOCOLO DE AUDITORIA CODE-SPEC-VALIDATOR (SPEC-003) ---"
    
    Test_01_MovingAverage_90DayBoundaryCalculation
    Test_02_Snapshot_LatestUpdateIsolation
    Test_03_IndeterminationAndZeroDivisionDefense
    Test_04_TaxonomicReflexiveCluster
    Test_05_OrphanSKUResolution
    Test_06_ProductMasterLeftJoinPreservation
    Test_07_FinalOutputMatrixSchemaFidelity
    Test_08_ApplicationStateRestorationTrap
    Test_09_CheckFileLockMechanism
    
    PrintTestSummary
End Sub

' ------------------------------------------------------------------------------
' Teste 1: Limite da Janela Móvel de 90 dias (Date - 89, -90 e -91)
' ------------------------------------------------------------------------------
Private Sub Test_01_MovingAverage_90DayBoundaryCalculation()
    Dim mockERP(1 To 3, 1 To 3) As Variant
    Dim today As Date: today = Date
    Dim dictAvg As Object
    Dim dictCount As Object
    
    ' Record 1: 89 dias atrás (DEVE ser incluído)
    mockERP(1, 1) = "SKU-BOUND": mockERP(1, 2) = today - 89: mockERP(1, 3) = 100#
    ' Record 2: Exatamente 90 dias atrás (DEVE ser incluído)
    mockERP(2, 1) = "SKU-BOUND": mockERP(2, 2) = today - 90: mockERP(2, 3) = 200#
    ' Record 3: 91 dias atrás (DEVE ser excluído)
    mockERP(3, 1) = "SKU-BOUND": mockERP(3, 2) = today - 91: mockERP(3, 3) = 999#
    
    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictAvg = ComputeSKUMovingAverages(mockERP, 90, 1, 2, 3, dictCount)
    
    ' Validação de retorno de instâncias
    If dictAvg Is Nothing Then
        AssertEqual "Nothing", "Object", "Corte 90D: dictAvg deve ser instanciado e retornado"
    ElseIf Not dictAvg.Exists("SKU-BOUND") Then
        AssertEqual False, True, "Corte 90D: Chave 'SKU-BOUND' deve existir em dictAvg"
    Else
        ' Média esperada: (100 + 200) / 2 = 150.00
        AssertDoubleApprox CDbl(dictAvg("SKU-BOUND")), 150#, 0.001, _
                           "Corte 90D: Considera apenas transações dentro da janela (Date - 89 e Date - 90)"
    End If
    
    If dictCount Is Nothing Then
        AssertEqual "Nothing", "Object", "Corte 90D: dictCount deve ser instanciado e retornado (Not Nothing)"
    ElseIf Not dictCount.Exists("SKU-BOUND") Then
        AssertEqual False, True, "Corte 90D: Chave 'SKU-BOUND' deve existir em dictCount"
    Else
        AssertEqual CLng(dictCount("SKU-BOUND")), 2&, _
                    "Corte 90D: Cardinalidade de lançamentos válidos deve ser exatamente 2"
    End If
End Sub

' ------------------------------------------------------------------------------
' Teste 2: Snapshot de Última Atualização (Maior Data por SKU)
' ------------------------------------------------------------------------------
Private Sub Test_02_Snapshot_LatestUpdateIsolation()
    Dim mockERP(1 To 3, 1 To 3) As Variant
    Dim today As Date: today = Date
    Dim dictAvg As Object
    Dim dictCount As Object
    
    ' Lançamentos em datas distintas para o mesmo SKU
    mockERP(1, 1) = "SKU-SNAP": mockERP(1, 2) = today - 30: mockERP(1, 3) = 50#
    mockERP(2, 1) = "SKU-SNAP": mockERP(2, 2) = today - 5:  mockERP(2, 3) = 175.5
    mockERP(3, 1) = "SKU-SNAP": mockERP(3, 2) = today - 10: mockERP(3, 3) = 120#
    
    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictAvg = ComputeSKULatestUpdate(mockERP, 1, 2, 3, dictCount)
    
    If dictAvg Is Nothing Then
        AssertEqual "Nothing", "Object", "Snapshot: dictAvg deve ser instanciado e retornado"
    ElseIf Not dictAvg.Exists("SKU-SNAP") Then
        AssertEqual False, True, "Snapshot: Chave 'SKU-SNAP' deve existir em dictAvg"
    Else
        AssertDoubleApprox CDbl(dictAvg("SKU-SNAP")), 175.5, 0.001, _
                           "Snapshot: Isola o custo da data mais recente (Date - 5)"
    End If
    
    If dictCount Is Nothing Then
        AssertEqual "Nothing", "Object", "Snapshot: dictCount deve ser instanciado e retornado (Not Nothing)"
    ElseIf Not dictCount.Exists("SKU-SNAP") Then
        AssertEqual False, True, "Snapshot: Chave 'SKU-SNAP' deve existir em dictCount"
    Else
        AssertEqual CLng(dictCount("SKU-SNAP")), 1&, _
                    "Snapshot: Cardinalidade de transações avaliadas é unitária (1)"
    End If
End Sub

' ------------------------------------------------------------------------------
' Teste 3: Tratamento de Indeterminação e Defesa contra Divisão por Zero (c = 0)
' ------------------------------------------------------------------------------
Private Sub Test_03_IndeterminationAndZeroDivisionDefense()
    Dim mockERP(1 To 1, 1 To 3) As Variant
    Dim dictAvg As Object
    Dim dictCount As Object
    Dim txCountVal As Long
    
    ' Registro fora da janela (120 dias atrás)
    mockERP(1, 1) = "SKU-ANTIGO": mockERP(1, 2) = Date - 120: mockERP(1, 3) = 500#
    
    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictAvg = ComputeSKUMovingAverages(mockERP, 90, 1, 2, 3, dictCount)
    
    ' SKU-ANTIGO não teve nenhum registro na janela de 90 dias
    Dim costOut As Double
    If Not dictAvg Is Nothing Then
        If dictAvg.Exists("SKU-ANTIGO") Then
            costOut = CDbl(dictAvg("SKU-ANTIGO"))
        Else
            costOut = 0#
        End If
    Else
        costOut = -1#
    End If
    
    AssertEqual costOut, 0#, _
                "Critério 1: SKU sem transações na janela avalia estritamente para 0.00 sem divisão por zero"
                
    If dictCount Is Nothing Then
        AssertEqual "Nothing", "Object", "Critério 1: dictCount deve ser instanciado e retornado"
    Else
        If dictCount.Exists("SKU-ANTIGO") Then
            txCountVal = CLng(dictCount("SKU-ANTIGO"))
        Else
            txCountVal = 0&
        End If
        AssertEqual txCountVal, 0&, _
                    "Critério 1: Quantidade de transações avaliadas deve ser 0"
    End If
End Sub

' ------------------------------------------------------------------------------
' Teste 4: Garantia Reflexiva do Cluster Familiar (Pai + Filhos)
' ------------------------------------------------------------------------------
Private Sub Test_04_TaxonomicReflexiveCluster()
    Dim mockMaster As Object
    Dim mockFamilyMap As Object
    Dim mockClusters As Object
    Dim mockSKUAvg As Object
    Dim dictFamilyAvg As Object
    
    Set mockMaster = CreateObject("Scripting.Dictionary")
    mockMaster("PAI-10") = "PRODUTO PAI"
    mockMaster("FILHO-11") = "PRODUTO FILHO 1"
    mockMaster("FILHO-12") = "PRODUTO FILHO 2"
    
    Set mockFamilyMap = CreateObject("Scripting.Dictionary")
    mockFamilyMap("PAI-10") = "PAI-10"
    mockFamilyMap("FILHO-11") = "PAI-10"
    mockFamilyMap("FILHO-12") = "PAI-10"
    
    Set mockClusters = CreateObject("Scripting.Dictionary")
    Set mockClusters("PAI-10") = CreateObject("Scripting.Dictionary")
    mockClusters("PAI-10")("PAI-10") = True   ' Membro #1 (Garantia Reflexiva)
    mockClusters("PAI-10")("FILHO-11") = True
    mockClusters("PAI-10")("FILHO-12") = True
    
    Set mockSKUAvg = CreateObject("Scripting.Dictionary")
    mockSKUAvg("PAI-10") = 100#
    mockSKUAvg("FILHO-11") = 150#
    mockSKUAvg("FILHO-12") = 200#
    
    Set dictFamilyAvg = ComputeFamilyAverages(mockClusters, mockSKUAvg, mockMaster, mockFamilyMap)
    
    ' Média esperada do cluster: (100 + 150 + 200) / 3 = 150.00
    AssertDoubleApprox CDbl(dictFamilyAvg("PAI-10")), 150#, 0.001, _
                       "Critério 8: Média da família computa o conjunto reflexivo unificado (Pai + Filhos)"
End Sub

' ------------------------------------------------------------------------------
' Teste 5: Resolução de Órfãos (Sem Associação no Banco)
' ------------------------------------------------------------------------------
Private Sub Test_05_OrphanSKUResolution()
    Dim mockMaster As Object
    Dim mockFamilyMap As Object
    Dim mockClusters As Object
    Dim mockSKUAvg As Object
    Dim dictFamilyAvg As Object
    
    Set mockMaster = CreateObject("Scripting.Dictionary")
    mockMaster("ORFAO-99") = "PRODUTO ISOLADO SEM VINCULO"
    
    Set mockFamilyMap = CreateObject("Scripting.Dictionary")
    Set mockClusters = CreateObject("Scripting.Dictionary")
    
    Set mockSKUAvg = CreateObject("Scripting.Dictionary")
    mockSKUAvg("ORFAO-99") = 82.5
    
    Set dictFamilyAvg = ComputeFamilyAverages(mockClusters, mockSKUAvg, mockMaster, mockFamilyMap)
    
    ' Critério 9: Cod_Familia = Cod_Produto e Custo_Medio_Familia = Custo_Medio_SKU
    AssertEqual CStr(mockFamilyMap("ORFAO-99")), "ORFAO-99", _
                "Critério 9: SKU órfão deve assumir Cod_Familia igual a Cod_Produto"
    AssertDoubleApprox CDbl(dictFamilyAvg("ORFAO-99")), 82.5, 0.001, _
                       "Critério 9: SKU órfão deve assumir Custo_Medio_Familia igual a Custo_Medio_SKU"
End Sub

' ------------------------------------------------------------------------------
' Teste 6: Left Join Semântico & Preservação Integral do Catálogo Mestre
' ------------------------------------------------------------------------------
Private Sub Test_06_ProductMasterLeftJoinPreservation()
    Dim mockMaster As Object
    Dim mockFamilyMap As Object
    Dim mockClusters As Object
    Dim mockSKUAvg As Object
    Dim mockFamilyAvg As Object
    Dim outMatrix As Variant
    
    Set mockMaster = CreateObject("Scripting.Dictionary")
    mockMaster("SKU-ATIVO-COM-MOV") = "PRODUTO COM HISTORICO"
    mockMaster("SKU-ATIVO-SEM-MOV") = "PRODUTO SEM NENHUM HISTORICO"
    
    Set mockFamilyMap = CreateObject("Scripting.Dictionary")
    mockFamilyMap("SKU-ATIVO-COM-MOV") = "SKU-ATIVO-COM-MOV"
    mockFamilyMap("SKU-ATIVO-SEM-MOV") = "SKU-ATIVO-SEM-MOV"
    
    Set mockSKUAvg = CreateObject("Scripting.Dictionary")
    mockSKUAvg("SKU-ATIVO-COM-MOV") = 250#
    ' SKU-ATIVO-SEM-MOV não consta em mockSKUAvg
    
    Set mockFamilyAvg = CreateObject("Scripting.Dictionary")
    mockFamilyAvg("SKU-ATIVO-COM-MOV") = 250#
    mockFamilyAvg("SKU-ATIVO-SEM-MOV") = 0#
    
    outMatrix = BuildFinalOutputMatrix(mockMaster, mockSKUAvg, mockFamilyAvg, mockFamilyMap, 90)
    
    ' Critério 6: Todos os 2 produtos do mestre devem constar na saída
    AssertEqual UBound(outMatrix, 1), 2&, _
                "Critério 6: Dimensão da matriz deve ser estritamente igual à quantidade do Cadastro Mestre"
    AssertEqual CStr(outMatrix(2, 1)), "SKU-ATIVO-SEM-MOV", _
                "Critério 6: Produto sem movimentação está preservado na matriz de saída"
    AssertEqual CDbl(outMatrix(2, 3)), 0#, _
                "Critério 7: Produto sem histórico é renderizado como 0.00 na matriz final"
End Sub

' ------------------------------------------------------------------------------
' Teste 7: Fidelidade de Schema da Matriz Final de Saída (Output Contract 01)
' ------------------------------------------------------------------------------
Private Sub Test_07_FinalOutputMatrixSchemaFidelity()
    Dim mockMaster As Object
    Dim mockFamilyMap As Object
    Dim mockSKUAvg As Object
    Dim mockFamilyAvg As Object
    Dim outMatrix As Variant
    
    Set mockMaster = CreateObject("Scripting.Dictionary")
    mockMaster("40001") = "TUBO PVC 100MM"
    
    Set mockFamilyMap = CreateObject("Scripting.Dictionary")
    mockFamilyMap("40001") = "40001"
    
    Set mockSKUAvg = CreateObject("Scripting.Dictionary")
    mockSKUAvg("40001") = 45.2
    
    Set mockFamilyAvg = CreateObject("Scripting.Dictionary")
    mockFamilyAvg("40001") = 45.2
    
    outMatrix = BuildFinalOutputMatrix(mockMaster, mockSKUAvg, mockFamilyAvg, mockFamilyMap, 90)
    
    ' Verifica se a matriz possui 7 colunas contratadas
    AssertEqual UBound(outMatrix, 2), 7&, _
                "Output Contract 01: Matriz de saída possui exatamente 7 colunas contratadas"
    AssertEqual CStr(outMatrix(1, 6)), "90D", _
                "Output Contract 01: Coluna 6 identificada com a janela '90D'"
End Sub

' ------------------------------------------------------------------------------
' Teste 8: State Restoration Trap em caso de Crash Simulado
' ------------------------------------------------------------------------------
Private Sub Test_08_ApplicationStateRestorationTrap()
    Dim initScreenUpdating As Boolean: initScreenUpdating = Application.ScreenUpdating
    Dim initCalculation As XlCalculation: initCalculation = Application.Calculation
    
    ' Congela estados
    Call FreezeAppState(setWaitCursor:=True)
    
    ' Simula bloco com erro
    On Error Resume Next
    Err.Raise vbObjectError + 9999, "SimulatedCrash", "Erro simulado"
    On Error GoTo 0
    
    ' Recuperação incondicional
    Call RestoreAppState
    
    AssertEqual Application.ScreenUpdating, True, _
                "Critério 4 / State Trap: ScreenUpdating restaurado para True"
    AssertEqual Application.Calculation, xlCalculationAutomatic, _
                "Critério 4 / State Trap: Calculation restaurado para Automático"
    AssertEqual Application.EnableEvents, True, _
                "Critério 4 / State Trap: EnableEvents restaurado para True"
End Sub

' ------------------------------------------------------------------------------
' Teste 9: Mecanismo de Trava de Concorrência (CheckFileLock)
' ------------------------------------------------------------------------------
Private Sub Test_09_CheckFileLockMechanism()
    Dim tempFile As String
    Dim fNum As Integer
    Dim isLocked As Boolean
    
    tempFile = ThisWorkbook.Path & Application.PathSeparator & "tmp_lock_test.txt"
    If Len(ThisWorkbook.Path) = 0 Then
        tempFile = "/Users/icaro/Documents/CWS/Nano Cores/tmp_lock_test.txt"
    End If
    
    ' Cria e trava arquivo exclusivamente
    fNum = FreeFile
    Open tempFile For Binary Access Read Write Lock Read Write As #fNum
    
    isLocked = CheckFileLock(tempFile)
    AssertEqual isLocked, True, _
                "CheckFileLock: Detecta arquivo bloqueado exclusivamente por outro processo"
    
    Close #fNum
    On Error Resume Next
    Kill tempFile
    On Error GoTo 0
    
    isLocked = CheckFileLock(tempFile)
    AssertEqual isLocked, False, _
                "CheckFileLock: Detecta que o arquivo liberado não está bloqueado"
End Sub
