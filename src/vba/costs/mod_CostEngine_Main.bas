Attribute VB_Name = "mod_CostEngine_Main"
' ==============================================================================
' Módulo: mod_CostEngine_Main
' Camada: Controller & Lifecycle (Orchestration Layer)
' Finalidade: Ponto de entrada público, orquestração sequencial do pipeline de
'             fechamento de custos, gestão de estados do Excel, medição de tempo
'             e barreira de contenção global de falhas.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' Conformidade: SPEC-003 (Seções 1.3, 4 Module 1 e 5.2 Step 5)
' Proibições: Não executa queries SQL diretas, não calcula médias e não itera sobre células.
' ==============================================================================

Option Explicit
Option Private Module

Private Const APP_TITLE As String = "Fechamento de Custos"

' ==============================================================================
' Procedimento: ExecuteCostClosingPipeline
' Finalidade: Ponto de entrada público do motor de cálculo de custos. Orquestra
'             as etapas de infraestrutura, ingestão, cálculo e persistência.
' Parâmetros:
'   - windowMode (Variant, Opcional): Janela temporal desejada (30, 60, 90)
'                                     ou "ULT" / "Última Atualização" (padrão: 90).
' ==============================================================================
Public Sub ExecuteCostClosingPipeline(Optional ByVal windowMode As Variant = 90)
    Dim windowDays As Long
    Dim modeDescription As String
    Dim erpFilePath As String
    Dim dbProductsPath As String
    Dim dbAssocPath As String
    Dim dbCostsPath As String
    
    Dim idxCodigo As Long
    Dim idxData As Long
    Dim idxCusto As Long
    
    Dim dictProductMaster As Object
    Dim dictFamilyMap As Object
    Dim dictClusters As Object
    Dim arrERPRaw As Variant
    Dim dictSKUAvg As Object
    Dim dictTxCount As Object
    Dim dictFamilyAvg As Object
    Dim outMatrix As Variant
    
    Dim elapsedSec As Double
    Dim skuCount As Long
    
    ' --------------------------------------------------------------------------
    ' 1. Normalização do parâmetro de janela temporal
    ' --------------------------------------------------------------------------
    Call NormalizeWindowParam(windowMode, windowDays, modeDescription)
    
    ' --------------------------------------------------------------------------
    ' 2. Validação prévia de diretórios e bases de dados externas
    ' --------------------------------------------------------------------------
    On Error GoTo PreconditionErrorHandler
    Call ValidatePreconditions(dbProductsPath, dbAssocPath, dbCostsPath)
    On Error GoTo 0
    
    ' --------------------------------------------------------------------------
    ' 3. Interação com o usuário: Seleção do Extrato de Custos ERP
    ' --------------------------------------------------------------------------
    erpFilePath = PromptSelectERPFile()
    If Len(erpFilePath) = 0 Then
        ' Usuário cancelou o diálogo de arquivo
        Exit Sub
    End If
    
    ' --------------------------------------------------------------------------
    ' 4. Congelamento de estados da aplicação e início do cronômetro (Step 1)
    ' --------------------------------------------------------------------------
    Call FreezeAppState(setWaitCursor:=True, _
                        statusMessage:="Executando fechamento de custos (" & modeDescription & ")...")
    Call StartStopwatch
    
    On Error GoTo MainErrorHandler
    
    ' --------------------------------------------------------------------------
    ' 5. Ingestão de Schemas Primários via ADODB (Step 2)
    ' --------------------------------------------------------------------------
    Application.StatusBar = "Carregando Cadastro Mestre de Produtos..."
    Set dictProductMaster = LoadProductMaster(dbProductsPath)
    
    If dictProductMaster Is Nothing Or dictProductMaster.Count = 0 Then
        Err.Raise vbObjectError + 5001, "mod_CostEngine_Main", _
                  "Nenhum produto ativo localizado no Cadastro Mestre (Cadastro_Produtos)."
    End If
    
    Application.StatusBar = "Carregando Taxonomia e Associações de Famílias..."
    Call LoadAssociationClusters(dictFamilyMap, dictClusters, dbAssocPath)
    
    ' --------------------------------------------------------------------------
    ' 6. Ingestão e Inspeção Dinâmica de Schema do ERP (Step 2)
    ' --------------------------------------------------------------------------
    Application.StatusBar = "Lendo e validando Extrato ERP..."
    arrERPRaw = IngestERPWorkbook(erpFilePath, idxCodigo, idxData, idxCusto)
    
    If IsEmpty(arrERPRaw) Then
        ' O módulo de dados já exibiu alerta amigável se faltaram colunas
        GoTo FinallyBlock
    End If
    
    ' --------------------------------------------------------------------------
    ' 7. Motor Matemático em Memória RAM (Step 3)
    ' --------------------------------------------------------------------------
    Application.StatusBar = "Processando médias individuais por SKU..."
    If windowDays <= 0 Then
        Set dictSKUAvg = ComputeSKULatestUpdate(arrERPRaw, idxCodigo, idxData, idxCusto, dictTxCount)
    Else
        Set dictSKUAvg = ComputeSKUMovingAverages(arrERPRaw, windowDays, idxCodigo, idxData, idxCusto, dictTxCount)
    End If
    
    Application.StatusBar = "Consolidando médias dos clusters familiares..."
    Set dictFamilyAvg = ComputeFamilyAverages(dictClusters, dictSKUAvg, dictProductMaster, dictFamilyMap)
    
    Application.StatusBar = "Montando matriz final consolidada (Left Join)..."
    outMatrix = BuildFinalOutputMatrix(dictProductMaster, dictSKUAvg, dictFamilyAvg, dictFamilyMap, windowDays, dictTxCount)
    
    ' --------------------------------------------------------------------------
    ' 8. Persistência Externa Estruturada (Step 4)
    ' --------------------------------------------------------------------------
    Application.StatusBar = "Gravando nova aba em Cerulean_DB_Custos.xlsx..."
    Call PersistToDatabase(outMatrix, windowDays, dbCostsPath)
    
    ' --------------------------------------------------------------------------
    ' 9. Conclusão e Relatório Executivo
    ' --------------------------------------------------------------------------
    elapsedSec = GetElapsedTime()
    skuCount = dictProductMaster.Count
    
    Call RestoreAppState()
    
    MsgBox "Fechamento de Custos concluído com sucesso!" & vbCrLf & vbCrLf & _
           "• Modo / Janela: " & modeDescription & vbCrLf & _
           "• Total de SKUs catalogados: " & Format$(skuCount, "#,##0") & vbCrLf & _
           "• Tempo de execução: " & Format$(elapsedSec, "0.00") & " segundos" & vbCrLf & _
           "• Destino: Cerulean_DB_Custos.xlsx", _
           vbInformation, APP_TITLE
           
    Exit Sub

FinallyBlock:
    Call RestoreAppState()
    Exit Sub

PreconditionErrorHandler:
    MsgBox "Falha de pré-condição operacional:" & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, APP_TITLE
    Exit Sub

MainErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Dim errSrc As String: errSrc = Err.Source
    
    Call RestoreAppState()
    
    MsgBox "Falha durante o processamento do pipeline de custos:" & vbCrLf & vbCrLf & _
           "• Origem: " & errSrc & vbCrLf & _
           "• Erro #" & errNum & vbCrLf & _
           "• Descrição: " & errDesc, _
           vbCritical, APP_TITLE
End Sub

' ==============================================================================
' Procedimento Privado: ValidatePreconditions
' Finalidade: Verificar a existência das bases e testar trava de arquivo.
' ==============================================================================
Private Sub ValidatePreconditions(ByRef outProductsPath As String, _
                                 ByRef outAssocPath As String, _
                                 ByRef outCostsPath As String)
    outProductsPath = ResolvePath("Cerulean_DB_Produtos.xlsx", DB_PRODUCTS_FILE)
    outAssocPath = ResolvePath("Cerulean_DB_Parametros_Associacao_Produtos.xlsx", DB_PARAM_ASSOC_FILE)
    outCostsPath = ResolvePath("Cerulean_DB_Custos.xlsx", DB_COSTS_FILE)
    
    If Len(Dir(outProductsPath)) = 0 Then
        Err.Raise vbObjectError + 5002, "mod_CostEngine_Main.ValidatePreconditions", _
                  "A base de produtos não foi localizada em:" & vbCrLf & outProductsPath
    End If
    
    If Len(Dir(outAssocPath)) = 0 Then
        Err.Raise vbObjectError + 5003, "mod_CostEngine_Main.ValidatePreconditions", _
                  "A base de parâmetros taxonômicos não foi localizada em:" & vbCrLf & outAssocPath
    End If
    
    If Len(Dir(outCostsPath)) = 0 Then
        Err.Raise vbObjectError + 5004, "mod_CostEngine_Main.ValidatePreconditions", _
                  "A base de destino Cerulean_DB_Custos.xlsx não foi localizada em:" & vbCrLf & outCostsPath
    End If
    
    If CheckFileLock(outCostsPath) Then
        Err.Raise vbObjectError + 5005, "mod_CostEngine_Main.ValidatePreconditions", _
                  "O arquivo 'Cerulean_DB_Custos.xlsx' encontra-se aberto por outro usuário na rede." & vbCrLf & _
                  "Feche o arquivo antes de iniciar o fechamento."
    End If
End Sub

' ==============================================================================
' Função Privada: PromptSelectERPFile
' Finalidade: Exibir caixa de diálogo corporativa para seleção do arquivo do ERP.
' ==============================================================================
Private Function PromptSelectERPFile() As String
    Dim fd As Office.FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o Extrato de Custos do ERP"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Pastas de Trabalho do Excel (*.xlsx; *.xlsm; *.xlsb)", "*.xlsx;*.xlsm;*.xlsb"
        .Filters.Add "Todos os Arquivos (*.*)", "*.*"
        
        If .Show = -1 Then
            PromptSelectERPFile = .SelectedItems(1)
        Else
            PromptSelectERPFile = vbNullString
        End If
    End With
    Set fd = Nothing
End Function

' ==============================================================================
' Procedimento Privado: NormalizeWindowParam
' Finalidade: Traduzir o parâmetro de entrada do usuário/UI para a convenção interna.
' ==============================================================================
Private Sub NormalizeWindowParam(ByVal windowMode As Variant, _
                                ByRef outDays As Long, _
                                ByRef outDesc As String)
    Dim modeStr As String
    modeStr = UCase$(Trim$(CStr(windowMode)))
    
    If modeStr = "0" Or InStr(modeStr, "ULT") > 0 Then
        outDays = 0
        outDesc = "Snapshot Última Atualização"
    ElseIf IsNumeric(modeStr) Then
        outDays = CLng(modeStr)
        Select Case outDays
            Case 30, 60, 90
                outDesc = "Média Móvel de " & outDays & " Dias"
            Case Else
                outDays = 90
                outDesc = "Média Móvel de 90 Dias (Padrão)"
        End Select
    Else
        outDays = 90
        outDesc = "Média Móvel de 90 Dias (Padrão)"
    End If
End Sub

' ==============================================================================
' Função Privada: ResolvePath
' ==============================================================================
Private Function ResolvePath(ByVal fileName As String, ByVal defaultPath As String) As String
    Dim candidate As String
    On Error Resume Next
    If Not ThisWorkbook Is Nothing Then
        If Len(ThisWorkbook.Path) > 0 Then
            candidate = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator & "sheets" & Application.PathSeparator & "Cerulean_DB" & Application.PathSeparator & "database" & Application.PathSeparator & fileName
            If Len(Dir(candidate)) > 0 Then
                ResolvePath = candidate
                Exit Function
            End If
            
            candidate = ThisWorkbook.Path & Application.PathSeparator & "sheets" & Application.PathSeparator & "Cerulean_DB" & Application.PathSeparator & "database" & Application.PathSeparator & fileName
            If Len(Dir(candidate)) > 0 Then
                ResolvePath = candidate
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0
    ResolvePath = defaultPath
End Function
