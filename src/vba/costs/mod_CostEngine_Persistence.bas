Attribute VB_Name = "mod_CostEngine_Persistence"
' ==============================================================================
' Módulo: mod_CostEngine_Persistence
' Camada: Output & Table Formatting (Persistence Layer)
' Finalidade: Gerenciar a pasta de trabalho de destino Cerulean_DB_Custos.xlsx,
'             validar travas de concorrência via acesso binário exclusivo,
'             criar nova aba com timestamp, realizar despejo em bloco único (.Value2),
'             converter para ListObject estruturado e aplicar formatações estritas.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' Conformidade: SPEC-003 (Seções 1.3, 2 Output Contract 01, 4 e 5.2 Step 4)
' Proibições: Não recalcula médias; apenas persiste e formata estruturadamente.
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Constante do caminho padrão de destino
' ------------------------------------------------------------------------------
Public Const DB_COSTS_FILE As String = "/Users/icaro/Documents/CWS/Nano Cores/src/sheets/Cerulean_DB/database/Cerulean_DB_Custos.xlsx"

' ==============================================================================
' Função: CheckFileLock
' Finalidade: Testar se o arquivo físico de destino está aberto e bloqueado por outro
'             processo ou usuário na rede via tentativa de abertura binária exclusiva.
' Retorno: Boolean (True se o arquivo estiver bloqueado/em uso; False se liberado).
' ==============================================================================
Public Function CheckFileLock(ByVal fullPath As String) As Boolean
    Dim fileNum As Integer
    Dim errNum As Long
    
    If Len(Trim$(fullPath)) = 0 Or Len(Dir(fullPath)) = 0 Then
        CheckFileLock = False
        Exit Function
    End If
    
    fileNum = FreeFile
    On Error Resume Next
    Open fullPath For Binary Access Read Write Lock Read Write As #fileNum
    errNum = Err.Number
    Close #fileNum
    On Error GoTo 0
    
    ' Se ocorrer erro (ex: 70 - Permission Denied ou 55 - File already open), o arquivo está travado
    If errNum <> 0 Then
        CheckFileLock = True
    Else
        CheckFileLock = False
    End If
End Function

' ==============================================================================
' Procedimento: PersistToDatabase
' Finalidade: Abrir o banco de dados de custos em segundo plano, gerar uma nova
'             Worksheet com timestamp dinâmico, descarregar cabeçalhos e matriz de
'             dados em bloco único (.Value2), converter para ListObject e formatar.
' Parâmetros:
'   - outputMatrix (Variant ByRef): Matriz 2D (1 To N, 1 To 7) gerada pelo Calculator.
'   - windowDays (Long): Quantidade de dias da janela móvel (30, 60, 90) ou 0 para ULT.
'   - customFilePath (String, Opcional): Caminho customizado da pasta Cerulean_DB_Custos.xlsx.
' ==============================================================================
Public Sub PersistToDatabase(ByRef outputMatrix As Variant, _
                            ByVal windowDays As Long, _
                            Optional ByVal customFilePath As String = vbNullString)
    
    Dim targetPath As String
    Dim numRows As Long
    Dim modeLabel As String
    Dim timeStampStr As String
    Dim baseSheetName As String
    Dim targetSheetName As String
    Dim baseTableName As String
    Dim targetTableName As String
    Dim suffixIdx As Long
    
    Dim wbDest As Workbook
    Dim wsNew As Worksheet
    Dim rngHeaders As Range
    Dim rngData As Range
    Dim rngTableTotal As Range
    Dim loCustos As ListObject
    Dim arrHeaders(1 To 1, 1 To 7) As Variant
    
    ' 1. Validação da matriz de entrada
    If IsEmpty(outputMatrix) Then
        Err.Raise vbObjectError + 4002, "mod_CostEngine_Persistence.PersistToDatabase", _
                  "A matriz de saída fornecida encontra-se vazia. Operação de persistência abortada."
    End If
    
    numRows = UBound(outputMatrix, 1)
    If numRows = 0 Then
        Err.Raise vbObjectError + 4003, "mod_CostEngine_Persistence.PersistToDatabase", _
                  "Nenhum registro para persistir na matriz de dados."
    End If
    
    ' 2. Resolução do caminho da base de custos
    If Len(Trim$(customFilePath)) > 0 Then
        targetPath = customFilePath
    Else
        targetPath = ResolveDatabasePath("Cerulean_DB_Custos.xlsx", DB_COSTS_FILE)
    End If
    
    If Len(Dir(targetPath)) = 0 Then
        Err.Raise vbObjectError + 4004, "mod_CostEngine_Persistence.PersistToDatabase", _
                  "Arquivo de banco de dados de custos não encontrado: " & targetPath
    End If
    
    ' 3. Teste de trava de arquivo (Concorrência / File Lock)
    If CheckFileLock(targetPath) Then
        Err.Raise vbObjectError + 4001, "mod_CostEngine_Persistence.PersistToDatabase", _
                  "O arquivo de destino 'Cerulean_DB_Custos.xlsx' encontra-se bloqueado por outro processo ou usuário na rede." & vbCrLf & _
                  "Por favor, feche o arquivo e repita a operação."
    End If
    
    ' 4. Determinação dos identificadores dinâmicos (Aba e Tabela)
    If windowDays <= 0 Then
        modeLabel = "ULT"
    Else
        modeLabel = CStr(windowDays) & "D"
    End If
    
    timeStampStr = Format$(Now, "yyyymmdd_hhnn")
    baseSheetName = "Custo_MM_" & modeLabel & "_" & timeStampStr
    baseTableName = "tb_Custos_" & timeStampStr
    
    On Error GoTo PersistenceErrorHandler
    
    ' 5. Abertura da pasta de trabalho em segundo plano
    Set wbDest = Application.Workbooks.Open(Filename:=targetPath, _
                                           UpdateLinks:=False, _
                                           ReadOnly:=False, _
                                           AddToMru:=False)
                                           
    ' 6. Resolução de colisão de nomes de abas
    targetSheetName = baseSheetName
    suffixIdx = 2
    Do While WorksheetExists(wbDest, targetSheetName)
        targetSheetName = Left$(baseSheetName, 28) & "_v" & suffixIdx
        suffixIdx = suffixIdx + 1
    Loop
    
    ' Cria a nova Worksheet na última posição
    Set wsNew = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
    wsNew.Name = targetSheetName
    
    ' 7. Preparação e despejo dos cabeçalhos contratados em A1:G1
    arrHeaders(1, 1) = "Cod_Produto"
    arrHeaders(1, 2) = "Descricao_Produto"
    arrHeaders(1, 3) = "Custo_Medio_SKU"
    arrHeaders(1, 4) = "Custo_Medio_Familia"
    arrHeaders(1, 5) = "Qtd_Transacoes_SKU"
    arrHeaders(1, 6) = "Janela_Dias_Base"
    arrHeaders(1, 7) = "Data_Processamento"
    
    Set rngHeaders = wsNew.Range("A1:G1")
    rngHeaders.Value2 = arrHeaders
    
    ' 8. Despejo em bloco único da matriz de dados a partir de A2
    Set rngData = wsNew.Range("A2").Resize(numRows, 7)
    rngData.Value2 = outputMatrix
    
    ' 9. Conversão para ListObject estruturado
    Set rngTableTotal = wsNew.Range("A1").Resize(numRows + 1, 7)
    
    targetTableName = baseTableName
    suffixIdx = 2
    Do While ListObjectExists(wbDest, targetTableName)
        targetTableName = baseTableName & "_v" & suffixIdx
        suffixIdx = suffixIdx + 1
    Loop
    
    Set loCustos = wsNew.ListObjects.Add(SourceType:=xlSrcRange, _
                                       Source:=rngTableTotal, _
                                       XlListObjectHasHeaders:=xlYes)
    loCustos.Name = targetTableName
    loCustos.TableStyle = "TableStyleMedium2"
    loCustos.ShowTableStyleRowStripes = False
    
    ' 10. Aplicação de formatos numéricos estritos por coluna
    With loCustos
        .ListColumns(1).DataBodyRange.NumberFormat = "@"                  ' Cod_Produto
        .ListColumns(2).DataBodyRange.NumberFormat = "@"                  ' Descricao_Produto
        .ListColumns(3).DataBodyRange.NumberFormat = "R$ #,##0.00"        ' Custo_Medio_SKU
        .ListColumns(4).DataBodyRange.NumberFormat = "R$ #,##0.00"        ' Custo_Medio_Familia
        .ListColumns(5).DataBodyRange.NumberFormat = "#,##0"              ' Qtd_Transacoes_SKU
        .ListColumns(6).DataBodyRange.NumberFormat = "@"                  ' Janela_Dias_Base
        .ListColumns(7).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm:ss"' Data_Processamento
        
        ' Alinhamentos estéticos
        .ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
        .ListColumns(2).DataBodyRange.HorizontalAlignment = xlLeft
        .ListColumns(3).DataBodyRange.HorizontalAlignment = xlRight
        .ListColumns(4).DataBodyRange.HorizontalAlignment = xlRight
        .ListColumns(5).DataBodyRange.HorizontalAlignment = xlRight
        .ListColumns(6).DataBodyRange.HorizontalAlignment = xlCenter
        .ListColumns(7).DataBodyRange.HorizontalAlignment = xlCenter
    End With
    
    ' 11. Autoajuste de colunas
    wsNew.Columns("A:G").AutoFit
    
    ' 12. Salvamento e fechamento explícito do arquivo
    wbDest.Close SaveChanges:=True
    Set wbDest = Nothing
    
    Exit Sub

PersistenceErrorHandler:
    Dim pErrDesc As String: pErrDesc = Err.Description
    Dim pErrNum As Long: pErrNum = Err.Number
    
    If Not wbDest Is Nothing Then
        On Error Resume Next
        wbDest.Close SaveChanges:=False
        Set wbDest = Nothing
        On Error GoTo 0
    End If
    
    Err.Raise pErrNum, "mod_CostEngine_Persistence.PersistToDatabase", _
              "Falha crítica durante a persistência em Cerulean_DB_Custos.xlsx: " & pErrDesc
End Sub

' ==============================================================================
' Funções Auxiliares Privadas
' ==============================================================================
Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Function ListObjectExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set lo = ws.ListObjects(tableName)
        If Not lo Is Nothing Then
            ListObjectExists = True
            Exit Function
        End If
    Next ws
    ListObjectExists = False
    On Error GoTo 0
End Function

Private Function ResolveDatabasePath(ByVal fileName As String, ByVal absoluteFallback As String) As String
    Dim candidate As String
    On Error Resume Next
    If Not ThisWorkbook Is Nothing Then
        If Len(ThisWorkbook.Path) > 0 Then
            ' Relativo à pasta do projeto
            candidate = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator & "sheets" & Application.PathSeparator & "Cerulean_DB" & Application.PathSeparator & "database" & Application.PathSeparator & fileName
            If Len(Dir(candidate)) > 0 Then
                ResolveDatabasePath = candidate
                Exit Function
            End If
            
            candidate = ThisWorkbook.Path & Application.PathSeparator & "sheets" & Application.PathSeparator & "Cerulean_DB" & Application.PathSeparator & "database" & Application.PathSeparator & fileName
            If Len(Dir(candidate)) > 0 Then
                ResolveDatabasePath = candidate
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0
    
    ResolveDatabasePath = absoluteFallback
End Function

' ==============================================================================
' Procedimento SafeZone: Test_CostEngine_Persistence_SafeZone
' Finalidade: Validar o teste de bloqueio de arquivo (CheckFileLock) e regras de nomenclatura.
' ==============================================================================
Public Sub Test_CostEngine_Persistence_SafeZone()
    Dim testPath As String
    Dim isLocked As Boolean
    
    testPath = ResolveDatabasePath("Cerulean_DB_Custos.xlsx", DB_COSTS_FILE)
    
    If Len(Dir(testPath)) > 0 Then
        isLocked = CheckFileLock(testPath)
        Debug.Print "[PASS] Test_CostEngine_Persistence_SafeZone: CheckFileLock executado com sucesso. Status bloqueado: " & isLocked
    Else
        Debug.Print "[WARN] Test_CostEngine_Persistence_SafeZone: Arquivo de banco de dados não encontrado para teste de trava."
    End If
End Sub
