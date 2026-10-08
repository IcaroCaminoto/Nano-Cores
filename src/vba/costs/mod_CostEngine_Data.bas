Attribute VB_Name = "mod_CostEngine_Data"
' ==============================================================================
' Módulo: mod_CostEngine_Data
' Camada: Ingestion & Hash Builders (Data Access Layer)
' Finalidade: Executar o transporte de dados externos (ADO e Leitor de Arquivo ERP)
'             e alimentar as estruturas de memória primárias (Scripting.Dictionary).
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' Conformidade: SPEC-003 (Seção 2, Input Contracts 01, 02 e 03)
' Proibições: Não realiza cálculos contábeis; entrega dados puros e validados por schema.
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Constantes de caminhos padrão
' ------------------------------------------------------------------------------
Public Const DB_PRODUCTS_FILE As String = "G:\Ícaro\Cerulean_DB\database\Cerulean_DB_Produtos.xlsx"
Public Const DB_PARAM_ASSOC_FILE As String = "G:\Ícaro\Cerulean_DB\database\Cerulean_DB_Parametros_Associacao_Produtos.xlsx"

' ==============================================================================
' Função Auxiliar Privada: BuildExcelConnectionString
' Finalidade: Montar a cadeia de conexão OLEDB do provedor Microsoft.ACE.OLEDB.12.0
'             com concatenação estrita e direta do caminho do arquivo.
' ==============================================================================
Private Function BuildExcelConnectionString(ByVal filePath As String) As String
    BuildExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                 "Data Source=" & filePath & ";" & _
                                 "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;"";"
End Function

' ==============================================================================
' Função: LoadProductMaster
' Finalidade: Ingerir o Cadastro Mestre de Produtos via consulta ADODB.
' Contrato SQL:
'   SELECT [Código], [Descrição] FROM [Cadastro_Produtos$]
'   WHERE ((Val([Código]) BETWEEN 40000 AND 45999) OR (Val([Código]) BETWEEN 50000 AND 50999))
'   AND [Ativo] = 'Sim'
' Retorno: Scripting.Dictionary contendo [Código] -> [Descrição]
' ==============================================================================
Public Function LoadProductMaster(Optional ByVal customFilePath As String = vbNullString) As Object
    Dim dbPath As String
    Dim connStr As String
    Dim sqlQuery As String
    Dim conn As Object
    Dim rs As Object
    Dim dictMaster As Object
    Dim rawCode As Variant
    Dim rawDesc As Variant
    Dim cleanCode As String
    Dim cleanDesc As String
    
    If Len(Trim$(customFilePath)) > 0 Then
        dbPath = customFilePath
    Else
        dbPath = ResolveDatabasePath("Cerulean_DB_Produtos.xlsx", DB_PRODUCTS_FILE)
    End If
    
    ' Validação de existência do arquivo físico
    If Len(Dir(dbPath)) = 0 Then
        Err.Raise vbObjectError + 2001, "mod_CostEngine_Data.LoadProductMaster", _
                  "Base de dados de produtos não encontrada: " & dbPath
    End If
    
    Set dictMaster = CreateObject("Scripting.Dictionary")
    dictMaster.CompareMode = vbTextCompare
    
    connStr = BuildExcelConnectionString(dbPath)
    sqlQuery = "SELECT [Código], [Descrição] " & _
               "FROM [Cadastro_Produtos$] " & _
               "WHERE ((Val([Código]) BETWEEN 40000 AND 45999) " & _
               "    OR (Val([Código]) BETWEEN 50000 AND 50999)) " & _
               "AND [Ativo] = 'Sim'"
               
    On Error GoTo DataErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connStr
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn, 0, 1 ' adOpenForwardOnly, adLockReadOnly
    
    Do Until rs.EOF
        rawCode = rs.Fields(0).Value
        rawDesc = rs.Fields(1).Value
        
        If Not IsNull(rawCode) Then
            cleanCode = Trim$(CStr(rawCode))
            
            If IsNull(rawDesc) Or Len(Trim$(CStr(rawDesc))) = 0 Then
                cleanDesc = "Sem Descrição"
            Else
                cleanDesc = Trim$(CStr(rawDesc))
            End If
            
            If Len(cleanCode) > 0 Then
                If dictMaster.Exists(cleanCode) Then
                    Debug.Print "[WARN] mod_CostEngine_Data: Código duplicado em Cadastro_Produtos: " & cleanCode
                End If
                dictMaster(cleanCode) = cleanDesc
            End If
        End If
        rs.MoveNext
    Loop
    
CleanupBlock:
    ' Liberação de memória no caminho de sucesso
    Call CloseAdoObjects(rs, conn)
    
    Set LoadProductMaster = dictMaster
    Exit Function

DataErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Call CloseAdoObjects(rs, conn)
    Err.Raise errNum, "mod_CostEngine_Data.LoadProductMaster", _
              "Falha na ingestão ADO do Cadastro Mestre: " & errDesc
End Function

' ==============================================================================
' Função: LoadAssociationClusters
' Finalidade: Ingerir a tabela de associações de produtos via ADODB e estruturar
'             em memória dois dicionários essenciais:
'             1. outFamilyMap (Scripting.Dictionary): Mapeamento direto Filho -> Pai e Pai -> Pai.
'             2. outClusters (Scripting.Dictionary): Clusters Pai -> Sub-Dicionário de Membros.
' Contrato SQL:
'   SELECT [Cod_Pai], [Cod_Filho] FROM [Assoc_MediasMoveis$] WHERE [Cod_Pai] IS NOT NULL
' Retorno: Boolean (True em caso de sucesso).
' ==============================================================================
Public Function LoadAssociationClusters(ByRef outFamilyMap As Object, _
                                       ByRef outClusters As Object, _
                                       Optional ByVal customFilePath As String = vbNullString) As Boolean
    Dim dbPath As String
    Dim connStr As String
    Dim sqlQuery As String
    Dim conn As Object
    Dim rs As Object
    Dim rawParent As Variant
    Dim rawChild As Variant
    Dim cleanParent As String
    Dim cleanChild As String
    Dim clusterMembers As Object
    
    If Len(Trim$(customFilePath)) > 0 Then
        dbPath = customFilePath
    Else
        dbPath = ResolveDatabasePath("Cerulean_DB_Parametros_Associacao_Produtos.xlsx", DB_PARAM_ASSOC_FILE)
    End If
    
    ' Validação de existência do arquivo físico
    If Len(Dir(dbPath)) = 0 Then
        Err.Raise vbObjectError + 2002, "mod_CostEngine_Data.LoadAssociationClusters", _
                  "Base de dados de associações não encontrada: " & dbPath
    End If
    
    Set outFamilyMap = CreateObject("Scripting.Dictionary")
    outFamilyMap.CompareMode = vbTextCompare
    
    Set outClusters = CreateObject("Scripting.Dictionary")
    outClusters.CompareMode = vbTextCompare
    
    connStr = BuildExcelConnectionString(dbPath)
    sqlQuery = "SELECT [Cod_Pai], [Cod_Filho] " & _
               "FROM [Assoc_MediasMoveis$] " & _
               "WHERE [Cod_Pai] IS NOT NULL " & _
               "AND [Cod_Filho] IS NOT NULL"
               
    On Error GoTo DataErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connStr
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn, 0, 1 ' adOpenForwardOnly, adLockReadOnly
    
    Do Until rs.EOF
        rawParent = rs.Fields(0).Value
        rawChild = rs.Fields(1).Value
        
        If Not IsNull(rawParent) And Not IsNull(rawChild) Then
            cleanParent = Trim$(CStr(rawParent))
            cleanChild = Trim$(CStr(rawChild))
            
            If Len(cleanParent) > 0 And Len(cleanChild) > 0 Then
                ' 1. Garantia reflexiva no cluster (Pai como membro número 1)
                If Not outClusters.Exists(cleanParent) Then
                    Set clusterMembers = CreateObject("Scripting.Dictionary")
                    clusterMembers.CompareMode = vbTextCompare
                    clusterMembers(cleanParent) = True ' Membro #1 é o próprio Pai
                    Set outClusters(cleanParent) = clusterMembers
                Else
                    Set clusterMembers = outClusters(cleanParent)
                End If
                
                ' 2. Adiciona o Filho ao cluster
                clusterMembers(cleanChild) = True
                
                ' 3. Registra no mapa direto de SKU -> Pai
                outFamilyMap(cleanParent) = cleanParent ' Reflexivo
                
                If outFamilyMap.Exists(cleanChild) Then
                    Debug.Print "[WARN] mod_CostEngine_Data: Cod_Filho duplicado: " & cleanChild & _
                                " (sobrescrevendo vínculo com novo Pai: " & cleanParent & ")"
                End If
                outFamilyMap(cleanChild) = cleanParent
            End If
        End If
        rs.MoveNext
    Loop
    
CleanupBlock:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    On Error GoTo 0
    
    LoadAssociationClusters = True
    Exit Function
    
DataErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Call CleanupBlock
    Err.Raise errNum, "mod_CostEngine_Data.LoadAssociationClusters", _
              "Falha na ingestão ADO de Associações: " & errDesc
End Function

' ==============================================================================
' Função: IngestERPWorkbook
' Finalidade: Abre a planilha do ERP em background, inspeciona horizontalmente
'             os cabeçalhos na linha 3 para resolver dinamicamente os índices de:
'             - 'Código'
'             - 'Data Atualização'
'             - 'Custo Gerencial'
'             Se faltar alguma coluna obrigatória, emite aviso via MsgBox, fecha
'             a pasta e retorna Empty.
'             Se válida, extrai a partir da linha 4 para uma matriz Variant em memória.
' Assinatura: Contratada na Seção 4 (Architectural Blueprint)
' ==============================================================================
Public Function IngestERPWorkbook(ByVal filePath As String, _
                                  ByRef outIdxCodigo As Long, _
                                  ByRef outIdxData As Long, _
                                  ByRef outIdxCusto As Long) As Variant
    Dim wbERP As Workbook
    Dim wsERP As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim headerRange As Variant
    Dim c As Long
    Dim colHeader As String
    Dim missingCols As String
    Dim dataRange As Range
    Dim rawMatrix As Variant
    
    outIdxCodigo = 0
    outIdxData = 0
    outIdxCusto = 0
    IngestERPWorkbook = Empty
    
    ' Validação de existência do arquivo ERP
    If Len(Trim$(filePath)) = 0 Or Len(Dir(filePath)) = 0 Then
        MsgBox "O arquivo do extrato ERP não foi encontrado no caminho especificado:" & vbCrLf & _
               filePath, vbCritical, "Ingestão Extrato ERP"
        Exit Function
    End If
    
    On Error GoTo OpenErrorHandler
    
    ' Abertura segura em background
    Set wbERP = Application.Workbooks.Open(Filename:=filePath, _
                                          UpdateLinks:=False, _
                                          ReadOnly:=True, _
                                          AddToMru:=False)
    Set wsERP = wbERP.Worksheets(1)
    
    ' 1. Varredura horizontal da Linha 3 para descoberta dinâmica de schema
    lastCol = wsERP.Cells(3, wsERP.Columns.Count).End(xlToLeft).Column
    
    If lastCol < 1 Then
        wbERP.Close SaveChanges:=False
        Set wbERP = Nothing
        MsgBox "Relatório ERP incompatível. A linha 3 de cabeçalhos encontra-se vazia.", _
               vbExclamation, "Ingestão Extrato ERP"
        Exit Function
    End If
    
    headerRange = wsERP.Range(wsERP.Cells(3, 1), wsERP.Cells(3, lastCol)).Value2
    
    For c = 1 To lastCol
        colHeader = NormalizeHeader(CStr(headerRange(1, c)))
        
        If colHeader = "CODIGO" Then
            outIdxCodigo = c
        ElseIf colHeader = "DATA ATUALIZACAO" Or colHeader = "DATA DA ATUALIZACAO" Then
            outIdxData = c
        ElseIf colHeader = "CUSTO GERENCIAL" Then
            outIdxCusto = c
        End If
    Next c
    
    ' 2. Verificação de obrigatoriedade do contrato de schema
    missingCols = vbNullString
    If outIdxCodigo = 0 Then missingCols = missingCols & vbCrLf & "• Código"
    If outIdxData = 0 Then missingCols = missingCols & vbCrLf & "• Data Atualização"
    If outIdxCusto = 0 Then missingCols = missingCols & vbCrLf & "• Custo Gerencial"
    
    If Len(missingCols) > 0 Then
        wbERP.Close SaveChanges:=False
        Set wbERP = Nothing
        MsgBox "Relatório ERP incompatível. As seguintes colunas obrigatórias não foram encontradas na linha 3:" & _
               missingCols, vbExclamation, "Ingestão Extrato ERP"
        Exit Function
    End If
    
    ' 3. Determinação da última linha preenchida com base na coluna chave
    lastRow = wsERP.Cells(wsERP.Rows.Count, outIdxCodigo).End(xlUp).Row
    
    If lastRow < 4 Then
        ' Não há registros de dados (apenas cabeçalhos ou vazio)
        wbERP.Close SaveChanges:=False
        Set wbERP = Nothing
        IngestERPWorkbook = Empty
        Exit Function
    End If
    
    ' 4. Extração em bloco único da linha 4 até lastRow para matriz Variant
    Set dataRange = wsERP.Range(wsERP.Cells(4, 1), wsERP.Cells(lastRow, lastCol))
    rawMatrix = dataRange.Value2
    
    ' Fechamento imediato do arquivo externo sem salvar
    wbERP.Close SaveChanges:=False
    Set wbERP = Nothing
    
    IngestERPWorkbook = rawMatrix
    Exit Function

OpenErrorHandler:
    Dim openErr As String: openErr = Err.Description
    If Not wbERP Is Nothing Then
        On Error Resume Next
        wbERP.Close SaveChanges:=False
        Set wbERP = Nothing
        On Error GoTo 0
    End If
    Err.Raise Err.Number, "mod_CostEngine_Data.IngestERPWorkbook", _
              "Erro ao abrir e processar o arquivo ERP: " & openErr
End Function

' ==============================================================================
' Função Auxiliar Privada: NormalizeHeader
' Finalidade: Padronizar nomes de cabeçalho para comparação case-insensitive
'             e sem caracteres diacríticos/acentos.
' ==============================================================================
Private Function NormalizeHeader(ByVal rawText As String) As String
    Dim cleaned As String
    cleaned = UCase$(Trim$(rawText))
    
    ' Substituição de acentuações comuns
    cleaned = Replace(cleaned, "Ó", "O")
    cleaned = Replace(cleaned, "Ò", "O")
    cleaned = Replace(cleaned, "Õ", "O")
    cleaned = Replace(cleaned, "Ô", "O")
    cleaned = Replace(cleaned, "Á", "A")
    cleaned = Replace(cleaned, "À", "A")
    cleaned = Replace(cleaned, "Ã", "A")
    cleaned = Replace(cleaned, "Â", "A")
    cleaned = Replace(cleaned, "É", "E")
    cleaned = Replace(cleaned, "Ê", "E")
    cleaned = Replace(cleaned, "Í", "I")
    cleaned = Replace(cleaned, "Ú", "U")
    cleaned = Replace(cleaned, "Ç", "C")
    
    NormalizeHeader = cleaned
End Function

' ==============================================================================
' Função Auxiliar Privada: ResolveDatabasePath
' Finalidade: Resolver dinamicamente a localização da base de dados, verificando
'             primeiro caminhos relativos ao ThisWorkbook e aplicando o fallback
'             absoluto contratado na SPEC caso necessário.
' ==============================================================================
Private Function ResolveDatabasePath(ByVal fileName As String, ByVal absoluteFallback As String) As String
    Dim candidate As String
    On Error Resume Next
    If Not ThisWorkbook Is Nothing Then
        If Len(ThisWorkbook.Path) > 0 Then
            ' 1. Tenta relativo ao caminho do projeto (Nano Cores/src/sheets/Cerulean_DB/database/...)
            candidate = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator & "sheets" & Application.PathSeparator & "Cerulean_DB" & Application.PathSeparator & "database" & Application.PathSeparator & fileName
            If Len(Dir(candidate)) > 0 Then
                ResolveDatabasePath = candidate
                Exit Function
            End If
            
            ' 2. Tenta a partir da pasta raiz sheets/
            candidate = ThisWorkbook.Path & Application.PathSeparator & "sheets" & Application.PathSeparator & "Cerulean_DB" & Application.PathSeparator & "database" & Application.PathSeparator & fileName
            If Len(Dir(candidate)) > 0 Then
                ResolveDatabasePath = candidate
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0
    
    ' 3. Aplica o caminho absoluto definido no contrato da SPEC
    ResolveDatabasePath = absoluteFallback
End Function

' ==============================================================================
' Função Auxiliar Privada: CleanupBlock
' Finalidade: Fechar e liberar com segurança instâncias de Recordset e 
'             Connection.
' ==============================================================================
Private Sub CloseAdoObjects(ByRef rs As Object, ByRef conn As Object)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    On Error GoTo 0
End Sub