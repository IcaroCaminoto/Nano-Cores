Attribute VB_Name = "mod_AdoExcel"
' ==============================================================================
' Módulo: mod_AdoExcel
' Camada: Data Access (PURA - pronta para DLL)
' Finalidade: Abrir conexões ADO/OLEDB com arquivos .xlsx ("bases burras") e
'             executar consultas SQL parametrizadas, devolvendo os registros
'             como Collection de Scripting.Dictionary (campo -> valor).
' Por que dicionário: o código lê os campos pelo NOME; uma coluna nova ou
'             ausente na base não quebra a leitura (tolerância de esquema).
' Regra DLL: nenhum objeto do Excel e nenhuma MsgBox. Erros sobem via
'             Err.Raise para a camada de interface decidir como avisar.
' Dependências: mod_DataUtils (Util_ArquivoExiste). Late binding (sem
'             referência obrigatória ao ADODB).
' ==============================================================================

Option Explicit

Public Const ERR_ADO_ARQUIVO_AUSENTE As Long = vbObjectError + 7101

Private Const AD_CMD_TEXT As Long = 1
Private Const AD_VAR_WCHAR As Long = 202
Private Const AD_PARAM_INPUT As Long = 1
Private Const AD_STATE_OPEN As Long = 1

' ==============================================================================
' Função: Ado_AbrirConexaoExcel
' Finalidade: Abrir conexão somente leitura com um .xlsx (cabeçalho na linha 1).
' ==============================================================================
Public Function Ado_AbrirConexaoExcel(ByVal caminhoArquivo As String) As Object
    Dim conn As Object

    If Not Util_ArquivoExiste(caminhoArquivo) Then
        Err.Raise ERR_ADO_ARQUIVO_AUSENTE, "mod_AdoExcel", _
                  "Base n" & ChrW(227) & "o encontrada: " & caminhoArquivo
    End If

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & caminhoArquivo & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Set Ado_AbrirConexaoExcel = conn
End Function

Public Sub Ado_FecharConexao(ByRef conn As Object)
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = AD_STATE_OPEN Then conn.Close
    End If
    Set conn = Nothing
    On Error GoTo 0
End Sub

' ==============================================================================
' Função: Ado_ConsultarRegistros
' Finalidade: Executar um SELECT com parâmetros "?" (opcionais, sempre texto).
'             Parâmetros evitam erro com apóstrofos nos valores e injeção SQL.
' Parâmetros:
'   - conn: conexão aberta por Ado_AbrirConexaoExcel
'   - sql: comando com marcadores "?" na ordem dos parâmetros
'   - parametros: Array de valores (ou omitido)
' Retorno: Collection de Dictionary (CompareMode texto). Vazia se não houver linhas.
' ==============================================================================
Public Function Ado_ConsultarRegistros(ByVal conn As Object, ByVal sql As String, _
                                       Optional ByVal parametros As Variant) As Collection
    Dim cmd As Object
    Dim rs As Object
    Dim resultado As Collection
    Dim arr As Variant
    Dim nomes() As String
    Dim reg As Object
    Dim p As Variant
    Dim valorTexto As String
    Dim nCampos As Long, tamanho As Long, k As Long
    Dim i As Long, j As Long
    Dim errNum As Long, errDesc As String

    On Error GoTo Falha

    Set resultado = New Collection
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = AD_CMD_TEXT

    If Not IsMissing(parametros) Then
        If IsArray(parametros) Then
            For Each p In parametros
                k = k + 1
                valorTexto = Util_Texto(p)
                tamanho = Len(valorTexto)
                If tamanho < 1 Then tamanho = 1
                cmd.Parameters.Append cmd.CreateParameter("p" & k, AD_VAR_WCHAR, AD_PARAM_INPUT, tamanho, valorTexto)
            Next p
        End If
    End If

    Set rs = cmd.Execute
    nCampos = rs.Fields.Count
    If nCampos > 0 Then
        ReDim nomes(0 To nCampos - 1)
        For j = 0 To nCampos - 1
            nomes(j) = rs.Fields(j).Name
        Next j
        If Not rs.EOF Then arr = rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    If IsArray(arr) Then
        For i = 0 To UBound(arr, 2)
            Set reg = CreateObject("Scripting.Dictionary")
            reg.CompareMode = vbTextCompare
            For j = 0 To nCampos - 1
                reg(nomes(j)) = arr(j, i)
            Next j
            resultado.Add reg
        Next i
    End If

    Set Ado_ConsultarRegistros = resultado
    Exit Function

Falha:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = AD_STATE_OPEN Then rs.Close
    End If
    On Error GoTo 0
    Err.Raise errNum, "mod_AdoExcel", errDesc & vbCrLf & "SQL: " & sql
End Function

' Nome de campo entre colchetes: [Grupo de Produto]
Public Function Ado_Campo(ByVal nomeCampo As String) As String
    Ado_Campo = "[" & nomeCampo & "]"
End Function

' Aba do Excel como tabela SQL: [Cadastro_Produtos$]
Public Function Ado_Aba(ByVal nomeAba As String) As String
    Ado_Aba = "[" & nomeAba & "$]"
End Function

' Lista de marcadores "?, ?, ?" para cláusulas IN
Public Function Ado_Marcadores(ByVal quantidade As Long) As String
    Dim i As Long
    Dim s As String
    For i = 1 To quantidade
        If i > 1 Then s = s & ", "
        s = s & "?"
    Next i
    Ado_Marcadores = s
End Function
