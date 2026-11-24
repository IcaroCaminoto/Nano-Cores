Attribute VB_Name = "mod_ImpostoRepository"
' ==============================================================================
' Módulo: mod_ImpostoRepository
' Camada: Data Repository (PURA - pronta para DLL)
' Finalidade: Ler Cerulean_DB_Impostos.xlsx via ADO/SQL e cruzar as duas abas:
'   - Impostos_Federais (impostos por produto): SKU, PIS, COFINS, IPI e as
'     exceções calc_icms / calc_pis_cofins / calc_ipi / calc_st ("Não" desliga)
'   - Matriz_ICMS_ST (impostos por UF): UF, SKU, ICMS, ST
' Unidades na base: PIS/COFINS em pontos percentuais (1,65 / 7,6); ICMS, IPI e
'   ST em fração. Tudo é normalizado para fração.
' Regra DLL: recebe caminho e UF por parâmetro; sem objetos do Excel, sem MsgBox.
' ==============================================================================

Option Explicit

Public Const ERR_IMPOSTOS_UF_INVALIDA As Long = vbObjectError + 7201

' ==============================================================================
' Função: Impostos_CarregarPorUF
' Finalidade: Duas consultas na mesma conexão (produto e UF) e cruzamento em
'             memória pelo SKU. Só entram SKUs presentes nas DUAS abas (mesma
'             regra do INNER JOIN anterior).
' Parâmetros:
'   - caminhoDb: caminho completo de Cerulean_DB_Impostos.xlsx
'   - siglaUF: "SP", "MG"... (validada aqui)
'   - outTemExcecoes (saída): True se a base tem as colunas calc_*
' Retorno: Dictionary chave SKU -> clsAliquotasDTO (pode vir vazio).
' ==============================================================================
Public Function Impostos_CarregarPorUF(ByVal caminhoDb As String, ByVal siglaUF As String, _
                                       Optional ByRef outTemExcecoes As Boolean) As Object
    Dim conn As Object
    Dim regsProduto As Collection
    Dim regsUF As Collection
    Dim porProduto As Object
    Dim resultado As Object
    Dim reg As Object
    Dim regProd As Object
    Dim dto As clsAliquotasDTO
    Dim chave As String
    Dim errNum As Long, errDesc As String

    siglaUF = UCase$(Trim$(siglaUF))
    If Not (Len(siglaUF) = 2 And siglaUF Like "[A-Z][A-Z]") Then
        Err.Raise ERR_IMPOSTOS_UF_INVALIDA, "mod_ImpostoRepository", "UF inv" & ChrW(225) & "lida: " & siglaUF
    End If

    Set conn = Ado_AbrirConexaoExcel(caminhoDb)
    On Error GoTo Falha
    Set regsProduto = Ado_ConsultarRegistros(conn, "SELECT * FROM " & Ado_Aba(DB_ABA_IMPOSTOS_PRODUTO))
    Set regsUF = Ado_ConsultarRegistros(conn, "SELECT * FROM " & Ado_Aba(DB_ABA_IMPOSTOS_UF) & _
                                              " WHERE Trim(" & Ado_Campo(DBIMP_UF) & ") = ?", Array(siglaUF))
    Ado_FecharConexao conn
    On Error GoTo 0

    ' ---- Índice por produto ----------------------------------------------------
    Set porProduto = CreateObject("Scripting.Dictionary")
    porProduto.CompareMode = vbTextCompare
    outTemExcecoes = False
    For Each reg In regsProduto
        If Not outTemExcecoes Then outTemExcecoes = reg.Exists(DBIMP_CALC_ICMS)
        chave = Util_ChaveSKU(Util_Campo(reg, DBIMP_SKU))
        If Len(chave) > 0 And Not porProduto.Exists(chave) Then porProduto.Add chave, reg
    Next reg

    ' ---- Cruzamento produto x UF ----------------------------------------------
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado.CompareMode = vbTextCompare
    For Each reg In regsUF
        chave = Util_ChaveSKU(Util_Campo(reg, DBIMP_SKU))
        If Len(chave) > 0 And Not resultado.Exists(chave) Then
            If porProduto.Exists(chave) Then
                Set regProd = porProduto(chave)
                Set dto = New clsAliquotasDTO
                dto.Chave = chave
                dto.ICMS = Util_NormalizarAliquota(Util_Campo(reg, DBIMP_ICMS))
                dto.STFora = Util_NormalizarAliquota(Util_Campo(reg, DBIMP_ST))
                dto.IPIFora = Util_NormalizarAliquota(Util_Campo(regProd, DBIMP_IPI))
                dto.PisCofins = PisCofinsDoRegistro(regProd)
                dto.CalcICMS = Not Util_EhNao(Util_Campo(regProd, DBIMP_CALC_ICMS))
                dto.CalcPisCofins = Not Util_EhNao(Util_Campo(regProd, DBIMP_CALC_PIS_COFINS))
                dto.CalcIPI = Not Util_EhNao(Util_Campo(regProd, DBIMP_CALC_IPI))
                dto.CalcST = Not Util_EhNao(Util_Campo(regProd, DBIMP_CALC_ST))
                resultado.Add chave, dto
            End If
        End If
    Next reg

    Set Impostos_CarregarPorUF = resultado
    Exit Function

Falha:
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_ImpostoRepository", errDesc
End Function

' ==============================================================================
' Função: Impostos_TemColunasExcecao
' Finalidade: Saber se a aba Impostos_Federais já tem as colunas calc_*
'             (pré-requisito para excluir a tabela local CfExcecoesImpostos).
' ==============================================================================
Public Function Impostos_TemColunasExcecao(ByVal caminhoDb As String) As Boolean
    Dim conn As Object
    Dim rs As Object
    Dim campos As Object
    Dim nome As Variant
    Dim j As Long
    Dim errNum As Long, errDesc As String

    Set conn = Ado_AbrirConexaoExcel(caminhoDb)
    On Error GoTo Falha
    Set rs = conn.Execute("SELECT * FROM " & Ado_Aba(DB_ABA_IMPOSTOS_PRODUTO) & " WHERE 1=0")
    Set campos = CreateObject("Scripting.Dictionary")
    campos.CompareMode = vbTextCompare
    For j = 0 To rs.Fields.Count - 1
        campos(rs.Fields(j).Name) = True
    Next j
    rs.Close
    Ado_FecharConexao conn
    On Error GoTo 0

    Impostos_TemColunasExcecao = True
    For Each nome In Array(DBIMP_CALC_ICMS, DBIMP_CALC_PIS_COFINS, DBIMP_CALC_IPI, DBIMP_CALC_ST)
        If Not campos.Exists(CStr(nome)) Then Impostos_TemColunasExcecao = False
    Next nome
    Exit Function

Falha:
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_ImpostoRepository", errDesc
End Function

' PIS + COFINS; sem nenhum dos dois na base, usa o padrão legado (9,25%)
Private Function PisCofinsDoRegistro(ByVal regProd As Object) As Double
    Dim pis As Variant, cofins As Variant
    pis = Util_Campo(regProd, DBIMP_PIS)
    cofins = Util_Campo(regProd, DBIMP_COFINS)
    If Len(Util_Texto(pis)) = 0 And Len(Util_Texto(cofins)) = 0 Then
        PisCofinsDoRegistro = ALIQ_PIS_COFINS_PADRAO
    Else
        PisCofinsDoRegistro = Util_NormalizarAliquota(pis) + Util_NormalizarAliquota(cofins)
    End If
End Function
