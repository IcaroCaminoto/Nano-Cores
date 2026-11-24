Attribute VB_Name = "mod_ProdutoRepository"
' ==============================================================================
' Módulo: mod_ProdutoRepository
' Camada: Data Repository (PURA - pronta para DLL)
' Finalidade: Consultas SQL na base de produtos (Cerulean_DB_Produtos.xlsx,
'             aba Cadastro_Produtos) e montagem dos clsProdutoDTO.
' Contrato da aba Cadastro_Produtos (cabeçalho na linha 1):
'   Código | Descrição | Qtd Padrão | Código de Barras | Grupo de Produto |
'   Peso Kg | Ativo ("Sim"/"Não"; vazio = ativo)
'   Colunas extras são ignoradas; colunas opcionais ausentes viram vazio.
' Regra DLL: recebe o caminho da base por parâmetro; sem objetos do Excel,
'            sem MsgBox. Erros sobem via Err.Raise.
' ==============================================================================

Option Explicit

' Quantos códigos por consulta na cláusula IN (limite seguro do ACE)
Private Const CODIGOS_POR_CONSULTA As Long = 40

' ==============================================================================
' Função: Produtos_ListarGrupos
' Finalidade: SELECT DISTINCT dos grupos de produto (para o formulário).
' Retorno: Array de String em ordem alfabética, ou Empty se não houver grupos.
' ==============================================================================
Public Function Produtos_ListarGrupos(ByVal caminhoDb As String) As Variant
    Dim conn As Object
    Dim registros As Collection
    Dim reg As Object
    Dim vistos As Object
    Dim grupos() As String
    Dim grupo As String
    Dim sql As String
    Dim n As Long
    Dim errNum As Long, errDesc As String

    sql = "SELECT DISTINCT " & Ado_Campo(DbProd_Grupo()) & _
          " FROM " & Ado_Aba(DB_ABA_PRODUTOS) & _
          " WHERE " & Ado_Campo(DbProd_Grupo()) & " IS NOT NULL" & _
          " ORDER BY " & Ado_Campo(DbProd_Grupo())

    Set conn = Ado_AbrirConexaoExcel(caminhoDb)
    On Error GoTo Falha
    Set registros = Ado_ConsultarRegistros(conn, sql)
    Ado_FecharConexao conn
    On Error GoTo 0

    Set vistos = CreateObject("Scripting.Dictionary")
    vistos.CompareMode = vbTextCompare
    If registros.Count = 0 Then Exit Function
    ReDim grupos(0 To registros.Count - 1)

    For Each reg In registros
        grupo = Util_Texto(Util_Campo(reg, DbProd_Grupo()))
        If Len(grupo) > 0 And Not vistos.Exists(grupo) Then
            vistos.Add grupo, True
            grupos(n) = grupo
            n = n + 1
        End If
    Next reg

    If n = 0 Then Exit Function
    ReDim Preserve grupos(0 To n - 1)
    Produtos_ListarGrupos = grupos
    Exit Function

Falha:
    ' Guarda o erro antes de fechar a conexão (On Error dentro dela zera o Err)
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_ProdutoRepository", errDesc
End Function

' ==============================================================================
' Função: Produtos_BuscarPorGrupo
' Finalidade: Todos os produtos ATIVOS de um grupo, em ordem de código.
' Retorno: Collection de clsProdutoDTO (vazia se o grupo não tiver produtos).
' ==============================================================================
Public Function Produtos_BuscarPorGrupo(ByVal caminhoDb As String, ByVal grupo As String) As Collection
    Dim conn As Object
    Dim registros As Collection
    Dim reg As Object
    Dim dto As clsProdutoDTO
    Dim resultado As New Collection
    Dim vistos As Object
    Dim sql As String
    Dim errNum As Long, errDesc As String

    sql = "SELECT * FROM " & Ado_Aba(DB_ABA_PRODUTOS) & _
          " WHERE Trim(" & Ado_Campo(DbProd_Grupo()) & ") = ?" & _
          " ORDER BY " & Ado_Campo(DbProd_Codigo())

    Set conn = Ado_AbrirConexaoExcel(caminhoDb)
    On Error GoTo Falha
    Set registros = Ado_ConsultarRegistros(conn, sql, Array(Trim$(grupo)))
    Ado_FecharConexao conn
    On Error GoTo 0

    Set vistos = CreateObject("Scripting.Dictionary")
    vistos.CompareMode = vbTextCompare

    For Each reg In registros
        Set dto = ProdutoDeRegistro(reg)
        If dto.Ativo And Len(dto.Chave) > 0 And Not vistos.Exists(dto.Chave) Then
            vistos.Add dto.Chave, True
            resultado.Add dto
        End If
    Next reg

    Set Produtos_BuscarPorGrupo = resultado
    Exit Function

Falha:
    ' Guarda o erro antes de fechar a conexão (On Error dentro dela zera o Err)
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_ProdutoRepository", errDesc
End Function

' ==============================================================================
' Função: Produtos_BuscarPorCodigos
' Finalidade: Buscar os produtos de uma lista de códigos (inclusão unitária).
'             Produtos inativos também são devolvidos: se o usuário digitou o
'             código, a descrição aparece (o filtro de ativos vale para o lote).
' Parâmetros:
'   - codigos: Collection de chaves (Util_ChaveSKU), sem repetição
' Retorno: Dictionary chave -> clsProdutoDTO (só os encontrados).
' ==============================================================================
Public Function Produtos_BuscarPorCodigos(ByVal caminhoDb As String, ByVal codigos As Collection) As Object
    Dim conn As Object
    Dim registros As Collection
    Dim reg As Object
    Dim dto As clsProdutoDTO
    Dim resultado As Object
    Dim lote() As Variant
    Dim sql As String
    Dim inicio As Long, fim As Long, i As Long
    Dim errNum As Long, errDesc As String

    Set resultado = CreateObject("Scripting.Dictionary")
    resultado.CompareMode = vbTextCompare
    Set Produtos_BuscarPorCodigos = resultado
    If codigos Is Nothing Then Exit Function
    If codigos.Count = 0 Then Exit Function

    Set conn = Ado_AbrirConexaoExcel(caminhoDb)
    On Error GoTo Falha

    For inicio = 1 To codigos.Count Step CODIGOS_POR_CONSULTA
        fim = inicio + CODIGOS_POR_CONSULTA - 1
        If fim > codigos.Count Then fim = codigos.Count
        ReDim lote(0 To fim - inicio)
        For i = inicio To fim
            lote(i - inicio) = codigos(i)
        Next i

        ' [Código] & '' converte número em texto (40001 -> "40001") e trata nulos
        sql = "SELECT * FROM " & Ado_Aba(DB_ABA_PRODUTOS) & _
              " WHERE Trim(" & Ado_Campo(DbProd_Codigo()) & " & '') IN (" & Ado_Marcadores(fim - inicio + 1) & ")"

        Set registros = Ado_ConsultarRegistros(conn, sql, lote)
        For Each reg In registros
            Set dto = ProdutoDeRegistro(reg)
            If Len(dto.Chave) > 0 And Not resultado.Exists(dto.Chave) Then resultado.Add dto.Chave, dto
        Next reg
    Next inicio

    Ado_FecharConexao conn
    Exit Function

Falha:
    ' Guarda o erro antes de fechar a conexão (On Error dentro dela zera o Err)
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_ProdutoRepository", errDesc
End Function

' ==============================================================================
' Função: Produtos_ColunasAusentes
' Finalidade: Validar o contrato da aba Cadastro_Produtos (usado na migração).
' Retorno: Collection com os nomes das colunas obrigatórias que faltam.
' ==============================================================================
Public Function Produtos_ColunasAusentes(ByVal caminhoDb As String) As Collection
    Dim conn As Object
    Dim rs As Object
    Dim presentes As Object
    Dim faltam As New Collection
    Dim nome As Variant
    Dim j As Long
    Dim errNum As Long, errDesc As String

    Set conn = Ado_AbrirConexaoExcel(caminhoDb)
    On Error GoTo Falha

    ' WHERE 1=0: só o cabeçalho (nomes dos campos), sem ler dados
    Set rs = conn.Execute("SELECT * FROM " & Ado_Aba(DB_ABA_PRODUTOS) & " WHERE 1=0")
    Set presentes = CreateObject("Scripting.Dictionary")
    presentes.CompareMode = vbTextCompare
    For j = 0 To rs.Fields.Count - 1
        presentes(rs.Fields(j).Name) = True
    Next j
    rs.Close
    Ado_FecharConexao conn
    On Error GoTo 0

    For Each nome In Array(DbProd_Codigo(), DbProd_Descricao(), DbProd_QtdPadrao(), DbProd_Grupo())
        If Not presentes.Exists(CStr(nome)) Then faltam.Add CStr(nome)
    Next nome
    Set Produtos_ColunasAusentes = faltam
    Exit Function

Falha:
    ' Guarda o erro antes de fechar a conexão (On Error dentro dela zera o Err)
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_ProdutoRepository", errDesc
End Function

' ==============================================================================
' Função Privada: ProdutoDeRegistro
' Finalidade: Converter um registro ADO (campo -> valor) em clsProdutoDTO.
' ==============================================================================
Private Function ProdutoDeRegistro(ByVal reg As Object) As clsProdutoDTO
    Dim dto As New clsProdutoDTO
    Dim v As Variant

    dto.Codigo = Util_Campo(reg, DbProd_Codigo())
    If VarType(dto.Codigo) = vbString Then
        ' Código numérico gravado como texto na base: grava número na planilha
        If Len(dto.Codigo) > 0 And Util_NumeroDaBase(dto.Codigo) > 0 And _
           CStr(Util_NumeroDaBase(dto.Codigo)) = Trim$(dto.Codigo) Then
            dto.Codigo = Util_NumeroDaBase(dto.Codigo)
        End If
    End If
    dto.Chave = Util_ChaveSKU(dto.Codigo)
    dto.Descricao = Util_Texto(Util_Campo(reg, DbProd_Descricao()))
    dto.CodigoBarras = Util_Texto(Util_Campo(reg, DbProd_CodigoBarras()))
    dto.Grupo = Util_Texto(Util_Campo(reg, DbProd_Grupo()))
    dto.Ativo = Not Util_EhNao(Util_Campo(reg, DbProd_Ativo()))

    v = Util_Campo(reg, DbProd_QtdPadrao())
    If Len(Util_Texto(v)) > 0 Then dto.QtdPadrao = Util_NumeroDaBase(v)

    v = Util_Campo(reg, DbProd_PesoKg())
    If Len(Util_Texto(v)) > 0 Then dto.PesoKg = Util_NumeroDaBase(v)

    Set ProdutoDeRegistro = dto
End Function
