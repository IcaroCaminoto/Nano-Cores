Attribute VB_Name = "mod_Pricing_CostPackageService"
' ==============================================================================
' Módulo: mod_Pricing_CostPackageService
' Camada: Application Service / Core (PURA - pronta para DLL)      [spec-009]
' Finalidade: Regras do "pacote de custos" (uma aba de fechamento dentro de
'             Cerulean_DB_Custos.xlsx), sem planilha e sem formulário:
'   - ListarPacotesCustosDisponiveis : abas válidas, lidas com OpenSchema (OLEDB)
'   - BuscarCustosPorPacote          : custo por SKU na aba escolhida (ADO/SQL)
'   - Pacote_*                       : validação do nome, coluna de custo por
'                                      prefixo, montagem do SQL, saneamento e
'                                      texto do cabeçalho de impressão
' O estado "pacote ativo" (Named Range) e o cabeçalho de impressão dependem do
' Excel e ficam em mod_PricingPacoteCustos (adapter). Na DLL, este módulo vira
' as Actions "custos.listarPacotes" e "custos.buscarPorPacote".
' Regra DLL: nenhum Range/Worksheet/Application/ThisWorkbook/MsgBox; o caminho
'            da base chega por parâmetro; erros sobem via Err.Raise.
' Dependências: mod_AdoExcel, mod_DataUtils, mod_PricingSchema.
' Conformidade: encoding-rules (strings funcionais 100% ASCII, acentos via ChrW).
' ==============================================================================

Option Explicit

Public Const ERR_PACOTE_LAYOUT_INVALIDO As Long = vbObjectError + 902
Public Const ERR_PACOTE_INEXISTENTE As Long = vbObjectError + 903

' Quantos SKUs por consulta na cláusula IN (limite seguro do ACE)
Private Const CODIGOS_POR_CONSULTA As Long = 40

' ==============================================================================
' Função: Pacote_EhValido
' Finalidade: True se o nome é de um pacote oficial do motor de custos:
'             Custo_MM_* (média móvel) ou Custo_EX_* (engenharia).
' ==============================================================================
Public Function Pacote_EhValido(ByVal nomePacote As String) As Boolean
    Dim nome As String
    Dim prefixo As String

    nome = Trim$(nomePacote)
    If Len(nome) <= Len(DB_ABA_CUSTOS_PREFIXO_MM) Then Exit Function

    prefixo = Left$(nome, Len(DB_ABA_CUSTOS_PREFIXO_MM))
    Pacote_EhValido = (StrComp(prefixo, DB_ABA_CUSTOS_PREFIXO_MM, vbTextCompare) = 0 Or _
                       StrComp(prefixo, DB_ABA_CUSTOS_PREFIXO_EX, vbTextCompare) = 0)
End Function

' ==============================================================================
' Função: Pacote_ResolverColunaCusto
' Finalidade: Coluna física de custo conforme o prefixo do pacote.
'   Custo_MM_* -> Custo_Medio_Familia      Custo_EX_* -> Custo_Final_Engenharia
' Erro: ERR_PACOTE_LAYOUT_INVALIDO (vbObjectError + 902) para outro prefixo.
' ==============================================================================
Public Function Pacote_ResolverColunaCusto(ByVal nomePacote As String) As String
    Dim prefixo As String

    prefixo = Left$(Trim$(nomePacote), Len(DB_ABA_CUSTOS_PREFIXO_MM))
    If StrComp(prefixo, DB_ABA_CUSTOS_PREFIXO_MM, vbTextCompare) = 0 Then
        Pacote_ResolverColunaCusto = DBCUS_CUSTO_MM
    ElseIf StrComp(prefixo, DB_ABA_CUSTOS_PREFIXO_EX, vbTextCompare) = 0 Then
        Pacote_ResolverColunaCusto = DBCUS_CUSTO_EX
    Else
        Err.Raise ERR_PACOTE_LAYOUT_INVALIDO, "mod_Pricing_CostPackageService", _
                  "Layout de pacote de custos n" & ChrW(227) & "o reconhecido: " & nomePacote
    End If
End Function

' ==============================================================================
' Função: Pacote_FiltrarAbas
' Finalidade: Ficar só com os pacotes válidos de uma lista de nomes de aba,
'             do mais recente para o mais antigo (os nomes terminam em
'             AAAAMMDD_HHMM, então a ordem decrescente de texto é a cronológica
'             dentro de cada tipo).
' Parâmetros:
'   - abas: Collection de String (nomes já limpos, ver Ado_ListarAbas)
' ==============================================================================
Public Function Pacote_FiltrarAbas(ByVal abas As Collection) As Collection
    Dim resultado As New Collection
    Dim validos() As String
    Dim nome As Variant
    Dim temp As String
    Dim n As Long, i As Long, j As Long

    Set Pacote_FiltrarAbas = resultado
    If abas Is Nothing Then Exit Function
    If abas.Count = 0 Then Exit Function

    ReDim validos(1 To abas.Count)
    For Each nome In abas
        If Pacote_EhValido(CStr(nome)) Then
            n = n + 1
            validos(n) = Trim$(CStr(nome))
        End If
    Next nome
    If n = 0 Then Exit Function

    ' Ordenação por inserção, decrescente (a lista é pequena)
    For i = 2 To n
        temp = validos(i)
        j = i - 1
        Do While j >= 1
            If StrComp(validos(j), temp, vbTextCompare) >= 0 Then Exit Do
            validos(j + 1) = validos(j)
            j = j - 1
        Loop
        validos(j + 1) = temp
    Next i

    For i = 1 To n
        resultado.Add validos(i)
    Next i
End Function

' ==============================================================================
' Função: Pacote_ExisteNaLista
' Finalidade: True se o pacote consta na lista de abas (sem diferenciar caixa).
' ==============================================================================
Public Function Pacote_ExisteNaLista(ByVal abas As Collection, ByVal nomePacote As String) As Boolean
    Dim nome As Variant

    If abas Is Nothing Then Exit Function
    For Each nome In abas
        If StrComp(Trim$(CStr(nome)), Trim$(nomePacote), vbTextCompare) = 0 Then
            Pacote_ExisteNaLista = True
            Exit Function
        End If
    Next nome
End Function

' ==============================================================================
' Função: ListarPacotesCustosDisponiveis
' Finalidade: Pacotes de custos válidos de Cerulean_DB_Custos.xlsx, lidos do
'             catálogo OLEDB (OpenSchema/adSchemaTables), sem abrir o Excel.
' Parâmetros:
'   - dbPath: caminho completo de Cerulean_DB_Custos.xlsx (sempre informado:
'             o núcleo não conhece pastas; quem chama resolve via mod_Config)
' Retorno: Collection de String, mais recentes primeiro (vazia se não houver).
' ==============================================================================
Public Function ListarPacotesCustosDisponiveis(ByVal dbPath As String) As Collection
    Dim conn As Object
    Dim abas As Collection
    Dim errNum As Long, errDesc As String

    Set conn = Ado_AbrirConexaoExcel(dbPath)
    On Error GoTo Falha
    Set abas = Ado_ListarAbas(conn)
    Ado_FecharConexao conn
    On Error GoTo 0

    Set ListarPacotesCustosDisponiveis = Pacote_FiltrarAbas(abas)
    Exit Function

Falha:
    ' Guarda o erro antes de fechar a conexão (On Error dentro dela zera o Err)
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_Pricing_CostPackageService", errDesc
End Function

' ==============================================================================
' Função: Pacote_MontarSql
' Finalidade: SELECT da spec-009 para um pacote, com "?" no lugar dos SKUs.
'   SELECT [Cod_Produto], [<coluna de custo>] AS [Custo_Resolvido]
'   FROM [<pacote>$] WHERE Trim([Cod_Produto] & '') IN (?, ?, ...)
' O "Trim(... & '')" compara o código como texto mesmo que a célula da base
' guarde número (40001) ou texto ("40001").
' ==============================================================================
Public Function Pacote_MontarSql(ByVal nomePacote As String, ByVal qtdCodigos As Long) As String
    Dim colunaCusto As String

    colunaCusto = Pacote_ResolverColunaCusto(nomePacote)
    Pacote_MontarSql = "SELECT " & Ado_Campo(DBCUS_COD_PRODUTO) & ", " & _
                       Ado_Campo(colunaCusto) & " AS " & Ado_Campo(DBCUS_CUSTO_RESOLVIDO) & _
                       " FROM " & Ado_Aba(Trim$(nomePacote)) & _
                       " WHERE Trim(" & Ado_Campo(DBCUS_COD_PRODUTO) & " & '') IN (" & _
                       Ado_Marcadores(qtdCodigos) & ")"
End Function

' ==============================================================================
' Função: Pacote_SanearCusto
' Finalidade: Custo válido para o motor. Nulo, não numérico ou negativo vira
'             0,00 e devolve o motivo em outAlerta (vazio quando o custo é bom).
' ==============================================================================
Public Function Pacote_SanearCusto(ByVal valor As Variant, ByRef outAlerta As String) As Double
    outAlerta = vbNullString

    If Not Util_EhNumeroDaBase(valor) Then
        outAlerta = "custo ausente ou n" & ChrW(227) & "o num" & ChrW(233) & "rico (assumido 0,00)"
        Exit Function
    End If

    Pacote_SanearCusto = Util_NumeroDaBase(valor)
    If Pacote_SanearCusto < 0# Then
        Pacote_SanearCusto = 0#
        outAlerta = "custo negativo (assumido 0,00)"
    End If
End Function

' ==============================================================================
' Função: BuscarCustosPorPacote
' Finalidade: Custo unitário de cada SKU na aba [<pacote>$] de
'             Cerulean_DB_Custos.xlsx (consulta única por lote de 40 SKUs).
' Parâmetros:
'   - dbPath: caminho completo de Cerulean_DB_Custos.xlsx
'   - nomePacote: aba escolhida (Custo_MM_* ou Custo_EX_*)
'   - skus: Collection de chaves (Util_ChaveSKU), sem repetição
'   - outAlertas (saída): Collection de String, um item por custo inválido
' Retorno: Dictionary chave -> Double, só com os SKUs encontrados no pacote.
'          SKU que não consta na aba simplesmente não aparece no resultado.
' Erros:
'   - ERR_PACOTE_LAYOUT_INVALIDO (902): prefixo desconhecido (antes de qualquer I/O)
'   - ERR_PACOTE_INEXISTENTE (903): a aba não existe mais na base
'   - ERR_ADO_ARQUIVO_AUSENTE (7101): base não encontrada (rede/pasta)
' ==============================================================================
Public Function BuscarCustosPorPacote(ByVal dbPath As String, ByVal nomePacote As String, _
                                      ByVal skus As Collection, _
                                      Optional ByRef outAlertas As Collection) As Object
    Dim conn As Object
    Dim registros As Collection
    Dim reg As Object
    Dim resultado As Object
    Dim lote() As Variant
    Dim sql As String
    Dim sku As String
    Dim alerta As String
    Dim custo As Double
    Dim inicio As Long, fim As Long, i As Long
    Dim errNum As Long, errDesc As String

    Set outAlertas = New Collection
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado.CompareMode = vbTextCompare
    Set BuscarCustosPorPacote = resultado

    nomePacote = Trim$(nomePacote)
    Call Pacote_ResolverColunaCusto(nomePacote)          ' falha cedo (902), sem tocar na base
    If skus Is Nothing Then Exit Function
    If skus.Count = 0 Then Exit Function

    Set conn = Ado_AbrirConexaoExcel(dbPath)
    On Error GoTo Falha

    If Not Pacote_ExisteNaLista(Ado_ListarAbas(conn), nomePacote) Then
        Err.Raise ERR_PACOTE_INEXISTENTE, "mod_Pricing_CostPackageService", _
                  "O pacote de custos """ & nomePacote & """ n" & ChrW(227) & "o existe mais em: " & dbPath
    End If

    For inicio = 1 To skus.Count Step CODIGOS_POR_CONSULTA
        fim = inicio + CODIGOS_POR_CONSULTA - 1
        If fim > skus.Count Then fim = skus.Count
        ReDim lote(0 To fim - inicio)
        For i = inicio To fim
            lote(i - inicio) = skus(i)
        Next i

        sql = Pacote_MontarSql(nomePacote, fim - inicio + 1)
        Set registros = Ado_ConsultarRegistros(conn, sql, lote)

        For Each reg In registros
            sku = Util_ChaveSKU(Util_Campo(reg, DBCUS_COD_PRODUTO))
            If Len(sku) > 0 And Not resultado.Exists(sku) Then
                custo = Pacote_SanearCusto(Util_Campo(reg, DBCUS_CUSTO_RESOLVIDO), alerta)
                If Len(alerta) > 0 Then outAlertas.Add sku & ": " & alerta
                resultado.Add sku, custo
            End If
        Next reg
    Next inicio

    Ado_FecharConexao conn
    Exit Function

Falha:
    ' Guarda o erro antes de fechar a conexão (On Error dentro dela zera o Err)
    errNum = Err.Number: errDesc = Err.Description
    Ado_FecharConexao conn
    Err.Raise errNum, "mod_Pricing_CostPackageService", errDesc
End Function

' ==============================================================================
' Função: Pacote_MontarCabecalho
' Finalidade: Texto do cabeçalho de impressão com a linha "Custos: [pacote]",
'             sem destruir o que já estiver no quadrante:
'   - vazio                        -> "Custos: [pacote]"
'   - já tem uma linha "Custos: "  -> só essa linha é trocada
'   - tem outro conteúdo           -> "Custos: [pacote]" entra numa nova linha
'   "&" no nome vira "&&" (no PageSetup, "&" inicia um código de formato).
' Parâmetros:
'   - textoAtual: conteúdo atual do quadrante (linhas separadas por vbLf)
' ==============================================================================
Public Function Pacote_MontarCabecalho(ByVal textoAtual As String, ByVal nomePacote As String) As String
    Dim linha As String
    Dim partes() As String
    Dim i As Long
    Dim achou As Boolean

    linha = CABECALHO_CUSTOS_PREFIXO & Replace(Trim$(nomePacote), "&", "&&")
    If Len(textoAtual) = 0 Then
        Pacote_MontarCabecalho = linha
        Exit Function
    End If

    partes = Split(textoAtual, vbLf)
    For i = LBound(partes) To UBound(partes)
        If StrComp(Left$(partes(i), Len(CABECALHO_CUSTOS_PREFIXO)), CABECALHO_CUSTOS_PREFIXO, vbBinaryCompare) = 0 Then
            partes(i) = linha
            achou = True
            Exit For
        End If
    Next i

    If achou Then
        Pacote_MontarCabecalho = Join(partes, vbLf)
    Else
        Pacote_MontarCabecalho = textoAtual & vbLf & linha
    End If
End Function

' ==============================================================================
' Função: Pacote_CodigosSemCusto
' Finalidade: SKUs pedidos que não constam no pacote (para avisar o usuário).
' ==============================================================================
Public Function Pacote_CodigosSemCusto(ByVal skus As Collection, ByVal custos As Object) As Collection
    Dim sku As Variant
    Dim lista As New Collection

    Set Pacote_CodigosSemCusto = lista
    If skus Is Nothing Then Exit Function
    For Each sku In skus
        If custos Is Nothing Then
            lista.Add CStr(sku)
        ElseIf Not custos.Exists(CStr(sku)) Then
            lista.Add CStr(sku)
        End If
    Next sku
End Function
