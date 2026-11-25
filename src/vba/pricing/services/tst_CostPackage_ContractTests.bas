Attribute VB_Name = "tst_CostPackage_ContractTests"
' ==============================================================================
' Módulo de Testes: tst_CostPackage_ContractTests                    [spec-009]
' Finalidade: Auditoria dos contratos da spec-009 (pacote de custos), conforme a
'             skill code-spec-validator.
'   - RunCostPackageContractTests      : testes em memória (sem base externa);
'                                        o único dado de planilha tocado é uma
'                                        aba temporária "tmp_test_Cabecalho" e o
'                                        Named Range cfg_PacoteCustosAtivo, que
'                                        voltam ao estado original no fim.
'   - RunCostPackageIntegrationTests   : consulta de verdade (ADO/OLEDB) a base
'                                        SINTÉTICA tests\fixtures\
'                                        mock_Cerulean_DB_Custos.xlsx; nunca
'                                        usa a base de produção.
' Execução: Editor VBA > Janela Imediata (Ctrl+G) >
'             RunCostPackageContractTests
'             RunCostPackageIntegrationTests
'             RunCostPackageIntegrationTests "C:\caminho\mock_Cerulean_DB_Custos.xlsx"
' Observação: os auxiliares de asserção são Private (não colidem com os de
'             tst_CostEngine_ContractTests).
' ==============================================================================

Option Explicit

Private m_Total As Long
Private m_Passou As Long
Private m_Falhou As Long

' ------------------------------------------------------------------------------
' Auxiliares de asserção
' ------------------------------------------------------------------------------
Private Sub Zerar()
    m_Total = 0: m_Passou = 0: m_Falhou = 0
End Sub

Private Sub Verificar(ByVal condicao As Boolean, ByVal nomeTeste As String)
    m_Total = m_Total + 1
    If condicao Then
        m_Passou = m_Passou + 1
        Debug.Print " [PASS] " & nomeTeste
    Else
        m_Falhou = m_Falhou + 1
        Debug.Print "![FAIL] " & nomeTeste
    End If
End Sub

Private Sub VerificarIgual(ByVal obtido As Variant, ByVal esperado As Variant, ByVal nomeTeste As String)
    m_Total = m_Total + 1
    If IsEmpty(obtido) And IsEmpty(esperado) Then
        m_Passou = m_Passou + 1
        Debug.Print " [PASS] " & nomeTeste
    ElseIf IsEmpty(obtido) Or IsEmpty(esperado) Then
        m_Falhou = m_Falhou + 1
        Debug.Print "![FAIL] " & nomeTeste & " | Esperado: [" & esperado & "] obtido: [" & obtido & "]"
    ElseIf CStr(obtido) = CStr(esperado) Then
        m_Passou = m_Passou + 1
        Debug.Print " [PASS] " & nomeTeste
    Else
        m_Falhou = m_Falhou + 1
        Debug.Print "![FAIL] " & nomeTeste & " | Esperado: [" & esperado & "] obtido: [" & obtido & "]"
    End If
End Sub

Private Sub VerificarAprox(ByVal obtido As Double, ByVal esperado As Double, ByVal nomeTeste As String)
    m_Total = m_Total + 1
    If Abs(obtido - esperado) <= 0.0001 Then
        m_Passou = m_Passou + 1
        Debug.Print " [PASS] " & nomeTeste
    Else
        m_Falhou = m_Falhou + 1
        Debug.Print "![FAIL] " & nomeTeste & " | Esperado ~[" & esperado & "] obtido: [" & obtido & "]"
    End If
End Sub

Private Sub Resumo(ByVal titulo As String)
    Debug.Print "=========================================================="
    Debug.Print titulo & ": " & m_Total & " asser" & ChrW(231) & ChrW(245) & "es | Aprovadas: " & m_Passou & " | Falhas: " & m_Falhou
    Debug.Print "=========================================================="
    If m_Falhou > 0 Then
        Err.Raise vbObjectError + 999, "tst_CostPackage_ContractTests", _
                  "Uma ou mais assercoes contratuais da SPEC-009 falharam."
    End If
End Sub

' Erro gerado por uma chamada que deve falhar. Cada teste chama o
' procedimento sob teste e captura Err.Number ANTES de qualquer On Error.
Private Function NumeroDoErro(ByVal quem As String, ByVal arg1 As String, Optional ByVal arg2 As String) As Long
    On Error Resume Next
    Select Case quem
        Case "coluna":   Call Pacote_ResolverColunaCusto(arg1)
        Case "sql":      Call Pacote_MontarSql(arg1, 1)
        Case "buscar":   Call BuscarCustosPorPacote(arg2, arg1, ColecaoDe("1"))
        Case "salvar":   Call SalvarPacoteCustosAtivo(arg1)
    End Select
    NumeroDoErro = Err.Number
    On Error GoTo 0
End Function

' Custo do SKU, ou -1 se ele não estiver no resultado (ler uma chave ausente de um
' Dictionary criaria uma chave vazia e esconderia a falha).
Private Function CustoDe(ByVal custos As Object, ByVal sku As String) As Double
    If custos.Exists(sku) Then CustoDe = CDbl(custos(sku)) Else CustoDe = -1#
End Function

Private Function ColecaoDe(ParamArray itens() As Variant) As Collection
    Dim c As New Collection
    Dim i As Long
    For i = LBound(itens) To UBound(itens)
        c.Add CStr(itens(i))
    Next i
    Set ColecaoDe = c
End Function

' ==============================================================================
' Ponto de entrada 1: contratos em memória
' ==============================================================================
Public Sub RunCostPackageContractTests()
    Zerar
    Debug.Print "--- AUDITORIA DE CONTRATOS (SPEC-009) ---"

    Test_01_NomeDaAbaOledb
    Test_02_FiltroDeDominio
    Test_03_ColunaPorPrefixo
    Test_04_SqlDaSpec
    Test_05_SaneamentoDeCusto
    Test_06_NumeroDaBase
    Test_07_GuardasSemIO
    Test_08_DtoEMatrizDoLote
    Test_09_EstadoPersistido
    Test_10_TextoDoCabecalho
    Test_11_CabecalhoNoPageSetup

    Resumo "EXECU" & ChrW(199) & ChrW(195) & "O CONCLU" & ChrW(205) & "DA (SPEC-009 / contratos)"
End Sub

' ------------------------------------------------------------------------------
' Teste 1: TABLE_NAME do catálogo OLEDB -> nome da aba (aspas, "$", não-abas)
' ------------------------------------------------------------------------------
Private Sub Test_01_NomeDaAbaOledb()
    VerificarIgual Ado_NomeDaAba("Custo_MM_90D_20261101_1200$"), "Custo_MM_90D_20261101_1200", "01a aba simples perde o $"
    VerificarIgual Ado_NomeDaAba("'Resumo Geral$'"), "Resumo Geral", "01b aspas simples removidas"
    VerificarIgual Ado_NomeDaAba("'Joe''s$'"), "Joe's", "01c aspa duplicada vira simples"
    VerificarIgual Ado_NomeDaAba("Custo_MM_90D$Print_Area"), "", "01d intervalo nomeado (sem $ final) descartado"
    VerificarIgual Ado_NomeDaAba("Tabela1"), "", "01e tabela/ListObject descartada"
    VerificarIgual Ado_NomeDaAba("_xlnm#_FilterDatabase"), "", "01f nome de sistema descartado"
    VerificarIgual Ado_NomeDaAba(""), "", "01g vazio"
    VerificarIgual Ado_NomeDaAba("$"), "", "01h so o cifrao"
End Sub

' ------------------------------------------------------------------------------
' Teste 2: filtro de domínio (só Custo_MM_* e Custo_EX_*) e ordem (recentes primeiro)
' ------------------------------------------------------------------------------
Private Sub Test_02_FiltroDeDominio()
    Dim abas As Collection
    Dim r As Collection

    Verificar Pacote_EhValido("Custo_MM_90D_20261101_1200"), "02a Custo_MM_* valido"
    Verificar Pacote_EhValido("Custo_EX_20261110_1530"), "02b Custo_EX_* valido"
    Verificar Pacote_EhValido("custo_mm_ult_20261105_0900"), "02c prefixo sem diferenciar caixa"
    Verificar Not Pacote_EhValido("Custo_XX_20261101_1200"), "02d prefixo desconhecido rejeitado"
    Verificar Not Pacote_EhValido("Notas_Custo_MM_90D"), "02e prefixo no meio rejeitado"
    Verificar Not Pacote_EhValido("Custo_MM_"), "02f so o prefixo rejeitado"
    Verificar Not Pacote_EhValido("Planilha1"), "02g aba comum rejeitada"
    Verificar Not Pacote_EhValido(""), "02h vazio rejeitado"

    Set abas = ColecaoDe("Planilha1", "Custo_EX_20261001_1530", "Custo_MM_90D_20261101_1200", _
                         "Resumo", "Custo_MM_90D_20261201_0800", "Custo_XX_20261101_1200")
    Set r = Pacote_FiltrarAbas(abas)
    VerificarIgual r.Count, 3, "02i so os 3 pacotes validos permanecem"
    VerificarIgual r(1), "Custo_MM_90D_20261201_0800", "02j ordem: MM mais recente primeiro"
    VerificarIgual r(2), "Custo_MM_90D_20261101_1200", "02k ordem: MM mais antigo depois"
    VerificarIgual r(3), "Custo_EX_20261001_1530", "02l EX depois dos MM"

    VerificarIgual Pacote_FiltrarAbas(New Collection).Count, 0, "02m lista vazia -> vazia"
    VerificarIgual Pacote_FiltrarAbas(Nothing).Count, 0, "02n Nothing -> vazia"
    Verificar Pacote_ExisteNaLista(abas, "CUSTO_EX_20261001_1530"), "02o existencia sem diferenciar caixa"
    Verificar Not Pacote_ExisteNaLista(abas, "Custo_EX_19990101_0000"), "02p aba ausente"
End Sub

' ------------------------------------------------------------------------------
' Teste 3: coluna física de custo por prefixo (e erro 902 para layout desconhecido)
' ------------------------------------------------------------------------------
Private Sub Test_03_ColunaPorPrefixo()
    VerificarIgual Pacote_ResolverColunaCusto("Custo_MM_90D_20261101_1200"), "Custo_Medio_Familia", "03a MM -> Custo_Medio_Familia"
    VerificarIgual Pacote_ResolverColunaCusto("Custo_MM_ULT_20261105_0900"), "Custo_Medio_Familia", "03b MM_ULT -> Custo_Medio_Familia"
    VerificarIgual Pacote_ResolverColunaCusto("Custo_EX_20261110_1530"), "Custo_Final_Engenharia", "03c EX -> Custo_Final_Engenharia"
    VerificarIgual NumeroDoErro("coluna", "Custo_XX_20261101_1200"), vbObjectError + 902, "03d layout desconhecido -> vbObjectError+902"
    VerificarIgual NumeroDoErro("coluna", ""), vbObjectError + 902, "03e nome vazio -> vbObjectError+902"
End Sub

' ------------------------------------------------------------------------------
' Teste 4: SQL literal da spec (alias Custo_Resolvido, FROM [aba$], IN com ?)
' ------------------------------------------------------------------------------
Private Sub Test_04_SqlDaSpec()
    VerificarIgual Pacote_MontarSql("Custo_MM_90D_20261101_1200", 3), _
        "SELECT [Cod_Produto], [Custo_Medio_Familia] AS [Custo_Resolvido] " & _
        "FROM [Custo_MM_90D_20261101_1200$] WHERE Trim([Cod_Produto] & '') IN (?, ?, ?)", _
        "04a SQL do pacote MM"
    VerificarIgual Pacote_MontarSql("Custo_EX_20261110_1530", 1), _
        "SELECT [Cod_Produto], [Custo_Final_Engenharia] AS [Custo_Resolvido] " & _
        "FROM [Custo_EX_20261110_1530$] WHERE Trim([Cod_Produto] & '') IN (?)", _
        "04b SQL do pacote EX"
    VerificarIgual NumeroDoErro("sql", "Planilha1"), vbObjectError + 902, "04c SQL de aba invalida -> 902"
End Sub

' ------------------------------------------------------------------------------
' Teste 5: nulo, não numérico ou < 0 => 0,00 + alerta
' ------------------------------------------------------------------------------
Private Sub Test_05_SaneamentoDeCusto()
    Dim alerta As String

    VerificarAprox Pacote_SanearCusto(12.5, alerta), 12.5, "05a custo valido"
    VerificarIgual alerta, "", "05b custo valido sem alerta"
    VerificarAprox Pacote_SanearCusto("12.5", alerta), 12.5, "05c texto com ponto"
    VerificarAprox Pacote_SanearCusto("12,5", alerta), 12.5, "05d texto com virgula"
    VerificarAprox Pacote_SanearCusto(0, alerta), 0#, "05e zero e valido"
    VerificarIgual alerta, "", "05f zero sem alerta"
    VerificarAprox Pacote_SanearCusto(Null, alerta), 0#, "05g Null -> 0"
    Verificar Len(alerta) > 0, "05h Null gera alerta"
    VerificarAprox Pacote_SanearCusto(Empty, alerta), 0#, "05i vazio -> 0"
    Verificar Len(alerta) > 0, "05j vazio gera alerta"
    VerificarAprox Pacote_SanearCusto("abc", alerta), 0#, "05k texto -> 0"
    Verificar Len(alerta) > 0, "05l texto gera alerta"
    VerificarAprox Pacote_SanearCusto(-3.2, alerta), 0#, "05m negativo -> 0"
    Verificar Len(alerta) > 0, "05n negativo gera alerta"
    VerificarAprox Pacote_SanearCusto("-3,2", alerta), 0#, "05o negativo em texto -> 0"
    Verificar Len(alerta) > 0, "05p negativo em texto gera alerta"
End Sub

' ------------------------------------------------------------------------------
' Teste 6: Util_EhNumeroDaBase (limites)
' ------------------------------------------------------------------------------
Private Sub Test_06_NumeroDaBase()
    Verificar Util_EhNumeroDaBase(0), "06a zero numerico"
    Verificar Util_EhNumeroDaBase("100"), "06b inteiro em texto"
    Verificar Util_EhNumeroDaBase("0.065"), "06c decimal com ponto"
    Verificar Util_EhNumeroDaBase(" 7,5 "), "06d decimal com virgula e espacos"
    Verificar Util_EhNumeroDaBase("-2"), "06e sinal negativo"
    Verificar Not Util_EhNumeroDaBase(""), "06f texto vazio"
    Verificar Not Util_EhNumeroDaBase("   "), "06g so espacos"
    Verificar Not Util_EhNumeroDaBase("abc"), "06h letras"
    Verificar Not Util_EhNumeroDaBase("1.2.3"), "06i dois pontos"
    Verificar Not Util_EhNumeroDaBase("R$ 5"), "06j moeda em texto"
    Verificar Not Util_EhNumeroDaBase("-"), "06k so o sinal"
    Verificar Not Util_EhNumeroDaBase(Null), "06l Null"
    Verificar Not Util_EhNumeroDaBase(Empty), "06m Empty"
End Sub

' ------------------------------------------------------------------------------
' Teste 7: guardas que falham antes de qualquer I/O (sem base, sem rede)
' ------------------------------------------------------------------------------
Private Sub Test_07_GuardasSemIO()
    VerificarIgual NumeroDoErro("buscar", "Custo_XX_20261101_1200", "Z:\nao\existe\base.xlsx"), vbObjectError + 902, _
        "07a layout invalido sobe 902 antes de abrir a base"
    VerificarIgual NumeroDoErro("buscar", "Custo_MM_90D_20261101_1200", "Z:\nao\existe\base.xlsx"), ERR_ADO_ARQUIVO_AUSENTE, _
        "07b base ausente sobe ERR_ADO_ARQUIVO_AUSENTE (7101)"
    VerificarIgual BuscarCustosPorPacote("Z:\nao\existe\base.xlsx", "Custo_MM_90D_20261101_1200", New Collection).Count, 0, _
        "07c sem SKUs: resultado vazio sem abrir a base"
    VerificarIgual BuscarCustosPorPacote("Z:\nao\existe\base.xlsx", "Custo_MM_90D_20261101_1200", Nothing).Count, 0, _
        "07d Nothing: resultado vazio sem abrir a base"
    Verificar Pacote_CodigosSemCusto(ColecaoDe("1", "2"), Nothing).Count = 2, "07e sem custos: todos ausentes"
End Sub

' ------------------------------------------------------------------------------
' Teste 8: DTO + matriz do lote (Custo e Operação padrão)
' ------------------------------------------------------------------------------
Private Sub Test_08_DtoEMatrizDoLote()
    Dim produtos As New Collection
    Dim custos As Object
    Dim dtoA As clsProdutoDTO, dtoB As clsProdutoDTO
    Dim matriz As Variant
    Dim colunas As Variant
    Dim iCusto As Long, iOper As Long, j As Long
    Dim semCusto As Collection

    Set dtoA = New clsProdutoDTO
    dtoA.Codigo = 90001: dtoA.Chave = "90001": dtoA.Descricao = "Produto A"
    Set dtoB = New clsProdutoDTO
    dtoB.Codigo = 90002: dtoB.Chave = "90002": dtoB.Descricao = "Produto B"
    produtos.Add dtoA
    produtos.Add dtoB

    Set custos = CreateObject("Scripting.Dictionary")
    custos.CompareMode = vbTextCompare
    custos.Add "90001", 12.5

    Call Catalogo_AplicarCustos(produtos, custos)
    Verificar dtoA.TemCusto, "08a produto com custo no pacote"
    VerificarAprox dtoA.Custo, 12.5, "08b valor do custo"
    Verificar Not dtoB.TemCusto, "08c produto sem custo no pacote"

    Set semCusto = Catalogo_CodigosSemCusto(produtos)
    VerificarIgual semCusto.Count, 1, "08d um produto sem custo"
    VerificarIgual semCusto(1), "90002", "08e e o 90002"
    VerificarIgual Catalogo_ChavesDosProdutos(produtos).Count, 2, "08f chaves para a consulta"

    colunas = Catalogo_ColunasLote()
    For j = LBound(colunas) To UBound(colunas)
        If CStr(colunas(j)) = Col_Custo() Then iCusto = j - LBound(colunas) + 1
        If CStr(colunas(j)) = Col_Operacao() Then iOper = j - LBound(colunas) + 1
    Next j
    Verificar iCusto > 0, "08g a coluna Custo faz parte do lote"
    Verificar iOper > 0, "08h a coluna Operacao continua no lote"

    matriz = Catalogo_MontarMatriz(produtos, colunas)
    VerificarAprox CDbl(matriz(1, iCusto)), 12.5, "08i matriz: custo do produto A"
    Verificar IsEmpty(matriz(2, iCusto)), "08j matriz: produto sem custo fica vazio"
    VerificarIgual matriz(1, iOper), OPERACAO_PADRAO_LOTE, "08k matriz: Operacao padrao do lote"

    ' Sem dicionario: ninguem tem custo
    Call Catalogo_AplicarCustos(produtos, Nothing)
    Verificar Not dtoA.TemCusto, "08l Nothing remove o custo"
End Sub

' ------------------------------------------------------------------------------
' Teste 9: estado persistido no Named Range oculto cfg_PacoteCustosAtivo
' ------------------------------------------------------------------------------
Private Sub Test_09_EstadoPersistido()
    Dim anterior As String
    Dim existia As Boolean
    Dim nm As Name
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As XlCalculation

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    anterior = ObterPacoteCustosAtivo()
    On Error Resume Next
    existia = Not (ThisWorkbook.Names(NOME_PACOTE_ATIVO) Is Nothing)
    On Error GoTo 0

    Call SalvarPacoteCustosAtivo("Custo_MM_90D_20261101_1200")
    VerificarIgual ObterPacoteCustosAtivo(), "Custo_MM_90D_20261101_1200", "09a grava e le o pacote"
    Set nm = ThisWorkbook.Names(NOME_PACOTE_ATIVO)
    Verificar Not nm.Visible, "09b o nome fica oculto"

    Call SalvarPacoteCustosAtivo("Custo_EX_20261110_1530")
    VerificarIgual ObterPacoteCustosAtivo(), "Custo_EX_20261110_1530", "09c regravar substitui o valor"

    VerificarIgual NumeroDoErro("salvar", "Planilha1"), vbObjectError + 902, "09d nome invalido e recusado (902)"
    VerificarIgual ObterPacoteCustosAtivo(), "Custo_EX_20261110_1530", "09e recusa nao altera o estado"

    ' Restaura o estado que existia antes do teste
    If existia And Len(anterior) > 0 Then
        Call SalvarPacoteCustosAtivo(anterior)
    ElseIf existia Then
        ThisWorkbook.Names(NOME_PACOTE_ATIVO).RefersTo = "="""""
    Else
        ThisWorkbook.Names(NOME_PACOTE_ATIVO).Delete
    End If
    VerificarIgual ObterPacoteCustosAtivo(), anterior, "09f estado original restaurado"

    VerificarIgual Application.ScreenUpdating, prevScreen, "09g ScreenUpdating preservado"
    VerificarIgual Application.EnableEvents, prevEvents, "09h EnableEvents preservado"
    VerificarIgual Application.Calculation, prevCalc, "09i Calculation preservado"
End Sub

' ------------------------------------------------------------------------------
' Teste 10: texto do cabeçalho "Custos: [pacote]" (função pura, sem PageSetup)
' ------------------------------------------------------------------------------
Private Sub Test_10_TextoDoCabecalho()
    Dim h As String

    VerificarIgual Pacote_MontarCabecalho("", "Custo_MM_90D_20261101_1200"), _
        "Custos: Custo_MM_90D_20261101_1200", "10a quadrante vazio: formato literal"

    h = Pacote_MontarCabecalho("Pagina &P", "Custo_EX_20261110_1530")
    VerificarIgual h, "Pagina &P" & vbLf & "Custos: Custo_EX_20261110_1530", "10b outro conteudo e preservado (nova linha)"

    h = Pacote_MontarCabecalho(h, "Custo_MM_ULT_20261105_0900")
    VerificarIgual h, "Pagina &P" & vbLf & "Custos: Custo_MM_ULT_20261105_0900", "10c so a linha Custos: e trocada"

    VerificarIgual Pacote_MontarCabecalho("Custos: Custo_EX_20261110_1530", "Custo_MM_90D_20261101_1200"), _
        "Custos: Custo_MM_90D_20261101_1200", "10d troca sem duplicar a linha"

    VerificarIgual Pacote_MontarCabecalho("", "Custo_MM_A&B"), "Custos: Custo_MM_A&&B", "10e '&' escapado como '&&'"
    VerificarIgual Pacote_MontarCabecalho("", "  Custo_EX_X  "), "Custos: Custo_EX_X", "10f nome aparado"
End Sub

' ------------------------------------------------------------------------------
' Teste 11: PageSetup real numa aba temporária (tmp_test_): o quadrante central
'           (nome do arquivo, &F) e o esquerdo não podem ser alterados
' ------------------------------------------------------------------------------
Private Sub Test_11_CabecalhoNoPageSetup()
    Dim ws As Worksheet
    Dim errNum As Long
    Dim prevAlerts As Boolean
    Dim centro As String

    prevAlerts = Application.DisplayAlerts
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "tmp_test_Cabecalho"
    centro = "&""Aptos,Regular""&16&U&F"

    On Error Resume Next
    ws.PageSetup.LeftHeader = "ESQ"
    ws.PageSetup.CenterHeader = centro
    ws.PageSetup.RightHeader = vbNullString
    errNum = Err.Number
    On Error GoTo 0

    If errNum <> 0 Then
        Debug.Print " [SKIP] 11 PageSetup indisponivel (sem impressora instalada?)"
    Else
        Call AtualizarCabecalhoImpressao(ws, "Custo_MM_90D_20261101_1200")
        VerificarIgual ws.PageSetup.RightHeader, "Custos: Custo_MM_90D_20261101_1200", "11a formato literal no quadrante direito"
        VerificarIgual ws.PageSetup.CenterHeader, centro, "11b quadrante central (&F) intacto"
        VerificarIgual ws.PageSetup.LeftHeader, "ESQ", "11c quadrante esquerdo intacto"

        Call AtualizarCabecalhoImpressao(ws, "Custo_EX_20261110_1530")
        VerificarIgual ws.PageSetup.RightHeader, "Custos: Custo_EX_20261110_1530", "11d troca de pacote atualiza o cabecalho"

        Call AtualizarCabecalhoImpressao(ws, "")
        VerificarIgual ws.PageSetup.RightHeader, "Custos: Custo_EX_20261110_1530", "11e pacote vazio nao altera o cabecalho"

        ws.PageSetup.RightHeader = "Pagina &P"
        Call AtualizarCabecalhoImpressao(ws, "Custo_MM_ULT_20261105_0900")
        VerificarIgual ws.PageSetup.RightHeader, "Pagina &P" & vbLf & "Custos: Custo_MM_ULT_20261105_0900", "11f conteudo anterior do quadrante preservado"
    End If

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = prevAlerts
End Sub

' ==============================================================================
' Ponto de entrada 2: integração ADO/OLEDB com a base sintética
' ==============================================================================
Public Sub RunCostPackageIntegrationTests(Optional ByVal caminhoFixture As String = vbNullString)
    Dim caminho As String
    Dim pacotes As Collection

    Zerar
    Debug.Print "--- INTEGRACAO ADO (SPEC-009) com base sintetica ---"

    caminho = caminhoFixture
    If Len(caminho) = 0 Then caminho = LocalizarFixture()
    If Len(caminho) = 0 Then
        Debug.Print " [SKIP] mock_Cerulean_DB_Custos.xlsx nao encontrado. Passe o caminho como argumento."
        Exit Sub
    End If
    Debug.Print " Base sintetica: " & caminho

    ' I1: descoberta via OpenSchema + filtro de dominio
    Set pacotes = ListarPacotesCustosDisponiveis(caminho)
    VerificarIgual pacotes.Count, 3, "I1a somente os 3 pacotes validos da base"
    If pacotes.Count = 3 Then
        VerificarIgual pacotes(1), "Custo_MM_ULT_20261105_0900", "I1b ordem: MM_ULT"
        VerificarIgual pacotes(2), "Custo_MM_90D_20261101_1200", "I1c ordem: MM_90D"
        VerificarIgual pacotes(3), "Custo_EX_20261110_1530", "I1d ordem: EX"
    End If

    Teste_I2_PacoteMediaMovel caminho
    Teste_I3_TrocaDePacote caminho
    Teste_I4_PacoteEngenharia caminho
    Teste_I5_LotesDe40 caminho
    Teste_I6_PacoteInexistente caminho

    Resumo "INTEGRA" & ChrW(199) & ChrW(195) & "O CONCLU" & ChrW(205) & "DA (SPEC-009 / ADO)"
End Sub

Private Function LocalizarFixture() As String
    Dim sep As String
    Dim base As String
    Dim cand(1 To 3) As String
    Dim i As Long

    sep = Application.PathSeparator
    On Error Resume Next
    base = ThisWorkbook.Path
    On Error GoTo 0
    If Len(base) = 0 Then Exit Function

    cand(1) = base & sep & "tests" & sep & "fixtures" & sep & "mock_Cerulean_DB_Custos.xlsx"
    cand(2) = base & sep & ".." & sep & "tests" & sep & "fixtures" & sep & "mock_Cerulean_DB_Custos.xlsx"
    cand(3) = base & sep & "data" & sep & "mock_Cerulean_DB_Custos.xlsx"
    For i = 1 To 3
        If Util_ArquivoExiste(cand(i)) Then LocalizarFixture = cand(i): Exit Function
    Next i
End Function

Private Sub Teste_I2_PacoteMediaMovel(ByVal caminho As String)
    Dim skus As Collection
    Dim custos As Object
    Dim alertas As Collection

    Set skus = ColecaoDe("90001", "90002", "90003", "90004", "90005", "90006", "90007", "A'B-01", "99999")
    Set custos = BuscarCustosPorPacote(caminho, "Custo_MM_90D_20261101_1200", skus, alertas)

    VerificarAprox CustoDe(custos, "90001"), 12.5, "I2a usa Custo_Medio_Familia (e nao Custo_Medio_SKU)"
    VerificarAprox CustoDe(custos, "90002"), 7.35, "I2b segundo SKU"
    VerificarAprox CustoDe(custos, "90003"), 0#, "I2c custo em branco -> 0,00"
    VerificarAprox CustoDe(custos, "90004"), 0#, "I2d custo negativo -> 0,00"
    VerificarAprox CustoDe(custos, "90005"), 0#, "I2e custo texto -> 0,00"
    VerificarAprox CustoDe(custos, "90006"), 100#, "I2f custo inteiro"
    VerificarAprox CustoDe(custos, "90007"), 20#, "I2g SKU duplicado: vale a primeira ocorrencia"
    VerificarAprox CustoDe(custos, "A'B-01"), 5.55, "I2h SKU com apostrofo (parametro, sem quebrar o SQL)"
    Verificar Not custos.Exists("99999"), "I2i SKU fora do pacote nao aparece"
    VerificarIgual alertas.Count, 3, "I2j tres alertas (branco, negativo, texto)"
    VerificarIgual Pacote_CodigosSemCusto(skus, custos).Count, 1, "I2k um SKU sem custo no pacote"
End Sub

Private Sub Teste_I3_TrocaDePacote(ByVal caminho As String)
    Dim custos As Object

    Set custos = BuscarCustosPorPacote(caminho, "Custo_MM_ULT_20261105_0900", ColecaoDe("90001", "90002"))
    VerificarAprox CustoDe(custos, "90001"), 13.75, "I3a outro pacote MM devolve o custo dele"
    VerificarAprox CustoDe(custos, "90002"), 7.35, "I3b SKU igual nos dois pacotes"
End Sub

Private Sub Teste_I4_PacoteEngenharia(ByVal caminho As String)
    Dim custos As Object

    Set custos = BuscarCustosPorPacote(caminho, "Custo_EX_20261110_1530", ColecaoDe("90001", "90002", "90009"))
    VerificarAprox CustoDe(custos, "90001"), 11.11, "I4a usa Custo_Final_Engenharia (e nao Custo_Medio_SKU)"
    VerificarAprox CustoDe(custos, "90002"), 22.22, "I4b codigo numerico na base comparado como texto"
    VerificarAprox CustoDe(custos, "90009"), 33.33, "I4c SKU so existente no EX"
End Sub

Private Sub Teste_I5_LotesDe40(ByVal caminho As String)
    Dim skus As New Collection
    Dim custos As Object
    Dim i As Long

    For i = 91001 To 91045
        skus.Add CStr(i)
    Next i
    Set custos = BuscarCustosPorPacote(caminho, "Custo_MM_90D_20261101_1200", skus)

    VerificarIgual custos.Count, 45, "I5a 45 SKUs (2 consultas: 40 + 5) todos encontrados"
    VerificarAprox CustoDe(custos, "91001"), 10#, "I5b primeiro SKU do primeiro lote"
    VerificarAprox CustoDe(custos, "91040"), 16#, "I5c ultimo SKU do primeiro lote"
    VerificarAprox CustoDe(custos, "91041"), 17.5, "I5d primeiro SKU do segundo lote"
    VerificarAprox CustoDe(custos, "91045"), 13#, "I5e ultimo SKU do segundo lote"
End Sub

Private Sub Teste_I6_PacoteInexistente(ByVal caminho As String)
    Dim errNum As Long

    On Error Resume Next
    Call BuscarCustosPorPacote(caminho, "Custo_MM_90D_19990101_0000", ColecaoDe("90001"))
    errNum = Err.Number
    On Error GoTo 0
    VerificarIgual errNum, vbObjectError + 903, "I6a aba que nao existe mais sobe vbObjectError+903"
End Sub
