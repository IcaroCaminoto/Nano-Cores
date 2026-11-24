Attribute VB_Name = "mod_Multidirectional_Engine"
' ==============================================================================
' Módulo: mod_Multidirectional_Engine
' Camada: Pricing Engine (SPEC-004 / SPEC-005)
' Finalidade: Recalcular a tabela CfPcfic a partir da coluna-âncora (margem,
'             preço final ou preço base), com os impostos lidos diretamente de
'             Cerulean_DB_Impostos.xlsx via ADO/SQL (sem CfImpostos e sem
'             fórmulas ÍNDICE/CORRESP/PROCX na planilha).
'             A leitura da base fica em mod_ImpostoRepository (camada pura);
'             este módulo só cuida do cache, da planilha e do cálculo.
' Fase 0 (refatoração para DLL): este módulo contém as regras que migrarão para
'             a Action "precificacao.recalcular". Nomes de tabela/coluna e modos
'             vêm de mod_PricingSchema; caminhos, de mod_Config.
'
' Modelo de cálculo (alíquotas em fração):
'   Encargos  E  = Frete + Acordo + Comissão + Outros + ICMS + Pis/Cofins + IPI + ICMS ST
'                  (impostos "por dentro", já com fator da Operação e exceções)
'   Preço Final Caixa  PFC = Custo / (1 - (E + M))
'   Margem Bruta       M   = 1 - E - Custo / PFC
'   Preço Final Unit.  PFU = PFC / Qtd Padrão
'   Preço Base Caixa   PBC = PFC / (1 + IPI_fora)
'                          = (Custo / (1 - E - M)) / (1 + IPI_fora)
'   Preço Base Unit.   PBU = PBC / Qtd Padrão
'                          = (Custo Unitário / (1 - E - M)) / (1 + IPI_fora)
'
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Cache das alíquotas por UF (evita uma consulta de rede a cada edição)
' ------------------------------------------------------------------------------
Private Const CACHE_SEGUNDOS As Long = 300

Private m_CacheUF As String
Private m_CacheImpostos As Object               ' chave SKU -> clsAliquotasDTO
Private m_CacheTemExcecoes As Boolean           ' base já tem colunas calc_*
Private m_CacheHora As Date

' ------------------------------------------------------------------------------
' Índices de coluna resolvidos uma única vez por execução
' ------------------------------------------------------------------------------
Private Type PricingColumns
    Codigo As Long
    Qtd As Long
    Custo As Long
    Frete As Long
    Acordo As Long
    Comissao As Long
    Outros As Long
    Operacao As Long
    ICMS As Long
    PisCofins As Long
    IPI As Long
    ST As Long
    IPIF As Long
    STF As Long
    Margem As Long
    PrecoFinalCaixa As Long
    PrecoFinalUnit As Long
    PrecoBaseCaixa As Long
    PrecoBaseUnit As Long
End Type

' ==============================================================================
' Procedimento: MotorPrecificacaoMultidirecional
' Finalidade: Ponto de entrada do motor. Não altera o estado do Application
'             (eventos/tela/cálculo): isso é responsabilidade do chamador
'             (ver Pricing_RecalcularTabela).
' Parâmetros:
'   - modoAlvo (String): Um dos identificadores MODO_* de mod_PricingSchema.
' ==============================================================================
Public Sub MotorPrecificacaoMultidirecional(ByVal modoAlvo As String)
    Dim LoPrecificacao As New clsListState
    Dim cols As PricingColumns
    Dim colunaFaltante As String
    Dim arrDados As Variant
    Dim siglaUF As String
    Dim dictImpostos As Object
    Dim dictFatores As Object
    Dim dictExcecoes As Object
    Dim ipiForaEfetivo As Double
    Dim i As Long

    If Not ModoValido(modoAlvo) Then
        MsgBox "Modo de cálculo desconhecido: " & modoAlvo, vbCritical
        Exit Sub
    End If

    If Not LoPrecificacao.Init() Then
        MsgBox "Tabela de precificação """ & TABELA_PRECIFICACAO & """ não encontrada.", vbCritical
        Exit Sub
    End If
    If Not LoPrecificacao.TemDados Then Exit Sub

    ' --------------------------------------------------------------------------
    ' 1. Mapeamento de colunas (todas obrigatórias: o motor lê ou escreve cada uma)
    ' --------------------------------------------------------------------------
    If Not MapearColunas(LoPrecificacao, cols, colunaFaltante) Then
        MsgBox "Erro de Mapeamento: a coluna obrigatória """ & colunaFaltante & _
               """ não está presente na tabela " & TABELA_PRECIFICACAO & ".", vbCritical
        Exit Sub
    End If

    ' Garante que Custo / Qtd Padrão (PROCV de outras abas) estejam atualizados
    ' mesmo com o cálculo em modo manual.
    LoPrecificacao.DataBodyPadrao.Calculate

    arrDados = LoPrecificacao.DataBodyPadrao.Value2
    If Not IsArray(arrDados) Then Exit Sub

    ' --------------------------------------------------------------------------
    ' 2. Parâmetros tributários: UF selecionada, alíquotas (ADO), operações e
    '    exceções (tabelas de configuração lidas em memória)
    ' --------------------------------------------------------------------------
    siglaUF = Pricing_ObterSiglaUF()
    If Len(siglaUF) = 0 Then
        MsgBox "Selecione um estado válido no campo REGIÃO (" & ABA_PARAMETROS & "!" & CELULA_UF & _
               ") antes de precificar.", vbExclamation
        Exit Sub
    End If

    Set dictImpostos = Pricing_ObterImpostosDaUF(siglaUF)
    If dictImpostos Is Nothing Then Exit Sub

    Set dictFatores = CarregarFatoresOperacao()

    ' Exceções (imposto desligado por SKU) vêm da base. Enquanto a base ainda
    ' não tiver as colunas calc_*, usa a tabela local CfExcecoesImpostos
    ' (transição; some quando a aba Impostos for excluída).
    If m_CacheTemExcecoes Then
        Set dictExcecoes = Nothing
    Else
        Set dictExcecoes = CarregarExcecoesImpostos()
    End If

    ' --------------------------------------------------------------------------
    ' 3. Cálculo linha a linha, 100% em memória
    ' --------------------------------------------------------------------------
    For i = 1 To UBound(arrDados, 1)
        ipiForaEfetivo = CalcularImpostosLinha(arrDados, i, cols, dictImpostos, dictFatores, dictExcecoes)
        Call CalcularPrecoMulti(arrDados, i, modoAlvo, cols, ipiForaEfetivo)
    Next i

    ' --------------------------------------------------------------------------
    ' 4. Injeção cirúrgica
    '    a) impostos e IPIF/STF: sempre (substituem as fórmulas legadas)
    '    b) preços/margem: todas as colunas derivadas, nunca a âncora
    ' --------------------------------------------------------------------------
    Call LoPrecificacao.InjetarLoteCirurgico(arrDados, Col_ICMS(), Col_PisCofins(), Col_IPI(), _
                                             Col_ICMSST(), Col_IPIF(), Col_STF())

    Select Case modoAlvo
        Case MODO_MARGEM_BRUTA
            Call LoPrecificacao.InjetarLoteCirurgico(arrDados, Col_PrecoFinalCaixa(), Col_PrecoFinalUnitario(), Col_PrecoBaseCaixa(), Col_PrecoBaseUnitario())
        Case MODO_PRECO_FINAL_CAIXA
            Call LoPrecificacao.InjetarLoteCirurgico(arrDados, Col_PrecoFinalUnitario(), Col_MargemBruta(), Col_PrecoBaseCaixa(), Col_PrecoBaseUnitario())
        Case MODO_PRECO_FINAL_UNITARIO
            Call LoPrecificacao.InjetarLoteCirurgico(arrDados, Col_PrecoFinalCaixa(), Col_MargemBruta(), Col_PrecoBaseCaixa(), Col_PrecoBaseUnitario())
        Case MODO_PRECO_BASE_CAIXA
            Call LoPrecificacao.InjetarLoteCirurgico(arrDados, Col_PrecoFinalCaixa(), Col_PrecoFinalUnitario(), Col_MargemBruta(), Col_PrecoBaseUnitario())
        Case MODO_PRECO_BASE_UNITARIO
            Call LoPrecificacao.InjetarLoteCirurgico(arrDados, Col_PrecoFinalCaixa(), Col_PrecoFinalUnitario(), Col_MargemBruta(), Col_PrecoBaseCaixa())
    End Select
End Sub

' ==============================================================================
' Procedimento: Pricing_RecalcularTabela
' Finalidade: Recalcular toda a CfPcfic com proteção do estado do Application.
'             Usado pelo evento da planilha, pelo botão Atualizar e pela troca
'             de UF na aba de parâmetros.
' Parâmetros:
'   - modoAlvo (opcional): âncora a usar; vazio = âncora atual (negrito).
' ==============================================================================
Public Sub Pricing_RecalcularTabela(Optional ByVal modoAlvo As String = vbNullString)
    Dim LoPrecificacao As New clsListState
    Dim prevEvents As Boolean, prevScreen As Boolean
    Dim prevCalc As XlCalculation

    If Not LoPrecificacao.Init() Then Exit Sub
    If Not LoPrecificacao.TemDados Then Exit Sub

    If Len(modoAlvo) = 0 Then modoAlvo = Pricing_DetectarModoAtual(LoPrecificacao)

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation

    On Error GoTo SaidaErro
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call MotorPrecificacaoMultidirecional(modoAlvo)
    Call Pricing_AplicarDestaqueAncora(LoPrecificacao, modoAlvo)

SaidaLimpa:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Exit Sub

SaidaErro:
    MsgBox "Falha no cálculo da precificação: " & Err.Description, vbCritical, "Erro " & Err.Number
    Resume SaidaLimpa
End Sub

' ==============================================================================
' Procedimento: Pricing_LimparCacheImpostos
' Finalidade: Forçar nova leitura da base de impostos na próxima execução
'             (troca de UF, botão Atualizar ou base alterada na rede).
' ==============================================================================
Public Sub Pricing_LimparCacheImpostos()
    m_CacheUF = vbNullString
    Set m_CacheImpostos = Nothing
    m_CacheTemExcecoes = False
End Sub

' ==============================================================================
' Procedimento: Pricing_AplicarDestaqueAncora
' Finalidade: Destacar (negrito azul cerúleo) a coluna-âncora ativa e
'             neutralizar as demais colunas-âncora possíveis.
' ==============================================================================
Public Sub Pricing_AplicarDestaqueAncora(ByVal LoPrecificacao As clsListState, ByVal modoAtual As String)
    Dim modo As Variant
    Dim idx As Long
    Dim rng As Range

    If Not LoPrecificacao.TemDados Then Exit Sub

    For Each modo In TodosOsModos()
        idx = LoPrecificacao.IndiceColuna(ColunaAncoraDoModo(CStr(modo)))
        If idx > 0 Then
            Set rng = LoPrecificacao.DataBodyPadrao.Columns(idx)
            If modo = modoAtual Then
                rng.Font.Color = RGB(20, 66, 114)
                rng.Font.Bold = True
                rng.Font.Size = 12
            Else
                rng.Font.ColorIndex = xlAutomatic
                rng.Font.Bold = False
                rng.Font.Size = 10
            End If
        End If
    Next modo
End Sub

' ==============================================================================
' Função: Pricing_DetectarModoAtual
' Finalidade: Descobrir a âncora ativa pela formatação (negrito) da primeira
'             linha. Fallback: célula A1 legada ("Preço Final" = preço
'             unitário, gravada por mod_Inputs); padrão = Margem Bruta.
' ==============================================================================
Public Function Pricing_DetectarModoAtual(ByVal LoPrecificacao As clsListState) As String
    Dim modo As Variant
    Dim idx As Long
    Dim legado As String

    If LoPrecificacao.TemDados Then
        For Each modo In TodosOsModos()
            idx = LoPrecificacao.IndiceColuna(ColunaAncoraDoModo(CStr(modo)))
            If idx > 0 Then
                If LoPrecificacao.DataBodyPadrao.Cells(1, idx).Font.Bold = True Then
                    Pricing_DetectarModoAtual = CStr(modo)
                    Exit Function
                End If
            End If
        Next modo
    End If

    legado = Util_Texto(LoPrecificacao.Planilha.Range("A1").Value2)
    If StrComp(legado, "Pre" & ChrW(231) & "o Final", vbTextCompare) = 0 Then
        Pricing_DetectarModoAtual = MODO_PRECO_FINAL_UNITARIO
    Else
        Pricing_DetectarModoAtual = MODO_MARGEM_BRUTA
    End If
End Function

' ==============================================================================
' Função Privada: MapearColunas
' Finalidade: Resolver todos os índices de coluna; devolve False e o nome da
'             primeira coluna ausente caso alguma não exista.
' ==============================================================================
Private Function MapearColunas(ByVal lo As clsListState, ByRef cols As PricingColumns, ByRef outFaltante As String) As Boolean
    MapearColunas = False

    If Not ResolverColuna(lo, Col_Codigo(), cols.Codigo, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_QtdPadrao(), cols.Qtd, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_Custo(), cols.Custo, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_Frete(), cols.Frete, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_Acordo(), cols.Acordo, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_Comissao(), cols.Comissao, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_Outros(), cols.Outros, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_Operacao(), cols.Operacao, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_ICMS(), cols.ICMS, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_PisCofins(), cols.PisCofins, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_IPI(), cols.IPI, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_ICMSST(), cols.ST, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_IPIF(), cols.IPIF, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_STF(), cols.STF, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_MargemBruta(), cols.Margem, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_PrecoFinalCaixa(), cols.PrecoFinalCaixa, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_PrecoFinalUnitario(), cols.PrecoFinalUnit, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_PrecoBaseCaixa(), cols.PrecoBaseCaixa, outFaltante) Then Exit Function
    If Not ResolverColuna(lo, Col_PrecoBaseUnitario(), cols.PrecoBaseUnit, outFaltante) Then Exit Function

    MapearColunas = True
End Function

Private Function ResolverColuna(ByVal lo As clsListState, ByVal nome As String, _
                                ByRef outIndice As Long, ByRef outFaltante As String) As Boolean
    outIndice = lo.IndiceColuna(nome)
    If outIndice = 0 Then outFaltante = nome
    ResolverColuna = (outIndice > 0)
End Function

' ==============================================================================
' Função Privada: IndiceColunaAncora
' ==============================================================================
Private Function IndiceColunaAncora(ByVal modoAlvo As String, ByRef cols As PricingColumns) As Long
    Select Case modoAlvo
        Case MODO_MARGEM_BRUTA:          IndiceColunaAncora = cols.Margem
        Case MODO_PRECO_FINAL_CAIXA:     IndiceColunaAncora = cols.PrecoFinalCaixa
        Case MODO_PRECO_FINAL_UNITARIO:  IndiceColunaAncora = cols.PrecoFinalUnit
        Case MODO_PRECO_BASE_CAIXA:      IndiceColunaAncora = cols.PrecoBaseCaixa
        Case MODO_PRECO_BASE_UNITARIO:   IndiceColunaAncora = cols.PrecoBaseUnit
    End Select
End Function

' ==============================================================================
' Função Privada: CalcularImpostosLinha
' Finalidade: Preencher ICMS, Pis/Cofins, IPI, ICMS ST, IPIF e STF da linha.
' Paridade com as fórmulas legadas:
'   - SKU sem cadastro na base       -> todas as 6 colunas vazias
'   - Operação vazia/não cadastrada  -> impostos vazios; IPIF/STF preenchidos
' Exceções: da própria base (clsAliquotasDTO.Calc*) ou, na transição, do
'   dicionário dictExcecoes lido da tabela local.
' Retorno: IPI "por fora" efetivo da linha (divisor do Preço Base).
' ==============================================================================
Private Function CalcularImpostosLinha(ByRef arrDados As Variant, ByVal i As Long, ByRef cols As PricingColumns, _
                                       ByVal dictImpostos As Object, ByVal dictFatores As Object, _
                                       ByVal dictExcecoes As Object) As Double
    Dim vSKU As String
    Dim vOperacao As String
    Dim aliq As clsAliquotasDTO
    Dim exc As Variant
    Dim Produto As clsProduto

    arrDados(i, cols.ICMS) = Empty
    arrDados(i, cols.PisCofins) = Empty
    arrDados(i, cols.IPI) = Empty
    arrDados(i, cols.ST) = Empty
    arrDados(i, cols.IPIF) = Empty
    arrDados(i, cols.STF) = Empty

    vSKU = Util_ChaveSKU(arrDados(i, cols.Codigo))
    If Len(vSKU) = 0 Then Exit Function
    If Not dictImpostos.Exists(vSKU) Then Exit Function

    Set aliq = dictImpostos(vSKU)

    Set Produto = New clsProduto
    Call Produto.Init(vSKU, aliq.ICMS, aliq.PisCofins, aliq.IPIFora, aliq.STFora)

    arrDados(i, cols.IPIF) = Produto.IPIFora
    arrDados(i, cols.STF) = Produto.STFora

    vOperacao = Util_Texto(arrDados(i, cols.Operacao))
    If Len(vOperacao) = 0 Then Exit Function
    If Not dictFatores.Exists(vOperacao) Then Exit Function

    exc = Array(aliq.CalcICMS, aliq.CalcPisCofins, aliq.CalcIPI, aliq.CalcST)
    If Not dictExcecoes Is Nothing Then
        If dictExcecoes.Exists(vSKU) Then exc = dictExcecoes(vSKU)
    End If

    Call Produto.AplicarOperacao(dictFatores(vOperacao), exc(0), exc(1), exc(2), exc(3))

    arrDados(i, cols.ICMS) = Produto.ICMS
    arrDados(i, cols.PisCofins) = Produto.PisCofins
    arrDados(i, cols.IPI) = Produto.IPI
    arrDados(i, cols.ST) = Produto.ST

    CalcularImpostosLinha = Produto.IPIForaEfetivo
End Function

' ==============================================================================
' Procedimento Privado: CalcularPrecoMulti
' Finalidade: Recalcular preços e margem de uma linha a partir da âncora.
' Regras:
'   - Custo <= 0          -> linha ignorada (mantém o que estiver na tela)
'   - Âncora vazia        -> limpa as colunas derivadas (sem âncora, sem preço)
'   - Âncora não numérica -> linha ignorada
'   - Preço resultante <= 0 ou divisor <= 0 -> limpa as colunas derivadas
' ==============================================================================
Private Sub CalcularPrecoMulti(ByRef arrDados As Variant, ByVal i As Long, ByVal modoAlvo As String, _
                               ByRef cols As PricingColumns, ByVal ipiForaEfetivo As Double)
    Dim colAncora As Long
    Dim vAncoraBruta As Variant
    Dim vAncora As Double
    Dim vCusto As Double, vQtd As Double, vEncargos As Double
    Dim vMargem As Double, vDivisor As Double
    Dim vFatorIPI As Double
    Dim vPrecoC As Double, vBaseC As Double

    vCusto = Util_ToDouble(arrDados(i, cols.Custo))
    If vCusto <= 0 Then Exit Sub

    colAncora = IndiceColunaAncora(modoAlvo, cols)
    vAncoraBruta = arrDados(i, colAncora)
    If IsError(vAncoraBruta) Then Exit Sub
    If Len(Trim$(CStr(vAncoraBruta))) = 0 Then
        Call LimparDerivados(arrDados, i, cols, colAncora)
        Exit Sub
    End If
    If Not IsNumeric(vAncoraBruta) Then Exit Sub
    vAncora = CDbl(vAncoraBruta)

    vQtd = Util_ToDouble(arrDados(i, cols.Qtd))
    vFatorIPI = 1# + ipiForaEfetivo

    vEncargos = Util_ToDouble(arrDados(i, cols.Frete)) + Util_ToDouble(arrDados(i, cols.Acordo)) + _
                Util_ToDouble(arrDados(i, cols.Comissao)) + Util_ToDouble(arrDados(i, cols.Outros)) + _
                Util_ToDouble(arrDados(i, cols.ICMS)) + Util_ToDouble(arrDados(i, cols.PisCofins)) + _
                Util_ToDouble(arrDados(i, cols.IPI)) + Util_ToDouble(arrDados(i, cols.ST))

    ' ---- Preço Final Caixa a partir da âncora --------------------------------
    Select Case modoAlvo
        Case MODO_MARGEM_BRUTA
            vMargem = vAncora
            vDivisor = 1# - (vEncargos + vMargem)
            If vDivisor <= 0 Then
                Call LimparDerivados(arrDados, i, cols, colAncora)
                Exit Sub
            End If
            vPrecoC = vCusto / vDivisor

        Case MODO_PRECO_FINAL_CAIXA
            vPrecoC = vAncora

        Case MODO_PRECO_FINAL_UNITARIO
            If vQtd <= 0 Then Exit Sub
            vPrecoC = vAncora * vQtd

        Case MODO_PRECO_BASE_CAIXA
            vPrecoC = vAncora * vFatorIPI

        Case MODO_PRECO_BASE_UNITARIO
            If vQtd <= 0 Then Exit Sub
            vPrecoC = vAncora * vQtd * vFatorIPI
    End Select

    If vPrecoC <= 0 Then
        Call LimparDerivados(arrDados, i, cols, colAncora)
        Exit Sub
    End If

    ' ---- Derivadas -----------------------------------------------------------
    vBaseC = vPrecoC / vFatorIPI

    arrDados(i, cols.PrecoFinalCaixa) = vPrecoC
    arrDados(i, cols.PrecoBaseCaixa) = vBaseC
    If modoAlvo <> MODO_MARGEM_BRUTA Then
        arrDados(i, cols.Margem) = 1# - vEncargos - vCusto / vPrecoC
    End If

    If vQtd > 0 Then
        arrDados(i, cols.PrecoFinalUnit) = vPrecoC / vQtd
        arrDados(i, cols.PrecoBaseUnit) = vBaseC / vQtd
    Else
        arrDados(i, cols.PrecoFinalUnit) = Empty
        arrDados(i, cols.PrecoBaseUnit) = Empty
    End If
End Sub

Private Sub LimparDerivados(ByRef arrDados As Variant, ByVal i As Long, ByRef cols As PricingColumns, ByVal colAncora As Long)
    If colAncora <> cols.Margem Then arrDados(i, cols.Margem) = Empty
    If colAncora <> cols.PrecoFinalCaixa Then arrDados(i, cols.PrecoFinalCaixa) = Empty
    If colAncora <> cols.PrecoFinalUnit Then arrDados(i, cols.PrecoFinalUnit) = Empty
    If colAncora <> cols.PrecoBaseCaixa Then arrDados(i, cols.PrecoBaseCaixa) = Empty
    If colAncora <> cols.PrecoBaseUnit Then arrDados(i, cols.PrecoBaseUnit) = Empty
End Sub

' ==============================================================================
' Função: Pricing_ObterSiglaUF
' Finalidade: Ler o estado escolhido em AreaCartoes!E9 (nome, ex. "SÃO PAULO")
'             e convertê-lo em sigla ("SP") pela tabela dUF.
' Retorno: sigla de 2 letras ou "" se não for possível determinar.
' ==============================================================================
Public Function Pricing_ObterSiglaUF() As String
    Dim wsParam As Worksheet
    Dim valor As String
    Dim lo As ListObject
    Dim arr As Variant
    Dim idxNome As Long, idxSigla As Long
    Dim i As Long

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(ABA_PARAMETROS)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    valor = UCase$(Util_Texto(wsParam.Range(CELULA_UF).Value2))
    If Len(valor) = 0 Then Exit Function

    If Len(valor) = 2 And valor Like "[A-Z][A-Z]" Then
        Pricing_ObterSiglaUF = valor
        Exit Function
    End If

    Set lo = LocalizarTabela(TABELA_UF)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxNome = IndiceListColumn(lo, ColUF_Nome())
    idxSigla = IndiceListColumn(lo, ColUF_Sigla())
    If idxNome = 0 Or idxSigla = 0 Then Exit Function

    arr = lo.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        If StrComp(Util_Texto(arr(i, idxNome)), valor, vbTextCompare) = 0 Then
            valor = UCase$(Util_Texto(arr(i, idxSigla)))
            If Len(valor) = 2 And valor Like "[A-Z][A-Z]" Then Pricing_ObterSiglaUF = valor
            Exit Function
        End If
    Next i
End Function

' ==============================================================================
' Função: Pricing_ObterImpostosDaUF
' Finalidade: Devolver o dicionário SKU -> clsAliquotasDTO da UF, usando o
'             cache quando ainda válido. A consulta em si (ADO/SQL nas abas de
'             impostos por produto e por UF) fica em mod_ImpostoRepository.
' Retorno: Nothing em caso de falha (o usuário já foi avisado).
' ==============================================================================
Public Function Pricing_ObterImpostosDaUF(ByVal siglaUF As String) As Object
    Dim caminhoDb As String
    Dim dict As Object
    Dim temExcecoes As Boolean

    If Not m_CacheImpostos Is Nothing Then
        If m_CacheUF = siglaUF And DateDiff("s", m_CacheHora, Now) < CACHE_SEGUNDOS Then
            Set Pricing_ObterImpostosDaUF = m_CacheImpostos
            Exit Function
        End If
    End If

    Call Pricing_LimparCacheImpostos

    caminhoDb = Config_ResolveDatabaseFile(DB_FILE_IMPOSTOS)
    On Error GoTo TrataErro
    Set dict = Impostos_CarregarPorUF(caminhoDb, siglaUF, temExcecoes)
    On Error GoTo 0

    If dict.Count = 0 Then
        MsgBox "A base de impostos não tem alíquotas para a UF " & siglaUF & "." & vbCrLf & caminhoDb, vbExclamation
        Exit Function
    End If

    Set m_CacheImpostos = dict
    m_CacheTemExcecoes = temExcecoes
    m_CacheUF = siglaUF
    m_CacheHora = Now
    Set Pricing_ObterImpostosDaUF = dict
    Exit Function

TrataErro:
    MsgBox "Erro ao consultar a base de impostos:" & vbCrLf & _
           caminhoDb & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function

' ==============================================================================
' Função Privada: CarregarFatoresOperacao
' Finalidade: Operação -> Fator (tabela CfOperações; ex. _OP_INTG = 1, _OP12 = 0,5)
' ==============================================================================
Private Function CarregarFatoresOperacao() As Object
    Dim dict As Object
    Dim lo As ListObject
    Dim arr As Variant
    Dim idxOp As Long, idxFator As Long
    Dim chave As String
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set CarregarFatoresOperacao = dict

    Set lo = LocalizarTabela(Tabela_Operacoes())
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxOp = IndiceListColumn(lo, Col_Operacao())
    idxFator = IndiceListColumn(lo, ColOp_Fator())
    If idxOp = 0 Or idxFator = 0 Then Exit Function

    arr = lo.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        chave = Util_Texto(arr(i, idxOp))
        If Len(chave) > 0 And Not dict.Exists(chave) Then dict.Add chave, Util_ToDouble(arr(i, idxFator))
    Next i
End Function

' ==============================================================================
' Função Privada: CarregarExcecoesImpostos  (TRANSIÇÃO)
' Finalidade: SKU -> Array(calcICMS, calcPisCofins, calcIPI, calcST)
'             (tabela local CfExcecoesImpostos; "Não" desliga o imposto).
'             Usada só enquanto a base de impostos não tiver as colunas calc_*.
' ==============================================================================
Private Function CarregarExcecoesImpostos() As Object
    Dim dict As Object
    Dim lo As ListObject
    Dim arr As Variant
    Dim idxCod As Long, idxICMS As Long, idxPC As Long, idxIPI As Long, idxST As Long
    Dim chave As String
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set CarregarExcecoesImpostos = dict

    Set lo = LocalizarTabela(TABELA_EXCECOES_IMPOSTOS)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxCod = IndiceListColumn(lo, Col_Codigo())
    idxICMS = IndiceListColumn(lo, ColExc_ICMS())
    idxPC = IndiceListColumn(lo, ColExc_PisCofins())
    idxIPI = IndiceListColumn(lo, ColExc_IPI())
    idxST = IndiceListColumn(lo, ColExc_ST())
    If idxCod = 0 Or idxICMS = 0 Or idxPC = 0 Or idxIPI = 0 Or idxST = 0 Then Exit Function

    arr = lo.DataBodyRange.Value2
    For i = 1 To UBound(arr, 1)
        chave = Util_ChaveSKU(arr(i, idxCod))
        If Len(chave) > 0 And Not dict.Exists(chave) Then
            dict.Add chave, Array(Not Util_EhNao(arr(i, idxICMS)), Not Util_EhNao(arr(i, idxPC)), _
                                  Not Util_EhNao(arr(i, idxIPI)), Not Util_EhNao(arr(i, idxST)))
        End If
    Next i
End Function

' ==============================================================================
' Utilitários
' ==============================================================================
Private Function LocalizarTabela(ByVal nomeTabela As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set LocalizarTabela = ws.ListObjects(nomeTabela)
        If Not LocalizarTabela Is Nothing Then Exit For
    Next ws
    On Error GoTo 0
End Function

Private Function IndiceListColumn(ByVal lo As ListObject, ByVal nomeColuna As String) As Long
    On Error Resume Next
    IndiceListColumn = lo.ListColumns(nomeColuna).Index
    On Error GoTo 0
End Function
