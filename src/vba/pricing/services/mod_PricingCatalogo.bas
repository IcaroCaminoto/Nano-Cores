Attribute VB_Name = "mod_PricingCatalogo"
' ==============================================================================
' Módulo: mod_PricingCatalogo
' Camada: Excel Adapter (Precificação) - NÃO vai para a DLL
' Finalidade: Ligar a planilha (CfPcfic) e o formulário às camadas puras:
'   - Pricing_PreencherCadastro      : gatilho unitário (código digitado)
'   - Pricing_AbrirAdicionarGrupo    : macro do botão do painel (sys_config)
'   - Pricing_AdicionarGrupoProdutos : inclusão em lote por grupo
' Este módulo resolve caminhos (mod_Config), lê/escreve células, controla o
' estado do Application e mostra mensagens. Toda consulta SQL, cruzamento e
' montagem de matriz acontece em mod_ProdutoRepository, mod_ImpostoRepository
' e mod_CatalogoService.
' ==============================================================================

Option Explicit

Public Const MACRO_ADICIONAR_GRUPO As String = "Pricing_AbrirAdicionarGrupo"
Private Const NOME_FORM_GRUPO As String = "frmAdicionarGrupo"

' ==============================================================================
' Função: Pricing_PreencherCadastro
' Finalidade: Buscar na base de produtos (ADO/SQL) os códigos das linhas
'             indicadas e gravar Descrição, Qtd Padrão, Grupo, Peso Kg (e
'             Código de Barras, se a coluna existir) como valores.
'             Os impostos são gravados em seguida pelo motor (recalcular).
' Parâmetros:
'   - linhas: Collection de índices (1 = primeira linha do corpo). Omitido =
'             todas as linhas (botão Atualizar, importação, migração).
'   - avisarAusentes: mostra os códigos que não existem na base.
' Retorno: False se a base não pôde ser lida (o recálculo deve ser pulado).
' ==============================================================================
Public Function Pricing_PreencherCadastro(Optional ByVal linhas As Collection, _
                                         Optional ByVal avisarAusentes As Boolean = True) As Boolean
    Dim estado As New clsListState
    Dim arr As Variant
    Dim colunas As Variant
    Dim indices() As Long
    Dim idxCod As Long
    Dim alvo As Collection
    Dim codigos As New Collection
    Dim pedidos As Object
    Dim produtos As Object
    Dim ausentes As New Collection
    Dim dto As clsProdutoDTO
    Dim linha As Variant
    Dim chave As String
    Dim valor As Variant
    Dim caminhoDb As String
    Dim i As Long, j As Long
    Dim prevEvents As Boolean, prevScreen As Boolean

    If Not estado.Init() Then Exit Function
    If Not estado.TemDados Then Pricing_PreencherCadastro = True: Exit Function

    idxCod = estado.IndiceColuna(Col_Codigo())
    If idxCod = 0 Then
        MsgBox "A coluna """ & Col_Codigo() & """ não existe na tabela " & TABELA_PRECIFICACAO & ".", vbCritical
        Exit Function
    End If

    ' Colunas cadastrais que existem nesta CfPcfic
    colunas = Catalogo_ColunasCadastro()
    ReDim indices(LBound(colunas) To UBound(colunas))
    For j = LBound(colunas) To UBound(colunas)
        indices(j) = estado.IndiceColuna(CStr(colunas(j)))
    Next j

    arr = estado.DataBodyPadrao.Value2
    If Not IsArray(arr) Then Exit Function

    ' Linhas a processar
    If linhas Is Nothing Then
        Set alvo = New Collection
        For i = 1 To UBound(arr, 1)
            alvo.Add i
        Next i
    Else
        Set alvo = linhas
    End If

    ' Códigos distintos a consultar
    Set pedidos = CreateObject("Scripting.Dictionary")
    pedidos.CompareMode = vbTextCompare
    For Each linha In alvo
        If linha >= 1 And linha <= UBound(arr, 1) Then
            chave = Util_ChaveSKU(arr(linha, idxCod))
            If Len(chave) > 0 And Not pedidos.Exists(chave) Then
                pedidos.Add chave, True
                codigos.Add chave
            End If
        End If
    Next linha

    ' Consulta (camada pura)
    caminhoDb = Config_ResolveDatabaseFile(DB_FILE_PRODUTOS)
    On Error GoTo TrataErro
    Set produtos = Produtos_BuscarPorCodigos(caminhoDb, codigos)
    On Error GoTo 0

    ' Atualiza a matriz em memória
    For Each linha In alvo
        If linha >= 1 And linha <= UBound(arr, 1) Then
            chave = Util_ChaveSKU(arr(linha, idxCod))
            Set dto = Nothing
            If Len(chave) > 0 Then
                If produtos.Exists(chave) Then
                    Set dto = produtos(chave)
                Else
                    ausentes.Add chave
                End If
            End If
            For j = LBound(colunas) To UBound(colunas)
                If indices(j) > 0 Then
                    If dto Is Nothing Then
                        arr(linha, indices(j)) = Empty
                    ElseIf Catalogo_ValorDaColuna(dto, CStr(colunas(j)), valor) Then
                        arr(linha, indices(j)) = valor
                    End If
                End If
            Next j
        End If
    Next linha

    ' Grava só as colunas cadastrais (eventos desligados: não redispara o Change)
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Call estado.InjetarColunas(arr, colunas)
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents

    If avisarAusentes And ausentes.Count > 0 Then
        MsgBox "C" & ChrW(243) & "digo(s) n" & ChrW(227) & "o encontrado(s) na base de produtos:" & vbCrLf & _
               Util_JuntarLista(ausentes) & vbCrLf & vbCrLf & caminhoDb, vbExclamation
    End If

    Pricing_PreencherCadastro = True
    Exit Function

TrataErro:
    MsgBox "Erro ao consultar a base de produtos:" & vbCrLf & _
           caminhoDb & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function

' ==============================================================================
' Função: Pricing_ListarGruposProdutos
' Finalidade: Grupos distintos da base de produtos (para o formulário).
' Retorno: Array de String ou Empty (com aviso ao usuário).
' ==============================================================================
Public Function Pricing_ListarGruposProdutos() As Variant
    Dim caminhoDb As String
    Dim grupos As Variant

    caminhoDb = Config_ResolveDatabaseFile(DB_FILE_PRODUTOS)
    On Error GoTo TrataErro
    grupos = Produtos_ListarGrupos(caminhoDb)
    On Error GoTo 0

    If Not IsArray(grupos) Then
        MsgBox "A base de produtos n" & ChrW(227) & "o tem grupos cadastrados (coluna """ & _
               DbProd_Grupo() & """)." & vbCrLf & caminhoDb, vbExclamation
        Exit Function
    End If
    Pricing_ListarGruposProdutos = grupos
    Exit Function

TrataErro:
    MsgBox "Erro ao consultar a base de produtos:" & vbCrLf & _
           caminhoDb & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function

' ==============================================================================
' Procedimento: Pricing_AbrirAdicionarGrupo
' Finalidade: Macro ligada ao botão "Adicionar Grupo de Produtos" do painel
'             (tabela TabelaRotinas da aba sys_config).
' Desacoplamento: o formulário só recebe a lista e devolve a escolha; não
'             consulta base nem escreve na planilha. O formulário é aberto por
'             nome (late binding) para o projeto compilar mesmo sem ele.
' ==============================================================================
Public Sub Pricing_AbrirAdicionarGrupo()
    Dim grupos As Variant
    Dim frm As Object
    Dim grupo As String

    grupos = Pricing_ListarGruposProdutos()
    If Not IsArray(grupos) Then Exit Sub

    On Error Resume Next
    Set frm = VBA.UserForms.Add(NOME_FORM_GRUPO)
    On Error GoTo 0
    If frm Is Nothing Then
        MsgBox "Formul" & ChrW(225) & "rio """ & NOME_FORM_GRUPO & """ n" & ChrW(227) & "o encontrado no projeto VBA.", vbCritical
        Exit Sub
    End If

    frm.DefinirGrupos grupos
    frm.Show vbModal
    If frm.Confirmado Then grupo = frm.GrupoSelecionado
    Unload frm
    Set frm = Nothing

    If Len(grupo) > 0 Then Call Pricing_AdicionarGrupoProdutos(grupo)
End Sub

' ==============================================================================
' Procedimento: Pricing_AdicionarGrupoProdutos
' Finalidade: Inserir na CfPcfic todos os produtos ativos de um grupo, já com
'             dados cadastrais e impostos da UF, e recalcular a tabela.
' Fluxo:
'   1. Produtos do grupo (SQL)            -> mod_ProdutoRepository
'   2. Impostos da UF (SQL, com cache)    -> Pricing_ObterImpostosDaUF
'   3. Cruzamento e filtro de repetidos   -> mod_CatalogoService
'   4. Linhas novas + gravação em bloco   -> clsListState
'      (Código, cadastro e Operação padrão OPERACAO_PADRAO_LOTE; uma linha
'       vazia reaproveitada mantém a Operação que já tiver)
'   5. Impostos e preços                  -> Pricing_RecalcularTabela
' ==============================================================================
Public Sub Pricing_AdicionarGrupoProdutos(ByVal grupo As String)
    Dim estado As New clsListState
    Dim caminhoDb As String
    Dim produtos As Collection
    Dim novos As Collection
    Dim semImpostos As Collection
    Dim impostosUF As Object
    Dim siglaUF As String
    Dim repetidos As Long
    Dim colunas As Variant
    Dim bloco As Variant
    Dim arr As Variant
    Dim indices() As Long
    Dim primeira As Long
    Dim i As Long, j As Long
    Dim resumo As String
    Dim prevEvents As Boolean, prevScreen As Boolean
    Dim prevCalc As XlCalculation

    If Not estado.Init() Then
        MsgBox "Tabela de precifica" & ChrW(231) & ChrW(227) & "o """ & TABELA_PRECIFICACAO & """ n" & ChrW(227) & "o encontrada.", vbCritical
        Exit Sub
    End If

    ' ---- 1. Produtos do grupo -------------------------------------------------
    caminhoDb = Config_ResolveDatabaseFile(DB_FILE_PRODUTOS)
    On Error GoTo ErroBase
    Set produtos = Produtos_BuscarPorGrupo(caminhoDb, grupo)
    On Error GoTo 0

    If produtos.Count = 0 Then
        MsgBox "Nenhum produto ativo no grupo """ & grupo & """.", vbInformation
        Exit Sub
    End If

    ' ---- 2 e 3. Impostos da UF e cruzamento ------------------------------------
    siglaUF = Pricing_ObterSiglaUF()
    If Len(siglaUF) > 0 Then Set impostosUF = Pricing_ObterImpostosDaUF(siglaUF)
    Call Catalogo_CruzarImpostos(produtos, impostosUF)

    Set novos = Catalogo_FiltrarNovos(produtos, estado.ChavesExistentes(), repetidos)
    If novos.Count = 0 Then
        MsgBox "Todos os " & produtos.Count & " produtos do grupo """ & grupo & """ j" & ChrW(225) & _
               " est" & ChrW(227) & "o na tabela.", vbInformation
        Exit Sub
    End If
    Set semImpostos = Catalogo_CodigosSemImpostos(novos)

    ' ---- 4. Linhas novas e gravação em bloco -----------------------------------
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    On Error GoTo ErroPlanilha
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Adicionando " & novos.Count & " produtos..."

    primeira = estado.AdicionarLinhas(novos.Count)

    colunas = Catalogo_ColunasLote()
    ReDim indices(LBound(colunas) To UBound(colunas))
    For j = LBound(colunas) To UBound(colunas)
        indices(j) = estado.IndiceColuna(CStr(colunas(j)))
    Next j

    bloco = Catalogo_MontarMatriz(novos, colunas)
    arr = estado.DataBodyPadrao.Value2
    For i = 1 To novos.Count
        For j = LBound(colunas) To UBound(colunas)
            If indices(j) > 0 Then
                If colunas(j) = Col_Operacao() And Len(Util_Texto(arr(primeira + i - 1, indices(j)))) > 0 Then
                    ' linha reaproveitada que já tinha Operação escolhida: mantém
                Else
                    arr(primeira + i - 1, indices(j)) = bloco(i, j - LBound(colunas) + 1)
                End If
            End If
        Next j
    Next i
    Call estado.InjetarColunas(arr, colunas)

    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    On Error GoTo 0

    ' ---- 5. Impostos e preços --------------------------------------------------
    Call Pricing_RecalcularTabela

    resumo = novos.Count & " produto(s) do grupo """ & grupo & """ adicionado(s) com Opera" & _
             ChrW(231) & ChrW(227) & "o " & OPERACAO_PADRAO_LOTE & "."
    If repetidos > 0 Then resumo = resumo & vbCrLf & repetidos & " j" & ChrW(225) & " estavam na tabela e foram ignorados."
    If semImpostos.Count > 0 Then
        resumo = resumo & vbCrLf & vbCrLf & "Sem impostos na base para a UF " & siglaUF & ":" & vbCrLf & _
                 Util_JuntarLista(semImpostos)
    End If
    MsgBox resumo, IIf(semImpostos.Count > 0, vbExclamation, vbInformation)
    Exit Sub

ErroBase:
    MsgBox "Erro ao consultar a base de produtos:" & vbCrLf & _
           caminhoDb & vbCrLf & vbCrLf & Err.Description, vbCritical
    Exit Sub

ErroPlanilha:
    MsgBox "Erro ao inserir os produtos na tabela: " & Err.Description, vbCritical, "Erro " & Err.Number
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
End Sub
