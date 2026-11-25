Attribute VB_Name = "mod_PricingPacoteCustos"
' ==============================================================================
' Módulo: mod_PricingPacoteCustos
' Camada: Excel Adapter (Precificação) - NÃO vai para a DLL        [spec-009]
' Finalidade: Ligar a planilha e o formulário ao serviço puro de pacotes de
'             custos (mod_Pricing_CostPackageService):
'   - ObterPacoteCustosAtivo / SalvarPacoteCustosAtivo : estado no Named Range
'     oculto "cfg_PacoteCustosAtivo" (vai junto com o arquivo ao salvar)
'   - AtualizarCabecalhoImpressao                      : "Custos: [pacote]" no
'     cabeçalho DIREITO (o central já mostra o nome do arquivo, &F)
'   - Pricing_AbrirSelecionarPacoteCustos              : macro do botão do painel
'   - Pricing_ObterPacoteOuAvisar                      : guarda de pré-condição
'   - Pricing_PreencherCustos                          : código digitado na CfPcfic,
'     importação e botão Atualizar; o lote usa as mesmas peças em
'     mod_PricingCatalogo
'   - Pricing_SincronizarCabecalhoImpressao            : chamada pelo BeforePrint
' O que precisa de Range/Worksheet/Application/MsgBox fica aqui; a consulta SQL,
' a regra de prefixo e o saneamento de custo ficam no serviço puro.
' ==============================================================================

Option Explicit

Public Const MACRO_SELECIONAR_PACOTE As String = "Pricing_AbrirSelecionarPacoteCustos"
Public Const NOME_PACOTE_ATIVO As String = "cfg_PacoteCustosAtivo"
Private Const NOME_FORM_PACOTE As String = "frmPacoteCustos"

' ==============================================================================
' Função: ObterPacoteCustosAtivo
' Finalidade: Nome do pacote escolhido pelo usuário ("" se nenhum).
' ==============================================================================
Public Function ObterPacoteCustosAtivo() As String
    Dim nm As Name
    Dim v As Variant

    On Error Resume Next
    Set nm = ThisWorkbook.Names(NOME_PACOTE_ATIVO)
    If Not nm Is Nothing Then v = Application.Evaluate(nm.RefersTo)
    On Error GoTo 0

    If Not IsError(v) Then ObterPacoteCustosAtivo = Util_Texto(v)
End Function

' ==============================================================================
' Procedimento: SalvarPacoteCustosAtivo
' Finalidade: Gravar o pacote ativo no Named Range oculto (criado na primeira
'             vez). O valor é um texto literal dentro do próprio nome.
' Erro: ERR_PACOTE_LAYOUT_INVALIDO se o nome não for Custo_MM_* / Custo_EX_*.
' ==============================================================================
Public Sub SalvarPacoteCustosAtivo(ByVal nomePacote As String)
    nomePacote = Trim$(nomePacote)
    If Not Pacote_EhValido(nomePacote) Then
        Err.Raise ERR_PACOTE_LAYOUT_INVALIDO, "mod_PricingPacoteCustos", _
                  "Layout de pacote de custos n" & ChrW(227) & "o reconhecido: " & nomePacote
    End If

    ThisWorkbook.Names.Add Name:=NOME_PACOTE_ATIVO, _
                           RefersTo:="=""" & Replace(nomePacote, """", """""") & """", _
                           Visible:=False
End Sub

' ==============================================================================
' Procedimento: AtualizarCabecalhoImpressao
' Finalidade: Gravar "Custos: [pacote]" no cabeçalho DIREITO da aba de
'             precificação. O quadrante central da aba já é usado pelo nome do
'             arquivo (&F, 16 pt sublinhado) e o esquerdo fica como está; no
'             direito só a linha "Custos: " é trocada (ver Pacote_MontarCabecalho).
' Erro: sobe o erro do PageSetup (ex.: sem impressora instalada).
' ==============================================================================
Public Sub AtualizarCabecalhoImpressao(ByVal wsTarget As Worksheet, ByVal nomePacote As String)
    Dim texto As String

    If wsTarget Is Nothing Then Exit Sub
    If Len(Trim$(nomePacote)) = 0 Then Exit Sub

    With wsTarget.PageSetup
        texto = Pacote_MontarCabecalho(.RightHeader, nomePacote)
        If .RightHeader <> texto Then .RightHeader = texto
    End With
End Sub

' ==============================================================================
' Procedimento: Pricing_SincronizarCabecalhoImpressao
' Finalidade: Chamado por Workbook_BeforePrint: garante o cabeçalho com o pacote
'             ativo. Nunca bloqueia a impressão (qualquer erro é ignorado).
' ==============================================================================
Public Sub Pricing_SincronizarCabecalhoImpressao()
    Dim estado As New clsListState
    Dim pacote As String

    pacote = ObterPacoteCustosAtivo()
    If Len(pacote) = 0 Then Exit Sub
    If Not estado.Init() Then Exit Sub

    On Error Resume Next
    Call AtualizarCabecalhoImpressao(estado.Planilha, pacote)
    On Error GoTo 0
End Sub

' ==============================================================================
' Função: Pricing_ObterPacoteOuAvisar
' Finalidade: Guarda de pré-condição da inclusão de produtos: devolve o pacote
'             ativo ou, se não houver, orienta o usuário e devolve "".
' ==============================================================================
Public Function Pricing_ObterPacoteOuAvisar() As String
    Dim pacote As String

    pacote = ObterPacoteCustosAtivo()
    If Len(pacote) = 0 Then
        MsgBox "Nenhum pacote de custos foi selecionado." & vbCrLf & vbCrLf & _
               "Abra o painel (Op" & ChrW(231) & ChrW(245) & "es), clique em ""Selecionar Pacote de Custos"" " & _
               "e escolha o fechamento antes de continuar.", vbExclamation
    End If
    Pricing_ObterPacoteOuAvisar = pacote
End Function

' ==============================================================================
' Função: Pricing_ObterCustosDoPacote
' Finalidade: Consultar os custos no pacote (serviço puro) e traduzir os erros
'             em mensagens para o usuário.
' Retorno: Dictionary SKU -> custo, ou Nothing se a consulta falhou (o chamador
'          deve interromper a inclusão).
' ==============================================================================
Public Function Pricing_ObterCustosDoPacote(ByVal pacote As String, ByVal skus As Collection, _
                                            ByRef outAlertas As Collection) As Object
    Dim caminhoDb As String
    Dim msg As String

    caminhoDb = Config_ResolveDatabaseFile(DB_FILE_CUSTOS)
    On Error GoTo TrataErro
    Set Pricing_ObterCustosDoPacote = BuscarCustosPorPacote(caminhoDb, pacote, skus, outAlertas)
    Exit Function

TrataErro:
    Select Case Err.Number
        Case ERR_PACOTE_INEXISTENTE
            msg = "O pacote de custos """ & pacote & """ n" & ChrW(227) & "o existe mais em " & DB_FILE_CUSTOS & "." & _
                  vbCrLf & vbCrLf & "Abra o painel (Op" & ChrW(231) & ChrW(245) & "es), clique em " & _
                  """Selecionar Pacote de Custos"" e escolha outro."
        Case ERR_PACOTE_LAYOUT_INVALIDO
            msg = Err.Description & vbCrLf & vbCrLf & "Selecione um pacote Custo_MM_* ou Custo_EX_*."
        Case Else
            msg = "Erro ao consultar a base de custos:" & vbCrLf & caminhoDb & vbCrLf & vbCrLf & Err.Description
    End Select
    MsgBox msg, vbCritical
End Function

' ==============================================================================
' Função: Pricing_PreencherCustos
' Finalidade: Gravar na coluna Custo o valor do pacote ativo (como valor fixo)
'             nas linhas indicadas ou, se omitidas, em TODAS (importação e botão
'             Atualizar).
'   - Sem pacote ativo, ou se a consulta falhar: o usuário é orientado e nada é
'     gravado. Com bloquearInclusao (código digitado), os códigos digitados são
'     apagados também, para nada entrar na CfPcfic sem custo.
'   - Código sem custo neste pacote: Custo fica vazio (nunca sobra o custo de
'     outro pacote) e o código é avisado.
'   - Linha sem código (apagado): Custo é limpo.
' Parâmetros:
'   - linhas: Collection de índices (1 = primeira linha do corpo); omitido = todas
'   - bloquearInclusao: True só no gatilho do código digitado. Em importação e
'     Atualizar fica False: nunca apaga códigos de uma tabela já preenchida.
' Retorno: False se não gravou (o recálculo deve ser pulado).
' ==============================================================================
Public Function Pricing_PreencherCustos(Optional ByVal linhas As Collection, _
                                        Optional ByVal bloquearInclusao As Boolean = False) As Boolean
    Dim estado As New clsListState
    Dim arr As Variant
    Dim idxCod As Long, idxCusto As Long
    Dim alvo As Collection
    Dim codigos As New Collection
    Dim pedidos As Object
    Dim custos As Object
    Dim alertas As New Collection
    Dim ausentes As New Collection
    Dim pacote As String
    Dim chave As String
    Dim msg As String
    Dim linha As Variant
    Dim i As Long
    Dim errDesc As String
    Dim prevEvents As Boolean, prevScreen As Boolean

    If Not estado.Init() Then Exit Function
    If Not estado.TemDados Then Pricing_PreencherCustos = True: Exit Function

    idxCod = estado.IndiceColuna(Col_Codigo())
    idxCusto = estado.IndiceColuna(Col_Custo())
    If idxCod = 0 Or idxCusto = 0 Then
        MsgBox "As colunas """ & Col_Codigo() & """ e """ & Col_Custo() & """ precisam existir na tabela " & _
               TABELA_PRECIFICACAO & ".", vbCritical
        Exit Function
    End If

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

    ' Códigos distintos nas linhas
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

    ' Pré-condição: pacote ativo e consulta possível
    If codigos.Count > 0 Then
        pacote = Pricing_ObterPacoteOuAvisar()
        If Len(pacote) = 0 Then
            If bloquearInclusao Then Call BloquearInclusao(estado, arr, alvo, idxCod)
            Exit Function
        End If
        Set custos = Pricing_ObterCustosDoPacote(pacote, codigos, alertas)
        If custos Is Nothing Then
            If bloquearInclusao Then Call BloquearInclusao(estado, arr, alvo, idxCod)
            Exit Function
        End If
    End If

    ' Atualiza a coluna Custo na matriz em memória
    For Each linha In alvo
        If linha >= 1 And linha <= UBound(arr, 1) Then
            chave = Util_ChaveSKU(arr(linha, idxCod))
            arr(linha, idxCusto) = Empty
            If Len(chave) > 0 Then
                If custos.Exists(chave) Then
                    arr(linha, idxCusto) = custos(chave)
                ElseIf pedidos(chave) Then
                    ausentes.Add chave
                    pedidos(chave) = False
                End If
            End If
        End If
    Next linha

    ' Grava só a coluna Custo (eventos desligados: não redispara o Change)
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    On Error GoTo ErroEscrita
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Call estado.InjetarColunas(arr, Array(Col_Custo()))
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    On Error GoTo 0

    If ausentes.Count > 0 Or alertas.Count > 0 Then
        msg = "Pacote de custos: " & pacote
        If ausentes.Count > 0 Then
            msg = msg & vbCrLf & vbCrLf & "C" & ChrW(243) & "digo(s) sem custo neste pacote:" & vbCrLf & _
                  Util_JuntarLista(ausentes)
        End If
        If alertas.Count > 0 Then
            msg = msg & vbCrLf & vbCrLf & "Custos inv" & ChrW(225) & "lidos na base:" & vbCrLf & _
                  Util_JuntarLista(alertas)
        End If
        MsgBox msg, vbExclamation
    End If

    Pricing_PreencherCustos = True
    Exit Function

ErroEscrita:
    errDesc = Err.Description
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    MsgBox "Erro ao gravar o custo na tabela: " & errDesc, vbCritical
End Function

' ==============================================================================
' Procedimento Privado: BloquearInclusao
' Finalidade: Desfazer a digitação do código (apaga o Código das linhas) quando
'             não há como buscar o custo, para nada entrar na CfPcfic sem custo.
' ==============================================================================
Private Sub BloquearInclusao(ByVal estado As clsListState, ByRef arr As Variant, _
                             ByVal linhas As Collection, ByVal idxCod As Long)
    Dim linha As Variant
    Dim prevEvents As Boolean, prevScreen As Boolean

    For Each linha In linhas
        If linha >= 1 And linha <= UBound(arr, 1) Then arr(linha, idxCod) = Empty
    Next linha

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error Resume Next
    Call estado.InjetarColunas(arr, Array(Col_Codigo()))
    On Error GoTo 0
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
End Sub

' ==============================================================================
' Função: Pricing_ListarPacotesCustos
' Finalidade: Pacotes válidos da base de custos (para o formulário).
' Retorno: Array de String ou Empty (com aviso ao usuário).
' ==============================================================================
Public Function Pricing_ListarPacotesCustos() As Variant
    Dim caminhoDb As String
    Dim pacotes As Collection
    Dim lista() As String
    Dim i As Long

    caminhoDb = Config_ResolveDatabaseFile(DB_FILE_CUSTOS)
    On Error GoTo TrataErro
    Set pacotes = ListarPacotesCustosDisponiveis(caminhoDb)
    On Error GoTo 0

    If pacotes.Count = 0 Then
        MsgBox "A base de custos n" & ChrW(227) & "o tem pacotes (abas " & DB_ABA_CUSTOS_PREFIXO_MM & "* ou " & _
               DB_ABA_CUSTOS_PREFIXO_EX & "*)." & vbCrLf & "Execute um fechamento de custos antes." & _
               vbCrLf & caminhoDb, vbExclamation
        Exit Function
    End If

    ReDim lista(0 To pacotes.Count - 1)
    For i = 1 To pacotes.Count
        lista(i - 1) = CStr(pacotes(i))
    Next i
    Pricing_ListarPacotesCustos = lista
    Exit Function

TrataErro:
    MsgBox "Erro ao consultar a base de custos:" & vbCrLf & _
           caminhoDb & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function

' ==============================================================================
' Procedimento: Pricing_AbrirSelecionarPacoteCustos
' Finalidade: Macro ligada ao botão "Selecionar Pacote de Custos" do painel
'             (tabela TabelaRotinas da aba sys_config).
' Desacoplamento: o formulário só recebe a lista (e o pacote ativo, para
'             pré-selecionar) e devolve a escolha; não consulta base nem grava
'             estado. É aberto por nome (late binding), como o frmAdicionarGrupo.
' ==============================================================================
Public Sub Pricing_AbrirSelecionarPacoteCustos()
    Dim pacotes As Variant
    Dim frm As Object
    Dim escolhido As String

    pacotes = Pricing_ListarPacotesCustos()
    If Not IsArray(pacotes) Then Exit Sub

    On Error Resume Next
    Set frm = VBA.UserForms.Add(NOME_FORM_PACOTE)
    On Error GoTo 0
    If frm Is Nothing Then
        MsgBox "Formul" & ChrW(225) & "rio """ & NOME_FORM_PACOTE & """ n" & ChrW(227) & "o encontrado no projeto VBA.", vbCritical
        Exit Sub
    End If

    frm.DefinirPacotes pacotes, ObterPacoteCustosAtivo()
    frm.Show vbModal
    If frm.Confirmado Then escolhido = frm.PacoteSelecionado
    Unload frm
    Set frm = Nothing

    If Len(escolhido) > 0 Then Call Pricing_DefinirPacoteCustos(escolhido)
End Sub

' ==============================================================================
' Procedimento: Pricing_DefinirPacoteCustos
' Finalidade: Confirmar a escolha: salva o pacote ativo e atualiza o cabeçalho
'             de impressão da aba de precificação.
' ==============================================================================
Public Sub Pricing_DefinirPacoteCustos(ByVal nomePacote As String)
    Dim estado As New clsListState

    On Error GoTo ErroSalvar
    Call SalvarPacoteCustosAtivo(nomePacote)
    Application.StatusBar = "Pacote de custos ativo: " & nomePacote

    On Error GoTo ErroCabecalho
    If estado.Init() Then Call AtualizarCabecalhoImpressao(estado.Planilha, nomePacote)
    Exit Sub

ErroSalvar:
    MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel salvar o pacote de custos: " & Err.Description, vbCritical
    Exit Sub

ErroCabecalho:
    MsgBox "Pacote de custos selecionado, mas o cabe" & ChrW(231) & "alho de impress" & ChrW(227) & "o n" & ChrW(227) & _
           "o p" & ChrW(244) & "de ser atualizado agora (" & Err.Description & ")." & vbCrLf & _
           "Ele ser" & ChrW(225) & " atualizado automaticamente ao imprimir.", vbExclamation
End Sub
