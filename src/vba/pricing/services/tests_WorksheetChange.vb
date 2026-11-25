' ==============================================================================
' Módulo de Planilha: wsPcfic (aba "TabelaCustos", tabela CfPcfic)
' Finalidade: Capturar a intenção do usuário e disparar o motor multidirecional:
'   - edição em coluna-âncora (Margem, Preço Final Cx/Un, Preço Base Cx/Un)
'     -> recalcula com aquela âncora;
'   - código digitado/colado -> grava o Custo do pacote de custos ativo
'     (spec-009; sem pacote, o código é apagado e o usuário é orientado) e busca
'     o cadastro do produto na base externa (ADO/SQL), gravando Descrição,
'     Qtd Padrão, Grupo e Peso Kg como valores;
'   - edição em insumo (Código, Qtd, Custo, Frete, Acordo, Comissão, Outros,
'     Operação) -> recalcula mantendo a âncora atual (impostos não são mais
'     fórmulas, então precisam ser recalculados pelo motor).
' Dependências: clsListState, mod_PricingSchema, mod_Multidirectional_Engine,
'   mod_PricingCatalogo, mod_PricingPacoteCustos;
'   Worksheet_Activate/Deactivate usam mod_Init (InicializarTodasTabelas,
'   AlternarTela) do arquivo de produção.
' Instalação: substituir TODO o código do módulo da aba wsPcfic por este.
' ==============================================================================
Option Explicit

Private Sub Worksheet_Activate()
    Call InicializarTodasTabelas
    On Error Resume Next
    Me.Shapes("botaoMaximizar").Visible = msoTrue
    On Error GoTo 0
End Sub

Private Sub Worksheet_Deactivate()
    If Application.DisplayFullScreen = True Then
        Call AlternarTela
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim EstadoTabela As New clsListState
    Dim alvo As Range
    Dim area As Range
    Dim coluna As Range
    Dim celula As Range
    Dim linhasCodigo As New Collection
    Dim nomeColuna As String
    Dim modo As String
    Dim modoAlvo As String
    Dim modosVistos As String
    Dim qtdAncoras As Long
    Dim temInsumo As Boolean

    If Not EstadoTabela.Init() Then Exit Sub
    If Not EstadoTabela.Planilha Is Me Then Exit Sub
    If Not EstadoTabela.TemDados Then Exit Sub

    Set alvo = Intersect(Target, EstadoTabela.DataBodyPadrao)
    If alvo Is Nothing Then Exit Sub

    ' --------------------------------------------------------------------------
    ' Leitura de intenção: o cabeçalho de cada coluna editada define o papel
    ' --------------------------------------------------------------------------
    For Each area In alvo.Areas
        For Each coluna In area.Columns
            nomeColuna = EstadoTabela.NomeColunaDaPlanilha(coluna.Column)
            modo = ModoPorColuna(nomeColuna)

            If Len(modo) > 0 Then
                If Not SomenteNumerosOuVazio(coluna) Then
                    MsgBox "A coluna """ & nomeColuna & """ aceita apenas números." & vbCrLf & _
                           "A edição foi desfeita.", vbExclamation
                    Call DesfazerEdicao
                    Exit Sub
                End If
                If InStr(1, modosVistos, "|" & modo & "|") = 0 Then
                    modosVistos = modosVistos & "|" & modo & "|"
                    qtdAncoras = qtdAncoras + 1
                    modoAlvo = modo
                End If
            ElseIf ColunaEhInsumo(nomeColuna) Then
                temInsumo = True
                If nomeColuna = Col_Codigo() Then
                    For Each celula In coluna.Cells
                        linhasCodigo.Add celula.Row - EstadoTabela.DataBodyPadrao.Row + 1
                    Next celula
                End If
            End If
        Next coluna
    Next area

    If qtdAncoras = 0 And Not temInsumo Then Exit Sub

    If qtdAncoras > 1 Then
        MsgBox "Mais de uma coluna-âncora foi alterada de uma vez." & vbCrLf & _
               "O recálculo vai manter a âncora atual (coluna em negrito).", vbExclamation
        modoAlvo = vbNullString
    End If

    ' Código novo: primeiro o custo do pacote ativo (sem pacote a inclusão é
    ' bloqueada), depois o cadastro da base externa antes do cálculo (Qtd Padrão
    ' entra no preço unitário). Se uma base não puder ser lida, não recalcula.
    If linhasCodigo.Count > 0 Then
        If Not Pricing_PreencherCustos(linhasCodigo, True) Then Exit Sub
        If Not Pricing_PreencherCadastro(linhasCodigo) Then Exit Sub
    End If

    ' Vazio = âncora atual (detectada pelo negrito)
    Call Pricing_RecalcularTabela(modoAlvo)
End Sub

' ==============================================================================
' Função Privada: SomenteNumerosOuVazio
' ==============================================================================
Private Function SomenteNumerosOuVazio(ByVal rng As Range) As Boolean
    Dim arr As Variant
    Dim v As Variant

    arr = rng.Value2
    If IsArray(arr) Then
        For Each v In arr
            If Not ValorAceito(v) Then Exit Function
        Next v
    Else
        If Not ValorAceito(arr) Then Exit Function
    End If
    SomenteNumerosOuVazio = True
End Function

Private Function ValorAceito(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then ValorAceito = True: Exit Function
    If VarType(v) = vbString Then
        ValorAceito = (Len(Trim$(v)) = 0)
        Exit Function
    End If
    ValorAceito = IsNumeric(v)
End Function

' ==============================================================================
' Procedimento Privado: DesfazerEdicao
' ==============================================================================
Private Sub DesfazerEdicao()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    Application.EnableEvents = prevEvents
End Sub
