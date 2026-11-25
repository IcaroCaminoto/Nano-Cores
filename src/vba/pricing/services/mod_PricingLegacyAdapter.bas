Attribute VB_Name = "mod_PricingLegacyAdapter"
' ==============================================================================
' Módulo: mod_PricingLegacyAdapter
' Camada: Compatibility (Precificação)
' Finalidade: Substituir o mod_Engine legado do arquivo de produção mantendo
'             as mesmas assinaturas públicas, para que mod_Inputs e os botões
'             já existentes continuem funcionando sem alteração.
'   - MotorPrecificacaoBidirecional(CfPcfic, "Preço Unitário")  (mod_Inputs)
'   - AplicarIndicacaoVisual(CfPcfic, "Preço Unitário")         (mod_Inputs)
'   - BotaoRefresh_AtualizarMotor                               (botão)
' Instalação: REMOVER mod_Engine do projeto antes de importar este módulo
'             (nomes iguais em dois módulos geram "Nome ambíguo detectado").
' ==============================================================================

Option Explicit

' Usado pela importação (mod_Inputs), que grava só Código e Preço: o Custo (do
' pacote de custos ativo, spec-009) e o cadastro (Descrição, Qtd Padrão...) vêm
' das bases externas antes do cálculo. Sem pacote ativo, a importação não calcula.
Public Sub MotorPrecificacaoBidirecional(ByVal CfPcfic As ListObject, ByVal modoAlvo As String)
    If Not Pricing_PreencherCustos() Then Exit Sub
    If Not Pricing_PreencherCadastro() Then Exit Sub
    Call MotorPrecificacaoMultidirecional(ModoDoTextoLegado(modoAlvo))
End Sub

Public Sub AplicarIndicacaoVisual(ByVal CfPcfic As ListObject, ByVal modoAtual As String)
    Dim LoPrecificacao As New clsListState
    If LoPrecificacao.Init() Then
        Call Pricing_AplicarDestaqueAncora(LoPrecificacao, ModoDoTextoLegado(modoAtual))
    End If
End Sub

' Botão Atualizar: relê custo (pacote ativo), cadastro e impostos das bases
' (ignora o cache)
Public Sub BotaoRefresh_AtualizarMotor()
    Call Pricing_LimparCacheImpostos
    If Not Pricing_PreencherCustos() Then Exit Sub
    If Not Pricing_PreencherCadastro() Then Exit Sub
    Call Pricing_RecalcularTabela
End Sub

' ==============================================================================
' Função Privada: ModoDoTextoLegado
' Finalidade: Traduzir os textos de modo usados pelo código antigo.
' ==============================================================================
Private Function ModoDoTextoLegado(ByVal texto As String) As String
    Dim prefixo As String
    prefixo = "Pre" & ChrW(231) & "o "                  ' "Preço "

    ' Textos legados: "Margem Bruta", "Preço Caixa", "Preço Unitário", "Preço Final"
    Select Case True
        Case ModoValido(texto)
            ModoDoTextoLegado = texto
        Case StrComp(texto, "Margem Bruta", vbTextCompare) = 0
            ModoDoTextoLegado = MODO_MARGEM_BRUTA
        Case StrComp(texto, prefixo & "Caixa", vbTextCompare) = 0
            ModoDoTextoLegado = MODO_PRECO_FINAL_CAIXA
        Case StrComp(texto, prefixo & "Unit" & ChrW(225) & "rio", vbTextCompare) = 0, _
             StrComp(texto, prefixo & "Final", vbTextCompare) = 0
            ModoDoTextoLegado = MODO_PRECO_FINAL_UNITARIO
        Case Else
            ModoDoTextoLegado = MODO_MARGEM_BRUTA
    End Select
End Function
