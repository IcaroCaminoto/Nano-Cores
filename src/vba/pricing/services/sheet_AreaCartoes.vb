' ==============================================================================
' Módulo de Planilha: AreaCartoes (aba de parâmetros)
' Finalidade: Recalcular a precificação quando o estado de destino (REGIÃO,
'             célula E9) muda. Antes, o ICMS acompanhava a UF por fórmula
'             (Impostos!G3 -> dUF); agora o motor consulta a base por UF.
' Dependências: mod_PricingSchema, mod_Multidirectional_Engine
' Instalação: colar no módulo da aba "AreaCartoes" (hoje vazio em produção).
' ==============================================================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range(CELULA_UF)) Is Nothing Then Exit Sub

    Call Pricing_LimparCacheImpostos
    Call Pricing_RecalcularTabela
End Sub
