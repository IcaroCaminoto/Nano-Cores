' ==============================================================================
' Módulo de Pasta de Trabalho: ThisWorkbook                          [spec-009]
' Finalidade: Antes de imprimir, garantir que o cabeçalho (quadrante direito) da
'             aba de precificação mostre o pacote de custos ativo, no formato
'             literal:
'                 Custos: [Pacote de Custos]
'             (a atualização também acontece ao confirmar o pacote no
'             frmPacoteCustos; aqui é a rede de segurança para a diretoria).
' Dependências: mod_PricingPacoteCustos (Pricing_SincronizarCabecalhoImpressao).
' Instalação: no Editor VBA, abra "ThisWorkbook" e cole a rotina abaixo. Se o
'   arquivo de produção JÁ tiver Workbook_BeforePrint, não crie outra: copie
'   apenas a linha "Call Pricing_SincronizarCabecalhoImpressao" para dentro dela.
' ==============================================================================
Option Explicit

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Call Pricing_SincronizarCabecalhoImpressao
End Sub
