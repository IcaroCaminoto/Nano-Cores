' =========================================================
' Módulo de Classe: clsProduto
' Descrição: Gerencia os atributos de um produto e organiza regras de negocio para precificação
' =========================================================
Option Explicit

Private p_sku As String
Private p_uf_destino As String
Private p_icms_fora As Double
Private p_pis_fora As Double
Private p_cofins_fora As Double
Private p_ipi_fora As Double
Private p_st_fora As Double
Private p_icms_dentro As Double
Private p_pis_dentro As Double
Private p_cofins_dentro As Double
Private p_ipi_dentro As Double
Private p_st_dentro As Double

' =========================================================
' MÉTODOS E PROPRIEDADES
' =========================================================

Public Sub Init(ByVal codigoSKU As String, ByVal uf As String, ByVal vICMS As Double,ByVal vPIS As Double,ByVal vCOFINS As Double, ByVal vIPI As Double, ByVal vST As Double)
    ' Construtor
    p_sku = codigoSKU
    p_uf_destino = uf
    p_icms_fora = vICMS
    p_pis_fora = vPIS
    p_cofins_fora = vCOFINS
    p_ipi_fora = vIPI
    p_st_fora = vST
    
    Call CalcularImpostosPorDentro()
End Sub

Public Sub CalcularImpostosPorDentro()
    ' Calcula os impostos por dentro para cada SKU
    Dim index As Double
    index = 1 + p_ipi_fora + p_st_fora

    p_ipi_dentro = p_ipi_fora / index
    p_st_dentro = p_st_fora / index
    p_icms_dentro = p_icms_fora / index
    p_pis_dentro = (p_pis_fora * (1 - p_icms_fora)) / index
    p_cofins_dentro = (p_cofins_fora * (1 - p_icms_fora)) / index
End Sub

Public Sub InjetarImpostosPrecificacao(ByRef arrDados As Variant, ByVal linha_alvo As Long, ByVal colIPI As Long, ByVal colST As Long, ByVal colICMS As Long, ByVal colPIS As Long, ByVal colCOF As Long)
    ' Plota os impostos calculados pela PF CalcularImpostosPorDentro na precificacaoSKU
    arrDados(linha_alvo, colICMS) = p_icms_dentro
    arrDados(linha_alvo, colPIS) = p_pis_dentro
    arrDados(linha_alvo, colCOF) = p_cofins_dentro
    arrDados(linha_alvo, colIPI) = p_ipi_dentro
    arrDados(linha_alvo, colST) = p_st_dentro
End Sub