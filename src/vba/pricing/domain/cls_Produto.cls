VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' Módulo de Classe: clsProduto
' Camada: Domain (Precificação)
' Finalidade: Regras tributárias de um SKU. Recebe as alíquotas "por fora" da
'             base Cerulean_DB_Impostos (via ADO) e devolve as alíquotas
'             "por dentro" exibidas na CfPcfic, já ajustadas pela Operação
'             (fator) e pelas exceções (CfExcecoesImpostos).
' Paridade: reproduz exatamente as fórmulas legadas da tabela CfImpostos e das
'             colunas ICMS / Pis/Cofins / IPI / ICMS ST da CfPcfic:
'   IPI_d  = IPIF / (1 + IPIF + STF)
'   ST_d   = STF  / (1 + IPIF + STF)
'   D      = 1 + IPI_d / (1 - IPI_d) + ST_d / (1 - ST_d)
'   ICMS_d = ICMS / D
'   PC_d   = PisCofins * (1 - ICMS) / D
'   Efetivo = X_d * Fator(Operação) * (Calc_X = "Não" ? 0 : 1)
' Unidades: todas as alíquotas em fração (0,12 = 12%).
' ==============================================================================
Option Explicit

Private p_sku As String

' Alíquotas "por fora" (entrada)
Private p_icms As Double
Private p_pis_cofins As Double
Private p_ipi_fora As Double
Private p_st_fora As Double

' Alíquotas "por dentro" brutas (antes de Operação/exceções)
Private p_icms_dentro As Double
Private p_pis_cofins_dentro As Double
Private p_ipi_dentro As Double
Private p_st_dentro As Double

' Multiplicadores da Operação e das exceções
Private p_fator As Double
Private p_calc_icms As Boolean
Private p_calc_pis_cofins As Boolean
Private p_calc_ipi As Boolean
Private p_calc_st As Boolean

' ==============================================================================
' Construtor
' ==============================================================================
Public Sub Init(ByVal codigoSKU As String, ByVal aliqICMS As Double, ByVal aliqPisCofins As Double, _
                ByVal aliqIPIFora As Double, ByVal aliqSTFora As Double)
    p_sku = codigoSKU
    p_icms = aliqICMS
    p_pis_cofins = aliqPisCofins
    p_ipi_fora = aliqIPIFora
    p_st_fora = aliqSTFora

    ' Padrão: operação integral, sem exceções
    p_fator = 1#
    p_calc_icms = True
    p_calc_pis_cofins = True
    p_calc_ipi = True
    p_calc_st = True

    Call CalcularImpostosPorDentro
End Sub

' ==============================================================================
' Define o fator da Operação (CfOperações) e as exceções (CfExcecoesImpostos)
' ==============================================================================
Public Sub AplicarOperacao(ByVal fatorOperacao As Double, ByVal calcICMS As Boolean, _
                           ByVal calcPisCofins As Boolean, ByVal calcIPI As Boolean, _
                           ByVal calcST As Boolean)
    p_fator = fatorOperacao
    p_calc_icms = calcICMS
    p_calc_pis_cofins = calcPisCofins
    p_calc_ipi = calcIPI
    p_calc_st = calcST
End Sub

Private Sub CalcularImpostosPorDentro()
    Dim somaFora As Double
    Dim divisor As Double

    somaFora = 1# + p_ipi_fora + p_st_fora
    p_ipi_dentro = p_ipi_fora / somaFora
    p_st_dentro = p_st_fora / somaFora

    ' Divisor da planilha legada (CfImpostos[ICMS] e CfImpostos[Pis/Cofins])
    divisor = 1# + p_ipi_dentro / (1# - p_ipi_dentro) + p_st_dentro / (1# - p_st_dentro)

    p_icms_dentro = p_icms / divisor
    p_pis_cofins_dentro = (p_pis_cofins * (1# - p_icms)) / divisor
End Sub

Private Function Multiplicador(ByVal calcula As Boolean) As Double
    If calcula Then Multiplicador = p_fator Else Multiplicador = 0#
End Function

' ==============================================================================
' Propriedades: valores efetivos exibidos na CfPcfic
' ==============================================================================
Public Property Get SKU() As String
    SKU = p_sku
End Property

Public Property Get ICMS() As Double
    ICMS = p_icms_dentro * Multiplicador(p_calc_icms)
End Property

Public Property Get PisCofins() As Double
    PisCofins = p_pis_cofins_dentro * Multiplicador(p_calc_pis_cofins)
End Property

Public Property Get IPI() As Double
    IPI = p_ipi_dentro * Multiplicador(p_calc_ipi)
End Property

Public Property Get ST() As Double
    ST = p_st_dentro * Multiplicador(p_calc_st)
End Property

Public Property Get TotalImpostos() As Double
    TotalImpostos = Me.ICMS + Me.PisCofins + Me.IPI + Me.ST
End Property

' Colunas IPIF / STF da CfPcfic: alíquota cadastral "por fora", sem fator
' (mesmo comportamento do XLOOKUP legado).
Public Property Get IPIFora() As Double
    IPIFora = p_ipi_fora
End Property

Public Property Get STFora() As Double
    STFora = p_st_fora
End Property

' IPI "por fora" efetivamente destacado na operação. É o divisor do Preço Base:
'   Preço Base = Preço Final / (1 + IPIForaEfetivo)
' Segue o mesmo fator/exceção aplicado ao IPI por dentro: se a operação não
' cobra IPI, não há IPI a descontar do preço.
Public Property Get IPIForaEfetivo() As Double
    IPIForaEfetivo = p_ipi_fora * Multiplicador(p_calc_ipi)
End Property
