VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAliquotasDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' Módulo de Classe: clsAliquotasDTO
' Camada: Domain DTO (PURA - pronta para DLL)
' Finalidade: Alíquotas brutas de UM produto para UMA UF, já cruzadas entre as
'             abas "impostos por produto" (Impostos_Federais) e "impostos por
'             UF" (Matriz_ICMS_ST). Todas em fração (0,18 = 18%).
'   ICMS        -> alíquota de ICMS da UF (Matriz_ICMS_ST.aliquota_icms)
'   PisCofins   -> PIS + COFINS (Impostos_Federais; 1,65 + 7,6 = 0,0925)
'   IPIFora     -> IPI "por fora" (Impostos_Federais.aliquota_ipi)
'   STFora      -> ICMS ST "por fora" da UF (Matriz_ICMS_ST.aliquota_st)
'   Calc*       -> False quando a base marca "Não" (exceção: imposto desligado)
' DTO = "objeto de transporte": só carrega dados, não acessa planilha nem base.
' ==============================================================================
Option Explicit

Public Chave As String
Public ICMS As Double
Public PisCofins As Double
Public IPIFora As Double
Public STFora As Double
Public CalcICMS As Boolean
Public CalcPisCofins As Boolean
Public CalcIPI As Boolean
Public CalcST As Boolean

Private Sub Class_Initialize()
    CalcICMS = True
    CalcPisCofins = True
    CalcIPI = True
    CalcST = True
End Sub
