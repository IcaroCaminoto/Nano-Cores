VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProdutoDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' Módulo de Classe: clsProdutoDTO
' Camada: Domain DTO (PURA - pronta para DLL)
' Finalidade: Dados cadastrais de um produto lidos de Cerulean_DB_Produtos.xlsx
'             e, após o cruzamento (mod_CatalogoService), as alíquotas da UF.
'   Codigo        -> valor original da base (número), gravado na CfPcfic
'   Chave         -> Codigo normalizado para comparação (Util_ChaveSKU)
'   Impostos      -> clsAliquotasDTO ou Nothing (produto sem imposto na base)
' ==============================================================================
Option Explicit

Public Codigo As Variant
Public Chave As String
Public Descricao As String
Public QtdPadrao As Variant
Public CodigoBarras As String
Public Grupo As String
Public PesoKg As Variant
Public Ativo As Boolean
Public Impostos As clsAliquotasDTO

Public Property Get TemImpostos() As Boolean
    TemImpostos = Not (Impostos Is Nothing)
End Property
