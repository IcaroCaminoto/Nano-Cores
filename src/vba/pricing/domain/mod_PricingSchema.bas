Attribute VB_Name = "mod_PricingSchema"
' ==============================================================================
' Módulo: mod_PricingSchema
' Camada: Domain Schema (Precificação)
' Finalidade: Fonte única dos identificadores de modo (âncora), dos nomes de
'             tabela/coluna e das células de parâmetro usados pela precificação.
'             Evita divergências de texto entre o evento da planilha, o motor e
'             a injeção de resultados.
' Referência: "@Demonstrativo de precificação.xlsm" (produção), tabela CfPcfic
'             na aba de codinome wsPcfic (nome visível "TabelaCustos").
' Observações:
'   - Os modos são identificadores ASCII estáveis; serão os mesmos valores
'     enviados à DLL no contrato JSON (Action "precificacao.recalcular").
'   - Nomes com acento são montados via ChrW() (encoding-rules).
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Modos de cálculo (coluna-âncora editada pelo usuário)
' ------------------------------------------------------------------------------
Public Const MODO_MARGEM_BRUTA As String = "MARGEM_BRUTA"
Public Const MODO_PRECO_FINAL_CAIXA As String = "PRECO_FINAL_CAIXA"
Public Const MODO_PRECO_FINAL_UNITARIO As String = "PRECO_FINAL_UNITARIO"
Public Const MODO_PRECO_BASE_CAIXA As String = "PRECO_BASE_CAIXA"
Public Const MODO_PRECO_BASE_UNITARIO As String = "PRECO_BASE_UNITARIO"

' ------------------------------------------------------------------------------
' Tabelas (ListObjects) e células de parâmetro do arquivo de produção
' ------------------------------------------------------------------------------
Public Const TABELA_PRECIFICACAO As String = "CfPcfic"
Public Const TABELA_EXCECOES_IMPOSTOS As String = "CfExcecoesImpostos"   ' transição: exceções migram para a base
Public Const TABELA_UF As String = "dUF"
Public Const ABA_PARAMETROS As String = "AreaCartoes"
Public Const CELULA_UF As String = "E9"                  ' campo "REGIÃO:" (nome do estado)

' Alíquota de fallback do workbook legado (CfImpostos: 0,0925 fixo) quando a
' base não trouxer PIS/COFINS para o SKU.
Public Const ALIQ_PIS_COFINS_PADRAO As Double = 0.0925

' Operação gravada nas linhas novas da inclusão em lote por grupo
' (deve existir na tabela CfOperações; senão os impostos ficam vazios)
Public Const OPERACAO_PADRAO_LOTE As String = "_OP_INTG"

' ------------------------------------------------------------------------------
' Bases externas ("bases burras" .xlsx no servidor, lidas via ADO/SQL)
'   Cerulean_DB_Produtos.xlsx  -> aba Cadastro_Produtos
'   Cerulean_DB_Impostos.xlsx  -> aba Impostos_Federais (impostos por produto)
'                                 aba Matriz_ICMS_ST    (impostos por UF)
' ------------------------------------------------------------------------------
Public Const DB_ABA_PRODUTOS As String = "Cadastro_Produtos"
Public Const DB_ABA_IMPOSTOS_PRODUTO As String = "Impostos_Federais"
Public Const DB_ABA_IMPOSTOS_UF As String = "Matriz_ICMS_ST"

' Campos da aba Impostos_Federais (impostos por produto)
Public Const DBIMP_SKU As String = "sku_codigo_produto"
Public Const DBIMP_PIS As String = "aliquota_pis"
Public Const DBIMP_COFINS As String = "aliquota_cofins"
Public Const DBIMP_IPI As String = "aliquota_ipi"
Public Const DBIMP_CALC_ICMS As String = "calc_icms"              ' "Não" desliga o imposto
Public Const DBIMP_CALC_PIS_COFINS As String = "calc_pis_cofins"
Public Const DBIMP_CALC_IPI As String = "calc_ipi"
Public Const DBIMP_CALC_ST As String = "calc_st"

' Campos da aba Matriz_ICMS_ST (impostos por UF)
Public Const DBIMP_UF As String = "uf_destino"
Public Const DBIMP_ICMS As String = "aliquota_icms"
Public Const DBIMP_ST As String = "aliquota_st"

Public Function Tabela_Operacoes() As String
    Tabela_Operacoes = "CfOpera" & ChrW(231) & ChrW(245) & "es"           ' CfOperações
End Function

' Campos da aba Cadastro_Produtos
Public Function DbProd_Codigo() As String
    DbProd_Codigo = "C" & ChrW(243) & "digo"                             ' Código
End Function

Public Function DbProd_Descricao() As String
    DbProd_Descricao = "Descri" & ChrW(231) & ChrW(227) & "o"            ' Descrição
End Function

Public Function DbProd_QtdPadrao() As String
    DbProd_QtdPadrao = "Qtd Padr" & ChrW(227) & "o"                      ' Qtd Padrão
End Function

Public Function DbProd_CodigoBarras() As String
    DbProd_CodigoBarras = "C" & ChrW(243) & "digo de Barras"             ' Código de Barras
End Function

Public Function DbProd_Grupo() As String
    DbProd_Grupo = "Grupo de Produto"
End Function

Public Function DbProd_PesoKg() As String
    DbProd_PesoKg = "Peso Kg"
End Function

Public Function DbProd_Ativo() As String
    DbProd_Ativo = "Ativo"
End Function

' ------------------------------------------------------------------------------
' Colunas da tabela CfPcfic
' ------------------------------------------------------------------------------
Public Function Col_Codigo() As String
    Col_Codigo = "C" & ChrW(243) & "digo"                                ' Código
End Function

Public Function Col_QtdPadrao() As String
    Col_QtdPadrao = "Qtd Padr" & ChrW(227) & "o"                         ' Qtd Padrão
End Function

' Colunas cadastrais (preenchidas a partir de Cerulean_DB_Produtos.xlsx)
Public Function Col_Descricao() As String
    Col_Descricao = "Descri" & ChrW(231) & ChrW(227) & "o"               ' Descrição
End Function

Public Function Col_Grupo() As String
    Col_Grupo = "Grupo"
End Function

Public Function Col_PesoKg() As String
    Col_PesoKg = "Peso Kg"
End Function

' Opcional: só é preenchida se a coluna existir na CfPcfic
Public Function Col_CodigoBarras() As String
    Col_CodigoBarras = "C" & ChrW(243) & "digo de Barras"                ' Código de Barras
End Function

Public Function Col_Custo() As String
    Col_Custo = "Custo"
End Function

Public Function Col_Frete() As String
    Col_Frete = "Frete"
End Function

Public Function Col_Acordo() As String
    Col_Acordo = "Acordo"
End Function

Public Function Col_Comissao() As String
    Col_Comissao = "Comiss" & ChrW(227) & "o"                            ' Comissão
End Function

Public Function Col_Outros() As String
    Col_Outros = "Outros"
End Function

Public Function Col_Operacao() As String
    Col_Operacao = "Opera" & ChrW(231) & ChrW(227) & "o"                 ' Operação
End Function

Public Function Col_ICMS() As String
    Col_ICMS = "ICMS"
End Function

Public Function Col_PisCofins() As String
    Col_PisCofins = "Pis/Cofins"
End Function

Public Function Col_IPI() As String
    Col_IPI = "IPI"
End Function

Public Function Col_ICMSST() As String
    Col_ICMSST = "ICMS ST"
End Function

Public Function Col_IPIF() As String
    Col_IPIF = "IPIF"                                                    ' IPI "por fora"
End Function

Public Function Col_STF() As String
    Col_STF = "STF"                                                      ' ST "por fora"
End Function

Public Function Col_MargemBruta() As String
    Col_MargemBruta = "Margem Bruta (%)"
End Function

Public Function Col_PrecoFinalCaixa() As String
    Col_PrecoFinalCaixa = "Pre" & ChrW(231) & "o Final Caixa"            ' Preço Final Caixa
End Function

Public Function Col_PrecoFinalUnitario() As String
    Col_PrecoFinalUnitario = "Pre" & ChrW(231) & "o Final Unit" & ChrW(225) & "rio"  ' Preço Final Unitário
End Function

Public Function Col_PrecoBaseCaixa() As String
    Col_PrecoBaseCaixa = "P. Base - Ex IPI/ST"                           ' com espaço após "P."
End Function

Public Function Col_PrecoBaseUnitario() As String
    Col_PrecoBaseUnitario = "P.Base Un - Ex IPI/ST"                      ' sem espaço após "P."
End Function

' ------------------------------------------------------------------------------
' Colunas das tabelas auxiliares (CfOperações, CfExcecoesImpostos, dUF)
' ------------------------------------------------------------------------------
Public Function ColOp_Fator() As String
    ColOp_Fator = "Fator"
End Function

Public Function ColExc_ICMS() As String
    ColExc_ICMS = "Calc_ICMS"
End Function

Public Function ColExc_PisCofins() As String
    ColExc_PisCofins = "Calc_PISCOFINS"
End Function

Public Function ColExc_IPI() As String
    ColExc_IPI = "Calc_IPI"
End Function

Public Function ColExc_ST() As String
    ColExc_ST = "Calc_ST"
End Function

Public Function ColUF_Nome() As String
    ColUF_Nome = "Nome_Estado"
End Function

Public Function ColUF_Sigla() As String
    ColUF_Sigla = "Sigla_Estado"
End Function

Public Function Texto_Nao() As String
    Texto_Nao = "N" & ChrW(227) & "o"                                    ' Não
End Function

' ==============================================================================
' Função: ModoPorColuna
' Finalidade: Traduzir o nome da coluna editada para o modo de cálculo.
' Retorno: Identificador MODO_* ou vbNullString se a coluna não for âncora.
' ==============================================================================
Public Function ModoPorColuna(ByVal nomeColuna As String) As String
    Select Case nomeColuna
        Case Col_MargemBruta():         ModoPorColuna = MODO_MARGEM_BRUTA
        Case Col_PrecoFinalCaixa():     ModoPorColuna = MODO_PRECO_FINAL_CAIXA
        Case Col_PrecoFinalUnitario():  ModoPorColuna = MODO_PRECO_FINAL_UNITARIO
        Case Col_PrecoBaseCaixa():      ModoPorColuna = MODO_PRECO_BASE_CAIXA
        Case Col_PrecoBaseUnitario():   ModoPorColuna = MODO_PRECO_BASE_UNITARIO
        Case Else:                      ModoPorColuna = vbNullString
    End Select
End Function

' ==============================================================================
' Função: ColunaAncoraDoModo
' Finalidade: Operação inversa de ModoPorColuna (usada na indicação visual).
' ==============================================================================
Public Function ColunaAncoraDoModo(ByVal modo As String) As String
    Select Case modo
        Case MODO_MARGEM_BRUTA:           ColunaAncoraDoModo = Col_MargemBruta()
        Case MODO_PRECO_FINAL_CAIXA:      ColunaAncoraDoModo = Col_PrecoFinalCaixa()
        Case MODO_PRECO_FINAL_UNITARIO:   ColunaAncoraDoModo = Col_PrecoFinalUnitario()
        Case MODO_PRECO_BASE_CAIXA:       ColunaAncoraDoModo = Col_PrecoBaseCaixa()
        Case MODO_PRECO_BASE_UNITARIO:    ColunaAncoraDoModo = Col_PrecoBaseUnitario()
        Case Else:                        ColunaAncoraDoModo = vbNullString
    End Select
End Function

' ==============================================================================
' Função: ColunaEhInsumo
' Finalidade: Indicar se a coluna editada é um insumo do cálculo (não-âncora)
'             cuja alteração exige recalcular a linha mantendo a âncora atual.
' ==============================================================================
Public Function ColunaEhInsumo(ByVal nomeColuna As String) As Boolean
    Select Case nomeColuna
        Case Col_Codigo(), Col_QtdPadrao(), Col_Custo(), Col_Frete(), Col_Acordo(), _
             Col_Comissao(), Col_Outros(), Col_Operacao()
            ColunaEhInsumo = True
        Case Else
            ColunaEhInsumo = False
    End Select
End Function

' ==============================================================================
' Função: ModoValido
' ==============================================================================
Public Function ModoValido(ByVal modo As String) As Boolean
    ModoValido = (Len(ColunaAncoraDoModo(modo)) > 0)
End Function

' ==============================================================================
' Função: TodosOsModos
' Finalidade: Lista dos modos na ordem de prioridade da detecção visual.
' ==============================================================================
Public Function TodosOsModos() As Variant
    TodosOsModos = Array(MODO_MARGEM_BRUTA, MODO_PRECO_FINAL_CAIXA, MODO_PRECO_FINAL_UNITARIO, _
                         MODO_PRECO_BASE_CAIXA, MODO_PRECO_BASE_UNITARIO)
End Function
