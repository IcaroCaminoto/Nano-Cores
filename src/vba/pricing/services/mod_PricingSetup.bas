Attribute VB_Name = "mod_PricingSetup"
' ==============================================================================
' Módulo: mod_PricingSetup
' Camada: Excel Adapter (rotinas de implantação, executar UMA vez via Alt+F8)
' Finalidade:
'   - Setup_RegistrarBotaoGrupo: cria a linha "Adicionar Grupo de Produtos" na
'     TabelaRotinas (aba sys_config). O painel do botão Opções
'     (frmPainelCerulean) gera o botão sozinho a partir dessa tabela.
'   - Setup_RegistrarBotaoPacoteCustos: idem para "Selecionar Pacote de Custos"
'     (spec-009).
'   - Setup_MigrarParaBasesExternas: troca as fórmulas da CfPcfic por valores
'     vindos das bases externas e exclui as abas locais "Produtos" (dProdutos)
'     e "Impostos" (CfImpostos + CfExcecoesImpostos), com verificações antes.
' ==============================================================================

Option Explicit

Private Const ABA_CONFIG As String = "sys_config"
Private Const TABELA_ROTINAS As String = "TabelaRotinas"
Private Const ABA_LEGADA_PRODUTOS As String = "Produtos"
Private Const ABA_LEGADA_IMPOSTOS As String = "Impostos"

' ==============================================================================
' Procedimento: Setup_RegistrarBotaoGrupo
' ==============================================================================
Public Sub Setup_RegistrarBotaoGrupo()
    Call RegistrarBotaoNoPainel("Adicionar Grupo de Produtos", MACRO_ADICIONAR_GRUPO)
End Sub

' ==============================================================================
' Procedimento: Setup_RegistrarBotaoPacoteCustos
' ==============================================================================
Public Sub Setup_RegistrarBotaoPacoteCustos()
    Call RegistrarBotaoNoPainel("Selecionar Pacote de Custos", MACRO_SELECIONAR_PACOTE)
End Sub

' ==============================================================================
' Procedimento Privado: RegistrarBotaoNoPainel
' Finalidade: Acrescentar (uma única vez) uma linha ativa na TabelaRotinas.
' ==============================================================================
Private Sub RegistrarBotaoNoPainel(ByVal nomeBotao As String, ByVal nomeMacro As String)
    Dim lo As ListObject
    Dim nova As ListRow
    Dim arr As Variant
    Dim idxID As Long, idxNome As Long, idxMacro As Long, idxStatus As Long
    Dim maiorID As Double
    Dim i As Long

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(ABA_CONFIG).ListObjects(TABELA_ROTINAS)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Tabela " & TABELA_ROTINAS & " n" & ChrW(227) & "o encontrada na aba " & ABA_CONFIG & ".", vbCritical
        Exit Sub
    End If

    idxID = IndiceDe(lo, "ID_Rotina")
    idxNome = IndiceDe(lo, "Nome_Botao")
    idxMacro = IndiceDe(lo, "Nome_Macro")
    idxStatus = IndiceDe(lo, "Status")
    If idxNome = 0 Or idxMacro = 0 Or idxStatus = 0 Then
        MsgBox "A " & TABELA_ROTINAS & " precisa das colunas Nome_Botao, Nome_Macro e Status.", vbCritical
        Exit Sub
    End If

    If Not lo.DataBodyRange Is Nothing Then
        arr = lo.DataBodyRange.Value2
        If Not IsArray(arr) Then
            Dim umaCelula(1 To 1, 1 To 1) As Variant
            umaCelula(1, 1) = arr
            arr = umaCelula
        End If
        For i = 1 To UBound(arr, 1)
            If StrComp(Util_Texto(arr(i, idxMacro)), nomeMacro, vbTextCompare) = 0 Then
                MsgBox "O bot" & ChrW(227) & "o """ & nomeBotao & """ j" & ChrW(225) & " est" & ChrW(225) & _
                       " cadastrado na " & TABELA_ROTINAS & ".", vbInformation
                Exit Sub
            End If
            If idxID > 0 Then
                If Util_ToDouble(arr(i, idxID)) > maiorID Then maiorID = Util_ToDouble(arr(i, idxID))
            End If
        Next i
    End If

    Set nova = lo.ListRows.Add
    If idxID > 0 Then nova.Range(1, idxID).Value2 = maiorID + 1
    nova.Range(1, idxNome).Value2 = nomeBotao
    nova.Range(1, idxMacro).Value2 = nomeMacro
    nova.Range(1, idxStatus).Value2 = "Ativo"

    MsgBox "Bot" & ChrW(227) & "o """ & nomeBotao & """ cadastrado. Feche e abra o painel (Op" & ChrW(231) & ChrW(245) & _
           "es) para v" & ChrW(234) & "-lo.", vbInformation
End Sub

' ==============================================================================
' Procedimento: Setup_MigrarParaBasesExternas
' Passos:
'   1. Confere as bases externas (existem? têm as colunas do contrato?)
'   2. Grava cadastro e impostos como valores na CfPcfic (remove PROCV/ÍNDICE)
'   3. Procura fórmulas e nomes que ainda apontem para as abas locais
'   4. Pede confirmação e exclui as abas "Produtos" e "Impostos"
' ==============================================================================
Public Sub Setup_MigrarParaBasesExternas()
    Dim caminhoProd As String, caminhoImp As String
    Dim faltam As Collection
    Dim referencias As Collection
    Dim abas As Collection
    Dim nomeAba As Variant
    Dim ws As Worksheet

    ' ---- 1. Bases externas -----------------------------------------------------
    caminhoProd = Config_ResolveDatabaseFile(DB_FILE_PRODUTOS)
    caminhoImp = Config_ResolveDatabaseFile(DB_FILE_IMPOSTOS)

    On Error GoTo ErroBase
    Set faltam = Produtos_ColunasAusentes(caminhoProd)
    If faltam.Count > 0 Then
        MsgBox "A base de produtos n" & ChrW(227) & "o tem as colunas: " & Util_JuntarLista(faltam) & _
               vbCrLf & caminhoProd & vbCrLf & vbCrLf & "Migra" & ChrW(231) & ChrW(227) & "o cancelada.", vbCritical
        Exit Sub
    End If
    If Not Impostos_TemColunasExcecao(caminhoImp) Then
        MsgBox "A base de impostos ainda n" & ChrW(227) & "o tem as colunas de exce" & ChrW(231) & ChrW(227) & "o " & _
               "(calc_icms, calc_pis_cofins, calc_ipi, calc_st) na aba " & DB_ABA_IMPOSTOS_PRODUTO & "." & vbCrLf & _
               "Sem elas, as exce" & ChrW(231) & ChrW(245) & "es da tabela CfExcecoesImpostos seriam perdidas." & vbCrLf & _
               caminhoImp & vbCrLf & vbCrLf & "Migra" & ChrW(231) & ChrW(227) & "o cancelada.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If MsgBox("A migra" & ChrW(231) & ChrW(227) & "o vai:" & vbCrLf & _
              "1. Gravar cadastro e impostos da CfPcfic como valores (bases externas);" & vbCrLf & _
              "2. Excluir as abas """ & ABA_LEGADA_PRODUTOS & """ e """ & ABA_LEGADA_IMPOSTOS & """." & vbCrLf & vbCrLf & _
              "Salve uma c" & ChrW(243) & "pia do arquivo antes. Continuar?", _
              vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

    ' ---- 2. Valores no lugar das fórmulas --------------------------------------
    Call Pricing_LimparCacheImpostos
    If Not Pricing_PreencherCadastro() Then Exit Sub
    Call Pricing_RecalcularTabela

    ' ---- 3. Referências restantes às abas locais -------------------------------
    Set abas = New Collection
    For Each nomeAba In Array(ABA_LEGADA_PRODUTOS, ABA_LEGADA_IMPOSTOS)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nomeAba))
        On Error GoTo 0
        If Not ws Is Nothing Then abas.Add ws
    Next nomeAba

    If abas.Count = 0 Then
        MsgBox "Cadastro e impostos gravados como valores. As abas locais j" & ChrW(225) & " tinham sido exclu" & ChrW(237) & "das.", vbInformation
        Exit Sub
    End If

    Set referencias = ReferenciasAsAbasLegadas(abas)
    If referencias.Count > 0 Then
        MsgBox "Ainda h" & ChrW(225) & " f" & ChrW(243) & "rmulas ou nomes usando as abas locais:" & vbCrLf & _
               Util_JuntarLista(referencias, 20) & vbCrLf & vbCrLf & _
               "Ajuste-os e rode a migra" & ChrW(231) & ChrW(227) & "o de novo. Nenhuma aba foi exclu" & ChrW(237) & "da.", vbExclamation
        Exit Sub
    End If

    ' ---- 4. Exclusão das abas --------------------------------------------------
    On Error GoTo ErroExclusao
    Application.DisplayAlerts = False
    For Each ws In abas
        ws.Delete
    Next ws
    Application.DisplayAlerts = True
    On Error GoTo 0

    MsgBox "Migra" & ChrW(231) & ChrW(227) & "o conclu" & ChrW(237) & "da. A planilha agora usa s" & ChrW(243) & _
           " as bases externas de produtos e impostos.", vbInformation
    Exit Sub

ErroBase:
    MsgBox "Erro ao ler as bases externas:" & vbCrLf & Err.Description, vbCritical
    Exit Sub

ErroExclusao:
    Application.DisplayAlerts = True
    MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel excluir a aba (estrutura do arquivo protegida?): " & _
           Err.Description, vbCritical
End Sub

' ==============================================================================
' Função Privada: ReferenciasAsAbasLegadas
' Finalidade: Listar células com fórmula e nomes definidos (fora das próprias
'             abas a excluir) que citam as abas ou suas tabelas.
' ==============================================================================
Private Function ReferenciasAsAbasLegadas(ByVal abas As Collection) As Collection
    Dim achados As New Collection
    Dim termos As New Collection
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim formulas As Range
    Dim celula As Range
    Dim nm As Name
    Dim termo As Variant
    Dim alvo As Worksheet
    Dim ehLegada As Boolean

    ' Termos procurados: "Produtos!", "'Produtos'!" e os nomes das tabelas
    For Each alvo In abas
        termos.Add alvo.Name & "!"
        termos.Add "'" & alvo.Name & "'!"
        For Each lo In alvo.ListObjects
            termos.Add lo.Name & "["
        Next lo
    Next alvo

    For Each ws In ThisWorkbook.Worksheets
        ehLegada = False
        For Each alvo In abas
            If alvo Is ws Then ehLegada = True
        Next alvo

        If Not ehLegada Then
            Set formulas = Nothing
            On Error Resume Next
            Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            If Not formulas Is Nothing Then
                For Each celula In formulas.Cells
                    For Each termo In termos
                        If InStr(1, celula.Formula, CStr(termo), vbTextCompare) > 0 Then
                            achados.Add ws.Name & "!" & celula.Address(False, False)
                            Exit For
                        End If
                    Next termo
                Next celula
            End If
        End If
    Next ws

    For Each nm In ThisWorkbook.Names
        ehLegada = False
        If TypeOf nm.Parent Is Worksheet Then
            For Each alvo In abas
                If nm.Parent Is alvo Then ehLegada = True
            Next alvo
        End If
        If Not ehLegada Then
            For Each termo In termos
                If InStr(1, nm.RefersTo, CStr(termo), vbTextCompare) > 0 Then
                    achados.Add "Nome " & nm.Name
                    Exit For
                End If
            Next termo
        End If
    Next nm

    Set ReferenciasAsAbasLegadas = achados
End Function

Private Function IndiceDe(ByVal lo As ListObject, ByVal nomeColuna As String) As Long
    On Error Resume Next
    IndiceDe = lo.ListColumns(nomeColuna).Index
    On Error GoTo 0
End Function
