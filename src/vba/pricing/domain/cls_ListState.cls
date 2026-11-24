VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' Módulo de Classe: clsListState
' Camada: Domain (Precificação)
' Finalidade: Localizar a tabela de precificação (CfPcfic) e oferecer acesso
'             seguro ao corpo de dados, aos índices de coluna e à injeção em
'             lote dos resultados do motor.
' Observação: a tabela é localizada pelo nome em qualquer aba do arquivo, sem
'             depender do codinome da aba (wsPcfic) para compilar.
' ==============================================================================
Option Explicit

Private m_Lo As ListObject

' ==============================================================================
' Init: devolve False (sem MsgBox) se a tabela não existir; o chamador decide
'       se deve avisar o usuário. Evita avisos repetidos dentro de eventos.
' ==============================================================================
Public Function Init() As Boolean
    Dim ws As Worksheet

    Set m_Lo = Nothing
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set m_Lo = ws.ListObjects(TABELA_PRECIFICACAO)
        If Not m_Lo Is Nothing Then Exit For
    Next ws
    On Error GoTo 0

    Init = Not (m_Lo Is Nothing)
End Function

Public Property Get Tabela() As ListObject
    Set Tabela = m_Lo
End Property

Public Property Get Planilha() As Worksheet
    Set Planilha = m_Lo.Parent
End Property

' ==============================================================================
' DataBodyPadrao: corpo de dados; se a tabela estiver vazia, devolve a linha
'                 logo abaixo do cabeçalho (evita erro no código chamador).
' ==============================================================================
Public Property Get DataBodyPadrao() As Range
    If m_Lo.DataBodyRange Is Nothing Then
        Set DataBodyPadrao = m_Lo.HeaderRowRange.Offset(1, 0)
    Else
        Set DataBodyPadrao = m_Lo.DataBodyRange
    End If
End Property

Public Property Get TemDados() As Boolean
    TemDados = Not (m_Lo.DataBodyRange Is Nothing)
End Property

' ==============================================================================
' IndiceColuna: posição da coluna dentro da tabela (0 se não existir)
' ==============================================================================
Public Function IndiceColuna(ByVal nomeColuna As String) As Long
    On Error Resume Next
    IndiceColuna = m_Lo.ListColumns(nomeColuna).Index
    On Error GoTo 0
End Function

' ==============================================================================
' NomeColunaDaPlanilha: nome da coluna da tabela que ocupa a coluna da planilha
'                       informada ("" se estiver fora da tabela).
' ==============================================================================
Public Function NomeColunaDaPlanilha(ByVal colunaPlanilha As Long) As String
    Dim idx As Long
    idx = colunaPlanilha - m_Lo.Range.Column + 1
    If idx >= 1 And idx <= m_Lo.ListColumns.Count Then
        NomeColunaDaPlanilha = m_Lo.ListColumns(idx).Name
    End If
End Function

' ==============================================================================
' InjetarLoteCirurgico: escreve, de uma só vez, as colunas pedidas a partir da
'   matriz em memória. A coluna é limpa antes da escrita para que fórmulas
'   legadas de coluna calculada (ÍNDICE/CORRESP, PROCX, PROCV) sejam removidas
'   e não voltem a ser replicadas em linhas novas.
' ==============================================================================
Public Sub InjetarLoteCirurgico(ByRef arrMatrizAtualizada As Variant, ParamArray colunasAlvo() As Variant)
    Dim lista As Variant
    lista = colunasAlvo
    Call InjetarColunas(arrMatrizAtualizada, lista)
End Sub

' InjetarColunas: mesma injeção, recebendo a lista de colunas como Array
' (usado quando a lista é montada em tempo de execução).
Public Sub InjetarColunas(ByRef arrMatrizAtualizada As Variant, ByVal nomesColunas As Variant)
    Dim nomeCol As Variant
    Dim colIndex As Long
    Dim i As Long
    Dim total_linhas As Long
    Dim arrColunaTemp() As Variant
    Dim destino As Range
    Dim temFormula As Variant

    If Not IsArray(arrMatrizAtualizada) Then Exit Sub
    If Not IsArray(nomesColunas) Then Exit Sub
    If m_Lo.DataBodyRange Is Nothing Then Exit Sub
    total_linhas = UBound(arrMatrizAtualizada, 1)
    If total_linhas < 1 Then Exit Sub
    ReDim arrColunaTemp(1 To total_linhas, 1 To 1)

    For Each nomeCol In nomesColunas
        colIndex = Me.IndiceColuna(CStr(nomeCol))

        If colIndex > 0 Then
            For i = 1 To total_linhas
                arrColunaTemp(i, 1) = arrMatrizAtualizada(i, colIndex)
            Next i
            Set destino = m_Lo.ListColumns(colIndex).DataBodyRange
            ' HasFormula devolve Null quando a coluna mistura fórmulas e valores
            temFormula = destino.HasFormula
            If IsNull(temFormula) Then
                destino.ClearContents
            ElseIf temFormula Then
                destino.ClearContents
            End If
            destino.Value2 = arrColunaTemp
        Else
            Debug.Print "Aviso: coluna não encontrada na " & TABELA_PRECIFICACAO & ": " & CStr(nomeCol)
        End If
    Next nomeCol
End Sub

' ==============================================================================
' ChavesExistentes: Dictionary com as chaves (Util_ChaveSKU) dos códigos que já
'                   estão na tabela (evita duplicar produtos no lote).
' ==============================================================================
Public Function ChavesExistentes() As Object
    Dim dict As Object
    Dim idxCod As Long
    Dim arr As Variant
    Dim chave As String
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set ChavesExistentes = dict

    idxCod = Me.IndiceColuna(Col_Codigo())
    If idxCod = 0 Or m_Lo.DataBodyRange Is Nothing Then Exit Function

    arr = ColunaComoMatriz(idxCod)
    For i = 1 To UBound(arr, 1)
        chave = Util_ChaveSKU(arr(i, 1))
        If Len(chave) > 0 Then dict(chave) = True
    Next i
End Function

' ==============================================================================
' AdicionarLinhas: garante "quantidade" linhas livres no fim da tabela.
'   Reaproveita as linhas finais sem Código (ex.: a linha em branco padrão) e
'   acrescenta o restante com ListRows.Add, que replica as colunas calculadas
'   que continuam na planilha (Custo, Comissão, Total Encargos, MKP DIV...).
' Retorno: índice (1 = primeira linha do corpo) da primeira linha livre.
' ==============================================================================
Public Function AdicionarLinhas(ByVal quantidade As Long) As Long
    Dim idxCod As Long
    Dim arr As Variant
    Dim total As Long, vazias As Long, faltam As Long
    Dim i As Long

    idxCod = Me.IndiceColuna(Col_Codigo())
    If m_Lo.DataBodyRange Is Nothing Then
        total = 0
    Else
        total = m_Lo.ListRows.Count
        If idxCod > 0 Then
            arr = ColunaComoMatriz(idxCod)
            For i = total To 1 Step -1
                If Len(Util_ChaveSKU(arr(i, 1))) > 0 Then Exit For
                vazias = vazias + 1
            Next i
        End If
    End If

    faltam = quantidade - vazias
    For i = 1 To faltam
        m_Lo.ListRows.Add
    Next i

    AdicionarLinhas = total - vazias + 1
End Function

' Valores de uma coluna sempre como matriz (1..n, 1..1), mesmo com 1 linha
Private Function ColunaComoMatriz(ByVal colIndex As Long) As Variant
    Dim v As Variant
    Dim arr(1 To 1, 1 To 1) As Variant
    v = m_Lo.ListColumns(colIndex).DataBodyRange.Value2
    If IsArray(v) Then
        ColunaComoMatriz = v
    Else
        arr(1, 1) = v
        ColunaComoMatriz = arr
    End If
End Function
