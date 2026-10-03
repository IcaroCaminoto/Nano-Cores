
' =========================================================
' Módulo de Classe: clsListState
' Descrição: Gerencia os ListObjects ativos. Configura dados para a sessão do usuário
' =========================================================
Option Explicit

Private m_LoPadrao As ListObject
Private m_LoPersonalizada As ListObject

Public Function Init() As Boolean
    On Error Resume Next
    Set m_LoPadrao = ws_pcfic1.ListObjects("LoPrecificacaoPadrao")
    Set m_LoPersonalizada = ws_pcfic2.ListObjects("LoPrecificacaoPersonalizada")
    On Error GoTo 0

    ' Se algo falhar durante a execução dos métodos ou se não encontrar, aborta
    If m_LoPadrao Is Nothing Or m_LoPersonalizada Is Nothing Then
        MsgBox "Tabelas de precificação não encontradas.", vbCritical
        Init = False
        Exit Function
    End If
    
    Init = True
End Function

' =========================================================
' PROPRIEDADES: Entrega Segura do DataBodyRange
' =========================================================

Public Property Get DataBodyPadrao() As Range
    ' Quando chamado, este método devolve o DataBodyRange do ListObject Padrão
    
    If m_LoPadrao.DataBodyRange Is Nothing Then
    ' Este If valida o método e devolve a primeira linha sob o cabeçalho se não houver dados na tabela
    ' Isso evita que o código principal quebre ao tentar acessar o alvo
        Set DataBodyPadrao = m_LoPadrao.HeaderRowRange.Offset(1, 0)
    Else
        Set DataBodyPadrao = m_LoPadrao.DataBodyRange
    End If
End Property

Public Property Get DataBodyPersonalizada() As Range
    ' Quando chamado, este método devolve o DataBodyRange do ListObject Personalizado
    
    If m_LoPersonalizada.DataBodyRange Is Nothing Then
    ' Este If valida o método e devolve a primeira linha sob o cabeçalho se não houver dados na tabela
    ' Isso evita que o código principal quebre ao tentar acessar o alvo
        Set DataBodyPersonalizada = m_LoPersonalizada.HeaderRowRange.Offset(1, 0)
    Else
        Set DataBodyPersonalizada = m_LoPersonalizada.DataBodyRange
    End If
End Property

Public Function IndiceColuna(ByVal NomeColuna As String) As Long
    ' Quando chamado, este método devolve o índice da coluna identificada
    ' O motor abservo o impacto se uma coluna não existir
    On Error Resume Next
    IndiceColuna = m_LoPadrao.ListColumns(NomeColuna).Index
    On Error GoTo 0
End Function

Public Sub InjetarLoteCirurgico(ByRef arrMatrizAtualizada As Variant, ParamArray colunasAlvo() As Variant)
' Identifica o alvo o usuário e injeta resultados na coluna alterada
    Dim nomeCol As Variant
    Dim colIndex As Long
    Dim i As Long
    Dim total_linhas As Long
    Dim arrColunaTemp() As Variant
    
    If Not IsArray(arrMatrizAtualizada) Then Exit Sub
    total_linhas = UBound(arrMatrizAtualizada, 1)
    ReDim arrColunaTemp(1 To totalLinhas, 1 To 1)
    
    ' Itera sobre cada nome de coluna que o router solicitou atualizar
    For Each nomeCol In colunasAlvo
        ' reuso do próprio método da classe pra mapear a coluna que foi alterada
        colIndex = Me.IndiceColuna(CStr(nomeCol))
        
        If colIndex > 0 Then
            ' Copia apenas a coluna especifica da matriz oara o tubo
            For i = 1 To totalLinhas
                arrColunaTemp(i, 1) = arrMatrizAtualizada(i, colIndex)
            Next i
            ' despejo de dados
            m_LoPadrao.ListColumns(colIndex).DataBodyRange.Value2 = arrColunaTemp
        Else
            Debug.Print "Aviso: uma falha de dumping ocorreu na Tabela de Custos porque uma coluna não encontrada: " & CStr(nomeCol)
        End If
    Next nomeCol
End Sub