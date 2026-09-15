Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet: Set ws = Me

    Dim idxMargem As Long, idxPrecoUn As Long, idxPrecoCx As Long
    Dim idxBaseUn As Long, idxBaseCx As Long
    Dim modoAlvo As String

    Dim EstadoTabela As New clsListState
    If Not EstadoTabela.Init() Then Exit Sub

    If Intersect(Target, EstadoTabela.DataBodyPadrao) Is Nothing Then Exit Sub

    idxMargem = EstadoTabela.IndiceColuna("Margem Bruta (%)")
    idxPrecoUn = EstadoTabela.IndiceColuna("Preço Final Unitário")
    idxPrecoCx = EstadoTabela.IndiceColuna("Preço Final Caixa")
    idxBaseUn = EstadoTabela.IndiceColuna("P.Base Un - Ex IPI/ST")
    idxBaseCx = EstadoTabela.IndiceColuna("P. Base - Ex IPI/ST")
    
    If idxMargem = 0 Or idxPrecoUn = 0 Or idxPrecoCx = 0 Or idxBaseUn = 0 Or idxBaseCx = 0 Then Exit Sub
    
    ' =========================================================================
    ' - LEITURA DE INTENÇÃO MATEMÁTICA
    ' A comparação é feita avaliando se a coluna alvo editada (Target.Column) 
    ' bate com o eixo mapeado pela classe. Rápido, leve e impossível de falhar.
    ' =========================================================================
    modoAlvo = ""
    
    If Target.Column = idxMargem Then
        modoAlvo = "Margem Bruta"
    ElseIf Target.Column = idxPrecoUn Then
        modoAlvo = "Preço Final Unitário"
    ElseIf Target.Column = idxPrecoCx Then
        modoAlvo = "Preço Final Caixa"
    ElseIf Target.Column = idxBaseUn Then
        modoAlvo = "Preço Base Unitário"
    ElseIf Target.Column = idxBaseCx Then
        modoAlvo = "Preço Base Caixa"
    End If
    
    If modoAlvo = "" Then Exit Sub

    On Error GoTo SaidaLimpa
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call MotorPrecificacaoMultidirecional(modoAlvo)
    
SaidaLimpa:
    Application.EnableEvents = True 
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number <> 0 Then
        MsgBox "Falha crítica no cálculo: " & Err.Description, vbCritical, "Erro " & Err.Number
    End If
End Sub

Private Sub AplicarIndicacaoVisual(CfPcfic As ListObject, modoAtual As String)
    Dim rngMargem As Range, rngPrecoU As Range, rngPrecoC As Range
    
    Set rngMargem = CfPcfic.ListColumns("Margem Bruta (%)").DataBodyRange
    Set rngPrecoU = CfPcfic.ListColumns("Preço Final Unitário").DataBodyRange
    Set rngPrecoC = CfPcfic.ListColumns("Preço Final Caixa").DataBodyRange
    
    ' 1. Limpa formatações de todas as 3 colunas envolvidas
    With Union(rngMargem, rngPrecoU, rngPrecoC)
        .Font.Color = RGB(200, 200, 200)
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
        .Font.Size = 10
    End With
    
    ' 2. Colore apenas a âncora atual
    If modoAtual = "Margem Bruta" Then
        With rngMargem
            .Font.Color = RGB(20, 66, 114)
            .Font.Bold = True
            .Font.Size = 12
        End With
    ElseIf modoAtual = "Preço Unitário" Or modoAtual = "Preço Final" Then
        With rngPrecoU
            .Font.Color = RGB(20, 66, 114)
            .Font.Bold = True
            .Font.Size = 12
        End With
    ElseIf modoAtual = "Preço Caixa" Then
        With rngPrecoC
            .Font.Color = RGB(20, 66, 114)
            .Font.Bold = True
            .Font.Size = 12
        End With
    End If
End Sub

Public Sub BotaoRefresh_AtualizarMotor()
    ' 1. Inicializa e valida as estruturas
    Call InicializarTodasTabelas
    
    If CfPcfic Is Nothing Then Exit Sub
    If CfPcfic.DataBodyRange Is Nothing Then
        MsgBox "Não há dados para atualizar.", vbExclamation, "@Cerulean"
        Exit Sub
    End If
    
    ' 2. Mapeamento das colunas vitais para descobrir o Estado Atual
    Dim colMargem As Long, colPrecoU As Long, colPrecoC As Long
    
    With CfPcfic.ListColumns
        colMargem = .Item("Margem Bruta (%)").Index
        colPrecoU = .Item("Preço Final Unitário").Index
        colPrecoC = .Item("Preço Final Caixa").Index
    End With
    
    ' =========================================================================
    ' LEITURA DE ESTADO VISUAL (Stateless Architecture)
    ' Inspeciona a formatação da primeira linha de dados para descobrir o foco
    ' =========================================================================
    Dim modoDetectado As String
    modoDetectado = ""
    
    ' Captura a primeira linha útil da tabela para testes de UI
    Dim primeiraLinhaRange As Range
    Set primeiraLinhaRange = CfPcfic.DataBodyRange.Rows(1)
    
    If primeiraLinhaRange.Cells(1, colMargem).Font.Bold = True Then
        modoDetectado = "Margem Bruta"
    ElseIf primeiraLinhaRange.Cells(1, colPrecoU).Font.Bold = True Then
        modoDetectado = "Preço Unitário"
    ElseIf primeiraLinhaRange.Cells(1, colPrecoC).Font.Bold = True Then
        modoDetectado = "Preço Caixa"
    End If
    
    ' Fallback de segurança: Se o arquivo acabou de ser aberto e ninguém digitou nada,
    ' lê a célula A1 ou assume "Margem Bruta" como padrão de fábrica.
    If modoDetectado = "" Then
        Dim estadoCidadao As String
        estadoCidadao = wsPcfic.Range("A1").Value
        If estadoCidadao = "Preço Final" Then
            modoDetectado = "Preço Unitário"
        Else
            modoDetectado = "Margem Bruta"
        End If
    End If
    
    ' =========================================================================
    ' 3. EXECUÇÃO DO RECALCULO MASSIVO
    ' =========================================================================
    On Error GoTo SaidaErro
    
    ' Maximiza a performance e blinda o ciclo de eventos
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Sincroniza a interface visual de segurança (Garante o Azul Cerúleo no lugar certo)
    Call AplicarIndicacaoVisual(CfPcfic, modoDetectado)
    
    ' Aciona o Motor em Memória passando a âncora correta detectada
    Call MotorPrecificacaoBidirecional(CfPcfic, modoDetectado)
    
    ' Força o recálculo final das fórmulas nativas de apoio (PROCVs, impostos da grade)
    wsPcfic.Calculate

SaidaLimpa:
    ' Devolve as chaves do ambiente de volta ao Excel
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

SaidaErro:
    MsgBox "Erro ao sincronizar motor de cálculo: " & Err.Description, vbCritical, "@Cerulean"
    Resume SaidaLimpa
End Sub