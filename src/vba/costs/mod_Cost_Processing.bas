' ==============================================================================
' Módulo: mod_Cost_Processing
' Finalidade: Ponto de entrada, orquestração de pipeline, blindagem de estado do
'             Excel (Zero UI Blocking), medição de tempo e recuperação de erros.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' ==============================================================================

Option Explicit

Private Const APP_TITLE As String = "Fechamento de Custos Mensal"

' ==============================================================================
' Procedimento: OrchestrateAverageCostsByProduct
' Finalidade: Ponto de entrada principal da automação. Orquestra a verificação
'             de dados e o controle de estado da aplicação.
' ==============================================================================

Public Sub OrchestrateAverageCostsByProduct()
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevDisplayAlerts As Boolean
    
    Dim startTime As Double
    Dim wbTarget As Workbook
    Dim processedData As Variant
    Dim recordCount As Long
    Dim elapsedTime As Double
    Dim isSuccess As Boolean
    
    ' --------------------------------------------------------------------------
    ' 1. Captura de estados globais para restauração incondicional se der erro
    ' --------------------------------------------------------------------------
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevDisplayAlerts = Application.DisplayAlerts
    
    ' --------------------------------------------------------------------------
    ' 2. Otimização de performance
    ' --------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    Set wbTarget = ThisWorkbook

    ' --------------------------------------------------------------------------
    ' 3. Auto-provisionamento de parâmetros e fontes auxiliares
    ' --------------------------------------------------------------------------
    Call EnsureCostsSheetExists(wbTarget)
    Call EnsureParameterSheetExists(wbTarget)
    
    ' --------------------------------------------------------------------------
    ' 4. Extração em memória, cálulo das médias e resolução taxonômica
    ' --------------------------------------------------------------------------
    isSuccess = CalculateAverageCostsByProduct(recordCount, wbTarget)
    
    If Not isSuccess Or recordCount = 0 Then
        MsgBox "Nenhum produto qualificado foi encontrado na matriz de preços para geração do catálogo.", _
               vbExclamation, APP_TITLE
        GoTo FinallyBlock
    End If

    ' --------------------------------------------------------------------------
    ' 5. Renderização do relatório formatado para impressão (A4 / 1 pág largura)
    ' --------------------------------------------------------------------------
    Call RenderOutputCatalog() ' <========================== REFATORAR
    
    elapsedTime = Round(Timer - startTime, 2)
    
    MsgBox "Pacote de custos processado com sucesso!" & vbCrLf & vbCrLf & _
           "• Total de custos catalogados: " & recordCount & vbCrLf & _
           "• Tempo total de processamento: " & elapsedTime & " segundos", _
           vbInformation, APP_TITLE

FinallyBlock:
    ' --------------------------------------------------------------------------
    ' 6. Restauração incondicional dos estados da aplicação
    ' --------------------------------------------------------------------------
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts
    Exit Sub

ErrorHandler:
    MsgBox "Falha crítica durante a execução do cálculo:" & vbCrLf & vbCrLf & _
           "• Módulo/Origem: " & Err.Source & vbCrLf & _
           "• Código de Erro: " & Err.Number & vbCrLf & _
           "• Descrição: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume FinallyBlock

End Sub

' ==============================================================================
' Procedimento: CalculateAverageCostsByProduct
' Finalidade: Motor de cálculo central da aplicação. Promove o processamento dos
'             dados brutos e resolução taxonômica tabela de custos.
' ==============================================================================

Private Sub CalculateAverageCostsByProduct(ByRef outCount As Long, Optional ByVal wb As Workbook = Nothing)
    ' R1: Faz média de entradas nos últimos 90 dias para cada sku
    

    ' R2: Faz média dos agrupamentos de produtos, considerando as associações de
    ' famílias na tabela tb_Param_Autocura, aba Param_AutoCura
    Dim rangeCustoOriginal As Range
    Dim rangeGrupo As Range
    Dim rangeFazMedia As Range
    
    Set rangeCustoOriginal = ws.Range(COL_CUSTO_ORIGINAL & LINHA_INICIAL & ":" & COL_CUSTO_ORIGINAL & lastRow)
    Set rangeGrupo = ws.Range(COL_GRUPO & LINHA_INICIAL & ":" & COL_GRUPO & lastRow)
    Set rangeFazMedia = ws.Range(COL_FAZ_MEDIA & LINHA_INICIAL & ":" & COL_FAZ_MEDIA & lastRow)
    
    For i = LINHA_INICIAL To lastRow
        If UCase(ws.Cells(i, COL_FAZ_MEDIA).Value) = "NÃO" Then
            ws.Cells(i, COL_CUSTO_FINAL).Value = ws.Cells(i, COL_CUSTO_ORIGINAL).Value
        ElseIf UCase(ws.Cells(i, COL_FAZ_MEDIA).Value) = "SIM" Then
            mediaGrupo = Application.AverageIfs(rangeCustoOriginal, rangeGrupo, ws.Cells(i, COL_GRUPO).Value, rangeFazMedia, "Sim")
            If IsError(mediaGrupo) Then
                ws.Cells(i, COL_CUSTO_FINAL).Value = "SCUSTO_PER"
            Else
                ws.Cells(i, COL_CUSTO_FINAL).Value = mediaGrupo
            End If
        Else
            ws.Cells(i, COL_CUSTO_FINAL).ClearContents
        End If
    Next i
End Sub