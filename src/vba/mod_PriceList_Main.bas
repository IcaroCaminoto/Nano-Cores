Attribute VB_Name = "mod_PriceList_Main"
' ==============================================================================
' Módulo: mod_PriceList_Main
' Finalidade: Ponto de entrada, orquestração de pipeline, blindagem de estado do
'             Excel (Zero UI Blocking), medição de tempo e recuperação de erros.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' ==============================================================================
Option Explicit

Private Const APP_TITLE As String = "Gerador de Tabela de Preços"

' ==============================================================================
' Procedimento: GeneratePriceList
' Finalidade: Ponto de entrada principal da automação. Orquestra a verificação
'             de dados, resolução taxonômica e renderização da tabela de preços.
' ==============================================================================
Public Sub GeneratePriceList()
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
    ' 1. Captura de estados globais para restauração incondicional
    ' --------------------------------------------------------------------------
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevDisplayAlerts = Application.DisplayAlerts
    
    ' --------------------------------------------------------------------------
    ' 2. Blindagem de performance (Zero UI Blocking)
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
    Call EnsureParameterSheetExists(wbTarget)
    
    ' --------------------------------------------------------------------------
    ' 4. Extração em memória, filtro de inclusão e resolução taxonômica
    ' --------------------------------------------------------------------------
    isSuccess = ExtractAndProcessPriceData(processedData, recordCount, wbTarget)
    
    If Not isSuccess Or recordCount = 0 Then
        MsgBox "Nenhum produto qualificado foi encontrado na matriz de preços para geração do catálogo.", _
               vbExclamation, APP_TITLE
        GoTo FinallyBlock
    End If
    
    ' --------------------------------------------------------------------------
    ' 5. Renderização do catálogo formatado para impressão (A4 / 1 pág largura)
    ' --------------------------------------------------------------------------
    Call RenderOutputCatalog(processedData, recordCount, wbTarget)
    
    elapsedTime = Round(Timer - startTime, 2)
    
    MsgBox "Tabela de Preços gerada com sucesso!" & vbCrLf & vbCrLf & _
           "• Total de produtos ativos catalogados: " & recordCount & vbCrLf & _
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
    MsgBox "Falha crítica durante a execução do motor de Tabela de Preços:" & vbCrLf & vbCrLf & _
           "• Módulo/Origem: " & Err.Source & vbCrLf & _
           "• Código de Erro: " & Err.Number & vbCrLf & _
           "• Descrição: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume FinallyBlock
End Sub
