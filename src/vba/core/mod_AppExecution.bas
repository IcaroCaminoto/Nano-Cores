Attribute VB_Name = "mod_AppExecution"
' ==============================================================================
' Módulo: mod_AppExecution
' Camada: Core Infrastructure & State Lifecycle Management
' Finalidade: Centralizar e blindar a gestão de estados da aplicação Excel
'             (Zero UI Blocking), suporte determinístico a medição de tempo
'             (Timer) e garantia incondicional de restauração de estado
'             (State Restoration Trap / FinallyBlock).
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' Conformidade: SPEC-003, vba-high-performance, vba-coding-standards, code-spec-validator
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Estrutura de dados para armazenamento do snapshot do estado do Excel
' ------------------------------------------------------------------------------
Private Type AppExecutionState
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    Cursor As XlMousePointer
    DisplayStatusBar As Boolean
    StatusBarText As Variant
    IsCaptured As Boolean
End Type

' ------------------------------------------------------------------------------
' Variáveis de escopo de módulo
' ------------------------------------------------------------------------------
Private m_SavedState As AppExecutionState
Private m_TimerStart As Double

' ==============================================================================
' Procedimento: FreezeAppState
' Finalidade: Captura o estado operacional corrente do Excel e aplica modo de
'             alta performance sem travamento de UI ou disparos acidentais de eventos.
' Parâmetros:
'   - setWaitCursor (Boolean): Se True, altera o cursor para ampulheta/ocupado.
'   - statusMessage (String): Mensagem opcional para exibição na barra de status.
' ==============================================================================
Public Sub FreezeAppState(Optional ByVal setWaitCursor As Boolean = True, _
                         Optional ByVal statusMessage As String = vbNullString)
    
    ' Captura o estado atual apenas se ainda não houver snapshot retido,
    ' garantindo proteção contra chamadas reentrantes ou aninhadas.
    If Not m_SavedState.IsCaptured Then
        With m_SavedState
            .ScreenUpdating = Application.ScreenUpdating
            .DisplayAlerts = Application.DisplayAlerts
            .EnableEvents = Application.EnableEvents
            .Calculation = Application.Calculation
            .Cursor = Application.Cursor
            .DisplayStatusBar = Application.DisplayStatusBar
            .StatusBarText = Application.StatusBar
            .IsCaptured = True
        End With
    End If
    
    ' Aplica configuração de alto desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    If setWaitCursor Then
        Application.Cursor = xlWait
    End If
    
    If Len(Trim$(statusMessage)) > 0 Then
        Application.DisplayStatusBar = True
        Application.StatusBar = statusMessage
    End If
    
End Sub

' ==============================================================================
' Procedimento: RestoreAppState
' Finalidade: Restaura incondicionalmente os estados do Excel capturados antes do
'             congelamento. Deve ser invocado no Finally/Exit de todas as rotinas.
' Parâmetros:
'   - forceReset (Boolean): Se True, força a redefinição para os valores padrão
'                           do Excel independentemente de snapshot prévio.
' ==============================================================================
Public Sub RestoreAppState(Optional ByVal forceReset As Boolean = False)
    On Error Resume Next
    
    If m_SavedState.IsCaptured And Not forceReset Then
        Application.ScreenUpdating = m_SavedState.ScreenUpdating
        Application.DisplayAlerts = m_SavedState.DisplayAlerts
        Application.EnableEvents = m_SavedState.EnableEvents
        Application.Calculation = m_SavedState.Calculation
        Application.Cursor = m_SavedState.Cursor
        Application.DisplayStatusBar = m_SavedState.DisplayStatusBar
        Application.StatusBar = m_SavedState.StatusBarText
    Else
        ' Fallback de segurança corporativo
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.Cursor = xlDefault
        Application.StatusBar = False
    End If
    
    ' Libera o snapshot
    m_SavedState.IsCaptured = False
    On Error GoTo 0
End Sub

' ==============================================================================
' Função: StartStopwatch
' Finalidade: Registra o carimbo inicial do Timer para medição determinística de tempo.
' Retorno: Double com a marca de segundos obtida pelo Timer.
' ==============================================================================
Public Function StartStopwatch() As Double
    m_TimerStart = Timer
    StartStopwatch = m_TimerStart
End Function

' ==============================================================================
' Função: GetElapsedTime
' Finalidade: Calcula o tempo decorrido em segundos desde o marco temporal indicado.
'             Possui compensação nativa para a virada de meia-noite do VBA Timer.
' Parâmetros:
'   - startMark (Double): Carimbo inicial. Se omitido/negativo, usa m_TimerStart.
' Retorno: Double com o tempo decorrido arredondado para duas casas decimais.
' ==============================================================================
Public Function GetElapsedTime(Optional ByVal startMark As Double = -1#) As Double
    Dim effectiveStart As Double
    Dim currentTimer As Double
    Dim elapsed As Double
    
    If startMark < 0# Then
        effectiveStart = m_TimerStart
    Else
        effectiveStart = startMark
    End If
    
    currentTimer = Timer
    elapsed = currentTimer - effectiveStart
    
    ' Compensação para descontinuidade temporal da meia-noite (Timer zera em 86400s)
    If elapsed < 0# Then
        elapsed = elapsed + 86400#
    End If
    
    GetElapsedTime = Round(elapsed, 2)
End Function

' ==============================================================================
' Função: IsAppStateFrozen
' Finalidade: Consulta se o ambiente encontra-se sob congelamento de estado ativo.
' Retorno: Boolean (True se congelado, False caso contrário).
' ==============================================================================
Public Function IsAppStateFrozen() As Boolean
    IsAppStateFrozen = m_SavedState.IsCaptured
End Function

' ==============================================================================
' Procedimento de Teste / SafeZone: Test_AppExecution_SafeZone
' Finalidade: Auditar o comportamento do manipulador de estado sob estresse e
'             simulação de erro em tempo de execução, garantindo que o State
'             Restoration Trap funcione 100% sem intervenção manual.
' ==============================================================================
Public Sub Test_AppExecution_SafeZone()
    Dim benchStart As Double
    Dim testPassed As Boolean
    Dim simErrorTriggered As Boolean
    Dim durationSec As Double
    
    testPassed = False
    simErrorTriggered = False
    
    ' 1. Inicia medição e congelamento de tela
    benchStart = StartStopwatch()
    Call FreezeAppState(setWaitCursor:=True, statusMessage:="Auditoria Safezone em andamento...")
    
    ' Verificação imediata do congelamento
    If Application.ScreenUpdating <> False Or _
       Application.Calculation <> xlCalculationManual Or _
       Application.EnableEvents <> False Then
        Call RestoreAppState(forceReset:=True)
        Err.Raise vbObjectError + 1001, "Test_AppExecution_SafeZone", "Falha de asserção: Estados não foram congelados."
    End If
    
    ' 2. Simulação de bloco de execução protegido com disparo de erro intencional
    On Error GoTo TestErrorHandler
    
    ' Simula operação e dispara erro deliberado
    simErrorTriggered = True
    Err.Raise vbObjectError + 9999, "Test_AppExecution_SafeZone", "Erro simulado para validação de recuperação incondicional."
    
TestFinallyBlock:
    ' 3. Bloco incondicional de recuperação (FinallyBlock)
    Call RestoreAppState()
    
    ' 4. Asserções pós-restauração
    durationSec = GetElapsedTime(benchStart)
    
    If Application.ScreenUpdating = True And _
       Application.Calculation = xlCalculationAutomatic And _
       Application.EnableEvents = True And _
       Application.Cursor = xlDefault And _
       Not IsAppStateFrozen() And _
       simErrorTriggered Then
        testPassed = True
    End If
    
    If testPassed Then
        Debug.Print "[PASS] Test_AppExecution_SafeZone concluído com sucesso em " & durationSec & "s. Estados íntegros."
    Else
        Debug.Print "[FAIL] Test_AppExecution_SafeZone falhou: restauração de estados incompleta."
        Err.Raise vbObjectError + 1002, "Test_AppExecution_SafeZone", "Falha crítica: Estados do Excel não foram restaurados."
    End If
    Exit Sub
    
TestErrorHandler:
    ' Captura o erro simulado e direciona ao bloco Finally
    Resume TestFinallyBlock
End Sub
