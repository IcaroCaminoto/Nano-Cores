Option Explicit

Private Type AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    DisplayAlerts As Boolean
End Type

Private m_SavedState As AppState
Private m_StartTime As Double

Public Sub BeginExecution(Optional ByVal DisableAlerts As Boolean = True)
    ' 1. Captura de estados anteriores
    With m_SavedState
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .Calculation = Application.Calculation
        .DisplayAlerts = Application.DisplayAlerts
    End With

    ' 2. Aplicação de estados de alta performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    If DisableAlerts Then Application.DisplayAlerts = False

    ' 3. Inicialização do cronômetro
    m_StartTime = Timer
End Sub

Public Function EndExecution(Optional ByVal RoutineName As String = "Processo") As Double
    Dim rawElapsed As Double

    ' 1. Cálculo do tempo decorrido com tratamento de Midnight Rollover
    rawElapsed = Timer - m_StartTime
    If rawElapsed < 0 Then rawElapsed = rawElapsed + 86400#
    EndExecution = Round(rawElapsed, 2)

    ' 2. Restauração incondicional dos estados
    On Error Resume Next
    With m_SavedState
        Application.ScreenUpdating = .ScreenUpdating
        Application.EnableEvents = .EnableEvents
        Application.Calculation = .Calculation
        Application.DisplayAlerts = .DisplayAlerts
    End With
    On Error GoTo 0

    Debug.Print "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & RoutineName & _
                " concluído em " & Format(EndExecution, "0.00") & "s."
End Function