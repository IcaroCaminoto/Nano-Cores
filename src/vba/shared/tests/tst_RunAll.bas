Attribute VB_Name = "tst_RunAll"
' ==============================================================================
' Modulo de Testes: tst_RunAll                       [infra de testes - ambos]
' Finalidade: Rodar, de uma vez, todos os pontos de entrada de teste do arquivo
'             aberto (Nano_Cores_CoreEngine.xlsm ou Demonstrativo de
'             Precificacao.xlsm) e imprimir um resumo unico no fim.
'   - Cada teste e chamado por Application.Run, entao este modulo compila
'     mesmo que o modulo de teste de outra spec nao esteja importado; nesse
'     caso o teste aparece como AUSENTE no resumo.
'   - Depois de cada teste, confere se ScreenUpdating, EnableEvents e
'     Calculation voltaram ao valor de antes. Se nao voltaram, o teste conta
'     como FALHA e o estado e restaurado para nao contaminar o proximo.
' Execucao: Editor VBA > Janela Imediata (Ctrl+G) >
'             RunAllTests                      (host detectado pelo nome)
'             RunAllTests "ENGINE"             (forca o host)
'             RunAllTests "DEMONSTRATIVO"
' Pre-requisito: Ferramentas > Opcoes > Geral > "Interromper em erros nao
'   tratados". Com "Interromper em todos os erros" o depurador para dentro
'   do teste em vez de devolver o erro para este modulo.
' Manutencao: toda spec nova com testes acrescenta uma linha em
'   RegistroDeTestes (ver docs/technical/Guia de Testes.md).
' Encoding: ASCII puro; acentos nas mensagens via ChrW (skill encoding-rules).
' ==============================================================================

Option Explicit

Private Const HOST_ENGINE As String = "ENGINE"
Private Const HOST_DEMO As String = "DEMONSTRATIVO"
Private Const HOST_AMBOS As String = "AMBOS"

Private Const ST_OK As String = "OK"
Private Const ST_FALHA As String = "FALHA"
Private Const ST_AUSENTE As String = "AUSENTE"

Private Const ERR_MACRO_INEXISTENTE As Long = 1004

' ------------------------------------------------------------------------------
' Registro: uma linha por ponto de entrada de teste (Public Sub sem argumentos
' obrigatorios). Formato: "HOST|Macro|Spec|Descricao"
'   HOST = ENGINE, DEMONSTRATIVO ou AMBOS (modulo compartilhado presente nos dois)
' Ordem: contratos (memoria) antes de integracao (base sintetica).
' ------------------------------------------------------------------------------
Private Function RegistroDeTestes() As Variant
    RegistroDeTestes = Array( _
        "AMBOS|Test_AppExecution_SafeZone|spec-003|Restauracao do estado do Application", _
        "ENGINE|RunAllCostEngineContractTests|spec-003|Motor de custos: contratos", _
        "ENGINE|Test_CostEngine_Persistence_SafeZone|spec-003|Motor de custos: trava de arquivo", _
        "DEMONSTRATIVO|RunCostPackageContractTests|spec-009|Pacote de custos: contratos", _
        "DEMONSTRATIVO|RunCostPackageIntegrationTests|spec-009|Pacote de custos: integracao (base sintetica)")
End Function

' ==============================================================================
' Ponto de entrada
' ==============================================================================
Public Sub RunAllTests(Optional ByVal hostForcado As String = vbNullString)
    Dim host As String
    Dim registro As Variant
    Dim partes() As String
    Dim i As Long
    Dim status As String
    Dim detalhe As String
    Dim linhas As String
    Dim nOk As Long, nFalha As Long, nAusente As Long, nIgnorado As Long

    host = UCase$(Trim$(hostForcado))
    If Len(host) = 0 Then host = DetectarHost()
    If host <> HOST_ENGINE And host <> HOST_DEMO And host <> HOST_AMBOS Then
        Err.Raise vbObjectError + 998, "tst_RunAll", _
                  "Host invalido: '" & hostForcado & "'. Use ENGINE, DEMONSTRATIVO ou AMBOS."
    End If

    Debug.Print String$(58, "#")
    Debug.Print "RunAllTests | " & ThisWorkbook.Name & " | host: " & host & _
                " | " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If host = HOST_AMBOS Then
        Debug.Print "(nome do arquivo nao indica o host: rodando o registro inteiro;" & _
                    " testes do outro arquivo vao aparecer como AUSENTE)"
    End If
    Debug.Print String$(58, "#")

    registro = RegistroDeTestes()
    For i = LBound(registro) To UBound(registro)
        partes = Split(registro(i), "|")
        If Not PertenceAoHost(partes(0), host) Then
            nIgnorado = nIgnorado + 1
        Else
            status = ExecutarTeste(partes(1), detalhe)
            Select Case status
                Case ST_OK:      nOk = nOk + 1
                Case ST_FALHA:   nFalha = nFalha + 1
                Case ST_AUSENTE: nAusente = nAusente + 1
            End Select
            linhas = linhas & vbCrLf & " " & Left$(status & Space$(8), 8) & _
                     Left$(partes(2) & Space$(9), 9) & partes(1) & " - " & partes(3) & _
                     " | " & detalhe
        End If
    Next i

    Debug.Print ""
    Debug.Print String$(58, "=")
    Debug.Print "RESUMO RunAllTests (" & host & ")" & linhas
    Debug.Print String$(58, "-")
    Debug.Print " OK: " & nOk & " | Falhas: " & nFalha & " | Ausentes: " & nAusente & _
                " | Fora deste arquivo: " & nIgnorado
    Debug.Print " Linhas [SKIP] / [WARN] acima indicam testes que rodaram sem base;" & _
                " confira antes de dar a spec por testada."
    Debug.Print String$(58, "=")

    If nFalha > 0 Or nAusente > 0 Then
        Err.Raise vbObjectError + 999, "tst_RunAll", _
                  nFalha & " teste(s) com falha e " & nAusente & _
                  " ausente(s). Veja o resumo na Janela Imediata."
    End If
End Sub

' ------------------------------------------------------------------------------
' Roda um ponto de entrada e devolve OK / FALHA / AUSENTE; detalhe recebe o
' tempo e, quando houver, o erro ou a alteracao de estado.
' ------------------------------------------------------------------------------
Private Function ExecutarTeste(ByVal macroNome As String, ByRef detalhe As String) As String
    Dim suAntes As Boolean, evAntes As Boolean, calcAntes As Long
    Dim t0 As Double, dur As Double
    Dim numErro As Long, descErro As String
    Dim estadoAlterado As String

    suAntes = Application.ScreenUpdating
    evAntes = Application.EnableEvents
    calcAntes = Application.Calculation

    Debug.Print ""
    Debug.Print ">>> " & macroNome
    t0 = Timer

    On Error Resume Next
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroNome
    numErro = Err.Number
    descErro = Err.Description
    Err.Clear
    On Error GoTo 0

    dur = Timer - t0
    If dur < 0 Then dur = dur + 86400#             ' virada da meia-noite
    detalhe = Format$(dur, "0.00") & "s"

    If Application.ScreenUpdating <> suAntes Then estadoAlterado = estadoAlterado & " ScreenUpdating"
    If Application.EnableEvents <> evAntes Then estadoAlterado = estadoAlterado & " EnableEvents"
    If Application.Calculation <> calcAntes Then estadoAlterado = estadoAlterado & " Calculation"
    If Len(estadoAlterado) > 0 Then
        Application.ScreenUpdating = suAntes
        Application.EnableEvents = evAntes
        Application.Calculation = calcAntes
        Application.Cursor = xlDefault
        Application.StatusBar = False
    End If

    If numErro = ERR_MACRO_INEXISTENTE And InStr(1, descErro, macroNome, vbTextCompare) > 0 Then
        ExecutarTeste = ST_AUSENTE
        detalhe = "macro n" & ChrW(227) & "o encontrada (m" & ChrW(243) & "dulo n" & ChrW(227) & "o importado?)"
        Debug.Print "![AUSENTE] " & macroNome & " | " & descErro
    ElseIf numErro <> 0 Then
        ExecutarTeste = ST_FALHA
        detalhe = detalhe & " | erro " & numErro & ": " & descErro
    ElseIf Len(estadoAlterado) > 0 Then
        ExecutarTeste = ST_FALHA
        detalhe = detalhe & " | estado n" & ChrW(227) & "o restaurado:" & estadoAlterado
    Else
        ExecutarTeste = ST_OK
    End If
End Function

Private Function DetectarHost() As String
    Dim nome As String
    nome = LCase$(ThisWorkbook.Name)
    If InStr(1, nome, "coreengine") > 0 Then
        DetectarHost = HOST_ENGINE
    ElseIf InStr(1, nome, "demonstrativo") > 0 Then
        DetectarHost = HOST_DEMO
    Else
        DetectarHost = HOST_AMBOS
    End If
End Function

Private Function PertenceAoHost(ByVal hostTeste As String, ByVal hostAtual As String) As Boolean
    PertenceAoHost = (hostAtual = HOST_AMBOS) Or (hostTeste = HOST_AMBOS) Or (hostTeste = hostAtual)
End Function
