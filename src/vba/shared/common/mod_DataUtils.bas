Attribute VB_Name = "mod_DataUtils"
' ==============================================================================
' Módulo: mod_DataUtils
' Camada: Shared (PURA - pronta para DLL)
' Finalidade: Conversões de valores comuns à camada de dados e ao motor.
' Regra DLL: nenhum objeto do Excel (Range, Worksheet, Application,
'            ThisWorkbook) e nenhuma MsgBox neste módulo.
' ==============================================================================

Option Explicit

' Texto aparado; "" para vazio, nulo ou erro (#N/D etc.)
Public Function Util_Texto(ByVal valor As Variant) As String
    If IsError(valor) Or IsEmpty(valor) Or IsNull(valor) Then Exit Function
    If IsObject(valor) Then Exit Function
    Util_Texto = Trim$(CStr(valor))
End Function

' Chave de SKU comparável entre planilha (texto ou número) e base (número):
' 40001 (Double) e "40001" (texto) geram a mesma chave.
Public Function Util_ChaveSKU(ByVal valor As Variant) As String
    Util_ChaveSKU = Util_Texto(valor)
End Function

' Número digitado numa célula (respeita a localidade: "0,18" = 0,18 em pt-BR).
' Substitui Val(), que ignora a vírgula decimal.
Public Function Util_ToDouble(ByVal valor As Variant) As Double
    If IsError(valor) Or IsEmpty(valor) Or IsNull(valor) Then Exit Function
    If IsObject(valor) Then Exit Function
    If IsNumeric(valor) Then Util_ToDouble = CDbl(valor)
End Function

' Número vindo de uma base via ADO. Colunas mistas (IMEX=1) chegam como texto
' com ponto decimal ("0.065"); CDbl leria "65" em pt-BR, por isso o texto é
' convertido com Val() (que sempre usa ponto), aceitando também vírgula.
Public Function Util_NumeroDaBase(ByVal valor As Variant) As Double
    Dim s As String
    If IsError(valor) Or IsEmpty(valor) Or IsNull(valor) Then Exit Function
    If IsObject(valor) Then Exit Function
    If VarType(valor) = vbString Then
        s = Trim$(valor)
        If Len(s) = 0 Then Exit Function
        If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then s = Replace(s, ",", ".")
        Util_NumeroDaBase = Val(s)
    ElseIf IsNumeric(valor) Then
        Util_NumeroDaBase = CDbl(valor)
    End If
End Function

' Alíquota em fração; valores > 1 são pontos percentuais (7,6 -> 0,076)
Public Function Util_NormalizarAliquota(ByVal valor As Variant) As Double
    Dim v As Double
    v = Util_NumeroDaBase(valor)
    If v > 1# Then v = v / 100#
    Util_NormalizarAliquota = v
End Function

' "Não" (qualquer caixa) -> True. Vazio, "Sim" ou ausente -> False.
Public Function Util_EhNao(ByVal valor As Variant) As Boolean
    Util_EhNao = (StrComp(Util_Texto(valor), "N" & ChrW(227) & "o", vbTextCompare) = 0)
End Function

Public Function Util_ArquivoExiste(ByVal caminho As String) As Boolean
    If Len(caminho) = 0 Then Exit Function
    On Error Resume Next
    Util_ArquivoExiste = (Len(Dir(caminho)) > 0)
    If Err.Number <> 0 Then Util_ArquivoExiste = False
    On Error GoTo 0
End Function

' Valor de um campo de registro ADO (Scripting.Dictionary campo -> valor);
' Empty se o campo não existir na base ou vier nulo.
Public Function Util_Campo(ByVal registro As Object, ByVal nomeCampo As String) As Variant
    If registro Is Nothing Then Exit Function
    If Not registro.Exists(nomeCampo) Then Exit Function
    If IsNull(registro(nomeCampo)) Then Exit Function
    Util_Campo = registro(nomeCampo)
End Function

' Junta os itens de uma Collection de textos ("a, b, c"), limitado a maxItens
Public Function Util_JuntarLista(ByVal itens As Collection, Optional ByVal maxItens As Long = 15) As String
    Dim item As Variant
    Dim n As Long
    Dim s As String

    For Each item In itens
        n = n + 1
        If n > maxItens Then
            s = s & " (+" & (itens.Count - maxItens) & ")"
            Exit For
        End If
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(item)
    Next item
    Util_JuntarLista = s
End Function
