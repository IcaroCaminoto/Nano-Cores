Attribute VB_Name = "mod_Config"
' ==============================================================================
' Módulo: mod_Config
' Camada: Shared Configuration
' Finalidade: Ponto único de configuração de caminhos e nomes de arquivos das
'             bases Cerulean_DB. Nenhum outro módulo deve conter caminhos fixos.
'             Na arquitetura compilada (DLL), estes valores serão enviados à DLL
'             no campo Parameters.PastaBanco do pedido JSON.
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' Conformidade: encoding-rules (strings funcionais 100% ASCII, acentos via ChrW)
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' Nomes físicos dos arquivos de base de dados
' ------------------------------------------------------------------------------
Public Const DB_FILE_PRODUTOS As String = "Cerulean_DB_Produtos.xlsx"
Public Const DB_FILE_PARAM_ASSOC As String = "Cerulean_DB_Parametros_Associacao_Produtos.xlsx"
Public Const DB_FILE_CUSTOS As String = "Cerulean_DB_Custos.xlsx"
Public Const DB_FILE_IMPOSTOS As String = "Cerulean_DB_Impostos.xlsx"

' ==============================================================================
' Função: Config_FallbackDatabaseFolder
' Finalidade: Pasta absoluta de contingência, usada apenas quando nenhuma pasta
'             relativa ao arquivo host contém a base. Equivale ao antigo
'             "G:\Ícaro\Cerulean_DB\sheets\database" (Í = ChrW(205)).
' ==============================================================================
Public Function Config_FallbackDatabaseFolder() As String
    Config_FallbackDatabaseFolder = "G:\" & ChrW(205) & "caro\Cerulean_DB\sheets\database"
End Function

' ==============================================================================
' Função: Config_ResolveDatabaseFile
' Finalidade: Localizar o caminho completo de uma base, testando em ordem:
'   1. <host>\src\sheets\Cerulean_DB\database\
'   2. <host>\sheets\Cerulean_DB\database\
'   3. <host>\sheets\database\           (layout legado da base de impostos)
'   4. <host>\data\                      (layout do repositório)
'   5. Config_FallbackDatabaseFolder()   (contingência absoluta)
' Retorno: Caminho do primeiro candidato existente; se nenhum existir, o caminho
'          de contingência (o chamador decide como tratar a ausência).
' ==============================================================================
Public Function Config_ResolveDatabaseFile(ByVal fileName As String) As String
    Dim sep As String
    Dim basePath As String
    Dim candidates(1 To 4) As String
    Dim i As Long

    sep = Application.PathSeparator

    On Error Resume Next
    basePath = ThisWorkbook.Path
    On Error GoTo 0

    If Len(basePath) > 0 Then
        candidates(1) = basePath & sep & "src" & sep & "sheets" & sep & "Cerulean_DB" & sep & "database" & sep & fileName
        candidates(2) = basePath & sep & "sheets" & sep & "Cerulean_DB" & sep & "database" & sep & fileName
        candidates(3) = basePath & sep & "sheets" & sep & "database" & sep & fileName
        candidates(4) = basePath & sep & "data" & sep & fileName

        For i = LBound(candidates) To UBound(candidates)
            If FileExists(candidates(i)) Then
                Config_ResolveDatabaseFile = candidates(i)
                Exit Function
            End If
        Next i
    End If

    Config_ResolveDatabaseFile = Config_FallbackDatabaseFolder() & sep & fileName
End Function

' ==============================================================================
' Função Privada: FileExists
' Finalidade: Testar a existência de um arquivo sem propagar erros de Dir()
'             (ex.: unidade de rede indisponível).
' ==============================================================================
Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(fullPath)) > 0)
    If Err.Number <> 0 Then FileExists = False
    On Error GoTo 0
End Function
