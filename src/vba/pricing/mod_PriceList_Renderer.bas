Attribute VB_Name = "mod_PriceList_Renderer"
' ==============================================================================
' Módulo: mod_PriceList_Renderer
' Finalidade: Renderização em lote dos dados na planilha Tabela_Preco,
'             injeção de quebras de seção por categoria/grupo, formatação visual
'             ink-efficient (baixo consumo de tinta) e parametrização de impressão
'             (A4 Retrato, ajustado a 1 página de largura).
' Ambiente: Microsoft Excel 2016+ / Microsoft 365 (VBA 7.1, 64-bit compatível)
' ==============================================================================
Option Explicit

Private Const SHEET_TABELA_PRECO As String = "Tabela_Preco"

' Paleta de cores corporativa Ink-Efficient (econômica para impressão física)
Private Const COLOR_NAVY_TITLE As Long = 7488020    ' RGB(20, 66, 114) - Azul Cerúleo
Private Const COLOR_SECTION_BG As Long = 15790320   ' RGB(240, 240, 240) - Cinza suave neutro
Private Const COLOR_SECTION_TEXT As Long = 2105376  ' RGB(32, 32, 32) - Quase preto
Private Const COLOR_HEADER_BG As Long = 16316664    ' RGB(248, 248, 248) - Fundo quase branco
Private Const COLOR_BORDER As Long = 13816530       ' RGB(210, 210, 210) - Linha de grade fina
Private Const COLOR_LIGHT_LINE As Long = 15132390   ' RGB(230, 230, 230) - Divisor sutil entre linhas

Private Const FONT_FAMILY As String = "Segoe UI"
Private Const FORMAT_CURRENCY As String = """R$ ""#,##0.00"

' ==============================================================================
' Procedimento: RenderOutputSheet / RenderOutputCatalog
' Finalidade: Prepara a planilha Tabela_Preco, organiza os produtos em grupos
'             taxonômicos homogêneos e despeja blocos de dados em memória,
'             aplicando layout pronto para impressão em papel.
' ==============================================================================
Public Sub RenderOutputSheet(ByRef outputData As Variant, ByVal rowCount As Long, Optional ByVal wb As Workbook = Nothing)
    Call RenderOutputCatalog(outputData, rowCount, wb)
End Sub

Public Sub RenderOutputCatalog(ByRef outputData As Variant, ByVal recordCount As Long, Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ThisWorkbook
    If recordCount <= 0 Or IsEmpty(outputData) Then Exit Sub
    
    Dim wsOut As Worksheet
    Dim dictGroups As Object
    Dim groupList() As String
    Dim groupCount As Long
    Dim i As Long
    Dim grpName As String
    Dim currentRow As Long
    
    ' --------------------------------------------------------------------------
    ' 1. Localização ou instanciação da aba de destino Tabela_Preco
    ' --------------------------------------------------------------------------
    On Error Resume Next
    Set wsOut = wb.Worksheets(SHEET_TABELA_PRECO)
    On Error GoTo 0
    
    If wsOut Is Nothing Then
        Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsOut.Name = SHEET_TABELA_PRECO
        wsOut.Tab.Color = COLOR_NAVY_TITLE
    Else
        ' Limpeza completa de conteúdos, formatos e configurações prévias
        wsOut.Cells.Clear
    End If
    
    ' Garante exibição em modo normal e habilitação de linhas de grade
    wsOut.Activate
    ActiveWindow.View = xlNormalView
    ActiveWindow.DisplayGridlines = True
    
    ' --------------------------------------------------------------------------
    ' 2. Agrupamento por Categoria/Taxonomia (Scripting.Dictionary de coleções)
    ' --------------------------------------------------------------------------
    Set dictGroups = CreateObject("Scripting.Dictionary")
    dictGroups.CompareMode = 1 ' vbTextCompare
    
    For i = 1 To recordCount
        grpName = Trim(CStr(outputData(i, 2))) ' Coluna 2 = Resolved_Group_Name
        If Len(grpName) = 0 Then grpName = "DIVERSOS"
        
        If Not dictGroups.Exists(grpName) Then
            Dim colRows As Collection
            Set colRows = New Collection
            dictGroups.Add grpName, colRows
        End If
        dictGroups(grpName).Add i
    Next i
    
    ' Extrai e ordena alfabeticamente os nomes dos grupos
    groupCount = dictGroups.Count
    ReDim groupList(1 To groupCount)
    
    Dim keyItem As Variant
    Dim k As Long: k = 1
    For Each keyItem In dictGroups.Keys
        groupList(k) = CStr(keyItem)
        k = k + 1
    Next keyItem
    
    Call SortStringArray(groupList, 1, groupCount)
    
    ' --------------------------------------------------------------------------
    ' 3. Renderização do Cabeçalho Institucional do Documento (Linhas 1 a 2)
    ' --------------------------------------------------------------------------
    currentRow = 1
    
    ' Título Principal
    With wsOut.Range("A1:D1")
        .Merge
        .Value = "CATÁLOGO DE PRODUTOS & TABELA DE PREÇOS"
        .Font.Name = FONT_FAMILY
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = COLOR_NAVY_TITLE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
    End With
    
    ' Subtítulo com Timestamp e Especificação da Base
    With wsOut.Range("A2:D2")
        .Merge
        .Value = "Emissão: " & Format(Now, "dd/mm/yyyy hh:nn") & _
                 "  |  Base: Preço Base (Ex IPI/ST)  |  Valores sujeitos a alteração"
        .Font.Name = FONT_FAMILY
        .Font.Size = 8.5
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 16
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = COLOR_BORDER
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    currentRow = 3 ' Linha de espaçamento inicial
    wsOut.Rows(currentRow).RowHeight = 8
    currentRow = 4
    
    ' --------------------------------------------------------------------------
    ' 4. Renderização dos Grupos e seus respectivos registros em bloco
    ' --------------------------------------------------------------------------
    Dim g As Long
    Dim currentGrpCollection As Collection
    Dim grpItemCount As Long
    Dim subArray() As Variant
    Dim itemIdx As Long
    Dim sourceRow As Long
    
    For g = 1 To groupCount
        grpName = groupList(g)
        Set currentGrpCollection = dictGroups(grpName)
        grpItemCount = currentGrpCollection.Count
        
        If grpItemCount > 0 Then
            ' --- 4.1 Linha de Quebra de Seção (Section Break Header) ---
            With wsOut.Range(wsOut.Cells(currentRow, 1), wsOut.Cells(currentRow, 4))
                .Merge
                .Value = " " & UCase(grpName) & " (" & grpItemCount & " ITENS)"
                .Font.Name = FONT_FAMILY
                .Font.Size = 10
                .Font.Bold = True
                .Font.Color = COLOR_SECTION_TEXT
                .Interior.Color = COLOR_SECTION_BG
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .RowHeight = 20
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = COLOR_NAVY_TITLE
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = COLOR_BORDER
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            currentRow = currentRow + 1
            
            ' --- 4.2 Cabeçalhos das Colunas ---
            wsOut.Cells(currentRow, 1).Value = "CÓDIGO"
            wsOut.Cells(currentRow, 2).Value = "DESCRIÇÃO DO PRODUTO"
            wsOut.Cells(currentRow, 3).Value = "QTD. EMB."
            wsOut.Cells(currentRow, 4).Value = "PREÇO BASE"
            
            With wsOut.Range(wsOut.Cells(currentRow, 1), wsOut.Cells(currentRow, 4))
                .Font.Name = FONT_FAMILY
                .Font.Size = 8.5
                .Font.Bold = True
                .Font.Color = RGB(60, 60, 60)
                .Interior.Color = COLOR_HEADER_BG
                .VerticalAlignment = xlCenter
                .RowHeight = 18
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = COLOR_BORDER
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            wsOut.Cells(currentRow, 1).HorizontalAlignment = xlCenter
            wsOut.Cells(currentRow, 2).HorizontalAlignment = xlLeft
            wsOut.Cells(currentRow, 3).HorizontalAlignment = xlCenter
            wsOut.Cells(currentRow, 4).HorizontalAlignment = xlRight
            currentRow = currentRow + 1
            
            ' --- 4.3 Despejo em Memória dos Produtos do Grupo (Single-shot) ---
            ReDim subArray(1 To grpItemCount, 1 To 4)
            For itemIdx = 1 To grpItemCount
                sourceRow = currentGrpCollection(itemIdx)
                subArray(itemIdx, 1) = outputData(sourceRow, 3) ' Codigo_Produto
                subArray(itemIdx, 2) = outputData(sourceRow, 4) ' Descricao_Produto
                subArray(itemIdx, 3) = outputData(sourceRow, 5) ' Qtd_Padrao
                subArray(itemIdx, 4) = outputData(sourceRow, 6) ' Preco_Venda
            Next itemIdx
            
            Dim dataStartRow As Long: dataStartRow = currentRow
            Dim dataEndRow As Long: dataEndRow = currentRow + grpItemCount - 1
            
            ' Atribuição em bloco direto à planilha (zero loop de células)
            wsOut.Cells(dataStartRow, 1).Resize(grpItemCount, 4).Value2 = subArray
            
            ' Formatação do bloco de dados deste grupo
            With wsOut.Range(wsOut.Cells(dataStartRow, 1), wsOut.Cells(dataEndRow, 4))
                .Font.Name = FONT_FAMILY
                .Font.Size = 9
                .VerticalAlignment = xlCenter
                .RowHeight = 17
                ' Divisores horizontais finos para economia de tinta
                .Borders(xlInsideH).LineStyle = xlContinuous
                .Borders(xlInsideH).Color = COLOR_LIGHT_LINE
                .Borders(xlInsideH).Weight = xlHairline
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = COLOR_BORDER
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            
            ' Alinhamentos de coluna e formato numérico do bloco
            wsOut.Range(wsOut.Cells(dataStartRow, 1), wsOut.Cells(dataEndRow, 1)).HorizontalAlignment = xlCenter
            wsOut.Range(wsOut.Cells(dataStartRow, 2), wsOut.Cells(dataEndRow, 2)).HorizontalAlignment = xlLeft
            wsOut.Range(wsOut.Cells(dataStartRow, 3), wsOut.Cells(dataEndRow, 3)).HorizontalAlignment = xlCenter
            With wsOut.Range(wsOut.Cells(dataStartRow, 4), wsOut.Cells(dataEndRow, 4))
                .HorizontalAlignment = xlRight
                .NumberFormat = FORMAT_CURRENCY
            End With
            
            currentRow = dataEndRow + 1
            
            ' Espaçamento entre grupos
            wsOut.Rows(currentRow).RowHeight = 8
            currentRow = currentRow + 1
        End If
    Next g
    
    ' --------------------------------------------------------------------------
    ' 5. Aplicação das diretrizes de impressão e ajustes finais de coluna
    ' --------------------------------------------------------------------------
    Call ApplyPrintFormatting(wsOut, currentRow - 1)
End Sub

' ==============================================================================
' Procedimento: ApplyPrintFormatting
' Finalidade: Aplica parametrização A4 Retrato, FitToPagesWide = 1,
'             repetição de títulos $1:$2 e larguras de coluna balanceadas.
' ==============================================================================
Public Sub ApplyPrintFormatting(ByVal targetWs As Worksheet, ByVal lastRow As Long)
    If targetWs Is Nothing Then Exit Sub
    
    ' 1. Larguras de colunas equilibradas para proporção A4 Retrato
    targetWs.Columns("A").ColumnWidth = 14  ' Código
    targetWs.Columns("B").ColumnWidth = 48  ' Descrição
    targetWs.Columns("C").ColumnWidth = 12  ' Qtd Embalagem
    targetWs.Columns("D").ColumnWidth = 17  ' Preço Base
    
    ' 2. Configurações de Impressão (A4, 1 página de largura)
    With targetWs.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        
        ' Ajuste para 1 página de largura sem travar a altura (páginas verticais livres)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        
        ' Repetição de cabeçalho em todas as páginas impressas
        .PrintTitleRows = "$1:$2"
        .PrintTitleColumns = ""
        
        ' Margens otimizadas para impressão em papel
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.6)
        .BottomMargin = Application.InchesToPoints(0.6)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        
        ' Centralização horizontal na página
        .CenterHorizontally = True
        .CenterVertically = False
        
        ' Rodapé com numeração de página e data
        .LeftFooter = "&8Cerulean Cores - Tabela Comercial"
        .RightFooter = "&8Página &P de &N"
        
        .PrintGridlines = False
    End With
    
    ' Seleciona a célula A1 para deixar a interface limpa
    targetWs.Range("A1").Select
End Sub

' ==============================================================================
' Função Auxiliar: SortStringArray
' Finalidade: Algoritmo QuickSort in-memory para ordenação alfabética de strings.
' ==============================================================================
Private Sub SortStringArray(ByRef arr() As String, ByVal leftIdx As Long, ByVal rightIdx As Long)
    Dim i As Long: i = leftIdx
    Dim j As Long: j = rightIdx
    Dim pivot As String: pivot = arr((leftIdx + rightIdx) \ 2)
    Dim temp As String
    
    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0 And i < rightIdx
            i = i + 1
        Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0 And j > leftIdx
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If leftIdx < j Then Call SortStringArray(arr, leftIdx, j)
    If i < rightIdx Then Call SortStringArray(arr, i, rightIdx)
End Sub
