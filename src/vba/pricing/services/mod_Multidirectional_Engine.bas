Public Sub MotorPrecificacaoMultidirecional(ByVal modoAlvo As String)
    Dim LoPrecificacao As New clsListState: If Not LoPrecificacao.Init() Then Exit Sub

    Dim colCusto As Long, colUF as Long, colMargem As Long, colPrecoCaixa As Long, colPrecoUnit As Long, colQtd As Long, colEncargos As Long, colSKU As Long, colICMS As Long, colPIS As Long, colCOF As Long, colIPI As Long, colST As Long, colPrecoBaseCx As Long, colPrecoBaseUn As Long

    Dim caminho_db As String
    Dim str_conn As String
    Dim ado_conn As Object
    Dim rs As Object
    Dim i As Long
    Dim dictFederais As Object, dictEstaduais As Object
    Dim query_federais As String, query_estaduais As String
    Dim arrFederais As Variant, arrEstaduais As Variant

    Dim sql_sku As String, chave_composta As String
    Dim sql_impostos As Variant

    With LoPrecificacao
        colCusto = .IndiceColuna("Custo")
        colUF = .IndiceColuna("UF")
        colMargem = .IndiceColuna("Margem Bruta (%)")
        colPrecoCaixa = .IndiceColuna("Preço Final Caixa")
        colPrecoUnit = .IndiceColuna("Preço Final Unitário")
        colQtd = .IndiceColuna("Qtd Padrão")
        colEncargos = .IndiceColuna("Total Encargos (%)")
        colSKU = .IndiceColuna("Código")
        colICMS = .IndiceColuna("ICMS")
        colPIS = .IndiceColuna("Pis")
        colCOF = .IndiceColuna("Cofins")
        colIPI = .IndiceColuna("IPI")
        colST = .IndiceColuna("ICMS ST")
        colPrecoBaseCx = .IndiceColuna("P.Base - Ex IPI/ST")
        colPrecoBaseUn = .IndiceColuna("P.Base Un - Ex IPI/ST")
    End With

    If colCusto = 0 Or colMargem = 0 Or colPrecoCaixa = 0 Or colPrecoUnit = 0 Or colQtd = 0 Or colEncargos = 0 Or colSKU = 0 Or colUF = 0  Or colPrecoBaseCx = 0 Or colPrecoBaseUn = 0 Then
        MsgBox "Erro de Mapeamento: Uma ou mais colunas obrigatórias não estão presentes.", vbCritical
        Exit Sub
    End If

    Dim arrDados As Variant
    arrDados = LoPrecificacao.DataBodyPadrao.Value2
    
    If Not IsArray(arrDados) Then 
        MsgBox "A tabela está vazia ou foi reduzida a uma célula.", vbInformation
        Exit Sub
    End If
    
    Set dictFederais = CreateObject("Scripting.Dictionary")
    Set dictEstaduais = CreateObject("Scripting.Dictionary")
    dictFederais.CompareMode = 1
    dictEstaduais.CompareMode = 1

    Set ado_conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    caminho_db = ThisWorkbook.Path & "\sheets\database\Cerulean_DB_Impostos.xlsx"
    
    str_conn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Data Source=" & caminho_db & ";" & _
               "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

    query_federais = "SELECT sku_codigo_produto, aliquota_pis, aliquota_cofins, aliquota_ipi FROM [Impostos_Federais$]"
    query_estaduais = "SELECT sku_codigo_produto, uf_destino, aliquota_icms, aliquota_st FROM [Matriz_ICMS_ST$]"
    
    On Error GoTo TrataErro

    ado_conn.Open str_conn
    ' --- FEDERAIS ---
    rs.Open query_federais, ado_conn, 0, 1 ' adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then arrFederais = rs.GetRows()
    rs.Close

    If Not IsEmpty(arrFederais) Then        
        For i = 0 To UBound(arrFederais, 2)
            sql_sku = CStr(arrFederais(0, i))
            sql_impostos = Array(arrFederais(1, i), arrFederais(2, i), arrFederais(3, i))
            
            If Not dictFederais.Exists(sql_sku) Then
                dictFederais.Add sql_sku, sql_impostos
            End If
        Next i
    End If

    ' --- ESTADUAIS ---
    rs.Open query_estaduais, ado_conn, 0, 1
    If Not rs.EOF Then arrEstaduais = rs.GetRows()
    rs.Close
    ado_conn.Close

    Set rs = Nothing
    Set ado_conn = Nothing

    If Not IsEmpty(arrEstaduais) Then        
        For i = 0 To UBound(arrEstaduais, 2)
            chave_composta = CStr(arrEstaduais(0, i)) & "|" & Trim(UCase(CStr(arrEstaduais(1, i)))) 'Uf, Sku
            sql_impostos = Array(arrEstaduais(2, i), arrEstaduais(3, i))
            
            If Not dictEstaduais.Exists(chave_composta) Then
                dictEstaduais.Add chave_composta, sql_impostos
            End If
        Next i
    End If

    On Error GoTo 0

    Dim vSKU As String
    Dim vUF As String
    Dim chave_busca As String
    Dim arrFed As Variant
    Dim arrEst As Variant
    Dim Produto As clsProduto
    Dim total_linhas As Long: total_linhas = UBound(arrDados, 1)

    For i = 1 To total_linhas
        If Not Trim(CStr(arrDados(i, colCusto))) = "" Then
            vSKU = CStr(arrDados(i, colSKU))
            vUF = CStr(arrDados(i, colUF))
            chave_busca = vSKU & "|" & vUF
            
            If dictEstaduais.Exists(chave_busca) And dictFederais.Exists(vSKU) Then
                arrFed = dictFederais(vSKU)
                arrEst = dictEstaduais(chave_busca)

                Set Produto = New clsProduto
                
                Call Produto.Init(vSKU, vUf, _ 
                CDbl(arrEst(0)), _ ' icms
                CDbl(arrFed(0)), _ ' pis
                CDbl(arrFed(1)), _ ' cofins
                CDbl(arrFed(2)), _ ' ipi
                CDbl(arrEst(1))) ' St

                Call Produto.InjetarImpostosPrecificacao(arrDados, i, colIPI, colST, colICMS, colPIS, colCOF)

                Set Produto = Nothing
            End If
            Call CalcularPrecoMulti(arrDados, i, modoAlvo, colCusto, colMargem, colPrecoCaixa, colPrecoUnit, colQtd, colEncargos, colIPI, colST, colPrecoBaseCx, colPrecoBaseUn)
        End If
    Next i

    If modoAlvo = "Margem Bruta" Then
        ' Âncora: Margem. Injetamos todos os 4 Preços derivados.
        Call LoPrecificacao.InjetarLoteCirurgico(arrDados, "Preço Final Caixa", "Preço Final Unitário", "P. Base - Ex IPI/ST", "P.Base Un - Ex IPI/ST")
        
    ElseIf modoAlvo = "Preço Final Caixa" Or modoAlvo = "Preço Final Unitário" Then
        ' Âncora: Preço Final Unitário. Injetamos Margem, Preço Caixa derivado e ambos os Bases.
        Call LoPrecificacao.InjetarLoteCirurgico(arrDados, "Preço Final Caixa", "Margem Bruta (%)", "P. Base - Ex IPI/ST", "P.Base Un - Ex IPI/ST")
        
    ElseIf modoAlvo = "Preço Final Caixa" Then
        ' Âncora: Preço Final Caixa. Injetamos Margem, Preço Unitário derivado e ambos os Bases.
        Call LoPrecificacao.InjetarLoteCirurgico(arrDados, "Preço Final Unitário", "Margem Bruta (%)", "P. Base - Ex IPI/ST", "P.Base Un - Ex IPI/ST")
        
    ElseIf modoAlvo = "Preço Base Caixa" Then
        ' Âncora: Preço Base Caixa. Forja Preços Finais com impostos, ajusta a Margem e deriva o Base Unitário.
        Call LoPrecificacao.InjetarLoteCirurgico(arrDados, "Preço Final Caixa", "Preço Final Unitário", "Margem Bruta (%)", "P.Base Un - Ex IPI/ST")
        
    ElseIf modoAlvo = "Preço Base Unitário" Then
        ' Âncora: Preço Base Unitário. Forja Preços Finais, ajusta a Margem e deriva o Base Caixa.
        Call LoPrecificacao.InjetarLoteCirurgico(arrDados, "Preço Final Caixa", "Preço Final Unitário", "Margem Bruta (%)", "P. Base - Ex IPI/ST")

    End If

    Exit Sub

TrataErro:
    MsgBox "Erro de Conexão com o Banco de Dados: " & Err.Description, vbCritical
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not ado_conn Is Nothing Then If ado_conn.State = 1 Then ado_conn.Close
End Sub

Private Sub CalcularPrecoMulti(ByRef arrDados As Variant, ByVal i As Long, ByVal modoAlvo As String, colCusto As Long, colMargem As Long, colPrecoCaixa As Long, colPrecoUnit As Long, colQtd As Long, colEncargos As Long, colIPI As Long, colST As Long, colPrecoBaseCx As Long, colPrecoBaseUn As Long)
    Dim vCusto As Double, vQtd As Double
    Dim vPfu As Double, vPfc As Double, vMbp As Double
    Dim vMkp As Double
    Dim vIPI As Double, vST As Double
    Dim vPbc As Double, vPbu As Double
    Dim vEncargos_dentro As Double, vEncargos_fora As Double

    vCusto = Val(arrDados(i, colCusto))
    If vCusto <= 0 Then Exit Sub
    
    vIPI = Val(arrDados(i, colIPI))
    vST = Val(arrDados(i, colST))
    vPrecoBaseC = Val(arrDados(i, colPrecoBaseCx))
    vPrecoBaseU = Val(arrDados(i, colPrecoBaseUn))
    
    vQtd = Val(arrDados(i, colQtd))
    vEncargos_dentro = Val(arrDados(i, colEncargos))
    vMargemB = Val(arrDados(i, colMargem))
    vPrecoU = Val(arrDados(i, colPrecoUnit))
    vPrecoC = Val(arrDados(i, colPrecoCaixa))

'     If itc = "Preço Final Caixa" Then
'         ' vPfc é a entrada informada
'         vPfu = vPfc / vQtd
'         vPbc = vPfc / (1 + encargos_por_fora)
'         vPbu = vPbc / vQtd
'         vMbp = 1 - ((vCusto / vPfc) + encargos_por_dentro)

'     ElseIf itc = "Preço Final Unitário" Then
'         ' vPfu é a entrada informada
'         vPfc = vPfu * vQtd
'         vPbc = vPfc / (1 + encargos_por_fora)
'         vPbu = vPbc / vQtd
'         vMbp = 1 - ((vCusto / vPfc) + encargos_por_dentro)

'     ElseIf itc = "Preço Base Caixa" Then
'         ' vPbc é a entrada informada
'         vPfc = vPbc * (1 + encargos_por_fora)
'         vPfu = vPfc / vQtd
'         vPbu = vPbc / vQtd
'         vMbp = 1 - ((vCusto / vPfc) + encargos_por_dentro)

'     ElseIf itc = "Preço Base Unitário" Then
'         ' vPbu é a entrada informada
'         vPbc = vPbu * vQtd
'         vPfc = vPbc * (1 + encargos_por_fora)
'         vPfu = vPfc / vQtd
'         vMbp = 1 - ((vCusto / vPfc) + encargos_por_dentro)

'     ElseIf itc = "Margem Bruta %" Then
'         ' vMbp é a entrada informada
'         vMkp = 1 - (encargos_por_dentro + vMbp)
'         vPfc = vCusto / vMkp
'         vPfu = vPfc / vQtd
'         vPbc = vPfc / (1 + encargos_por_fora)
'         vPbu = vPbc / vQtd
'     End If

    If modoAlvo = "Margem Bruta" Then
        vDivisorBase = 1 - (vEncargos_dentro + vMargemB)
        If vDivisorBase <> 0 Then
            vPrecoC = vCusto / vDivisorBase
            arrDados(i, colPrecoCaixa) = vPrecoC
            If vQtd > 0 Then
                vPrecoU = vPrecoC / vQtd
                arrDados(i, colPrecoUnit) = vPrecoU
            End If
        End If
    ElseIf modoAlvo = "Preço Caixa" Then
        If vPrecoC <> 0 Then
        vMargemB = (vPrecoC - vCusto - (vPrecoC * vEncargos)) / vPrecoC
            arrDados(i, colMargem) = vMargemB
        End If
        If vQtd > 0 Then
            vPrecoU = vPrecoC / vQtd
            arrDados(i, colPrecoUnit) = vPrecoU
        End If
    ElseIf modoAlvo = "Preço Unitário" Then
        vPrecoC = vPrecoU * vQtd
        arrDados(i, colPrecoCaixa) = vPrecoC
        If vPrecoC <> 0 Then
            vMargemB = (vPrecoC - vCusto - (vPrecoC * vEncargos)) / vPrecoC
            arrDados(i, colMargem) = vMargemB
        End If
    ElseIf modoAlvo = "Preço Base Caixa" Then
        vDivisorBaseImp = 1 - (vIPI + vST)
        
        If vDivisorBaseImp <> 0 Then
            vPrecoC = vPrecoBaseC / vDivisorBaseImp
            arrDados(i, colPrecoCaixa) = vPrecoC
            
            vMargemB = (vPrecoC - vCusto - (vPrecoC * vEncargos)) / vPrecoC
            arrDados(i, colMargem) = vMargemB
            
            If vQtd > 0 Then
                vPrecoU = vPrecoC / vQtd
                arrDados(i, colPrecoUnit) = vPrecoU
            End If
        End If
    End If

    ' Despejo reverso
    arrDados(i, colPrecoBaseCx) = vPrecoC * (1 - vIPI - vST)
    arrDados(i, colPrecoBaseUn) = vPrecoU * (1 - vIPI - vST)
End Sub