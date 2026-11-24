' ==============================================================================
' UserForm: frmAdicionarGrupo
' Camada: UI (apenas apresentação)
' Finalidade: Escolher um grupo de produtos para inclusão em lote.
'   - Recebe a lista pronta (DefinirGrupos) e devolve a escolha
'     (Confirmado / GrupoSelecionado). NÃO consulta base, NÃO escreve na
'     planilha: isso fica em mod_PricingCatalogo e nas camadas puras.
'   - Os controles são criados em tempo de execução (mesmo padrão do
'     frmPainelCerulean), então o formulário não depende de um arquivo .frx.
' Instalação: no Editor VBA, Inserir > UserForm; na janela Propriedades, mude
'   (Name) para frmAdicionarGrupo; cole este código no módulo do formulário.
' ==============================================================================
Option Explicit

Private Const COR_CERULEO As Long = 7488020          ' RGB(20, 66, 114)
Private Const MARGEM As Single = 12

Private WithEvents txtFiltro As MSForms.TextBox
Private WithEvents lstGrupos As MSForms.ListBox
Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancelar As MSForms.CommandButton
Private lblContagem As MSForms.Label

Private m_Grupos As Variant
Private m_Confirmado As Boolean
Private m_GrupoSelecionado As String

' ------------------------------------------------------------------------------
' Interface pública (usada por Pricing_AbrirAdicionarGrupo)
' ------------------------------------------------------------------------------
Public Sub DefinirGrupos(ByVal grupos As Variant)
    m_Grupos = grupos
    Call AplicarFiltro(vbNullString)
End Sub

Public Property Get Confirmado() As Boolean
    Confirmado = m_Confirmado
End Property

Public Property Get GrupoSelecionado() As String
    GrupoSelecionado = m_GrupoSelecionado
End Property

' ------------------------------------------------------------------------------
' Montagem da tela
' ------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim lbl As MSForms.Label
    Dim largura As Single

    Me.Caption = "Adicionar Grupo de Produtos"
    Me.Width = 340
    Me.Height = 380
    largura = Me.InsideWidth - 2 * MARGEM

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblFiltro")
    With lbl
        .Caption = "Filtrar grupos:"
        .Left = MARGEM: .Top = MARGEM: .Width = largura: .Height = 14
        .Font.Name = "Aptos": .Font.Size = 9
    End With

    Set txtFiltro = Me.Controls.Add("Forms.TextBox.1", "txtFiltro")
    With txtFiltro
        .Left = MARGEM: .Top = MARGEM + 16: .Width = largura: .Height = 20
        .Font.Name = "Aptos": .Font.Size = 10
    End With

    Set lstGrupos = Me.Controls.Add("Forms.ListBox.1", "lstGrupos")
    With lstGrupos
        .Left = MARGEM: .Top = MARGEM + 42: .Width = largura: .Height = 220
        .Font.Name = "Aptos": .Font.Size = 10
    End With

    Set lblContagem = Me.Controls.Add("Forms.Label.1", "lblContagem")
    With lblContagem
        .Left = MARGEM: .Top = lstGrupos.Top + lstGrupos.Height + 4: .Width = largura: .Height = 14
        .Font.Name = "Aptos": .Font.Size = 8
        .ForeColor = RGB(90, 90, 90)
    End With

    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar")
    With btnCancelar
        .Caption = "Cancelar"
        .Width = 80: .Height = 26
        .Left = Me.InsideWidth - MARGEM - .Width
        .Top = lblContagem.Top + 22
        .Cancel = True
        .Font.Name = "Aptos": .Font.Size = 9
    End With

    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With btnOK
        .Caption = "OK"
        .Width = 80: .Height = 26
        .Left = btnCancelar.Left - 8 - .Width
        .Top = btnCancelar.Top
        .Default = True
        .BackColor = COR_CERULEO
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Aptos": .Font.Size = 9: .Font.Bold = True
    End With

    Me.Height = Me.Height - Me.InsideHeight + btnOK.Top + btnOK.Height + MARGEM
End Sub

' ------------------------------------------------------------------------------
' Filtro por texto (contém, sem diferenciar maiúsculas)
' ------------------------------------------------------------------------------
Private Sub AplicarFiltro(ByVal texto As String)
    Dim g As Variant

    lstGrupos.Clear
    If IsArray(m_Grupos) Then
        For Each g In m_Grupos
            If Len(texto) = 0 Or InStr(1, CStr(g), texto, vbTextCompare) > 0 Then
                lstGrupos.AddItem CStr(g)
            End If
        Next g
    End If

    lblContagem.Caption = lstGrupos.ListCount & " grupo(s)"
    If lstGrupos.ListCount = 1 Then lstGrupos.ListIndex = 0
End Sub

Private Sub txtFiltro_Change()
    Call AplicarFiltro(Trim$(txtFiltro.Text))
End Sub

' ------------------------------------------------------------------------------
' Confirmação / cancelamento (o formulário só se esconde; quem chamou lê a
' escolha e descarrega)
' ------------------------------------------------------------------------------
Private Sub Confirmar()
    If lstGrupos.ListIndex < 0 Then
        MsgBox "Selecione um grupo na lista.", vbExclamation
        Exit Sub
    End If
    m_GrupoSelecionado = CStr(lstGrupos.List(lstGrupos.ListIndex))
    m_Confirmado = True
    Me.Hide
End Sub

Private Sub btnOK_Click()
    Call Confirmar
End Sub

Private Sub lstGrupos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call Confirmar
End Sub

Private Sub btnCancelar_Click()
    m_Confirmado = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        m_Confirmado = False
        Me.Hide
    End If
End Sub
