' ==============================================================================
' UserForm: frmPacoteCustos                                          [spec-009]
' Camada: UI (apenas apresentação)
' Finalidade: Escolher o pacote de custos (aba de Cerulean_DB_Custos.xlsx) que
'             será usado na precificação.
'   - Recebe a lista pronta e o pacote ativo (DefinirPacotes) e devolve a
'     escolha (Confirmado / PacoteSelecionado). NÃO consulta base, NÃO grava o
'     estado, NÃO mexe no cabeçalho: isso fica em mod_PricingPacoteCustos e no
'     serviço puro mod_Pricing_CostPackageService.
'   - Segue o padrão visual do frmAdicionarGrupo: mesmos tamanhos, fonte, cor do
'     botão principal e controles criados em tempo de execução (sem arquivo .frx).
' Instalação: no Editor VBA, Inserir > UserForm; na janela Propriedades, mude
'   (Name) para frmPacoteCustos; cole este código no módulo do formulário.
' ==============================================================================
Option Explicit

Private Const COR_CERULEO As Long = 7488020          ' RGB(20, 66, 114)
Private Const MARGEM As Single = 12

Private WithEvents lstPacoteCustos As MSForms.ListBox
Private WithEvents btnConfirmar As MSForms.CommandButton
Private WithEvents btnCancelar As MSForms.CommandButton

Private m_Confirmado As Boolean
Private m_PacoteSelecionado As String

' ------------------------------------------------------------------------------
' Interface pública (usada por Pricing_AbrirSelecionarPacoteCustos)
' ------------------------------------------------------------------------------
' pacotes: Array de String; pacoteAtivo: se estiver na lista, já vem selecionado
Public Sub DefinirPacotes(ByVal pacotes As Variant, Optional ByVal pacoteAtivo As String = vbNullString)
    Dim p As Variant
    Dim i As Long

    lstPacoteCustos.Clear
    If IsArray(pacotes) Then
        For Each p In pacotes
            lstPacoteCustos.AddItem CStr(p)
        Next p
    End If

    ' Feedback imediato: seleciona o pacote ativo (ou o único da lista)
    If Len(pacoteAtivo) > 0 Then
        For i = 0 To lstPacoteCustos.ListCount - 1
            If StrComp(CStr(lstPacoteCustos.List(i)), pacoteAtivo, vbTextCompare) = 0 Then
                lstPacoteCustos.ListIndex = i
                Exit For
            End If
        Next i
    End If
    If lstPacoteCustos.ListIndex < 0 And lstPacoteCustos.ListCount = 1 Then lstPacoteCustos.ListIndex = 0
End Sub

Public Property Get Confirmado() As Boolean
    Confirmado = m_Confirmado
End Property

Public Property Get PacoteSelecionado() As String
    PacoteSelecionado = m_PacoteSelecionado
End Property

' ------------------------------------------------------------------------------
' Montagem da tela
' ------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim lbl As MSForms.Label
    Dim largura As Single

    Me.Caption = "Selecionar Pacote de Custos"
    Me.Width = 340
    Me.Height = 330
    largura = Me.InsideWidth - 2 * MARGEM

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblInstrucao")
    With lbl
        .Caption = "Selecione a aba de fechamento de custos (Cerulean_DB_Custos):"
        .Left = MARGEM: .Top = MARGEM: .Width = largura: .Height = 28
        .WordWrap = True
        .Font.Name = "Aptos": .Font.Size = 9
    End With

    Set lstPacoteCustos = Me.Controls.Add("Forms.ListBox.1", "lstPacoteCustos")
    With lstPacoteCustos
        .Left = MARGEM: .Top = MARGEM + 34: .Width = largura: .Height = 190
        .Font.Name = "Aptos": .Font.Size = 10
    End With

    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar")
    With btnCancelar
        .Caption = "Cancelar"
        .Width = 80: .Height = 26
        .Left = Me.InsideWidth - MARGEM - .Width
        .Top = lstPacoteCustos.Top + lstPacoteCustos.Height + 12
        .Cancel = True
        .Font.Name = "Aptos": .Font.Size = 9
    End With

    Set btnConfirmar = Me.Controls.Add("Forms.CommandButton.1", "btnConfirmar")
    With btnConfirmar
        .Caption = "Confirmar"
        .Width = 80: .Height = 26
        .Left = btnCancelar.Left - 8 - .Width
        .Top = btnCancelar.Top
        .Default = True
        .BackColor = COR_CERULEO
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Aptos": .Font.Size = 9: .Font.Bold = True
    End With

    Me.Height = Me.Height - Me.InsideHeight + btnConfirmar.Top + btnConfirmar.Height + MARGEM
End Sub

' ------------------------------------------------------------------------------
' Confirmação / cancelamento (o formulário só se esconde; quem chamou lê a
' escolha e descarrega)
' ------------------------------------------------------------------------------
Private Sub Confirmar()
    If lstPacoteCustos.ListIndex < 0 Then
        MsgBox "Selecione um pacote de custos na lista.", vbExclamation
        Exit Sub
    End If
    m_PacoteSelecionado = CStr(lstPacoteCustos.List(lstPacoteCustos.ListIndex))
    m_Confirmado = True
    Me.Hide
End Sub

Private Sub btnConfirmar_Click()
    Call Confirmar
End Sub

Private Sub lstPacoteCustos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
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
