VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form formDescascadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descascadores"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "formDescascadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   795
      Left            =   6720
      Picture         =   "formDescascadores.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sair"
      Top             =   2580
      Width           =   915
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   795
      Left            =   6720
      Picture         =   "formDescascadores.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Excluir usuário selecionado"
      Top             =   1740
      Width           =   915
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Cadastrar"
      Height          =   795
      Left            =   6720
      Picture         =   "formDescascadores.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Incluir novo usuário"
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Selecione um usuário existente ou clique em Cadastrar"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      TabIndex        =   28
      Top             =   240
      Width           =   6495
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   4635
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5580
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   29
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   4185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4860
         TabIndex        =   30
         Top             =   540
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Dados Cadastrais"
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   60
      TabIndex        =   16
      Top             =   1680
      Width           =   6495
      Begin VB.TextBox txtTelefone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cboSexo 
         Height          =   315
         ItemData        =   "formDescascadores.frx":30C0
         Left            =   5640
         List            =   "formDescascadores.frx":30CA
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtRG 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCPF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   60
         TabIndex        =   1
         Top             =   360
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4020
         MaxLength       =   25
         TabIndex        =   3
         Top             =   720
         Width           =   2235
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDNasc 
         Height          =   285
         Left            =   3420
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCEP 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   9
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tel."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3900
         TabIndex        =   27
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RG"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sexo"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4740
         TabIndex        =   25
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D. Nasc"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         TabIndex        =   24
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPF"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bairro"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CEP"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1740
         TabIndex        =   18
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Órgão"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3900
         TabIndex        =   17
         Top             =   1440
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdConcluir 
      Caption         =   "&Concluir"
      Enabled         =   0   'False
      Height          =   795
      Left            =   6720
      Picture         =   "formDescascadores.frx":30D4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Concluir inclusões ou alterações"
      Top             =   900
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   795
      Left            =   6720
      Picture         =   "formDescascadores.frx":399E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sair"
      Top             =   3420
      Width           =   915
   End
End
Attribute VB_Name = "formDescascadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()

    Unload Me

End Sub

Private Sub cmdConcluir_Click()
Dim dData As String
    
    'Uso essa gambiarra para evitar erros no UPDATE
    dData = IIf(IsDate(txtDNasc), txtDNasc, "")
    con.Execute ("UPDATE tbl_descascador SET telefone = '" & txtTelefone & "', nascimento = IF('" & dData & "' = '', '', '" & Format(dData, "yyyy/mm/dd") & "'), endereco = '" & txtEndereco & "', bairro = '" & txtBairro & "', cidade = '" & txtCidade & "', estado = '" & txtEstado & "', cep = '" & txtCEP & "', rg = '" & txtRG & "', orgao_rg = '" & txtOrgao & "', cpf = '" & txtCPF & "', sexo = IF('" & cboSexo.Text & "' = 'F', 0, 1) WHERE codigo = " & txtCodigo & "")
    LimpaDescascador
    Combo1.ListIndex = -1
    Combo1.Locked = False
    cmdNovo.Enabled = True
    cmdConcluir.Enabled = False
    
End Sub

Private Sub cmdExcluir_Click()
    
    If Combo1.Text = "" Then
        MsgBox "selecione um descascador para exclusão!", vbCritical
        Exit Sub
    End If
    If MsgBox("Tem certeza de que deseja excluir este Descascador?", vbYesNo + vbQuestion) = vbYes Then
        con.Execute ("DELETE FROM tbl_descascador WHERE codigo = " & txtCodigo & "")
        cmdExcluir.Enabled = False
        'Exclui do combo o item selecionado
        Combo1.RemoveItem Combo1.ListIndex
        cmdConcluir.Enabled = False
        cmdNovo.Enabled = True
        Combo1.Locked = False
        LimpaDescascador
    End If
    
End Sub

Private Sub cmdImprimir_Click()

    Me.Hide
    formPrincipal.Enabled = False
    FormRelatorio1.Config
    formPrincipal.Enabled = True
    Me.Show 1

End Sub

Private Sub cmdNovo_Click()
        
    Titulo = "Descascadores"
    Mensagem = "Digite o nome do descascador."
    Caixa = InputBox(Mensagem, Titulo)
    If Caixa <> "" Then
        'Verifica se o usuario está cadastrado
        Set RS = con.Execute("SELECT * FROM tbl_descascador WHERE nome = '" & Caixa & "'")
        If RS.EOF = False Then
            If MsgBox("Já existe um descascador cadastrado com esse nome! Deseja prosseguir?", vbYesNo + vbQuestion) = vbNo Then
                RS.Close
                Exit Sub
            End If
        End If
        RS.Close
        'con.Execute "INSERT INTO tbl_descascador (nome) values('" & Caixa & "')"
        con.Execute "INSERT IGNORE INTO TBL_DESCASCADOR (NOME, CODIGO) SELECT '" & Caixa & "', CASE WHEN MAX(CODIGO) IS NULL THEN '1' ELSE MAX(CODIGO) +1 END FROM TBL_DESCASCADOR"
        'Coloca o nome do descascador no combo
        Combo1.AddItem Caixa
        'Adiciona seu nome na ultima posição do combo
        Combo1.ListIndex = Combo1.ListCount - 1
        Combo1.Locked = True
        cmdNovo.Enabled = False
        cmdConcluir.Enabled = True
    End If

End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub Combo1_Click()
    
    Me.MousePointer = vbHourglass
    If Combo1.Text <> "" Then
        Set RS = con.Execute("SELECT * FROM tbl_descascador WHERE nome = '" & Combo1.Text & "'")
        PreencheCampos
        RS.Close
        cmdExcluir.Enabled = True
        cmdConcluir.Enabled = True
        txtEndereco.SetFocus
    End If
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    
    Set RS = con.Execute("SELECT nome FROM tbl_descascador ORDER BY nome")
    If RS.EOF = False Then
        While Not RS.EOF
            Combo1.AddItem RS!nome
            RS.MoveNext
        Wend
    End If
    RS.Close

End Sub

Private Sub PreencheCampos()
    
    If Not IsNull(RS!nascimento) Then txtDNasc.Text = Format(RS!nascimento, "DD/MM/YYYY")
    txtEndereco.Text = RS!endereco & ""
    txtBairro.Text = RS!bairro & ""
    txtCidade.Text = RS!cidade & ""
    txtEstado.Text = RS!estado & ""
    txtTelefone.Text = RS!telefone & ""
    txtCEP.AllowPrompt = True
    If RS!cep <> "" Then txtCEP.Text = RS!cep
    txtCEP.AllowPrompt = False
    txtRG.Text = RS!RG & ""
    txtOrgao.Text = RS!orgao_rg & ""
    txtCPF.Text = RS!CPF & ""
    cboSexo.ListIndex = IIf(RS!sexo = 1, 0, 1)
    txtCodigo = RS!codigo & ""
            
End Sub

Private Sub LimpaDescascador()
Dim X As Control

    For Each X In formDescascadores
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
    formDescascadores.cboSexo.ListIndex = -1
    formDescascadores.txtDNasc = "__/__/____"
    formDescascadores.txtCEP = "_____-___"
    
End Sub

