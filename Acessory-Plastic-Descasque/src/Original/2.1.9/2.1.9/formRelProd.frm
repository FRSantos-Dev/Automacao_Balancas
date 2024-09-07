VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form formRelProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios de Produção"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "formRelProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      ScaleHeight     =   1005
      ScaleWidth      =   3945
      TabIndex        =   7
      Top             =   1260
      Width           =   3975
      Begin MSMask.MaskEdBox txtDataInicial 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDataFinal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   540
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Data Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Data Inicial"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   3945
      TabIndex        =   6
      Top             =   180
      Width           =   3975
      Begin VB.ComboBox cboProduto 
         Height          =   315
         ItemData        =   "formRelProd.frx":030A
         Left            =   1200
         List            =   "formRelProd.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboTipoRela 
         Height          =   315
         ItemData        =   "formRelProd.frx":030E
         Left            =   1200
         List            =   "formRelProd.frx":0327
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   60
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Produto"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Relatório"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Gráfico"
      Height          =   795
      Left            =   4440
      Picture         =   "formRelProd.frx":03BF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
   Begin VB.CommandButton cmdRelatorio 
      Caption         =   "Relatório"
      Default         =   -1  'True
      Height          =   795
      Left            =   4440
      Picture         =   "formRelProd.frx":06C9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   795
      Left            =   4440
      Picture         =   "formRelProd.frx":0F93
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   915
   End
End
Attribute VB_Name = "formRelProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Referencias As Variant
Private Sub cboTipoRela_Click()
    
    If cboTipoRela.Text <> "" Then
        If cboTipoRela.ListIndex = 1 Then
            Label2.Visible = True
            cboProduto.Visible = True
        Else
            Label2.Visible = False
            cboProduto.Visible = False
        End If
    End If
    
    txtDataInicial.SetFocus

End Sub

Private Sub cmdGrafico_Click()
    
    MsgBox "Este recurso ainda não foi implementado!"
    Exit Sub
    formGrafico.Show 1

End Sub

Private Sub cmdRelatorio_Click()
    
    If cboTipoRela.Text = "" Then
        MsgBox "Selecione um Tipo de relatório!", vbCritical
        Exit Sub
    End If
    If txtDataInicial = "__/__/____" Or txtDataFinal = "__/__/____" Then
        MsgBox "Digite um intervalo de datas!", vbCritical
        Exit Sub
    End If
    If txtDataInicial = "__/__/____" Or Not IsDate(txtDataInicial) Or txtDataFinal = "__/__/____" Or Not IsDate(txtDataFinal) Then
        MsgBox "Digite um intervalo de datas válido!", vbCritical
        Exit Sub
    End If
    
    DataIni = txtDataInicial
    DataFim = txtDataFinal
    Me.Hide
    DoEvents
    Select Case cboTipoRela.Text
    Case "Todas as Máquinas"
        formAllMachinesReport.Config
    Case "Produtos"
        If cboProduto <> "TODOS" Then
            CRITERIO = "REFERENCIA = '" & cboProduto.Text & "' "
        Else
            CRITERIO = ""
        End If
        formProductReport.Config
    Case "Todos os Produtos"
        FormRelaProducao2.Config
    Case "Máquinas Detalhadas"
        formDetailedMachinesReport.Config
    Case "Matéria-Prima Analítico"
        formRelaProducao5.Config
    Case "Matéria-Prima Sintético"
        formRelaMPSintetico.Config
    Case "Diferença Devolução X Galha"
        'formRelaGalha.Config
        MsgBox "Em Construção", vbExclamation
    End Select
    Me.Show 1

End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim TMP() As String
Dim i As Integer
            
    DoDates txtDataInicial, txtDataFinal
    
    'Preenche combo Produtos
    'VERIFICA SE A TABELA DE REFERENCIAS ESTÁ VAZIA. SE ESTIVER, NÃO PREENCHE O COMBO
    If (RSReferencias.EOF And RSReferencias.BOF) Then
        Exit Sub
    End If
    
    'PREENCHE O COMBO
    RSReferencias.MoveFirst
    cboProduto.AddItem "TODOS"
    While Not RSReferencias.EOF
        cboProduto.AddItem RSReferencias!REFERENCIA
        RSReferencias.MoveNext
    Wend
    cboProduto.ListIndex = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.Hide
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'formPrincipal.Enabled = True
    'MDIForm1.Enabled = True
    formProducao.Show
    'formPrincipal.SetFocus

End Sub
