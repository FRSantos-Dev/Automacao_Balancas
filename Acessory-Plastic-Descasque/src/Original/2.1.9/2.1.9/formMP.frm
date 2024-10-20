VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form formMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matéria Prima"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "formMP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEstoque 
      Caption         =   "&Estoque"
      Height          =   795
      Left            =   5160
      Picture         =   "formMP.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sair"
      Top             =   1080
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Cadastro de Matéria Prima"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      TabIndex        =   5
      Top             =   180
      Width           =   4935
      Begin VB.TextBox txtMP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   60
         TabIndex        =   0
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Matéria Prima"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   795
      Left            =   5160
      Picture         =   "formMP.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sair"
      Top             =   1920
      Width           =   915
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   795
      Left            =   5160
      Picture         =   "formMP.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Incluir nova Matéria Prima"
      Top             =   240
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1635
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1140
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2884
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      HighLight       =   2
      Appearance      =   0
      FormatString    =   "Matéria Prima                                                      |Estoque          "
   End
End
Attribute VB_Name = "formMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastrar_Click()

On Error GoTo TrataErro

    If txtMP <> "" Then
        con.Execute "INSERT INTO TBL_MP (NOME) values('" & txtMP.Text & "')"
    Else
        MsgBox "É necessário preencher o nome da matéria prima!", vbExclamation
        Exit Sub
    End If
    Grid1.Rows = 1
    PreencheGrid
    txtMP.Text = ""
    txtMP.SetFocus
    Exit Sub
    
'Caso o usuário tente inserir uma MP já existente no BD (campo unique. mais rápido que uma consulta)
TrataErro:
    If Err.Number = -2147217900 Then
        MsgBox "Nome de Matéria Prima já existente", vbExclamation
    Else
        MsgBox "Erro: " & Err.Number & "." & Err.Description
        End
    End If

End Sub

Private Sub cmdEstoque_Click()

    formEstoque.Show 1

End Sub

Private Sub cmdSair_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    PreencheGrid

End Sub

Private Sub Grid1_Click()

    Grid1.CellBackColor = vbBlue
    Grid1.CellForeColor = vbWhite

End Sub

Private Sub Grid1_LeaveCell()

    Grid1.CellBackColor = &H80000005
    Grid1.CellForeColor = &H80000012
    
End Sub

Public Sub PreencheGrid()
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenForwardOnly

    'CRIEI UM CAMPO ESTOQUE NA TABELA MP PARA ARMAZENAR O TOTAL DO ESTOQUE DE CADA MP
    'É MENOS ELEGANTE, MAS É MUITO MAIS RÁPIDO
'    Set RS = con.Execute("SELECT nome, SUM(peso) AS SOMA FROM tbl_mp LEFT JOIN tbl_entrada_mp ON tbl_mp.ID=tbl_entrada_mp.cod_mp GROUP BY nome")
'    RS2.Open "SELECT nome, SUM(MP_CONSUMIDA) AS SOMA2 FROM tbl_mp LEFT JOIN tbl_DADOS ON tbl_mp.ID=tbl_DADOS.cod_mp GROUP BY nome", con
    RS.Open "SELECT * FROM TBL_MP GROUP BY nome", con
    While Not RS.EOF
        Grid1.AddItem RS!nome & vbTab & Round(CDbl(IIf((Left(RS!ESTOQUE, 1) = "-"), (RS!ESTOQUE), (0 & RS!ESTOQUE))), 3)
        RS.MoveNext
    Wend
    RS.Close

End Sub

