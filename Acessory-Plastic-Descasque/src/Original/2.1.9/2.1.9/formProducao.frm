VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form formProducao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produção"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "formProducao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFichas 
      Caption         =   "Fichas"
      Height          =   855
      Left            =   6900
      Picture         =   "formProducao.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   6900
      Picture         =   "formProducao.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   915
   End
   Begin VB.CommandButton cmdRelatorios 
      Caption         =   "R&elatórios"
      Height          =   855
      Left            =   6900
      Picture         =   "formProducao.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   60
      ScaleHeight     =   1245
      ScaleWidth      =   6495
      TabIndex        =   23
      Top             =   4800
      Width           =   6525
      Begin VB.ComboBox cboMoldes 
         Height          =   315
         ItemData        =   "formProducao.frx":5558
         Left            =   1020
         List            =   "formProducao.frx":555A
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   900
         Width           =   1515
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   855
         Left            =   5520
         Picture         =   "formProducao.frx":555C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   915
      End
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
         Left            =   1020
         TabIndex        =   20
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
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
         Left            =   1020
         TabIndex        =   21
         Top             =   540
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblRelatorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produto"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblMP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   408
         Left            =   3696
         TabIndex        =   29
         Top             =   180
         Width           =   1752
      End
      Begin VB.Label lblProducao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   408
         Left            =   3696
         TabIndex        =   28
         Top             =   660
         Width           =   1752
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total De MP"
         ForeColor       =   &H00FFFFFF&
         Height          =   408
         Left            =   2676
         TabIndex        =   27
         Top             =   180
         Width           =   1032
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total De Produção"
         ForeColor       =   &H00FFFFFF&
         Height          =   408
         Left            =   2676
         TabIndex        =   26
         Top             =   660
         Width           =   1032
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data Inicial"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   6900
      Picture         =   "formProducao.frx":5866
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4260
      Width           =   915
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Default         =   -1  'True
      Height          =   855
      Left            =   6900
      Picture         =   "formProducao.frx":6130
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   180
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Entrada de Produção"
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   60
      TabIndex        =   31
      Top             =   120
      Width           =   6705
      Begin VB.TextBox txtMPPrin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox txtMPSec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   10
         Top             =   1380
         Width           =   1065
      End
      Begin VB.ComboBox cboMPSec 
         Height          =   315
         ItemData        =   "formProducao.frx":69FA
         Left            =   1260
         List            =   "formProducao.frx":69FC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3180
         Width           =   4935
      End
      Begin VB.TextBox txtTempo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   12
         Top             =   2100
         Width           =   1065
      End
      Begin VB.TextBox txtCiclos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   13
         Top             =   2460
         Width           =   1065
      End
      Begin VB.ComboBox cboMP 
         Height          =   315
         ItemData        =   "formProducao.frx":69FE
         Left            =   1260
         List            =   "formProducao.frx":6A00
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2820
         Width           =   4935
      End
      Begin VB.TextBox txtMolde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   0
         Top             =   300
         Width           =   1635
      End
      Begin VB.ComboBox cboTurno 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "formProducao.frx":6A02
         Left            =   1260
         List            =   "formProducao.frx":6A0C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1380
         Width           =   1635
      End
      Begin VB.TextBox txtCavidades 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   44
         Top             =   1740
         Width           =   495
      End
      Begin VB.TextBox txtProducao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   660
         Width           =   1065
      End
      Begin VB.TextBox txtInjecoes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   4
         Top             =   2100
         Width           =   1065
      End
      Begin VB.TextBox txtMPConsumida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1740
         Width           =   1065
      End
      Begin VB.ComboBox cboMaquina 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "formProducao.frx":6A1E
         Left            =   5160
         List            =   "formProducao.frx":6A9A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1260
         MaxLength       =   140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3540
         Width           =   4935
      End
      Begin MSMask.MaskEdBox txtData 
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
         Left            =   1260
         TabIndex        =   1
         Top             =   660
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHora 
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
         Left            =   1260
         TabIndex        =   2
         Top             =   1020
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label24 
         Caption         =   "Kg"
         Height          =   255
         Left            =   6240
         TabIndex        =   53
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Kg"
         Height          =   255
         Left            =   6240
         TabIndex        =   52
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Kg"
         Height          =   255
         Left            =   6240
         TabIndex        =   51
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MP Princ. Consumida"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   50
         Top             =   1020
         Width           =   1635
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MP Sec. Consumida"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   49
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MP Secundária"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   48
         Top             =   3180
         Width           =   1170
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora entrada"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tempo de Inj."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   46
         Top             =   2100
         Width           =   1635
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total de Ciclos"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   45
         Top             =   2460
         Width           =   1635
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MP Principal"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   42
         Top             =   2820
         Width           =   1170
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Molde"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cavidades"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produção"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   36
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Injeções"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MP Total Consumida"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   34
         Top             =   1740
         Width           =   1635
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Máquina"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   33
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observação"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   3540
         Width           =   1155
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Total Semanal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   2700
      TabIndex        =   41
      Top             =   4500
      Width           =   1530
   End
End
Attribute VB_Name = "formProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Turno As Byte
Dim dData As Date
Dim Peso As Currency



Private Sub cboMPSec_Click()
    
    'HABILITA E DESABILITA MP SEC, PARA EVITAR QUE O CAMPO SEJA PREENCHIDO ERRADAMENTE
    If cboMPSec.Text <> "" Then
        txtMPSec.Enabled = True
    Else
        txtMPSec.Text = ""
        txtMPSec.Enabled = False
    End If
    cboMaquina.SetFocus

End Sub

Private Sub cmdAtualizar_Click()
        
    If cboMoldes.Text = "" Then
        MsgBox "Selecione um Molde antes de prosseguir!"
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    TotalSemanal
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdExcluir_Click()
Dim MPPrinCons As String
Dim MPSecCons As String
Dim CodigoMP As String
Dim CodigoMPSec As String

    'EXCLUI A ENTRADA DE PRODUÇÃO, DEFINIDA PELOS CAMPOS PREENCHIDOS
    If txtMolde = "" Then
        MsgBox "Preencha o campo Molde!", vbCritical
        Exit Sub
    End If
    If cboTurno = "" Then
        MsgBox "Selecione o campo Turno!", vbCritical
        Exit Sub
    End If
    If txtData = "__/__/____" Then
        MsgBox "Preencha o campo data!", vbCritical
        Exit Sub
    End If
    If cboMaquina.Text = "" Then
        MsgBox "Preencha o campo Máquina!", vbCritical
        Exit Sub
    End If
    If cboMP.Text = "" Then
        MsgBox "Preencha o campo Matéria Prima!", vbCritical
        Exit Sub
    End If
    
    'O CAMPO MP SECUNDARIA NÃO É OBRIGATÓRIO
    
    dData = txtData
    Turno = IIf((cboTurno.Text = "Manhã"), 0, 1)
    CodigoMP = CLng(Left(cboMP, 3))
    If cboMPSec.Text <> "" Then CodigoMPSec = CLng(Left(cboMPSec, 3))
    
    If MsgBox("Será excluído qualquer lançamento que corresponder aos dados digitados! Tem certeza de que deseja prosseguir?", vbYesNo + vbCritical) = vbYes Then
        'COLETA A MP_PRIN_CONSUMIDA E MP_SEC_CONSUMIDA NO BD, PARA DAR RETORNO NO ESTOQUE, E VERIFICA SE JÁ FOI FEITO O LANCAMENTO
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT MP_PRIN_CONSUMIDA, MP_SEC_CONSUMIDA FROM TBL_DADOS WHERE data = '" & Format(dData, "yyyy/mm/dd") & "' AND turno = " & Turno & " and maquina = " & cboMaquina.Text & " and referencia = '" & txtMolde & "'", con
        If RS.EOF = False Then
            MPPrinCons = CDbl(RS!MP_PRIN_CONSUMIDA)
            MPSecCons = CDbl(RS!MP_Sec_Consumida)
            RS.Close
            con.Execute "DELETE FROM TBL_DADOS WHERE REFERENCIA = '" & txtMolde & "' AND data = '" & Format(dData, "YYYY/MM/DD") & "' AND turno = " & Turno & " and maquina = " & cboMaquina.Text & ""
            'LANÇA DE VOLTA NO ESTOQUE A MP PRINCIPAL UTILIZADA
            con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE + '" & FormatFloatForDB(MPPrinCons) & "') WHERE ID = '" & CodigoMP & "'"
            'LANÇA DE VOLTA NO ESTOQUE A MP PRINCIPAL UTILIZADA
            If MPSecCons > 0 Then
                con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE + '" & FormatFloatForDB(MPSecCons) & "') WHERE ID = '" & CodigoMPSec & "'"
            End If
            MsgBox "O registro foi excluído do Banco de Dados!"
        Else
            MsgBox "Nenhum registro que corresponda a consulta foi localizado!"
        End If
        LimpaCampos
    End If
    
End Sub

Private Sub cmdFicha_Click()

    formFichas.Show 1

End Sub

Private Sub cmdFichas_Click()

    Me.Hide
    formFichas.Show 1
    
End Sub

Private Sub cmdRelatorios_Click()
    
    Me.Hide
    formRelProd.Show 1
    
End Sub

Private Sub cmdSair_Click()

    Unload Me
    formPrincipal.SetFocus
    
End Sub

Private Sub cmdSalvar_Click()
Dim CodigoMP As Integer
Dim CodigoMPPrin As Integer
Dim CodigoMPSec As Integer
Dim MPCons As String
Dim MPPrinCons As String
Dim MPSecCons As String
Dim sHora As String
Dim sRef As String

    If txtMolde = "" Then
        MsgBox "Preencha o campo Molde!", vbCritical
        Exit Sub
    End If
    If cboTurno = "" Then
        MsgBox "Selecione o campo Turno!", vbCritical
        Exit Sub
    End If
    If txtData = "__/__/____" Then
        MsgBox "Preencha o campo data!", vbCritical
        Exit Sub
    End If
    If txtCavidades = "" Then
        MsgBox "Preencha o campo Cavidades!", vbCritical
        Exit Sub
    End If
    If txtInjecoes = "" Then
        MsgBox "Preencha o campo Injeções!", vbCritical
        Exit Sub
    End If
    If cboMaquina.Text = "" Then
        MsgBox "Preencha o campo Máquina!", vbCritical
        Exit Sub
    End If
    If txtTempo = "" Then
        MsgBox "Preencha o campo Tempo de Injeção!", vbCritical
        Exit Sub
    End If
    If txtCiclos = "" Then
        MsgBox "Preencha o campo Total de Ciclos!", vbCritical
        Exit Sub
    End If
    If cboMaquina.Text = "" Then
        MsgBox "Preencha o campo Máquina!", vbCritical
        Exit Sub
    End If
    If cboMP.Text = "" Then
        MsgBox "Preencha o campo Matéria Prima!", vbCritical
        Exit Sub
    End If
    
    dData = txtData
    Turno = IIf((cboTurno.Text = "Manhã"), 0, 1)
    If cboMP.Text <> "" Then CodigoMP = CLng(Left(cboMP, 3))
    If cboMPSec.Text <> "" Then CodigoMPSec = CLng(Left(cboMPSec, 3))
       
    MPCons = CDbl(txtMPConsumida)
    
    'MP Principal Consumida
    MPPrinCons = CDbl(0 & txtMPPrin)
    
    'MP Secundaria Consumida
    MPSecCons = CDbl(0 & txtMPSec)
    
    'FORMATA A HORA PARA SER GRAVADA CORRETAMENTE
    sHora = txtHora
      
    'Referência que será passada para CheckIfExist
    sRef = txtMolde
    
    'Fazer o calculo para confirmar se a soma de Prin e Sec dá a MP Total, somente se a MPSec estiver populada
    If cboMPSec.Text <> "" Then
        If (CDbl(MPPrinCons) + CDbl(MPSecCons)) <> CDbl(MPCons) Then
            MsgBox "O Somatório de MP Principal e MP Secundária está diferindo do valor de MP Total Consumida. Verifique e tente novamente!" & _
            vbCrLf & " Matéria Prima Principal = " & MPPrinCons & " Kg" & _
            vbCrLf & " Matéria Prima Secundária = " & MPSecCons & " Kg" & _
            vbCrLf & " Matéria Prima Total = " & MPCons & " Kg" & _
            vbCrLf & " Diferença = " & (CDbl(MPCons)) - (CDbl(MPPrinCons) + CDbl(MPSecCons)) & " Kg", vbExclamation
            Exit Sub
        End If
    Else
        'ATRIBUI O VALOR DE MPTOTAL A MPPRINCIPAL, POIS O USUÁRIO PODE PREENCHER ERRADO OU SEQUER PREENCHER.
        If txtMPPrin <> txtMPConsumida Then
            MsgBox "O valor de MP Principal foi corrigido, pois  estava diferente do Valor de MP Total Consumida!", vbExclamation
            MPPrinCons = MPCons
        End If
    End If
    'VERIFICA SE JÁ FOI FEITO ESTE LANCAMENTO. SE FOI, QUESTIONA SE QUER ALTERAR OS VALORES
    If CheckIfExist(CodigoMP, CodigoMPSec, MPCons, MPPrinCons, MPSecCons, sHora, sRef) = True Then
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    
    'INSERE O LANÇAMENTO DOS DADOS
    con.Execute "INSERT INTO tbl_dados (referencia, Turno, Data, Cavidades, Injecoes, maquina, Producao, obs, cod_mp, cod_mp_sec, MP_consumida, MP_Prin_Consumida, MP_Sec_Consumida, Hora, Ciclos, Tempo_Injecao) Values('" & txtMolde & "', '" & Turno & "', '" & Format(dData, "yyyy/mm/dd") & "', " & txtCavidades & ", " & txtInjecoes & ", " & cboMaquina.Text & ", " & 0 & FormatFloatForDB(txtProducao) & ", '" & txtObs & "', " & CodigoMP & ", " & CodigoMPSec & ", '" & FormatFloatForDB(MPCons) & "',  '" & FormatFloatForDB(MPPrinCons) & "',  '" & FormatFloatForDB(MPSecCons) & "', '" & sHora & "', " & txtCiclos & ", '" & txtTempo & "')"
    'DÁ BAIXA NO ESTOQUE DA MP PRINCIPAL UTILIZADA
    con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE - '" & FormatFloatForDB(MPPrinCons) & "') WHERE ID = '" & CodigoMP & "'"
    'DÁ BAIXA NO ESTOQUE DA MP SECUNDÁRIA UTILIZADA
    If cboMPSec.Text <> "" Then
        con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE - '" & FormatFloatForDB(MPSecCons) & "') WHERE ID = '" & CodigoMPSec & "'"
    End If

    LimpaCampos
    txtMolde.SetFocus
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    
    formPrincipal.Enabled = False
    MDIForm1.Enabled = False
    
    'USO A TABELA DE REFERENCIAS PARA FAZER ALGUMAS ROTINAS.
    'USEI O REQUERY PARA PEGAR A TABELA TOTALMENTE ATUALIZADA
    RSReferencias.Requery
    
    'Alimenta referencias
    PreencheMoldes
    
    'Preenche o combo com as MP's
    Carrega_MP
    
    txtData = DateAdd("d", -1, Date)
    DoDates txtDataInicial, txtDataFinal
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    formPrincipal.Enabled = True
    MDIForm1.Enabled = True

End Sub

Private Sub txtCavidades_LostFocus()

    Calcula

End Sub

Private Sub txtCiclos_KeyPress(KeyAscii As Integer)

    KeyAscii = OnlyNumbers(txtInjecoes, KeyAscii)

End Sub

Private Sub txtInjecoes_KeyPress(KeyAscii As Integer)

    KeyAscii = OnlyNumbers(txtInjecoes, KeyAscii)

End Sub

Private Sub txtInjecoes_LostFocus()

    Calcula
    
End Sub

Private Sub txtMolde_LostFocus()
    
    'LOCALIZA E EXIBE OS DADOS DO MOLDE
    If txtMolde = "" Then Exit Sub
    RSReferencias.MoveFirst
    RSReferencias.Find "referencia = '" & CStr(txtMolde) & "'"
    'Se não achar, exibe uma crítica
    If RSReferencias.EOF = True Then
        MsgBox "Digite uma referência válida!", vbCritical
        txtMolde.SetFocus
        Exit Sub
    End If
    'Se achar, exibe as cavidades e calcula
    If Not IsNull(RSReferencias!Cavidades) Then txtCavidades = RSReferencias!Cavidades
    Calcula

End Sub

Private Sub txtMPConsumida_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtMPConsumida, KeyAscii)

End Sub

Private Sub LimpaCampos()
Dim X As Control

    For Each X In formProducao
        If TypeOf X Is TextBox Then X.Text = ""
    Next
    txtHora.Text = "__:__"
    cboTurno.ListIndex = -1
    cboMaquina.ListIndex = -1
    cboMP.ListIndex = -1
    cboMPSec.ListIndex = -1
    
End Sub

Private Function CheckIfExist(ByVal Code_MP_Pri As Long, ByVal Code_MP_Sec As Long, ByVal MP_Consumida As Double, ByVal MP_Prim_Cons As Double, ByVal MP_Sec_Cons As Double, ByVal sHora As String, ByVal sReferencia As String) As Boolean
Dim Resposta As Integer
Dim DB_MPPrinCons As String
Dim DB_MPSecCons As String

    'Pode haver mais de um lançamento de Moldes para a mesma máquina, no mesmo turno, na mesma data.
    'Verifica se já foi feito um lancamento para este turno nesta data
    CheckIfExist = False
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT ID, MP_PRIN_CONSUMIDA, MP_SEC_CONSUMIDA FROM TBL_DADOS WHERE data = '" & Format(dData, "yyyy/mm/dd") & "' AND turno = " & Turno & " and maquina = " & cboMaquina.Text & " and referencia = '" & sReferencia & "'", con
    If RS.EOF = False Then
        Resposta = MsgBox("Este lançamento já foi feito! Deseja sobrescrevê-lo?", vbYesNo)
        DB_MPPrinCons = CDbl(RS!MP_PRIN_CONSUMIDA)
        DB_MPSecCons = CDbl(RS!MP_Sec_Consumida)
        If Resposta = vbYes Then
            con.Execute "UPDATE TBL_DADOS SET REFERENCIA = '" & txtMolde & "', TURNO = '" & Turno & "', DATA = '" & Format(dData, "yyyy/mm/dd") & "', CAVIDADES = " & txtCavidades & ", INJECOES = " & txtInjecoes & ", MAQUINA = " & cboMaquina.Text & ", PRODUCAO =  " & 0 & FormatFloatForDB(txtProducao) & ", COD_MP = " & Code_MP_Pri & ", Cod_MP_Sec = " & Code_MP_Sec & ", MP_CONSUMIDA = '" & FormatFloatForDB(MP_Consumida) & "', MP_Prin_Consumida = '" & FormatFloatForDB(MP_Prim_Cons) & "', MP_Sec_Consumida = '" & FormatFloatForDB(MP_Sec_Cons) & "', HORA = '" & sHora & "', CICLOS = " & txtCiclos & ", TEMPO_INJECAO = '" & txtTempo & "', OBS = '" & txtObs & "' WHERE ID= " & RS!ID & ""
            'DEVOLVE AO ESTOQUE A MP PRINCIPAL UTILIZADA ANTERIORMENTE
            con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE + '" & FormatFloatForDB(DB_MPPrinCons) & "') WHERE ID = '" & Code_MP_Pri & "'"
            'DÁ BAIXA NO ESTOQUE DA MP PRINCIPAL UTILIZADA ATUALMENTE
            con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE - '" & FormatFloatForDB(MP_Prim_Cons) & "') WHERE ID = '" & Code_MP_Pri & "'"
            
            'DEVOLVE AO ESTOQUE A MP SECUNDÁRIA UTILIZADA ANTERIORMENTE, E DEPOIS DÁ BAIXA NA ATUAL
            If cboMPSec.Text <> "" Then
                con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE + '" & FormatFloatForDB(DB_MPSecCons) & "') WHERE ID = '" & Code_MP_Sec & "'"
                con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE - '" & FormatFloatForDB(MP_Sec_Cons) & "') WHERE ID = '" & Code_MP_Sec & "'"
            End If
            LimpaCampos
        End If
        CheckIfExist = True
    End If
    RS.Close
        
End Function

Private Sub TotalSemanal()
Dim mData As Date
Dim nData As Date

    'Calcula o Total semanal do produto que estiver selecionado no cboProdutos, pois foi uma opcao do usuario
    mData = txtDataInicial
    nData = txtDataFinal
    
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT CAST(IFNULL(SUM(mp_consumida),0) AS CHAR) AS MPConsumida, CAST(IFNULL(SUM(producao),0) AS CHAR) AS Production FROM TBL_DADOS WHERE referencia = '" & cboMoldes.Text & "' AND data BETWEEN '" & Format(mData, "yyyy/mm/dd") & "' AND '" & Format(nData, "yyyy/mm/dd") & "'", con
    If RS.EOF = False Then
        lblMP = Format(RS!MPConsumida, "#,##0.00") & " Kg"
        lblProducao = Format(RS!Production, "#,##0")
    Else
        MsgBox "Não ouve lançamento durante o período selecionado!", vbExclamation
    End If
    
    RS.Close

End Sub

Private Sub Calcula()
Dim Producao As Double
Dim MP As Double
Dim Cavidades As Double
Dim Injecoes As Double

    txtCavidades = IIf(txtCavidades = "", 0, txtCavidades)
    txtInjecoes = IIf(txtInjecoes = "", 0, txtInjecoes)
    If txtMolde = "" Then Exit Sub
    Cavidades = CDbl(txtCavidades)
    Injecoes = CDbl(txtInjecoes)
    Producao = Cavidades * Injecoes
    If Not IsNull(RSReferencias!Peso) Then MP = CDbl(RSReferencias!Peso) * Injecoes
    txtProducao = Format(Producao, "#,##0")
    txtMPConsumida = Format(MP, "#,##0.000")
    
End Sub

Private Sub Carrega_MP()
    
    'Carrega os dados do combo com as matérias-primas
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT ID, NOME FROM TBL_MP", con
    If RS.EOF = False Then
    'O CBOMPSEC DEVE TER O PRIMEIRO ÍTEM EM BRANCO, POIS O USUÁRIO PRECISA TER A OPCAO DE CORRIGIR
    'A SELECAO CASO TENHA SIDO FEITO ERRADAMENTE
    cboMPSec.AddItem ""
        While Not RS.EOF
            cboMP.AddItem Format(RS!ID, "000") & " - " & RS!nome
            cboMPSec.AddItem Format(RS!ID, "000") & " - " & RS!nome
            RS.MoveNext
        Wend
    End If
    RS.Close

End Sub

Private Sub txtMPPrin_KeyPress(KeyAscii As Integer)

        KeyAscii = TypeCurrency(txtMPPrin, KeyAscii)

End Sub

Private Sub txtMPPrin_LostFocus()
        
    txtMPPrin = Format(txtMPPrin, "#,##0.000")

End Sub

Private Sub txtMPSec_KeyPress(KeyAscii As Integer)

        KeyAscii = TypeCurrency(txtMPSec, KeyAscii)

End Sub

Private Sub txtMPSec_LostFocus()

    txtMPSec = Format(txtMPSec, "#,##0.000")

End Sub


Private Sub txtProducao_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtProducao, KeyAscii)
    
End Sub

Private Sub PreencheMoldes()

    'Carrega o RS de moldes
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT referencia FROM TBL_PRECOS ORDER BY REFERENCIA", con
    While Not RS.EOF
        cboMoldes.AddItem RS!REFERENCIA
        RS.MoveNext
    Wend
    RS.Close
    
End Sub
