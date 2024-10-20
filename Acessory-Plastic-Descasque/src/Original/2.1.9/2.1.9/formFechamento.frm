VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form formFechamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "formFechamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Relatórios"
      Height          =   1185
      Left            =   180
      TabIndex        =   6
      Top             =   1140
      Width           =   4215
      Begin VB.OptionButton Option4 
         Caption         =   "Pagas"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   750
         Width           =   870
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Fechamento"
         Height          =   255
         Left            =   2625
         TabIndex        =   8
         Top             =   360
         Width           =   1395
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pendências"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Totais"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Emitir"
      Height          =   915
      Left            =   3480
      Picture         =   "formFechamento.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox txtDIni 
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDFim 
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
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "formFechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para manter a janela modal sem ser modal
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Sub Command1_Click()
    
    If Not IsDate(txtDIni) Or Not IsDate(txtDFim) Then Exit Sub

    pDataIni = txtDIni
    pDataFim = txtDFim
    formFechamento.Tag = "DESCASQUE DO PERÍODO DE " & pDataIni & " A " & pDataFim
    
    'Uso o hide e um doevents porque a janela estava demorando para sair
    Me.Hide
    DoEvents
    Screen.MousePointer = vbHourglass
    
    'TOTAIS
    If Option1.Value Then
        'TOTAIS
        formRelatorio2.Config
    Else
        If Option2.Value Then
            'PENDENCIAS
            formPendencias.Config
        Else
            If Option3.Value Then
                'FECHAMENTO
                MsgBox "Este Relatório exibe todos os produtos que ainda não foram pagos. imprima-o antes de fazer o fechamento.", vbInformation
                formRelatorio3.Config
            Else
                'PAGAS
                formPagas.Config
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Unload Me
    
End Sub

Private Sub Form_Activate()

    txtDIni.SetFocus

End Sub

Private Sub Form_Load()

'Para manter a janela modal sem ser modal
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
MDIForm1.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    MDIForm1.Enabled = True
    MDIForm1.SetFocus
    FormDescasque.Visible = True

End Sub

Private Sub Option6_Click()

End Sub
