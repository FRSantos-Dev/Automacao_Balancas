VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.2#0"; "ReportX.ocx"
Begin VB.Form formRelaReferencias 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   56.356
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   207.434
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   2115
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   450
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   5160
         TabIndex        =   0
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   344
         Campo           =   "PRECO"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   344
         Campo           =   "referencia"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   195
         Left            =   8340
         TabIndex        =   2
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   344
         Campo           =   "PESO"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9480
         TabIndex        =   9
         Top             =   0
         Width           =   210
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   2115
      Left            =   0
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   3731
      Tipo            =   2
      Begin ReportX.ReportField rpf 
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         Campo           =   "=Página [Pagina]"
         Formato         =   "0"
         Caption         =   ""
         TipoCampo       =   6
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   370
         Campo           =   "=Impresso em [Hoje]"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   600
         X2              =   9960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "LISTAGEM DOS PREÇOS DO DESCASQUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   9375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "DESCASQUE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         TabIndex        =   6
         Top             =   180
         Width           =   9375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PREÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5640
         TabIndex        =   5
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "REFERENCIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   300
         TabIndex        =   4
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "PESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8820
         TabIndex        =   3
         Top             =   1800
         Width           =   420
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   8
      Top             =   2520
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
End
Attribute VB_Name = "formRelaReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Config()

    Set RS = New ADODB.Recordset
    'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS RÁPIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    RS.Open "SELECT * FROM tbl_precos", con
    
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "LISTAGEM DE PREÇOS DO DESCASQUE"
    Relatorio.Ativar
    RS.Close
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de código
    ' do formulário. Isso libera os recursos utilizados
    ' pelo formulário. É uma boa prática no VB, pois o Unload
    ' libera apenas a parte visual do formulário.
    Set formRelaReferencias = Nothing

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    ' Mostra a mensagem de erro para o usuário.
    ' Essa mensagem é opcional, porém é um boa idéia
    ' deixar o usuário saber o que está acontecendo.
    Rpx_MsgErro Numero
    
End Sub


