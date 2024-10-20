VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form formRelaPagas 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   88.371
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   202.406
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      Top             =   3510
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   1085
      Tipo            =   6
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   6
         Left            =   6840
         TabIndex        =   21
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         Campo           =   "TotalPagas"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Label lblRel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total Pago"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   22
         Top             =   270
         Width           =   2925
      End
      Begin VB.Line Line1 
         X1              =   2760
         X2              =   8460
         Y1              =   240
         Y2              =   240
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   5530
      Tipo            =   2
      Begin ReportX.ReportField rpf 
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   0
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
         TabIndex        =   1
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
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   344
         Campo           =   "nome"
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
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   7200
         TabIndex        =   7
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   344
         Campo           =   "cod_desc"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   195
         Left            =   7200
         TabIndex        =   8
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   344
         Campo           =   "telefone"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   344
         Campo           =   "endereco"
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Referência"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pagas"
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
         TabIndex        =   16
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Preço por Kg"
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
         Left            =   7080
         TabIndex        =   15
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Peso Líquido"
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
         Left            =   5160
         TabIndex        =   14
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Saída"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "End:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   7200
         TabIndex        =   10
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
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
         Left            =   7200
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   480
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
         TabIndex        =   3
         Top             =   180
         Width           =   9375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "RELATÓRIO DAS PAGAS"
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
         TabIndex        =   2
         Top             =   720
         Width           =   9375
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   600
         X2              =   9960
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   3135
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   661
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Campo           =   "data_saida"
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   5280
         TabIndex        =   18
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   344
         Campo           =   "peso_liq"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField6 
         Height          =   195
         Left            =   7200
         TabIndex        =   19
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   344
         Campo           =   "preco"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField7 
         Height          =   195
         Left            =   9120
         TabIndex        =   20
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
         Campo           =   "pagar"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField8 
         Height          =   195
         Left            =   1560
         TabIndex        =   24
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   344
         Campo           =   "referencia"
         Caption         =   ""
         Alignment       =   2
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
   End
End
Attribute VB_Name = "formRelaPagas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variavel local para acumular total
Private pTotalPagas@

Public Sub Config(ByVal Valor As Integer, ByVal DataDe As Date, ByVal DataAte As Date)

    Set RS2 = New ADODB.Recordset
    'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS RÁPIDO
    RS2.CursorLocation = adUseClient
    RS2.CursorType = adOpenStatic
    RS2.Open "SELECT T2.nome, T2.telefone, T2.endereco, T1.pagar, T1.Referencia, T1.peso_liq, T1.preco, T1.cod_desc, T1.data_saida FROM tbl_entrega AS T1 INNER JOIN tbl_descascador AS T2 ON T2.codigo = T1.cod_desc WHERE cod_desc = " & Valor & " and T1.data_dev <> 0 AND T1.quitado = 1 AND data_saida BETWEEN '" & Format(DataDe, "yyyy/mm/dd") & "' AND '" & Format(DataAte, "yyyy/mm/dd") & "' ORDER BY data_dev", con
    Set Relatorio.Recordset = RS2
    Relatorio.Titulo = "Cadastro de Descascadores"
    Relatorio.Ativar
    RS2.Close
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de código
    ' do formulário. Isso libera os recursos utilizados
    ' pelo formulário. É uma boa prática no VB, pois o Unload
    ' libera apenas a parte visual do formulário.
    Set formRelaPagas = Nothing

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    ' Mostra a mensagem de erro para o usuário.
    ' Essa mensagem é opcional, porém é um boa idéia
    ' deixar o usuário saber o que está acontecendo.
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
    
    ' Esse evento é disparado para todos os campos
    ' com Formula = True (para o caso do ReportField)
    ' ou quando um Label tem um Tag começando com '@'
    Select Case Campo
        Case "TotalPagas": Valor = pTotalPagas
    End Select
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    pTotalPagas = pTotalPagas + Relatorio.Recordset!pagar
    
End Sub
