VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "REPORTX.OCX"
Begin VB.Form formRelatorio3 
   Caption         =   "Relatório de Retiradas e Devoluções"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   56.356
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   1515
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   1500
         TabIndex        =   0
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
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
         Left            =   300
         TabIndex        =   6
         Top             =   0
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
      Begin ReportX.ReportField ReportField5 
         Height          =   195
         Left            =   8460
         TabIndex        =   9
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   344
         Campo           =   "Total_PG"
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
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1515
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2672
      Tipo            =   2
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   435
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
         Left            =   300
         TabIndex        =   7
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label1 
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
         Left            =   1500
         TabIndex        =   3
         Top             =   1200
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
         TabIndex        =   2
         Top             =   180
         Width           =   9375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Label4"
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
         TabIndex        =   1
         Top             =   720
         Width           =   9375
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   4
      Top             =   2520
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Total 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   1770
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   688
      Tipo            =   7
      Ordem           =   1
      Begin ReportX.ReportField ReportField6 
         Height          =   285
         Left            =   8280
         TabIndex        =   10
         Top             =   60
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Campo           =   "Total_PG"
         Caption         =   ""
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
         Caption         =   "TOTAL:"
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
         Left            =   6960
         TabIndex        =   5
         Top             =   60
         Width           =   1320
      End
   End
End
Attribute VB_Name = "formRelatorio3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variavel local para acumular total
Private pTotal_PG As Currency

Public Sub Config()
    
    Set RS = New ADODB.Recordset
    
    'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS RÁPIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic

    RS.Open "SELECT tbl_descascador.nome, tbl_entrega.cod_desc, ROUND(Sum(tbl_entrega.pagar), 2) AS Total_PG " & _
                         "FROM tbl_descascador LEFT JOIN tbl_entrega ON tbl_descascador.codigo = tbl_entrega.cod_desc " & _
                         "Where data_saida BETWEEN '" & Format(pDataIni, "yyyy/mm/dd") & "' AND '" & Format(pDataFim, "yyyy/mm/dd") & "' And data_dev <> 0 And quitado = 0 " & _
                         "GROUP BY tbl_descascador.nome, tbl_entrega.cod_desc", con
    Label4.Caption = formFechamento.Tag
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relatório de Fechamento"
    Relatorio.Ativar
    RS.Close
    Unload Me
            
End Sub


Private Sub Form_Load()

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de código
    ' do formulário. Isso libera os recursos utilizados
    ' pelo formulário. É uma boa prática no VB, pois o Unload
    ' libera apenas a parte visual do formulário.
    Set formRelatorio3 = Nothing

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
        Case "Total_PG": Valor = pTotal_PG
    End Select
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    ' O ReportField Campo(4) é o Campo que será acumulado
    pTotal_PG = CDbl(pTotal_PG) + CDbl(Relatorio.Recordset("Total_PG"))
    
    ' Poderia ser utilizado o campo diretamente:
    ' pTotalGeral = pTotalGeral + Relatorio.Recordset("ExtendedPrice")
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    ' Reset na variavel que acumula o grupo.
    ' Isso é necessário, porque depois de imprimir para a Tela
    ' a variável deve ser zerada para começar a acumular
    ' corretamente para a saída na impressora.
    pTotal_PG = 0
    
End Sub
