VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form formRelatorio2 
   Caption         =   "Relat�rio de Retiradas e Devolu��es"
   ClientHeight    =   3390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   59.796
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   198.438
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   1515
      Width           =   11250
      _ExtentX        =   19844
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
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   6840
         TabIndex        =   1
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   344
         Campo           =   "p_bruto"
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
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   300
         TabIndex        =   9
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
         TabIndex        =   12
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   344
         Campo           =   "p_liq"
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
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   2672
      Tipo            =   2
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Peso L�q."
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
         TabIndex        =   11
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
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
         TabIndex        =   10
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
         TabIndex        =   5
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto"
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
         Left            =   7140
         TabIndex        =   4
         Top             =   1200
         Width           =   915
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
         TabIndex        =   2
         Top             =   720
         Width           =   9375
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   6
      Top             =   2520
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Total 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   1770
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   688
      Tipo            =   7
      Ordem           =   1
      Begin ReportX.ReportField ReportField3 
         Height          =   285
         Left            =   6660
         TabIndex        =   7
         Top             =   60
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Campo           =   "TotalBruto"
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
      Begin ReportX.ReportField ReportField6 
         Height          =   285
         Left            =   8280
         TabIndex        =   13
         Top             =   60
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Campo           =   "TotalLiq"
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
         Left            =   5340
         TabIndex        =   8
         Top             =   60
         Width           =   1320
      End
   End
End
Attribute VB_Name = "formRelatorio2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variavel local para acumular total
Private pTotalBruto As Currency, pTotalLiq As Currency


Public Sub Config()
    
    Set RS = New ADODB.Recordset
    
    'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS R�PIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic

    RS.Open "SELECT tbl_descascador.nome, tbl_entrega.cod_desc, ROUND(Sum(tbl_entrega.peso_bruto), 2) AS P_Bruto, ROUND(Sum(tbl_entrega.peso_liq), 2) AS P_Liq " & _
                         "FROM tbl_descascador LEFT JOIN tbl_entrega ON tbl_descascador.codigo = tbl_entrega.cod_desc " & _
                         "Where data_saida BETWEEN '" & Format(pDataIni, "yyyy/mm/dd") & "' AND '" & Format(pDataFim, "yyyy/mm/dd") & "' " & _
                         "GROUP BY tbl_descascador.nome, tbl_entrega.cod_desc", con
    Label4.Caption = formFechamento.Tag
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relat�rio de Retirada e Devolu��o"
    Relatorio.Ativar
    RS.Close
    Unload Me
            
End Sub


Private Sub Form_Load()

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de c�digo
    ' do formul�rio. Isso libera os recursos utilizados
    ' pelo formul�rio. � uma boa pr�tica no VB, pois o Unload
    ' libera apenas a parte visual do formul�rio.
    Set formRelatorio2 = Nothing

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    ' Mostra a mensagem de erro para o usu�rio.
    ' Essa mensagem � opcional, por�m � um boa id�ia
    ' deixar o usu�rio saber o que est� acontecendo.
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
    
    ' Esse evento � disparado para todos os campos
    ' com Formula = True (para o caso do ReportField)
    ' ou quando um Label tem um Tag come�ando com '@'
    Select Case Campo
        Case "TotalBruto": Valor = pTotalBruto
        Case "TotalLiq": Valor = pTotalLiq
    End Select
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    ' O ReportField Campo(4) � o Campo que ser� acumulado
    pTotalBruto = CDbl(pTotalBruto) + CDbl(Relatorio.Recordset("P_Bruto"))
    pTotalLiq = CDbl(pTotalLiq) + CDbl(Relatorio.Recordset("P_Liq"))
    
    ' Poderia ser utilizado o campo diretamente:
    ' pTotalGeral = pTotalGeral + Relatorio.Recordset("ExtendedPrice")
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    ' Reset na variavel que acumula o grupo.
    ' Isso � necess�rio, porque depois de imprimir para a Tela
    ' a vari�vel deve ser zerada para come�ar a acumular
    ' corretamente para a sa�da na impressora.
    pTotalBruto = 0
    pTotalLiq = 0
    
End Sub
