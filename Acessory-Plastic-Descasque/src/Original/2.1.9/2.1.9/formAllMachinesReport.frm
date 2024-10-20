VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form formAllMachinesReport 
   Caption         =   "Relatório de Produção por Máquinas"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5.636
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   26.882
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1395
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   1140
         TabIndex        =   0
         Top             =   0
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   344
         Campo           =   "data"
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
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   4560
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
         Campo           =   "iproducao"
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
         Left            =   5700
         TabIndex        =   2
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
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
      Begin ReportX.ReportField ReportField5 
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   344
         Campo           =   "iturno"
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
      Begin ReportX.ReportField ReportField7 
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   344
         Campo           =   "hora"
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
      Begin ReportX.ReportField ReportField8 
         Height          =   195
         Left            =   7440
         TabIndex        =   11
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   344
         Campo           =   "tempo_injecao"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   195
         Left            =   8640
         TabIndex        =   12
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Campo           =   "ciclos"
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
      Begin ReportX.ReportField ReportField10 
         Height          =   195
         Left            =   9480
         TabIndex        =   13
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   344
         Campo           =   "OBS"
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
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Campo           =   "maquina"
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
      Begin VB.Line Line12 
         X1              =   1140
         X2              =   1140
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line9 
         X1              =   9480
         X2              =   9480
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line5 
         X1              =   8580
         X2              =   8580
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line15 
         X1              =   7380
         X2              =   7380
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line13 
         X1              =   5700
         X2              =   5700
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line11 
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line8 
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line7 
         X1              =   2520
         X2              =   2520
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Shape Shape4 
         Height          =   255
         Left            =   60
         Top             =   0
         Width           =   16755
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1395
      Left            =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   2461
      Tipo            =   2
      Begin VB.Line Line10 
         X1              =   1140
         X2              =   1140
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Máquina"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   900
      End
      Begin VB.Shape Shape5 
         Height          =   255
         Left            =   60
         Top             =   1080
         Width           =   16875
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1620
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.Line Line25 
         X1              =   2520
         X2              =   2520
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2820
         TabIndex        =   23
         Top             =   1080
         Width           =   510
      End
      Begin VB.Line Line24 
         X1              =   3600
         X2              =   3600
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Line Line23 
         X1              =   7380
         X2              =   7380
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produção"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4620
         TabIndex        =   22
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3840
         TabIndex        =   21
         Top             =   1080
         Width           =   600
      End
      Begin VB.Line Line22 
         X1              =   5700
         X2              =   5700
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   20
         Top             =   1080
         Width           =   780
      End
      Begin VB.Line Line17 
         X1              =   4560
         X2              =   4560
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Line Line16 
         X1              =   8580
         X2              =   8580
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo Inj."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   19
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Line Line1 
         X1              =   9480
         X2              =   9480
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8640
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9540
         TabIndex        =   17
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Produção por Máquinas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2280
         TabIndex        =   14
         Top             =   480
         Width           =   12915
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   9120
         X2              =   9120
         Y1              =   180
         Y2              =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5220
         TabIndex        =   9
         Top             =   180
         Width           =   3915
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   5100
         X2              =   5100
         Y1              =   180
         Y2              =   480
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   2280
         X2              =   2280
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   2280
         X2              =   16800
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   915
         Left            =   60
         Top             =   180
         Width           =   16755
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Produção"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6105
         TabIndex        =   5
         Top             =   1620
         Width           =   75
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Produtos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         TabIndex        =   4
         Top             =   180
         Width           =   2475
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   7
      Top             =   2640
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   1
      Regua           =   -1  'True
      Escala          =   7
      Orientacao      =   2
      Titulo          =   ""
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      Top             =   1665
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1429
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   6
         Left            =   12975
         TabIndex        =   15
         Top             =   210
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         Campo           =   "TotalProducao"
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
         Caption         =   "TOTAL DA PRODUÇÃO:"
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
         Left            =   10035
         TabIndex        =   16
         Top             =   210
         Width           =   2925
      End
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   6420
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Todos os produtos"
      Height          =   195
      Left            =   17100
      TabIndex        =   8
      Top             =   2700
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "formAllMachinesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variavel local para acumular total
Private pTotalProducao@

Public Sub Config()
    
    Set RS = New ADODB.Recordset
   'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS RÁPIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
        
    RS.Open "SELECT maquina, referencia, data, CAST(ifnull(SUM(producao),0) AS CHAR) AS iProducao, DATE_FORMAT(hora, GET_FORMAT(TIME, 'ISO')) AS HORA, tempo_injecao, ciclos, OBS, IF(turno = 1, 'Noite', 'Manhã') AS iTurno FROM TBL_DADOS WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' GROUP BY maquina, data, turno, referencia, producao", con
        
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relatório de produção"
    Relatorio.Ativar
    RS.Close
    Unload Me
            
End Sub

Private Sub Form_Load()

    Label10 = "Período: " & DataIni & " A " & DataFim
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de código
    ' do formulário. Isso libera os recursos utilizados
    ' pelo formulário. É uma boa prática no VB, pois o Unload
    ' libera apenas a parte visual do formulário.
    Set formAllMachinesReport = Nothing

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
        Case "TotalProducao": Valor = pTotalProducao
    End Select
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    ' O ReportField Campo(4) é o Campo que será acumulado
    pTotalProducao = pTotalProducao + Relatorio.Recordset!iProducao
    
    ' Poderia ser utilizado o campo diretamente:
    ' pTotalGrupo = pTotalGrupo + Relatorio.Recordset("ExtendedPrice")
    
End Sub



