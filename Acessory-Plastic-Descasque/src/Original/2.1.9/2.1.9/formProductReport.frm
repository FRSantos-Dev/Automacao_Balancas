VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form formProductReport 
   Caption         =   "Relat�rio de Produtos"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.255
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   26.882
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1425
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   60
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
         Left            =   3480
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
         Campo           =   "producao"
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
         Left            =   4620
         TabIndex        =   2
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   344
         Campo           =   "MP_consumida"
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
         Left            =   2520
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
      Begin ReportX.ReportField ReportField6 
         Height          =   195
         Left            =   6240
         TabIndex        =   4
         Top             =   0
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   344
         Campo           =   "Nome"
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
         Left            =   1560
         TabIndex        =   18
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   10200
         TabIndex        =   27
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   11280
         TabIndex        =   28
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   12240
         TabIndex        =   29
         Top             =   0
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   344
         Campo           =   "obs"
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
      Begin VB.Line Line21 
         X1              =   12180
         X2              =   12180
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line19 
         X1              =   11280
         X2              =   11280
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line17 
         X1              =   10080
         X2              =   10080
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line15 
         X1              =   6240
         X2              =   6240
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line13 
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line11 
         X1              =   3420
         X2              =   3420
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line8 
         X1              =   2460
         X2              =   2460
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line7 
         X1              =   1500
         X2              =   1500
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
      Height          =   1155
      Left            =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   2037
      Tipo            =   2
      Begin ReportX.ReportField ReportField4 
         Height          =   600
         Left            =   2280
         TabIndex        =   23
         Top             =   480
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   1058
         Campo           =   "referencia"
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
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
         Caption         =   "Per�odo:"
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
         TabIndex        =   15
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
         Caption         =   "Produ��o"
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1620
         Width           =   75
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Relat�rio de Produtos"
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
         TabIndex        =   5
         Top             =   180
         Width           =   2475
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   8
      Top             =   3000
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
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1155
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      Tipo            =   3
      Ordem           =   1
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observa��o"
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
         Left            =   12240
         TabIndex        =   26
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label Label13 
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
         Left            =   11340
         TabIndex        =   25
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line20 
         X1              =   12180
         X2              =   12180
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label Label11 
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
         Left            =   10140
         TabIndex        =   24
         Top             =   0
         Width           =   1080
      End
      Begin VB.Line Line18 
         X1              =   11280
         X2              =   11280
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP Utilizada"
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
         Left            =   6300
         TabIndex        =   22
         Top             =   0
         Width           =   3765
      End
      Begin VB.Line Line16 
         X1              =   10080
         X2              =   10080
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line14 
         X1              =   3420
         X2              =   3420
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP Consumida"
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
         TabIndex        =   21
         Top             =   0
         Width           =   1545
      End
      Begin VB.Line Line12 
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label Label5 
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
         Left            =   2700
         TabIndex        =   20
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produ��o"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   0
         Width           =   930
      End
      Begin VB.Line Line10 
         X1              =   6240
         X2              =   6240
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line9 
         X1              =   2460
         X2              =   2460
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label Label12 
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
         Left            =   1680
         TabIndex        =   17
         Top             =   0
         Width           =   510
      End
      Begin VB.Line Line5 
         X1              =   1500
         X2              =   1500
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   16
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape3 
         Height          =   255
         Left            =   60
         Top             =   0
         Width           =   16755
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      Top             =   1695
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1535
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   6
         Left            =   14595
         TabIndex        =   9
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
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   0
         Left            =   14580
         TabIndex        =   10
         Top             =   540
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         Campo           =   "TotalMP"
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
         Caption         =   "TOTAL DA PRODU��O:"
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
         Left            =   11655
         TabIndex        =   13
         Top             =   210
         Width           =   2925
      End
      Begin VB.Label lblRel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL DE MP CONSUMIDA:"
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
         Index           =   0
         Left            =   11040
         TabIndex        =   12
         Top             =   540
         Width           =   3525
      End
      Begin VB.Label lblRel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   16140
         TabIndex        =   11
         Top             =   540
         Width           =   300
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
      TabIndex        =   14
      Top             =   2700
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "formProductReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variavel local para acumular total
Private pTotalProducao@
Private pTotalMP@

Public Sub Config()
    
    CabGrupo.Tipo = secCabecalhoGrupo
    RodGrupo.Tipo = secRodapeGrupo
    
    ' Coloca QuebraDepois no Rodap� de Grupo Indicando
    ' Que ser� feito um grupo por p�gina, ou n�o, dependedo
    ' do par�metro passado.
    RodGrupo.QuebraDepois = True
    
   'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS R�PIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    
    If CRITERIO = "" Then
        RS.Open "SELECT referencia, data, IF(turno = 1, 'Noite', 'Manh�') AS iTurno, producao, MP_consumida, Nome, DATE_FORMAT(hora, GET_FORMAT(TIME, 'ISO')) AS HORA, tempo_injecao, ciclos, OBS FROM TBL_DADOS LEFT JOIN TBL_MP ON TBL_DADOS.COD_MP=TBL_MP.ID WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' ORDER BY referencia, data, turno", con
    Else
        RS.Open "SELECT referencia, data, IF(turno = 1, 'Noite', 'Manh�') AS iTurno, producao, MP_consumida, Nome, DATE_FORMAT(hora, GET_FORMAT(TIME, 'ISO')) AS HORA, tempo_injecao, ciclos, OBS FROM TBL_DADOS LEFT JOIN TBL_MP ON TBL_DADOS.COD_MP=TBL_MP.ID WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' AND " & CRITERIO & " ORDER BY referencia, data, turno"
    End If
    
    'Label4.Caption = formFechamento.Tag
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relat�rio de Produtos"
    Relatorio.Ativar
    RS.Close
    Unload Me
            
End Sub


Private Sub Form_Load()

    Label10 = "Per�odo:" & DataIni & " A " & DataFim
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de c�digo
    ' do formul�rio. Isso libera os recursos utilizados
    ' pelo formul�rio. � uma boa pr�tica no VB, pois o Unload
    ' libera apenas a parte visual do formul�rio.
    Set formProductReport = Nothing

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
        Case "TotalProducao": Valor = pTotalProducao
        Case "TotalMP": Valor = pTotalMP
    End Select
    
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)

    ' Define a formula para o Grupo 1
    
    ' Se existir apenas um grupo, n�o � necess�rio o IF
    ' Aqui vc define qual o campo, ou campos que fazem a quebra do grupo
    ' Para que a quebra seja homogenea, o campo que faz a quebra
    ' deve estar ordenado no SQL, no caso o SQL tem que ter um ORDER BY OrderID
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("referencia")
    End If
    
    ' Para que vc tenha uma quebra de p�gina ap�s cada grupo,
    ' � s� definir a propriedade QuebraDepois = True na se��o
    ' do Rodap� de Grupo.

End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    ' O ReportField Campo(4) � o Campo que ser� acumulado
    pTotalProducao = pTotalProducao + Relatorio.Recordset!Producao
    pTotalMP = pTotalMP + Relatorio.Recordset!mp_consumida
    
    ' Poderia ser utilizado o campo diretamente:
    ' pTotalGrupo = pTotalGrupo + Relatorio.Recordset("ExtendedPrice")
    
    'Zebrar as linhas
    If ReportSection2.BackColor = vbWhite Then
        ReportSection2.BackColor = &HE0E0E0
    Else
        ReportSection2.BackColor = vbWhite
    End If
    
    
End Sub

Private Sub Relatorio_IniciarGrupo(ByVal Ordem As Byte)
    ' Reset na variavel que acumula o grupo.
    ' Isso � necess�rio, porque depois de imprimir para a Tela
    ' a vari�vel deve ser zerada para come�ar a acumular
    ' corretamente para a sa�da na impressora.
    pTotalProducao = 0
    pTotalMP = 0
End Sub


