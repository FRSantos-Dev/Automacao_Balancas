VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form formRelaProducao3 
   Caption         =   "Relatório de Produção por Máquinas"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6.429
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   25.718
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      Top             =   1875
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   529
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   1320
         TabIndex        =   0
         Top             =   0
         Width           =   2835
         _ExtentX        =   5001
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
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   6240
         TabIndex        =   1
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
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
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Campo           =   "maquina"
         Caption         =   "maq"
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
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   8040
         TabIndex        =   11
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
         Campo           =   "iTurno"
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
         Left            =   4380
         TabIndex        =   13
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   344
         Campo           =   "iproducao"
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
      Height          =   1875
      Left            =   0
      Top             =   0
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   3307
      Tipo            =   2
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   540
         Width           =   11475
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produção"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4605
         TabIndex        =   12
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8265
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6600
         TabIndex        =   6
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Produção"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   11475
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Máquina"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   315
         TabIndex        =   3
         Top             =   1560
         Width           =   885
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   7
      Top             =   3120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   1
      Regua           =   -1  'True
      Escala          =   7
      Titulo          =   ""
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   1170
      Left            =   0
      Top             =   2175
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   2064
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   6
         Left            =   6495
         TabIndex        =   8
         Top             =   690
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
      Begin VB.Line Line1 
         X1              =   2940
         X2              =   8640
         Y1              =   660
         Y2              =   660
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
         Left            =   3555
         TabIndex        =   9
         Top             =   690
         Width           =   2925
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Máquinas"
      Height          =   195
      Left            =   5700
      TabIndex        =   14
      Top             =   3300
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "formRelaProducao3"
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
        
    RS.Open "SELECT maquina, referencia, data, SUM(producao) AS iProducao, IF(turno = 1, 'Noite', 'Manhã') AS iTurno FROM TBL_DADOS WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' GROUP BY maquina, data, turno, referencia, producao", con
        
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relatório de produção"
    Relatorio.Ativar
    RS.Close
    Unload Me
            
End Sub

Private Sub Form_Load()

    Label8 = "Período:" & DataIni & " A " & DataFim
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de código
    ' do formulário. Isso libera os recursos utilizados
    ' pelo formulário. É uma boa prática no VB, pois o Unload
    ' libera apenas a parte visual do formulário.
    Set formRelaProducao3 = Nothing

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


