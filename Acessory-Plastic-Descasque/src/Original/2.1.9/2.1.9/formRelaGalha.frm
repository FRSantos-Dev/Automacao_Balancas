VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "REPORTX.OCX"
Begin VB.Form formRelaGalha 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7.752
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   21.034
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      Top             =   1935
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   529
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   9180
         TabIndex        =   0
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   344
         Campo           =   "MP_TOTAL"
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
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5145
         _ExtentX        =   9075
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
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   3413
      Tipo            =   2
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Galha"
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
         Left            =   9720
         TabIndex        =   9
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devolução"
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
         Left            =   7080
         TabIndex        =   8
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Diferença Devolução X Galha"
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
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   11475
      End
      Begin VB.Label Label7 
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
         Left            =   90
         TabIndex        =   2
         Top             =   600
         Width           =   11475
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   4
      Top             =   3000
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Divisao         =   1
      Regua           =   -1  'True
      Escala          =   7
      Titulo          =   ""
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   1110
      Left            =   0
      Top             =   2235
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1958
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   0
         Left            =   9120
         TabIndex        =   5
         Top             =   780
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
         Left            =   5580
         TabIndex        =   7
         Top             =   780
         Width           =   3525
      End
      Begin VB.Line Line1 
         X1              =   5820
         X2              =   11520
         Y1              =   420
         Y2              =   420
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
         Left            =   10680
         TabIndex        =   6
         Top             =   780
         Width           =   300
      End
   End
End
Attribute VB_Name = "formRelaGalha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variavel local para acumular total
Private pTotalMP@

Public Sub Config()
    
    Set RS = New ADODB.Recordset
   'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS RÁPIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    
    RS.Open "SELECT SUM(MP_consumida) AS MP_TOTAL, Nome FROM TBL_DADOS LEFT JOIN TBL_MP ON TBL_DADOS.COD_MP=TBL_MP.ID WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' GROUP BY NOME, MP_CONSUMIDA ORDER BY nome", con
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relatório de Diferença Devolução X Galha"
    Relatorio.Ativar
    RS.Close
    Unload Me
            
End Sub


Private Sub Form_Load()

    Label7 = "Período:" & DataIni & " A " & DataFim
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Essa linha retira a parte de código
    ' do formulário. Isso libera os recursos utilizados
    ' pelo formulário. É uma boa prática no VB, pois o Unload
    ' libera apenas a parte visual do formulário.
    Set formRelaGalha = Nothing

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
        Case "TotalMP": Valor = pTotalMP
    End Select
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    ' O ReportField Campo(4) é o Campo que será acumulado
    pTotalMP = pTotalMP + Relatorio.Recordset!MP_TOTAL
    
    ' Poderia ser utilizado o campo diretamente:
    ' pTotalGrupo = pTotalGrupo + Relatorio.Recordset("ExtendedPrice")
    
End Sub

Private Sub Relatorio_IniciarGrupo(ByVal Ordem As Byte)
    ' Reset na variavel que acumula o grupo.
    ' Isso é necessário, porque depois de imprimir para a Tela
    ' a variável deve ser zerada para começar a acumular
    ' corretamente para a saída na impressora.
    pTotalMP = 0
End Sub




