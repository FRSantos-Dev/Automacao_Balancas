VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "REPORTX.OCX"
Begin VB.Form formRelaProducao5 
   Caption         =   "Relatório de Produção por Matérias-primas"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.176
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   23.178
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      Top             =   2535
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   529
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   5280
         TabIndex        =   0
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
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
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   9180
         TabIndex        =   1
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
         Left            =   6930
         TabIndex        =   2
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
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
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   3413
      Tipo            =   2
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
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   11475
      End
      Begin VB.Label Label2 
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
         Left            =   5610
         TabIndex        =   6
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP Consumida"
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
         Left            =   9180
         TabIndex        =   5
         Top             =   1620
         Width           =   1545
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
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   11475
      End
      Begin VB.Label Label8 
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
         Left            =   7095
         TabIndex        =   3
         Top             =   1620
         Width           =   615
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
      Divisao         =   1
      Regua           =   -1  'True
      Escala          =   7
      Titulo          =   ""
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Top             =   1935
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   1058
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField6 
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   344
         Campo           =   "Nome"
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Matéria-prima"
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
         Left            =   240
         TabIndex        =   14
         Top             =   0
         Width           =   1485
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   1110
      Left            =   0
      Top             =   2835
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   1958
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   0
         Left            =   9360
         TabIndex        =   9
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
         Left            =   5820
         TabIndex        =   11
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
         Left            =   10920
         TabIndex        =   10
         Top             =   780
         Width           =   300
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Todos os produtos"
      Height          =   195
      Left            =   5040
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "formRelaProducao5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variavel local para acumular total
Private pTotalMP@

Public Sub Config()
    
    CabGrupo.Tipo = secCabecalhoGrupo
    RodGrupo.Tipo = secRodapeGrupo
    
    ' Coloca QuebraDepois no Rodapé de Grupo Indicando
    ' Que será feito um grupo por página, ou não, dependedo
    ' do parâmetro passado.
    RodGrupo.QuebraDepois = True
    
    Set RS = New ADODB.Recordset
   'CURSOR NO CLIENTE FAZ O RELATORIO SER EXIBIDO MAIS RÁPIDO
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    
    RS.Open "SELECT data, IF(turno = 1, 'Noite', 'Manhã') AS iTurno, MP_consumida, Nome FROM TBL_DADOS LEFT JOIN TBL_MP ON TBL_DADOS.COD_MP=TBL_MP.ID WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' ORDER BY nome, data, turno", con
    'Label4.Caption = formFechamento.Tag
    Set Relatorio.Recordset = RS
    Relatorio.Titulo = "Relatório de produção"
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
    Set formRelaProducao5 = Nothing

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

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)

    ' Define a formula para o Grupo 1
    
    ' Se existir apenas um grupo, não é necessário o IF
    ' Aqui vc define qual o campo, ou campos que fazem a quebra do grupo
    ' Para que a quebra seja homogenea, o campo que faz a quebra
    ' deve estar ordenado no SQL, no caso o SQL tem que ter um ORDER BY OrderID
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("nome")
    End If
    
    ' Para que vc tenha uma quebra de página após cada grupo,
    ' é só definir a propriedade QuebraDepois = True na seção
    ' do Rodapé de Grupo.

End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' Acumula o total para o grupo
    ' O ReportField Campo(4) é o Campo que será acumulado
    pTotalMP = pTotalMP + Relatorio.Recordset!MP_consumida
    
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


