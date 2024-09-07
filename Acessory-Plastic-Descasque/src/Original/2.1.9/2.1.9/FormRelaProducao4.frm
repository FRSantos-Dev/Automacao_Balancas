VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.2#0"; "ReportX.ocx"
Begin VB.Form FormRelaProducao4 
   Caption         =   "Relatório de Produção Detalhado por Máquinas"
   ClientHeight    =   4695
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.281
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   16.722
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      Top             =   2355
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1799
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   1680
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
         Left            =   6600
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
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   8040
         TabIndex        =   2
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
         Left            =   4860
         TabIndex        =   3
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
      Begin ReportX.ReportField ReportField6 
         Height          =   975
         Left            =   9360
         TabIndex        =   13
         Top             =   0
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1720
         Linhas          =   5
         Campo           =   "OBS"
         Caption         =   ""
         WordWrap        =   -1  'True
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
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   1110
      Left            =   0
      Top             =   3375
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1958
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   6
         Left            =   4995
         TabIndex        =   16
         Top             =   630
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
         X1              =   1440
         X2              =   7140
         Y1              =   600
         Y2              =   600
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
         Left            =   2055
         TabIndex        =   17
         Top             =   630
         Width           =   2925
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1875
      Left            =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   3307
      Tipo            =   2
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs."
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
         Left            =   9540
         TabIndex        =   12
         Top             =   1560
         Width           =   495
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
         TabIndex        =   9
         Top             =   180
         Width           =   11475
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
         Left            =   2700
         TabIndex        =   8
         Top             =   1560
         Width           =   795
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
         Left            =   6900
         TabIndex        =   7
         Top             =   1560
         Width           =   465
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
         TabIndex        =   6
         Top             =   1560
         Width           =   615
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
         Left            =   5205
         TabIndex        =   5
         Top             =   1500
         Width           =   945
      End
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
         TabIndex        =   4
         Top             =   540
         Width           =   11475
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   10
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
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   1875
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   847
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   240
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
            Weight          =   700
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
         Left            =   240
         TabIndex        =   14
         Top             =   -60
         Width           =   885
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Máquinas Detalhadas"
      Height          =   195
      Left            =   5700
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "FormRelaProducao4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variavel local para acumular total
Private pTotalProducao@

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
    RS.Open "SELECT maquina, referencia, data, OBS, SUM(producao) AS iProducao, IF(turno = 1, 'Noite', 'Manhã') AS iTurno FROM TBL_DADOS WHERE data BETWEEN '" & Format(DataIni, "yyyy/mm/dd") & "' AND '" & Format(DataFim, "yyyy/mm/dd") & "' GROUP BY maquina, data, turno, referencia, producao, OBS ORDER BY maquina", con
    
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
    Set FormRelaProducao4 = Nothing

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


Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)

    ' Define a formula para o Grupo 1
    
    ' Se existir apenas um grupo, não é necessário o IF
    ' Aqui vc define qual o campo, ou campos que fazem a quebra do grupo
    ' Para que a quebra seja homogenea, o campo que faz a quebra
    ' deve estar ordenado no SQL, no caso o SQL tem que ter um ORDER BY OrderID
    
    If Ordem = 1 Then
        Valor = CInt(Relatorio.Recordset("maquina"))
    End If
    
    ' Para que vc tenha uma quebra de página após cada grupo,
    ' é só definir a propriedade QuebraDepois = True na seção
    ' do Rodapé de Grupo.

End Sub

Private Sub Relatorio_IniciarGrupo(ByVal Ordem As Byte)
    ' Reset na variavel que acumula o grupo.
    ' Isso é necessário, porque depois de imprimir para a Tela
    ' a variável deve ser zerada para começar a acumular
    ' corretamente para a saída na impressora.
    pTotalProducao = 0
    
End Sub


