VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form formControle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle do Descasque"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabControle 
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Entrega"
      TabPicture(0)   =   "formControle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSair(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalvar(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPesar(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Devolução"
      TabPicture(1)   =   "formControle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPesar(1)"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "cmdSalvar(1)"
      Tab(1).Control(3)=   "cmdSair(1)"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdPesar 
         Caption         =   "&Peso (F12)"
         Height          =   855
         Index           =   1
         Left            =   -70200
         Picture         =   "formControle.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   915
      End
      Begin VB.CommandButton cmdPesar 
         Caption         =   "&Peso (F12)"
         Height          =   855
         Index           =   0
         Left            =   4800
         Picture         =   "formControle.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   4575
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            Caption         =   "Selecione a Retirada"
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   2775
            Begin VB.TextBox txtPesoBruto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000001&
               ForeColor       =   &H80000005&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   600
               Width           =   1275
            End
            Begin VB.ComboBox cboDataRetirada 
               Appearance      =   0  'Flat
               BackColor       =   &H80000001&
               ForeColor       =   &H80000005&
               Height          =   315
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Peso Bruto"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   1290
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Data da Retirada"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.TextBox txtAPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtGalha 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   9
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtPesoLiq 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   8
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total a pagar"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3120
            TabIndex        =   26
            Top             =   1500
            Width           =   1275
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Galha"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Peso Líquido"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pendências"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   4095
         Begin VB.TextBox txtPreco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   3
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtPeso 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   2
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtReferencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   1
            Top             =   720
            Width           =   1095
         End
         Begin MSMask.MaskEdBox txtDataSaida 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   0
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Preço / Kilo:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   180
            TabIndex        =   20
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data de Saída:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   195
            TabIndex        =   19
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Peso Bruto:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   180
            TabIndex        =   18
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Referência:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "R&eceber"
         Default         =   -1  'True
         Height          =   855
         Index           =   1
         Left            =   -70200
         Picture         =   "formControle.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   420
         Width           =   915
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "&Entregar"
         Height          =   855
         Index           =   0
         Left            =   4800
         Picture         =   "formControle.frx":0F16
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   420
         Width           =   915
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sai&r"
         Height          =   855
         Index           =   1
         Left            =   -70200
         Picture         =   "formControle.frx":17E0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2220
         Width           =   915
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sai&r"
         Height          =   855
         Index           =   0
         Left            =   4800
         Picture         =   "formControle.frx":20AA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2220
         Width           =   915
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   180
      X2              =   5940
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   180
      TabIndex        =   15
      Top             =   60
      Width           =   5790
   End
End
Attribute VB_Name = "formControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variavel da Balanca
Dim hScale As Long
Dim Valor As Currency
Dim btn As Integer
Dim Origemcbo As Boolean 'INFORMA SE O EVENTO VEIO DO COMBO1 PARA EVITAR UMA NOVA CONSULTA NO CBODATARETIRADA

Private Sub cboDataRetirada_Click()
    
    'SE O CHAMADO PARTIU DO COMBO1, A VAR ORIGEMCBO EVITA QUE O CÓDIGO SEJA EXECUTADO
    If cboDataRetirada.ListIndex <> -1 And Origemcbo = False Then
        Set RS2 = con.Execute("SELECT peso_bruto, preco FROM tbl_entrega WHERE cod_desc = " & Caixa & " AND referencia = '" & Combo1.Text & "' AND data_saida = '" & Format(cboDataRetirada.Text, "yyyy/mm/dd") & "' AND data_dev = 0")
        If RS2.EOF = False Then
            txtPesoBruto = RS2!peso_bruto
            Valor = RS2!preco
        End If
        RS2.Close
        txtPesoLiq.SetFocus
    End If
    
End Sub

Private Sub cmdVoltar_Click()

    Unload Me

End Sub

Private Sub cmdPesar_Click(Index As Integer)

    btn = Index
    Abre_Porta
    Fecha_Porta

End Sub

Private Sub cmdSair_Click(Index As Integer)

    Unload Me
    
End Sub

Private Sub cmdSalvar_Click(Index As Integer)

    Select Case Index
    Case 0
        Entregar
    Case 1
        Receber
    End Select

End Sub

Private Sub Combo1_Click()
    
    If Combo1.Text <> "" Then
        'Esta rotina preenche o comboDataRetirada com as datas em que a referencia selecionada
        'no combo1 foram retiradas, assim como os outros campos
        'Nao preciso checar se ele achou algo, pois para ele ser exibido no combo1, ele tem que ter sido achado na base de dados
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT data_saida, peso_bruto, preco FROM tbl_entrega WHERE cod_desc = " & Caixa & " AND referencia = '" & Combo1.Text & "' AND data_dev = 0", con
        cboDataRetirada.Clear
        Valor = RS!preco
        txtPesoBruto = RS!peso_bruto
        While Not RS.EOF
            cboDataRetirada.AddItem Format(RS!data_saida, "dd/mm/yyyy")
            RS.MoveNext
        Wend
        RS.Close
        'QUANDO O CBODATARETIRADA.TEXT FOR CHAMADO, ELE IRÁ EXECUTAR TODO SEU CÓDIGO.
        'ESSA VARIÁVEL EVITA ISSO
        Origemcbo = True
        cboDataRetirada.ListIndex = 0
        Origemcbo = False
        txtPesoLiq.SetFocus
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Rotina para facilitar a digitacao
    If KeyCode = vbKeyF12 Then
        KeyCode = 0
        If tabControle.Tab = 0 Then
            cmdPesar_Click 0
        Else
            cmdPesar_Click 1
        End If
    End If

End Sub

Private Sub Form_Load()

    lblCodigo = lblCodigo & Caixa & " - " & formPrincipal.Tag
    Set RS = con.Execute("SELECT referencia, cod_desc FROM tbl_entrega WHERE cod_desc = " & Caixa & " AND data_dev = 0")
    If RS.EOF = False Then
        While Not RS.EOF
            Combo1.AddItem RS!REFERENCIA
            RS.MoveNext
        Wend
    End If
    RS.Close
    txtDataSaida.Text = Date

End Sub

Private Sub txtDataSaida_LostFocus()

    'Se o usuario nao entrar com uma data, sai da rotina
    If txtDataSaida.Text = "__/__/____" Then Exit Sub
    CheckIfIsDate txtDataSaida
    
End Sub

Private Sub txtGalha_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtPeso, KeyAscii)
    
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    
    KeyAscii = TypeCurrency(txtPeso, KeyAscii)

End Sub

Private Sub txtPesoLiq_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtPesoLiq, KeyAscii)

End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    
    KeyAscii = TypeCurrency(txtPreco, KeyAscii)

End Sub

Private Sub txtPesoLiq_LostFocus()
Dim Galha As Currency, APagar As Currency, CCurPBruto As Currency, CCurPLiq As Currency, CCurValor As Currency
    
    'Efetua os calculos apenas se os pesos bruto e liquido estiverem preenchidos
    If txtPesoBruto <> "" And Not IsNull(txtPesoBruto) And txtPesoLiq <> "" And Not IsNull(txtPesoLiq) Then
        'Atribui o valor às variaveis para calculo
        CCurPBruto = CCur(0 & txtPesoBruto)
        CCurPLiq = CCur(0 & txtPesoLiq)
        'Verifica quanto de galha é retornada
        Galha = CCurPBruto - CCurPLiq 'usar CCur para calculos c\ centavos
        With txtGalha
            .Text = Galha
            .SelStart = 0
            .SelLength = .MaxLength
        End With
        'Valor = Preco da referencia no BD
        APagar = Format(CCurPLiq * CCur(Valor), "#0.000")
        txtAPagar = APagar
    End If
    
End Sub

Private Sub txtReferencia_LostFocus()
    
    If txtReferencia.Text = "" Then Exit Sub
    'VERIFICA SE A TABELA DE REFERENCIAS ESTÁ VAZIA
    If (RSReferencias.EOF And RSReferencias.BOF) Then
        MsgBox "A tabela de referencias está vazia. Por favor, cadastre algo antes de prosseguir!", vbCritical
        Exit Sub
    Else
        RSReferencias.MoveFirst
    End If
        
    RSReferencias.Find "referencia = '" & txtReferencia & "'"
    If RSReferencias.EOF = True Then
        MsgBox "Esta referência não está cadastrada", vbExclamation
        With txtReferencia
            .SetFocus
            .SelStart = .SelLength
        End With
    Else
        txtPreco = RSReferencias!preco
    End If
    
End Sub

Private Sub Receber()
'RECEBE O PRODUTO DAS MÃOS DO DESCASCADOR

    If Combo1.Text = "" Or txtPesoLiq.Text = "" Or txtGalha.Text = "" Or cboDataRetirada.Text = "" Then
        MsgBox "É necessário preencher todos os campos!", vbExclamation
        Exit Sub
    End If
    If CLng(txtPesoLiq) > CLng(txtPesoBruto) Then
        Resposta = MsgBox("O Peso Liquido é maior que o Peso Bruto! Pode ter havido algum erro de digitação. Deseja prosseguir assim mesmo?", vbExclamation + vbYesNo)
        If Resposta = vbNo Then
            txtPesoLiq.SetFocus
            Exit Sub
        End If
    End If
    'LOCALIZA O PRODUTO NO BD, E DÁ ENTRADA DELE
    con.Execute ("UPDATE tbl_entrega SET peso_liq = '" & FormatFloatForDB(txtPesoLiq) & "', data_dev = '" & Format(Date, "yyyy/mm/dd") & "', pagar = '" & FormatFloatForDB(txtAPagar) & "', Galha = '" & FormatFloatForDB(txtGalha) & "' " & _
                 "WHERE data_saida = '" & Format(cboDataRetirada.Text, "yyyy/mm/dd") & "' AND cod_desc = " & Caixa & " AND referencia = '" & Combo1.Text & "' AND data_dev = 0")
    Combo1.RemoveItem Combo1.ListIndex
    If cboDataRetirada.ListIndex <> -1 Then
        cboDataRetirada.RemoveItem cboDataRetirada.ListIndex
    Else
        cboDataRetirada.Text = ""
    End If
    LimpaCampos

End Sub

Private Sub Entregar()

    If txtDataSaida.Text <> "__/__/____" And txtPreco <> "" And txtPeso <> "" And txtReferencia <> "" Then
        
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        'XXXX VERIFICA SE O PRODUTO FOI ENTREGUE PARA O DESCASCADOR NESTE MESMO DIA XXXXXX
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        Set RS = con.Execute("SELECT * FROM tbl_entrega WHERE referencia = '" & txtReferencia.Text & "' AND data_saida = '" & Format(txtDataSaida.Text, "yyyy/mm/dd") & "' AND cod_desc = " & Caixa & " AND data_dev =  0")
        If RS.EOF = False Then
            Resposta = MsgBox("Este produto já se encontra com o descascador. Deseja sobrescrevê-lo?", vbYesNo + vbQuestion)
            Select Case Resposta
                Case vbYes
                    'Atualiza os campos
                    con.Execute "UPDATE tbl_entrega SET cod_desc = " & Caixa & ", preco = '" & FormatFloatForDB(txtPreco) & "', peso_bruto = '" & FormatFloatForDB(txtPeso) & "', data_saida = '" & Format(txtDataSaida.Text, "yyyy/mm/dd") & "' WHERE referencia = '" & txtReferencia.Text & "' AND data_saida = '" & Format(txtDataSaida.Text, "yyyy/mm/dd") & "' AND cod_desc = " & Caixa & " AND data_dev =  0"
                    LimpaCampos
                    txtDataSaida.SetFocus
                    RS.Close
                    Exit Sub
                Case vbNo
                    MsgBox "Não é possível lançar 2 produtos iguais para o mesmo cliente no mesmo dia. Primeiro faça a devolução do material que estiver pendente!"
                    RS.Close
                    Exit Sub
            End Select
        End If
        RS.Close
        con.Execute "INSERT INTO tbl_entrega (cod_desc, preco, peso_bruto, referencia, data_saida) Values(" & Caixa & ", '" & FormatFloatForDB(txtPreco) & "', '" & FormatFloatForDB(txtPeso) & "', '" & txtReferencia & "', '" & Format(txtDataSaida.Text, "yyyy/mm/dd") & "')"
        Combo1.AddItem txtReferencia.Text
        LimpaCampos
        txtDataSaida.SetFocus
    Else
        MsgBox "É necessário preencher todos os campos!", vbExclamation
    End If

End Sub

Private Sub LimpaCampos()
Dim X As Control

    'limpa todos os textbox
    For Each X In formControle
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
End Sub

Private Sub Abre_Porta()
Dim COM, BaudRate, ByteSize, StopBits, Parity, Status As Long

    COM = 1         '  // porta de comunicacao
    BaudRate = 9600 '  // taxa de transferencia - velocidade
    ByteSize = 8    '  // numero de bits/byte, 4-8
    StopBits = 0    '  // 0,1,2 = 1, 1.5, 2 -- ATENCAO
    Parity = 0      '  // 0-4=no,odd,even,mark,space -- ATENCAO
    
    'Se for gerado um erro, ele nao abre a porta
    On Error GoTo TrataErro
    
    Status = AttachScale(hScale, COM, BaudRate, ByteSize, StopBits, Parity)
    If (Status < 0) Then
      Select Case Status
        Case -2
          MsgBox "Erro abrindo COM"
        Case -3
          MsgBox "COM inválida"
        Case -4
          MsgBox "Muitos arquivos abertos"
        Case -5
          MsgBox "Acesso à COM não permitido"
        Case -6
          MsgBox "Erro setando parâmetros de comunicação"
        Case -7
          MsgBox "Erro setando timeout de comunicação"
        Case -100
          MsgBox "Arquivo vazio"
        Case -101
          MsgBox "Disco cheio"
        Case -106
          MsgBox "Input inválido"
        Case Else
          MsgBox "Erro Genérico"
      End Select
    Else
      LeBalanca
    End If
    Exit Sub
TrataErro:
    Select Case Err.Number
    Case 48
        MsgBox "Está faltando o arquivo para comunicação com a balança. Por favor, entre em ctt com o Suporte Técnico.", vbCritical
    Case Else
        MsgBox "Erro: " & Err.Number & vbCrLf & "Descrição: " & Err.Description
        End
    End Select

End Sub

Private Sub LeBalanca()
Dim Erro As String

'contantes de erro
Const ScaleError As Long = 0
Const ScaleInMotion As Long = 1
Const ScaleStable As Long = 2
Const ScaleOutOfRange As Long = 3
Const ProteqNotFound As Long = 4

'---------------------------------------
'Modelo do indicador ou balança conectado ao canal de leitura
Const BIDS As Long = 6
'0 = Indicador IDS Filizola
'1 = Indicador IQ Plus 810 Filizola
'2 = Indicador 9091 Toledo
'3 = Indicador 8132 Toledo
'4 = Indicador ID10000 ou IDS com protocolo demanda da Filizola
'5 = Balanças da linha BP da Filizola
'6 = Balança MF Filizola
'7 = Indicador IDC Filizola
'8 = Balanças da linha E
'---------------------------------------

Dim Status As Long
Dim grossbuf As String, tarebuf As String, netbuf As String, unitbuf As String

    'Esses buffers DEVEM ser limpos antes da chamada da função ReadScale.
    'char *gross - Ponteiro para buffer onde será retornado o valor de peso bruto.
    grossbuf = String(9, Chr(0))
    'char *tare - Ponteiro para buffer onde será retornado o valor de tara.
    tarebuf = String(9, Chr(0))
    'char *net - Ponteiro para buffer onde será retornado o valor de peso líquido.
    netbuf = String(9, Chr(0))
    'char *unit - Ponteiro para buffer onde será retornado a unidade programada na balança.
    unitbuf = String(3, Chr(0))
     
    'Le a balanca e atribui o valor do status
    Status = ReadScale(hScale, BIDS, grossbuf, tarebuf, netbuf, unitbuf)
     
    Select Case Status
       Case ScaleError
         MsgBox "Erro de leitura da balança"
       Case ScaleOutOfRange
         MsgBox "Sobrecarga"
       Case ProteqNotFound
         MsgBox "Chave de licença não encontrada"
       Case ScaleInMotion
         MsgBox "Peso oscilando"
       Case ScaleStable
         'Peso estável"
On Error GoTo TrataErro
         If btn = 0 Then
            If txtPeso <> "" Then
                Erro = "Erro 1: " & grossbuf
                txtPeso = Format(CSng(0 & txtPeso) + CSng(grossbuf), "#0.000")
            Else
                Erro = "Erro 2: " & grossbuf
                txtPeso = Format(grossbuf, "#0.000")
            End If
         Else
            If txtPesoLiq <> "" Then
                Erro = "Erro 3: " & grossbuf
                txtPesoLiq = Format(CSng(0 & txtPesoLiq) + CSng(grossbuf), "#0.000")
                
            Else
                Erro = "Erro 4: " & grossbuf
                txtPesoLiq = Format(grossbuf, "#0.000")
            End If
            'Tira o foco para que o calculo seja feito
            txtPesoLiq.SetFocus
            txtGalha.SetFocus
        End If
    End Select
    Exit Sub
    
TrataErro:
    MsgBox Erro & vbCrLf & Err.Number & ": " & Err.Description
    End

End Sub

Private Sub Fecha_Porta()

    'Fecha a porta
    On Error GoTo TrataErro
    DettachScale (hScale)
    Exit Sub
    
TrataErro:
    Select Case Err.Number
    Case 48 'Arquivo de Comunicação com a balança não encontrado. Usuário já foi informado por outra crítica
        Exit Sub
    End Select

End Sub

