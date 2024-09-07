VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form formRelatorios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "formRelatorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabRelatorios 
      Height          =   4335
      Left            =   150
      TabIndex        =   3
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "A Pagar"
      TabPicture(0)   =   "formRelatorios.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GridAPagar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pagas"
      TabPicture(1)   =   "formRelatorios.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridPagas"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   6495
         Begin VB.CommandButton cmdApagar 
            Caption         =   "&Apagar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4920
            TabIndex        =   16
            Top             =   1140
            Width           =   1455
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000001&
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   525
            Width           =   1125
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4920
            TabIndex        =   11
            Top             =   675
            Width           =   1455
         End
         Begin VB.CommandButton cmdProcurar 
            Caption         =   "&Procurar"
            Height          =   375
            Left            =   4920
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtDataDe 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   525
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDataAte 
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Top             =   525
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "A Pagar"
            Height          =   195
            Left            =   3120
            TabIndex        =   15
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "De:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Até:"
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            Top             =   285
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   6495
         Begin VB.CommandButton cmdImprimirPagas 
            Caption         =   "&Imprimir"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4920
            TabIndex        =   22
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdProcurarPagas 
            Caption         =   "&Procurar"
            Height          =   375
            Left            =   4920
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtTotalPagas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000001&
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Top             =   600
            Width           =   1125
         End
         Begin MSMask.MaskEdBox txtDePagas 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAtePagas 
            Height          =   285
            Left            =   1560
            TabIndex        =   18
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            Caption         =   "Até:"
            Height          =   255
            Left            =   1560
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "De:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblPagas 
            AutoSize        =   -1  'True
            Caption         =   "Pagas"
            Height          =   195
            Left            =   3120
            TabIndex        =   9
            Top             =   360
            Width           =   450
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridAPagar 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid GridPagas 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   5
         Top             =   2160
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
      End
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo:"
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
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   6780
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "formRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApagar_Click()
Dim Dia As Date
Dim Ref As String
Dim Resposta As Integer

    If GridAPagar.Col = 0 Then
        Ref = GridAPagar.Text
        GridAPagar.Col = 5
        Dia = GridAPagar.Text
        Resposta = MsgBox("Você tem certeza de que deseja excluir esta referência deste descascador?", vbYesNo + vbQuestion)
        If Resposta = vbYes Then
            con.Execute "DELETE FROM TBL_ENTREGA WHERE cod_desc = " & Caixa & " AND referencia = '" & Ref & "' AND data_dev <> 0 AND quitado = 0 AND data_saida = '" & Format(Dia, "yyyy/mm/dd") & "'"
            cmdApagar.Enabled = False
            cmdProcurar_Click
        End If
    End If
    
End Sub

Private Sub cmdImprimir_Click()
Dim DataDe As Date
Dim DataAte As Date
    
If txtDataDe.Text <> "__/__/____" And txtDataAte.Text <> "__/__/____" Then
    DataDe = txtDataDe.Text
    DataAte = txtDataAte.Text
    
    'carrega dados para exibir relatorio
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXX Seleciona nome, pagar(quanto sera pago),      XXXXXXX
    'XXXXXXXX peso liq, preco do prod, e cod_desc, Onde     XXXXXXX
    'XXXXXXXX material = devolvido e nao tenha sido quitado XXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT tbl_descascador.nome FROM tbl_entrega INNER JOIN tbl_descascador ON tbl_descascador.codigo = tbl_entrega.cod_desc WHERE tbl_entrega.data_dev <> 0 AND tbl_entrega.quitado = 0", con
    
    'se selecionado "YES", preenche o campo "quitado" da tbl_entrega, confirmando o pagamento do descasque ao cliente
    If MsgBox("Deseja quitar os débitos referentes a este intervalo de datas?", vbYesNo) = vbYes Then
        'exibe relatorio dos artigos nao quitados
        If RS.EOF = False Then
            'USO ME.HIDE PARA PODER EXIBIR O RELATORIO
            'ELE NÃO É EXIBIDO SE UMA JANELA MODAL ESTIVER VISÍVEL
            Me.Hide
            DoEvents
            'FECHO O RS POIS ELE É USADO NO PROXIMO FORM
            RS.Close
            formRelaAPagar.Config Caixa, DataDe, DataAte
            Me.Show 1
        End If
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        'XXXXXXXX Todos os artigos do descascador atual que  XXXXXXXXXX
        'XXXXXXXX tenham sido devolvidos e que a data de     XXXXXXXXXX
        'XXXXXXXX saida esteja entre determinada data.       XXXXXXXXXX
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        con.Execute "UPDATE tbl_entrega SET quitado = 1 WHERE cod_desc = " & Caixa & " AND data_dev <> 0 AND data_saida BETWEEN '" & Format(DataDe, "yyyy/mm/dd") & "' AND '" & Format(DataAte, "yyyy/mm/dd") & "'"
        
        GridAPagar.Rows = 1
        txtDataDe.Text = "__/__/____"
        txtDataAte.Text = "__/__/____"
        txtTotal.Text = ""
    Else
        'se selecionado "NO", nao altera o campo quitado, apenas exibe os artigos nao pagos
        'exibe relatorio dos artigos nao quitados
        If RS.EOF = False Then
            Me.Hide
            'USO ME.HIDE PARA PODER EXIBIR O RELATORIO
            'ELE NÃO É EXIBIDO SE UMA JANELA MODAL ESTIVER VISÍVEL
            DoEvents
            'FECHO O RS POIS ELE É USADO NO PROXIMO FORM
            RS.Close
            formRelaAPagar.Config Caixa, DataDe, DataAte
            Me.Show 1
        End If
        txtDataDe.Text = "__/__/____"
        txtDataAte.Text = "__/__/____"
        txtTotal.Text = ""
        GridAPagar.Rows = 1
    End If
End If
    
End Sub
Private Sub cmdImprimirPagas_Click()
Dim vID As Integer
Dim DataDe As Date
Dim DataAte As Date

    If txtDePagas.Text <> "__/__/____" And txtAtePagas.Text <> "__/__/____" Then
        DataDe = txtDePagas.Text
        DataAte = txtAtePagas.Text
    
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT tbl_descascador.nome, tbl_entrega.pagar, tbl_entrega.peso_liq, tbl_entrega.preco, tbl_entrega.cod_desc FROM tbl_entrega INNER JOIN tbl_descascador ON tbl_descascador.codigo = tbl_entrega.cod_desc WHERE tbl_entrega.data_dev <> 0 AND tbl_entrega.quitado = 1", con
        If RS.EOF = False Then
            'USO ME.HIDE PARA PODER EXIBIR O RELATORIO
            'ELE NÃO É EXIBIDO SE UMA JANELA MODAL ESTIVER VISÍVEL
            Me.Hide
            DoEvents
            'FECHO O RS POIS ELE É USADO NO PROXIMO FORM
            RS.Close
            formRelaPagas.Config Caixa, DataDe, DataAte
            Me.Show 1
            'SAIO DA ROTINA PARA EVITAR ERRO NO FECHAMENTO DO RS
            Exit Sub
        End If
        RS.Close
    End If
    
End Sub

Private Sub cmdProcurar_Click()
    Dim DataDe As Date
    Dim DataAte As Date
    
    If txtDataDe.Text <> "__/__/____" And txtDataAte.Text <> "__/__/____" Then
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT referencia, data_dev, data_saida, peso_bruto, preco, pagar FROM tbl_entrega WHERE cod_desc = " & Caixa & " AND data_dev <> 0 and quitado = 0 AND data_saida BETWEEN '" & Format(txtDataDe, "yyyy/mm/dd") & "' AND '" & Format(txtDataAte, "yyyy/mm/dd") & "' ORDER BY data_dev", con
        If RS.EOF = False Then
            GridAPagar.Rows = 1
            cmdImprimir.Enabled = True
            While Not RS.EOF
                GridAPagar.AddItem RS!REFERENCIA & vbTab & Format(RS!data_dev, "dd/mm/yyyy") & vbTab & RS!peso_bruto & vbTab & RS!preco & vbTab & RS!pagar & vbTab & Format(RS!data_saida, "dd/mm/yyyy")
                Valor = CSng(RS!pagar) + CSng(Valor)
                RS.MoveNext
            Wend
            txtTotal = Valor
            Valor = 0
            cmdApagar.Enabled = True
        End If
        RS.Close
    Else
        cmdApagar.Enabled = False
        MsgBox "Preencha os dois intervalos de datas!", vbExclamation
    End If

    
End Sub

Private Sub cmdProcurarPagas_Click()
    
    If txtDePagas.Text <> "__/__/____" And txtAtePagas.Text <> "__/__/____" Then
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT referencia, data_dev, data_saida, peso_bruto, peso_liq, preco, pagar FROM tbl_entrega WHERE cod_desc = " & Caixa & " AND data_dev <> 0 and quitado = 1 AND data_saida BETWEEN '" & Format(txtDePagas, "yyyy/mm/dd") & "' AND '" & Format(txtAtePagas, "yyyy/mm/dd") & "' ORDER BY data_dev", con
        If RS.EOF = False Then
            GridPagas.Rows = 1
            cmdImprimirPagas.Enabled = True
            While Not RS.EOF
                GridPagas.AddItem RS!REFERENCIA & vbTab & Format(RS!data_dev, "dd/mm/yyyy") & vbTab & RS!peso_liq & vbTab & RS!peso_bruto & vbTab & RS!preco & vbTab & RS!pagar & vbTab & Format(RS!data_saida, "dd/mm/yyyy")
                Valor = CSng(RS!pagar) + CSng(Valor)
                RS.MoveNext
            Wend
            txtTotalPagas = Valor
            Valor = 0
        Else
            GridPagas.Rows = 1
            txtTotalPagas.Text = ""
        End If
        RS.Close
    Else
        MsgBox "Preencha os dois intervalos de datas!", vbExclamation
    End If

End Sub

Private Sub Form_Load()
Dim Limite As String
Dim Valor As Currency
Dim i As Long

    'preenche GridAPagar
    With GridAPagar
        .ColWidth(0) = 1200
        .ColWidth(1) = 1300
        .TextMatrix(0, 0) = "Referência"
        .TextMatrix(0, 1) = "Data Devolução"
        .TextMatrix(0, 2) = "Peso Liq."
        .TextMatrix(0, 3) = "Preço/Kg"
        .TextMatrix(0, 4) = "Val. a Pagar"
        .TextMatrix(0, 5) = "Data Saída"
    End With
    
    'preenche GridPagas
    With GridPagas
        .ColWidth(0) = 1200
        .ColWidth(1) = 1300
        .TextMatrix(0, 0) = "Referência"
        .TextMatrix(0, 1) = "Data Devolução"
        .TextMatrix(0, 2) = "Peso Liq."
        .TextMatrix(0, 3) = "Peso Bruto"
        .TextMatrix(0, 4) = "Preço/Kg"
        .TextMatrix(0, 5) = "Val. a Pagar"
    End With
    lblCodigo = Caixa & ": " & NomeDescascador
    Limite = IIf(GetSetting("Descasque", "Fechamento", "Show Records") = "", 100, GetSetting("Descasque", "Fechamento", "Show Records"))
    
    'preenche tabPagas
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT REFERENCIA, DATA_DEV, PESO_LIQ, PESO_BRUTO, PRECO, DATA_SAIDA, PAGAR FROM tbl_entrega WHERE cod_desc = " & Caixa & " AND data_dev <> 0 AND quitado = 1 ORDER BY data_dev DESC", con
    While Not RS.EOF
        Valor = CSng(RS!pagar) + CSng(Valor)
        'Só exibe até o limite configurado em ferramentas
        If i <= Limite Then
            GridPagas.AddItem RS!REFERENCIA & vbTab & Format(RS!data_dev, "dd/mm/yyyy") & vbTab & RS!peso_liq & vbTab & RS!peso_bruto & vbTab & RS!preco & vbTab & RS!pagar & vbTab & Format(RS!data_saida, "dd/mm/yyyy")
        End If
        RS.MoveNext
        i = i + 1
    Wend
    txtTotalPagas = Valor
    Valor = 0
    RS.Close
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set formRelatorios = Nothing

End Sub

Private Sub txtDataAte_LostFocus()

    If Not IsDate(txtDataAte.FormattedText) And txtDataAte.Text <> "__/__/____" Then
        MsgBox "Data Inválida. Digite novamente para prosseguir."
        txtDataAte.SetFocus
        txtDataDe.SelStart = txtDataDe.SelLength
    End If

End Sub

Private Sub txtDataDe_LostFocus()

    If Not IsDate(txtDataDe.FormattedText) And txtDataDe.Text <> "__/__/____" Then
        MsgBox "Data Inválida. Digite novamente para prosseguir."
        txtDataDe.SetFocus
        txtDataDe.SelStart = txtDataDe.SelLength
    End If
    
End Sub
