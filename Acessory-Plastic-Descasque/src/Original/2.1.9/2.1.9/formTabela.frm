VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formTabela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de Pre�os e Mat�ria Prima"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ForeColor       =   &H8000000F&
   Icon            =   "formTabela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAlterarTudo 
      Caption         =   "Alt. Tudo"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Alterar dados de produtos"
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdMP 
      Caption         =   "&MP"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sair"
      Top             =   4380
      Width           =   915
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Incluir novo produto"
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":17A8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Excluir produto selecionado"
      Top             =   2700
      Width           =   915
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":1AB2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir relat�rio de produtos"
      Top             =   3540
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":4254
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sair"
      Top             =   5220
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4545
      Left            =   120
      TabIndex        =   13
      Top             =   1500
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   8017
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      HighLight       =   2
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Cadastro de Produtos e Pre�os"
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      TabIndex        =   12
      Top             =   180
      Width           =   5475
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtReferencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   15
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4620
         MaxLength       =   6
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtCavidades 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4620
         MaxLength       =   2
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C�digo"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Refer�ncia"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pre�o"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cavidades"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Top             =   720
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":4B1E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Alterar dados de produtos"
      Top             =   1020
      Width           =   915
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "&Concluir"
      Height          =   795
      Left            =   5820
      Picture         =   "formTabela.frx":4E28
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Concluir inclus�es ou altera��es"
      Top             =   1020
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "formTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim codBuff As Integer 'ARMAZENA O VALOR DO CODIGO DA REFERENCIA ANTES DE ALTERAR OS VALORES DELA

Private Sub cmdImprimir_Click()
    
    Me.Hide
    formPrincipal.Enabled = False
    formRelaReferencias.Config
    formPrincipal.Enabled = True
    Me.Show 1
    
End Sub

Private Sub cmdAlterar_Click()
Dim X As Integer

    'ANTERIORMENTE, AQUI HAVIA UMA CONSULTA AO BD, PARA PEGAR OS DADOS DA REFER�NCIA.
    'POR�M FOI RETIRADA, POIS A PROBABILIDADE DE ALGU�M ALTERAR OU EXCLUIR ESTA REFER�NCIA AO MESMO TEMPO � MUITO REMOTA.
    Me.MousePointer = vbHourglass
    If Grid1.Text <> "" And Grid1.Col = 0 And Grid1.Row <> 0 Then
        txtReferencia = Grid1.Text
        With Grid1
            X = Grid1.RowSel
             txtCodigo = .TextMatrix(X, 1)
             'USADO PARA SABER SE O C�DIGO FOI ALTERADO OU N�O.
             codBuff = .TextMatrix(X, 1)
             txtPreco = .TextMatrix(X, 2)
             txtPeso = .TextMatrix(X, 3)
             txtCavidades = .TextMatrix(X, 4)
        End With
        txtReferencia.Locked = True
        cmdCadastrar.Enabled = False
        cmdAlterarTudo.Enabled = False
        cmdExcluir.Enabled = False
        cmdMP.Enabled = False
        cmdImprimir.Enabled = False
        cmdAlterar.Visible = False
        cmdFinalizar.Visible = True
        'Evita que o usuario tire o foco da celula
        Grid1.Enabled = False
    Else
        MsgBox "Selecione a Refer�ncia na tabela para prosseguir!", vbExclamation
    End If
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdAlterarTudo_Click()
Dim Retorno As String
Dim N As Integer
    
    N = MsgBox("Aten��o! voc� est� prestes a atualizar toda a tabela de pre�os. Tem certeza de que deseja prosseguir?", vbYesNo + vbQuestion)
    If N = vbNo Then Exit Sub
    Retorno = InputBox("Digite o percentual de corre��o:", "Atualizar toda a Tabela")
    If Not IsNumeric(Retorno) Then
        MsgBox "Digite um valor v�lido!", vbCritical
        Exit Sub
    End If
    con.Execute ("UPDATE TBL_PRECOS SET PRECO = PRECO + ((PRECO * '" & FormatFloatForDB(Retorno) & "') / 100)")
    RSReferencias.Requery
    MsgBox "Tabela atualizada! A janela atual ser� fechada. Abra-a em seguida.", vbInformation
    Unload Me

End Sub

Private Sub cmdCadastrar_Click()
    
    'VERIFICA SE A REFERENCIA OU O C�DIGO J� EXISTEM
    If txtReferencia <> "" And txtPreco <> "" Then
        Set RS = con.Execute("SELECT * FROM tbl_precos WHERE REFERENCIA = '" & txtReferencia & "' OR CODIGO = " & txtCodigo & "")
        'SE A REFERENCIA OU O C�DIGO NAO EXISTIREM, CADASTRA
        If RS.EOF Then
            con.Execute ("INSERT INTO tbl_precos (referencia, preco, peso, cavidades, codigo) values('" & txtReferencia.Text & "', '" & FormatFloatForDB(txtPreco.Text) & "', '" & FormatFloatForDB(txtPeso.Text) & "', '" & txtCavidades & "', " & txtCodigo & ")")
            txtReferencia = ""
            txtPreco = ""
            txtPeso = ""
            txtCavidades = ""
            txtCodigo = ""
            RS.Close
        Else
            'SE A REFERENCIA OU O C�DIGO EXISTIREM, AVISA E SAI DA ROTINA
            MsgBox "Esta refer�ncia ou o c�digo j� est�o cadastrados! Verifique e tente novamente.", vbCritical
            RS.Close
            Exit Sub
        End If
    Else
        MsgBox "� necess�rio preencher todos os campos!", vbExclamation
        Exit Sub
    End If
    Grid1.Rows = 1
    RSReferencias.Requery
    If RSReferencias.EOF = False Then
        While Not RSReferencias.EOF
            Grid1.AddItem RSReferencias!REFERENCIA & vbTab & RSReferencias!codigo & vbTab & RSReferencias!preco & vbTab & RSReferencias!Peso & vbTab & RSReferencias!Cavidades
            RSReferencias.MoveNext
        Wend
    End If
    txtReferencia.SetFocus
    RSReferencias.MoveFirst
    
End Sub

Private Sub cmdExcluir_Click()
    
    If Grid1.Text <> "" And Grid1.Col = 0 And Grid1.Row <> 0 Then
        Resposta = MsgBox("Tem certeza de que deseja excluir a refer�ncia " & Grid1.Text, vbYesNo + vbQuestion)
        If Resposta = vbNo Then Exit Sub
        con.Execute ("DELETE FROM tbl_precos WHERE referencia = '" & Grid1.Text & "'")
        If Grid1.Rows = 2 Then
            Grid1.Rows = 1
        Else
            Grid1.RemoveItem Grid1.RowSel
        End If
        RSReferencias.Requery
    Else
        MsgBox "Selecione uma refer�ncia v�lida!"
    End If
    
End Sub

Private Sub cmdFinalizar_Click()
Dim X As Integer

    Me.MousePointer = vbHourglass
    'CASO O USUARIO TENHA ALTERADO O C�DIGO, � OBRIGAT�RIO VERIFICAR SE J� EXISTE UMA
    'REFERENCIA CADASTRADA COM ESSE MESMO CODIGO, PARA EVITAR DUPLICA��ES.
    'SE O USUARIO N�O ALTEROU O CODIGO, FINALIZA NORMALMENTE
    If txtCodigo = "" Then txtCodigo = 0 'evita erro se o usuario esquecer o campo vazio
    If txtCodigo = codBuff Then
        con.Execute ("UPDATE tbl_PRECOS SET preco = '" & IIf(txtPreco = "", 0, FormatFloatForDB(txtPreco)) & "', Peso = '" & IIf(txtPeso = "", 0, FormatFloatForDB(txtPeso)) & "', Cavidades = '" & IIf(txtCavidades = "", 0, txtCavidades) & "' WHERE REFERENCIA = '" & txtReferencia & "'")
    Else
        'SE O USUARIO ALTEROU O CODIGO, FAZ UMA CONSULTA NO BD PARA SABER SE O CODIGO J� EST� CADASTRADO
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT * FROM tbl_precos WHERE CODIGO = " & txtCodigo, con
        If RS.EOF Or txtCodigo = 0 Then
            'O CODIGO PODE SER ALTERADO SE MUDAR DE POSI��O NO CABIDEIRO
            'SE NAO TIVER CODIGO CADASTRADO, O VALOR PADR�O ZERO ESTAR� ATRIBU�DO
            con.Execute ("UPDATE tbl_PRECOS SET preco = '" & IIf(txtPreco = "", 0, FormatFloatForDB(txtPreco)) & "', Peso = '" & IIf(txtPeso = "", 0, FormatFloatForDB(txtPeso)) & "', Cavidades = '" & IIf(txtCavidades = "", 0, txtCavidades) & "', codigo = '" & IIf(txtCodigo = 0, 0, txtCodigo) & "' " & _
                         "WHERE REFERENCIA = '" & txtReferencia & "'")
            RS.Close
        Else
            'FECHO O RS E SAIO DA SUB PARA PERMITIR AO USUARIO TENTAR ALTERAR O VALOR NOVAMENTE
            RS.Close
            MsgBox "Este c�digo j� est� cadastrado! Verifique e tente novamente.", vbCritical
            Me.MousePointer = vbDefault
            txtCodigo.SetFocus
            Exit Sub
        End If
    End If
    
    cmdFinalizar.Visible = False
    cmdExcluir.Visible = True
    cmdCadastrar.Enabled = True
    cmdExcluir.Enabled = True
    cmdAlterar.Visible = True
    cmdAlterarTudo.Enabled = True
    cmdMP.Enabled = True
    cmdImprimir.Enabled = True
    txtReferencia.Locked = False
    'Altera os dados exibidos no grid
    With Grid1
        X = Grid1.RowSel
        .TextMatrix(X, 1) = txtCodigo
        .TextMatrix(X, 2) = txtPreco
        .TextMatrix(X, 3) = txtPeso
        .TextMatrix(X, 4) = txtCavidades
    End With
    txtReferencia = ""
    txtPreco = ""
    txtPeso = ""
    txtCavidades = ""
    txtCodigo = ""
    'devolve o foco � celula
    Grid1.Enabled = True
    
    RSReferencias.Requery
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdMP_Click()

    formMP.Show 1

End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    With Grid1
        .ColWidth(0) = 1400
        .TextMatrix(0, 0) = "Refer�ncia"
        .TextMatrix(0, 1) = "C�digo"
        .TextMatrix(0, 2) = "Pre�os"
        .TextMatrix(0, 3) = "Peso"
        .TextMatrix(0, 4) = "Cavidades"
    End With
    'Se o RSReferencias estiver vazio, n�o preenche o grid e sai da rotina
    If (RSReferencias.BOF And RSReferencias.EOF) Then Exit Sub
    RSReferencias.MoveFirst
    If RSReferencias.EOF = False Then
        While Not RSReferencias.EOF
            Grid1.AddItem RSReferencias!REFERENCIA & vbTab & RSReferencias!codigo & vbTab & RSReferencias!preco & vbTab & RSReferencias!Peso & vbTab & RSReferencias!Cavidades
            RSReferencias.MoveNext
        Wend
    End If
    RSReferencias.MoveFirst
    
End Sub

Private Sub Grid1_Click()

    Grid1.CellBackColor = vbBlue
    Grid1.CellForeColor = vbWhite

End Sub

Private Sub Grid1_LeaveCell()

    Grid1.CellBackColor = &H80000005
    Grid1.CellForeColor = &H80000012
    
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtPeso, KeyAscii)

End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    
    KeyAscii = TypeCurrency(txtPreco, KeyAscii)

End Sub

