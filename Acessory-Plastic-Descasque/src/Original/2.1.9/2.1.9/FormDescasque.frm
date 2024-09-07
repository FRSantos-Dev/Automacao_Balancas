VERSION 5.00
Begin VB.Form FormDescasque 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descasque"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   1590
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "FormDescasque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3855
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Fechamento"
      Height          =   915
      Index           =   3
      Left            =   60
      Picture         =   "FormDescasque.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Cheques"
      Height          =   915
      Index           =   2
      Left            =   2610
      Picture         =   "FormDescasque.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Controle"
      Height          =   915
      Index           =   0
      Left            =   60
      Picture         =   "FormDescasque.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Descascador"
      Height          =   915
      Index           =   1
      Left            =   1320
      Picture         =   "FormDescasque.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Relatórios"
      Height          =   915
      Index           =   4
      Left            =   1320
      Picture         =   "FormDescasque.frx":14F2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Sair"
      Height          =   915
      Index           =   5
      Left            =   2610
      Picture         =   "FormDescasque.frx":17FC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   1155
   End
End
Attribute VB_Name = "FormDescasque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'por: HELIO CANDIDO

'Função melhorada e adaptada por mim

'Os códigos abaixo colocam uma borda com um efeito 3D muito interessante
'vale a pena usá-la, confira

'Para poder usar esse efeito use a função abaixo

'A opção Ativando pode ser igual a True ou False (Inner ou OuterBevel)
'Use Autoredraw = true no formulario (added by Marcos Alessandro in 10/09/2001)

Public Sub Efeito3d(Objeto As Control, LarguraBorda As Integer, EspacoBorda As Integer, Ativando%)

    MakeIt3D Objeto, LarguraBorda, EspacoBorda, Ativando
    
End Sub

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
    Case 0
        Me.Hide
        Abre_Controle
    Case 1
        Me.Hide
        formDescascadores.Show 1
        Me.Visible = True
    Case 2
        Me.Hide
        Abre_Cheque
        Me.Visible = True
    Case 3
        Me.Hide
        Abre_Relatorios
        Me.Visible = True
    Case 4
        Me.Hide
        formFechamento.Show
    Case 5
        Unload Me
        formPrincipal.Visible = True
    End Select
    
End Sub

Private Sub Form_Activate()

    For i = 0 To 5
        Efeito3d Command1(i), 3, 3, False
        DoEvents
    Next

End Sub


Private Sub Abre_Cheque()

    'Utilizo InputBox para entrada do codigo do descascador
    Titulo = "Impressão de cheques"
    Mensagem = "Digite o código do descascador."
    Caixa = InputBox(Mensagem, Titulo)
    If Caixa <> "" Then
        Set RS = con.Execute("SELECT nome FROM tbl_descascador WHERE codigo = " & Caixa & "")
        If RS.EOF = False Then
            formCheques.txtNome = RS!nome
        Else
            MsgBox "Código não cadastrado!", vbExclamation
            Me.Visible = True
            Exit Sub
        End If
        RS.Close
        formCheques.Show 1
    End If
    Me.Visible = True

End Sub

Private Sub Abre_Controle()

'Trocar a consulta por uma matriz com dois vetores na inicializacao do sistema,
'contendo um array com o codigo e nome da pessoa, para acelerar a abertura da janela.
'Caso seja feito o cadastramento de um descascador em outra máquina, basta que seja reinicializado o programa

    'Utilizo InputBox para entrada do codigo do descascador
    Titulo = "Controle do Descasque"
    Mensagem = "Digite o código do descascador."
    Caixa = InputBox(Mensagem, Titulo)
    If Caixa <> "" Then
        Set RS = con.Execute("SELECT nome FROM tbl_descascador WHERE codigo = " & Caixa & "")
        If RS.EOF = False Then
            formPrincipal.Tag = RS!nome
        Else
            MsgBox "Código não cadastrado!", vbExclamation
            RS.Close
            Me.Visible = True
            Exit Sub
        End If
        RS.Close
        formControle.Show 1
    End If
    Me.Visible = True

End Sub

Private Sub Abre_Relatorios()

    'Utilizo InputBox para entrada do codigo do descascador
    Titulo = "Relatórios"
    Mensagem = "Digite o código do descascador."
    Caixa = InputBox(Mensagem, Titulo)
    If Caixa <> "" Then
        Set RS = con.Execute("SELECT nome FROM tbl_descascador WHERE codigo = " & Caixa & "")
        If RS.EOF = False Then
            NomeDescascador = RS!nome
        Else
            MsgBox "Código não cadastrado!", vbExclamation
            Me.Visible = True
            RS.Close
            Exit Sub
        End If
        RS.Close
        formRelatorios.Show 1
    End If
    Me.Visible = True

End Sub

Private Sub MakeIt3D(Ctrl As Control, nBevel%, nSpace%, bInset%)
'Parte do código que gera o efeito nos botões

     PixX% = Screen.TwipsPerPixelX
     PixY% = Screen.TwipsPerPixelY

    With Ctrl
     CTop% = .Top - PixX%
     CLft% = .Left - PixY%
     CRgt% = .Left + .Width
     CBtm% = .Top + .Height

     If bInset% Then

     For i% = nSpace% To (nBevel% + nSpace% - 1)
        AddX% = i% * PixX%
        AddY% = i% * PixY%
        .Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), 8421504
        .Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), 8421504
        .Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), &HFFFFFF
        .Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), 16777215
     Next

     Else

     For i% = nSpace% To (nBevel% + nSpace% - 1)
        AddX% = i% * PixX%
        AddY% = i% * PixY%
        .Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), 8421504
        .Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), 8421504
        .Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), &HFFFFFF
        .Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), 16777215
     Next

     End If
    
    End With
    
End Sub

Private Sub Form_Load()

    'Nao é possivel setar esta propriedade em tempo de design pois ele é
    'um MDIChild
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2

End Sub
