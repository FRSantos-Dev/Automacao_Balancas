VERSION 5.00
Begin VB.Form formPrincipal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accessory Plastic - Menu Principal"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Sair"
      Height          =   915
      Index           =   5
      Left            =   2640
      Picture         =   "formPrincipal2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Produtos"
      Height          =   915
      Index           =   1
      Left            =   1380
      Picture         =   "formPrincipal2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Ferramentas"
      Height          =   915
      Index           =   4
      Left            =   1380
      Picture         =   "formPrincipal2.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Height          =   915
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Produção"
      Height          =   915
      Index           =   2
      Left            =   2640
      Picture         =   "formPrincipal2.frx":3376
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Descasque"
      Height          =   915
      Index           =   0
      Left            =   120
      Picture         =   "formPrincipal2.frx":5B18
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "formPrincipal"
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
        'Abre_Controle
        Me.Hide
        FormDescasque.Show
    Case 1
        formTabela.Show 1
    Case 2
        formProducao.Show 'Existe uma rotina para transforma-lo em modal
    Case 3
        'formGrafico.Show
    Case 4
        formFerramentas.Show 1
    Case 5
        End
    End Select

End Sub

Private Sub Form_Activate()

    For i = 0 To 5
        Efeito3d Command1(i), 3, 3, False
        DoEvents
    Next

End Sub

Private Sub Form_Load()

    'Nao é possivel setar esta propriedade em tempo de design pois ele é
    'um MDIChild
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    AbrirConexao 'adicionado para o mysql++++++++++++++++++++++++++++++++++++++++++++

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
