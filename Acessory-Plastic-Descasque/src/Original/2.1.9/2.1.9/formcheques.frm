VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form formCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de cheques"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formcheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Porta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2340
      Picture         =   "formcheques.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   1260
      Width           =   7575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ejetar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1260
      Picture         =   "formcheques.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox txtAno 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8340
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "2001"
      Top             =   2400
      Width           =   675
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "formcheques.frx":2EEE
      Left            =   6120
      List            =   "formcheques.frx":2F16
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtDia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "01"
      Top             =   2400
      Width           =   435
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7500
      TabIndex        =   0
      Text            =   "0,00"
      Top             =   240
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   180
      Picture         =   "formcheques.frx":2F7F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2100
      Width           =   1035
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   60
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Line Line6 
      X1              =   180
      X2              =   9360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label6 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Tag             =   "a"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7920
      TabIndex        =   11
      Top             =   2460
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   10
      Top             =   2460
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Rio,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   9
      Top             =   2460
      Width           =   435
   End
   Begin VB.Line Line5 
      X1              =   9240
      X2              =   4500
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label lblValor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   720
      Width           =   7575
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   9300
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   9300
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line2 
      X1              =   9300
      X2              =   9300
      Y1              =   180
      Y2              =   660
   End
   Begin VB.Line Line1 
      X1              =   7380
      X2              =   7380
      Y1              =   180
      Y2              =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Pague por este cheque a quantia de"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "formCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Favorecido As String
Dim Dia As String
Dim mes As String
Dim Ano As String

     If txtValor = "" Or txtValor = "0,00" Then
         MsgBox "Preencha o Valor antes de prosseguir!", vbCritical
         Exit Sub
     End If
         If txtDia = "" Then
         MsgBox "Preencha o Dia antes de prosseguir!", vbCritical
         Exit Sub
     End If
     If txtAno = "" Then
         MsgBox "Preencha o Ano antes de prosseguir!", vbCritical
         Exit Sub
     End If
     If Combo1.Text = "" Then
         MsgBox "Selecione o Mês antes de prosseguir!", vbCritical
         Exit Sub
     End If
    
    ' Buffer to hold input string
    Dim InString As String
    ' Use COM1.
    MSComm1.CommPort = GetSetting("Descasque", "Portas", "Configuracao", "2")
    ' 9600 baud, no parity, 8 data, and 1 stop bit.
    MSComm1.Settings = "9600,N,8,1"
    ' Tell the control to read entire buffer when Input
    ' is used.
    MSComm1.InputLen = 0
    ' Abre a porta
    MSComm1.PortOpen = True
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXX INICIO DA ATRIBUICAO DOS DADOS XXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Banco
    Send_Signal "341", "Banco" 'ITAU
    'CIDADE Rio de Janeiro
    MSComm1.Output = Chr(27) & Chr(67) & Chr(82) & Chr(105) & Chr(111) & Chr(32) & Chr(100) & Chr(101) & Chr(32) & Chr(106) & Chr(97) & Chr(110) & Chr(101) & Chr(105) & Chr(114) & Chr(111) & Chr(36)
    'DATA
    Dia = Format(txtDia, "00")
    mes = Format(CheckMonth(Combo1.Text), "00")
    Ano = Format(Right(txtAno, 2), "00")
    Send_Signal CStr(Dia & mes & Ano), "Data"
    'Favorecido
    Send_Signal txtNome, "Favorecido"
    'VALOR
    Send_Signal txtValor, "Valor"
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXX FIM DA ATRIBUICAO DOS DADOS XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    'Fecha a porta
    MSComm1.PortOpen = False
    
End Sub

Private Sub Command2_Click()
   
    'Ejetar o cheque
On Error GoTo TrataErro

    MSComm1.PortOpen = True
    MSComm1.Output = Chr(12) & Chr(36)
    MSComm1.PortOpen = False
    Exit Sub

TrataErro:
    MsgBox "Houve um erro na abertura ou no fechamento da Porta Serial! Se o erro continuar ocorrendo, reinicie o computador.", vbCritical
    
End Sub

Private Sub Command3_Click()
Dim Resposta As Integer

    Resposta = MsgBox("Somente altere estas configurações se a COM2 não estiver funcionando ou seja necessário usar outra porta serial! Deseja prosseguir?", vbCritical + vbYesNo)
    If Resposta = vbYes Then formPortas.Show 1
    
End Sub

Private Sub Form_Load()
Dim mes As Integer
    
    'rotina para colocar o mes corrente no combo
    txtDia = Day(Date)
    'A varivel mes será o valor do listindex, portanto deve ser sempre -1
    mes = Month(Date) - 1
    Combo1.ListIndex = mes
    txtAno = Year(Date)

End Sub


Private Sub txtAno_KeyPress(KeyAscii As Integer)

    KeyAscii = OnlyNumbers(txtAno, KeyAscii)
    
End Sub

Private Sub txtAno_LostFocus()

    If txtAno = "" Then Exit Sub
    If CInt(txtAno.Text) <> Year(Date) Then
        MsgBox "O ano digitado é diferente que o ano atual!"
    End If

End Sub

Private Sub txtDia_KeyPress(KeyAscii As Integer)

    KeyAscii = OnlyNumbers(txtDia, KeyAscii)

End Sub

Private Sub txtDia_LostFocus()
    
    If txtDia = "" Then Exit Sub
    If CInt(txtDia) > 31 Then
        MsgBox "Dia inválido!"
        txtDia.SetFocus
    End If

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtValor, KeyAscii)

End Sub

Private Function CheckMonth(mes As String) As Byte

    Select Case mes
    Case "Janeiro"
        CheckMonth = 1
    Case "Fevereiro"
        CheckMonth = 2
    Case "Março"
        CheckMonth = 3
    Case "Abril"
        CheckMonth = 4
    Case "Maio"
        CheckMonth = 5
    Case "Junho"
        CheckMonth = 6
    Case "Julho"
        CheckMonth = 7
    Case "Agosto"
        CheckMonth = 8
    Case "Setembro"
        CheckMonth = 9
    Case "Outubro"
        CheckMonth = 10
    Case "Novembro"
        CheckMonth = 11
    Case "Dezembro"
        CheckMonth = 12
    End Select

End Function

Private Sub txtValor_LostFocus()

    If txtValor = "" Then Exit Sub
    lblValor = ValExtenso(txtValor)
    
End Sub

'Converte para o valor decimal
Private Sub Send_Signal(ByVal STR As String, ByVal Campo As String)
Dim Code As String
Dim i As Integer

    'Exclui os caracteres que nao devem ser processados
    STR = GetRid(STR)
    Select Case Campo
    Case "Banco"
        'Envia ESC F
        MSComm1.Output = Chr(27) & Chr(66)
        For i = 1 To Len(STR)
            Code = Asc(Mid(STR, i, 1))
            MSComm1.Output = Chr(Code)
            DoEvents
        Next
        MSComm1.Output = Chr(36)
    Case "Favorecido"
        'Envia ESC F
        MSComm1.Output = Chr(27) & Chr(70)
        For i = 1 To Len(STR)
            Code = Asc(Mid(STR, i, 1))
            MSComm1.Output = Chr(Code)
            DoEvents
        Next
        MSComm1.Output = Chr(36)
    Case "Cidade"
        'Envia ESC C
        MSComm1.Output = Chr(27) & Chr(67)
        For i = 1 To Len(STR)
            Code = Asc(Mid(STR, i, 1))
            MSComm1.Output = Chr(Code)
            DoEvents
        Next
        MSComm1.Output = Chr(36)
    Case "Valor"
        'Envia ESC V
        MSComm1.Output = Chr(27) & Chr(118)
        For i = 1 To Len(STR)
            Code = Asc(Mid(STR, i, 1))
            MSComm1.Output = Chr(Code)
            DoEvents
        Next
        MSComm1.Output = Chr(36)
    Case "Data"
        'Envia ESC D
        MSComm1.Output = Chr(27) & Chr(68)
        For i = 1 To Len(STR)
            Code = Asc(Mid(STR, i, 1))
            MSComm1.Output = Chr(Code)
            DoEvents
        Next
    End Select
    
End Sub

Function GetRid(ByVal A As String) As String

    'substitui os caracteres selecionados por expaco vazio
    Dim CHARS As String, i As Integer, B As String

    CHARS = "()<>][}{',.@#$%&*+\|/?!"
    B = A
    For i = 1 To Len(CHARS)
        B = Replace(B, Mid(CHARS, i, 1), "")
    Next

    GetRid = B

End Function

