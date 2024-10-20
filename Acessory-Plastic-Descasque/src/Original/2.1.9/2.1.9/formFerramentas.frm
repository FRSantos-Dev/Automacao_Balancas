VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form formFerramentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ferramentas"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "formFerramentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5636
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Backup       "
      TabPicture(0)   =   "formFerramentas.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Drive1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Dir1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Configurar"
      TabPicture(1)   =   "formFerramentas.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Caminho para o BD"
      TabPicture(2)   =   "formFerramentas.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Excluir Registros"
      TabPicture(3)   =   "formFerramentas.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Reparar Tabelas"
      TabPicture(4)   =   "formFerramentas.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "E-Mail BD"
      TabPicture(5)   =   "formFerramentas.frx":0396
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame7"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Digite o Endereço de e-mail do destinatário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   480
         TabIndex        =   32
         Top             =   720
         Width           =   6135
         Begin VB.ListBox lstStatus 
            BackColor       =   &H8000000F&
            Height          =   1035
            Left            =   120
            TabIndex        =   36
            Top             =   780
            Width           =   4200
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Enviar"
            Height          =   915
            Index           =   3
            Left            =   4800
            Picture         =   "formFerramentas.frx":03B2
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            TabIndex        =   33
            Text            =   "marcos@itcase.com.br"
            Top             =   435
            Width           =   3210
         End
         Begin VB.Label lblProgress 
            Alignment       =   2  'Center
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1725
            TabIndex        =   37
            Top             =   1875
            Width           =   1290
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "E-Mail"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   435
            Width           =   1035
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Reparar Tabelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74520
         TabIndex        =   30
         Top             =   720
         Width           =   5895
         Begin VB.CommandButton cmdVerificar 
            Caption         =   "&Reparar"
            Height          =   855
            Left            =   4800
            Picture         =   "formFerramentas.frx":06BC
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Excluir usuário selecionado"
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "Excluir registros de Produção:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -74760
         TabIndex        =   22
         Top             =   1800
         Width           =   6495
         Begin VB.ComboBox cboTurno 
            Height          =   315
            ItemData        =   "formFerramentas.frx":09C6
            Left            =   3720
            List            =   "formFerramentas.frx":09D3
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluirProd 
            Caption         =   "E&xcluir"
            Height          =   795
            Left            =   5400
            Picture         =   "formFerramentas.frx":09EC
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Excluir usuário selecionado"
            Top             =   240
            Width           =   915
         End
         Begin MSMask.MaskEdBox txtDataIni 
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
            TabIndex        =   24
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
         Begin MSMask.MaskEdBox txtDataFim 
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
            TabIndex        =   25
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Turno:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2640
            TabIndex        =   29
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data Fim:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data Início:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
         Height          =   915
         Index           =   1
         Left            =   -69000
         Picture         =   "formFerramentas.frx":0CF6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   915
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Digite o Caminho para o SGBD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   -74520
         TabIndex        =   17
         Top             =   720
         Width           =   6135
         Begin VB.TextBox txtEndereco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Top             =   735
            Width           =   3210
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Salvar"
            Height          =   915
            Index           =   0
            Left            =   4800
            Picture         =   "formFerramentas.frx":1000
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Endereço IP"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   735
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Exibir Registros do Estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   13
         Top             =   660
         Width           =   2655
         Begin VB.OptionButton optTodos 
            Caption         =   "    Todos"
            Height          =   255
            Left            =   375
            TabIndex        =   15
            Top             =   690
            Width           =   1095
         End
         Begin VB.TextBox txtQTD 
            Height          =   285
            Left            =   150
            TabIndex        =   14
            Text            =   "100"
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Últimos"
            Height          =   195
            Left            =   825
            TabIndex        =   16
            Top             =   375
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Exibir Registros de Fechamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -72075
         TabIndex        =   9
         Top             =   660
         Width           =   2955
         Begin VB.TextBox txtQTD2 
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Text            =   "100"
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton optTodos2 
            Caption         =   "    Todos"
            Height          =   255
            Left            =   375
            TabIndex        =   10
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Últimos"
            Height          =   195
            Left            =   825
            TabIndex        =   12
            Top             =   375
            Width           =   510
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Excluir registros do Descasque com data anterior a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   6495
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Excluir"
            Height          =   795
            Left            =   5400
            Picture         =   "formFerramentas.frx":130A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Excluir usuário selecionado"
            Top             =   240
            Width           =   915
         End
         Begin MSMask.MaskEdBox txtData 
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
            TabIndex        =   6
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1080
         End
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   3
         Top             =   1260
         Width           =   2715
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -74760
         TabIndex        =   2
         Top             =   900
         Width           =   2715
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Backup"
         Height          =   915
         Index           =   2
         Left            =   -70140
         Picture         =   "formFerramentas.frx":1614
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selecione o local do Backup"
         Height          =   195
         Left            =   -74700
         TabIndex        =   4
         Top             =   660
         Width           =   2040
      End
   End
End
Attribute VB_Name = "formFerramentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs As FileSystemObject
Dim Email As String
Dim WithEvents vSendMail As vbSendMail.clsSendMail
Attribute vSendMail.VB_VarHelpID = -1

'------------declarações do ShellAndWait------------------
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&
'------------declarações do ShellAndWait------------------

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdExcluir_Click()
Dim dData As Date
Dim a As Long

    If Not IsDate(txtData) Then
        MsgBox "Preencha o campo Data antes de prosseguir!"
        Exit Sub
    End If
    dData = txtData

    If MsgBox("Tem certeza de que deseja excluir todos os registros com data anterior a " & txtData & " ?" & Chr(13) & "Os registros excluídos não poderão mais ser recuperados.", vbYesNo + vbExclamation) = vbYes Then
        Me.MousePointer = vbHourglass
        con.Execute "DELETE FROM TBL_ENTREGA WHERE DATA_DEV < '" & Format(dData, "yyyy/mm/dd") & "' AND QUITADO = 1", a
        Me.MousePointer = vbDefault
        MsgBox a & " Registros excluídos!", vbExclamation
    End If

End Sub

Private Sub cmdExcluirProd_Click()
Dim Turno As String
Dim a As Long

    If (IsDate(txtDataIni) = False) Or (IsDate(txtDataFim) = False) Then
        MsgBox "Digite as datas de início e fim do intervalo!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Você tem certeza de que deseja excluir estes registros de Produção?", vbYesNo + vbQuestion) = vbYes Then
        Me.MousePointer = vbHourglass
        If cboTurno = "Manhã" Then Turno = " AND TURNO = '0'"
        If cboTurno = "Noite" Then Turno = " AND TURNO = '1'"
        
        'RECUPERO A MATÉRIA PRIMA USADA E DEVOLVO AO ESTOQUE
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.CursorLocation = adUseClient
        RS.Open "SELECT SUM(MP_CONSUMIDA) AS MP_CONS, COD_MP FROM TBL_DADOS WHERE DATA BETWEEN '" & Format(txtDataIni, "yyyy/mm/dd") & "' AND '" & Format(txtDataFim, "yyyy/mm/dd") & "'" & Turno & " GROUP BY MP_CONSUMIDA", con
        con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE + '" & FormatFloatForDB(RS!MP_CONS) & "') WHERE ID = '" & RS!COD_MP & "'"
        RS.Close
        
        'DEPOIS DE DEVOLVER A MP AO ESTOQUE, EXCLUI DA TABELA DADOS
        con.Execute "DELETE FROM TBL_DADOS WHERE DATA BETWEEN '" & Format(txtDataIni, "yyyy/mm/dd") & "' AND '" & Format(txtDataFim, "yyyy/mm/dd") & "'" & Turno, a
        MsgBox a & " Registros de Produção foram excluídos.", vbInformation
    End If
    Me.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdVerificar_Click()
Dim BUFF As String

    'VERIFICA SE A TABELA ESTÁ UP TO DATE
    Set RS = con.Execute("CHECK TABLE TBL_ENTREGA FAST QUICK")
    If RS!MSG_TEXT = "Table is already up to date" Then
        MsgBox "NÃO HÁ QUALQUER PROBLEMA COM A TABELA!"
    Else
        'SE HOUVER PROBLEMA COM A TABELA, FAZ O REAPRO E EXIBE UMA MSG
        Set RS = con.Execute("REPAIR TABLE TBL_ENTREGA")
        While Not RS.EOF
            BUFF = BUFF & RS!MSG_TYPE & ": " & RS!MSG_TEXT & vbCrLf
            RS.MoveNext
        Wend
        MsgBox BUFF
    End If
    RS.Close

End Sub

Private Sub Command1_Click(Index As Integer)
Dim IP As String

    Select Case Index
    Case 0
        'ENDERECO DO SERVIDOR
        IP = IIf(txtEndereco = "", "127.0.0.1", txtEndereco)
        SaveSetting "Descasque", "BaseDados", "PathMySQL", IP
        MsgBox "A alteração no caminho para o Banco de Dados foi efetuada com sucesso!", vbInformation
    Case 1
        'salva no registro a qtd de linhas a serem exibidas no estoque
        If optTodos.Value = True Then
            SaveSetting "Descasque", "Estoque", "Show Records", "999999"
        Else
            'Tem que haver algum valor no registro
            If txtQTD = "" Or Not IsNumeric(txtQTD) Then
                MsgBox "Não foi possível gravar no registro. Não foi inserido nenhum valor ou não é um valor válido"
                Exit Sub
            End If
            SaveSetting "Descasque", "Estoque", "Show Records", txtQTD
        End If
        
        'salva no registro a qtd de linhas a serem exibidas no relatorio de fechamento
        If optTodos2.Value = True Then
            SaveSetting "Descasque", "Fechamento", "Show Records", "999999"
        Else
            'Tem que haver algum valor no registro
            If txtQTD2 = "" Or Not IsNumeric(txtQTD2) Then
                MsgBox "Não foi possível gravar no registro. Não foi inserido nenhum valor ou não é um valor válido"
                Exit Sub
            End If
            SaveSetting "Descasque", "Fechamento", "Show Records", txtQTD2
        End If
    Case 2
        'Backup do BD. só pode ser feito a partir do servidor
        Me.MousePointer = vbHourglass
        MsgBox "Aguarde alguns instantes até que o Backup seja efetuado. O arquivo pode ser localizado na pasta " & Dir1.Path
        CompactBD Dir1.Path & "\"
        Me.MousePointer = vbDefault
        MsgBox "Operação realizada com sucesso!", vbInformation, "Backup"
    Case 3
        If MsgBox("Esta implementação só permite que o BD seja enviado a partir do Servidor de Arquivos. Deseja continuar?", vbYesNo + vbCritical) = vbYes Then
            CompactBD
            Email = IIf(txtEmail = "", "marcos@itcase.com.br", txtEmail)
            
            'Se ocorrer o erro Cant Create Object, registre a DLL vbsendmail.
            Set vSendMail = New clsSendMail
            vSendMail.Attachment = DataDirPath & "\descasque\descasque.cab"
            vSendMail.EmailAddressValidation = VALIDATE_SYNTAX
            vSendMail.UseAuthentication = True
            vSendMail.From = "marcos@itcase.com.br"
            vSendMail.Recipient = IIf(Email = "", "marcos@itcase.com.br", Email)
            vSendMail.Subject = "BD do Descasque"
            vSendMail.SMTPHost = "pop3.itcase.com.br"
            vSendMail.Username = "marcos@itcase.com.br"
            vSendMail.Password = "181073"
            vSendMail.FromDisplayName = "Accessory Plastic"
            If vSendMail.IsValidEmailAddress(Email) Then
                lstStatus.Clear
                vSendMail.Send
            Else
                MsgBox "Não foi possível enviar o E-Mail. Favor verificar corretamente o endereço de e-mail digitado!", vbExclamation
            End If
            vSendMail.Shutdown
            Set vSendMail = Nothing
        End If
    End Select
    
    
    
End Sub

Private Sub Drive1_Change()

On Error GoTo TrataErro
    'Muda o DirListBox
    Dir1.Path = Drive1.Drive
    Exit Sub

TrataErro:
Select Case Err.Number
Case 68
    MsgBox "O disco não está disponível! Insira um disco e tente novamente.", vbCritical
    Drive1.Drive = "c:"
End Select

End Sub

Private Sub Form_Load()
Dim qtd As String
Dim qtd2 As String
    
    Set fs = New FileSystemObject
    Drive1.Drive = "c:"
    'Atribui os valores do registro aos labels
    txtEndereco = GetSetting("Descasque", "BaseDados", "PathMySQL")
    qtd = GetSetting("Descasque", "Estoque", "Show Records")
    qtd2 = GetSetting("Descasque", "Fechamento", "Show Records")
    'QUANTIDADE DE REGISTROS DO ESTOQUE
    If qtd = "999999" Then
        txtQTD = ""
        optTodos.Value = True
    Else
        txtQTD = qtd
    End If
    
    'QUANTIDADE DE REGISTROS DO FECHAMENTO
    If qtd2 = "999999" Then
        txtQTD2 = ""
        optTodos2.Value = True
    Else
        txtQTD2 = qtd2
    End If
    lblProgress = ""
    lstStatus.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Set fs = Nothing
    
End Sub

Private Sub optTodos_Click()
    
    txtQTD = ""

End Sub

Private Sub optTodos2_Click()

    txtQTD2 = ""

End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)

    KeyAscii = ValidaIP(txtEndereco, KeyAscii)

End Sub

Private Sub txtQTD_GotFocus()

    optTodos.Value = False
    
End Sub

Private Sub txtQTD_KeyPress(KeyAscii As Integer)

    KeyAscii = OnlyNumbers(txtQTD, KeyAscii)

End Sub


Private Sub txtQTD2_GotFocus()

    optTodos2.Value = False
    
End Sub

Private Sub txtQTD2_KeyPress(KeyAscii As Integer)
    
    KeyAscii = OnlyNumbers(txtQTD2, KeyAscii)

End Sub

Private Sub lstStatus_Click()

lstStatus.ListIndex = -1

End Sub


Public Function ValidaIP(Controle As TextBox, Tecla As Integer) As Integer
'Verfifica se o valor digitado é somente numeros ou ponto
'Usar no evento KeyPress

    'testa as teclas. Somente aceita numerico e backspace
    Select Case Tecla
        Case 46   'Ponto .
        Case 8    'backspace
        Case 48 To 57 'numeros
        Case Else
            ValidaIP = 0  'nada
            Exit Function
    End Select
        ValidaIP = Tecla
    
End Function


Private Sub CompactBD(Optional sPath As String)
'A compactação por makecab, precisa de dois passos:
'-Escrever o caminho completo para os arquivos no arquivo .DDF
'-Executar o comando: makecab /F arquivo.ddf
Dim fs As New FileSystemObject
Dim Arquivo As TextStream
Dim sFiles() As String
Dim i As Long

    

    Me.MousePointer = vbHourglass
    'PEGA OS ARQUIVOS DO DIRETÓRIO DO DESCASQUE
    sFiles = AllFiles(DataDirPath & "\descasque")
    
    'CRIA UM ARQUIVO UDF E SOBRESCREVE O EXISTENTE
    Set Arquivo = fs.CreateTextFile(DataDirPath & "\descasque\descasque.ddf", True)
    
    'SETA AS PROPRIEDADES DO ARQUIVO
    With Arquivo
        .WriteLine (".Set Cabinet=ON")
        .WriteLine (".Set Compress=ON")
        .WriteLine (".Set CompressionType=LZX")
        .WriteLine (".Set CompressionMemory=21")
        .WriteLine (".Set CabinetNameTemplate=DESCASQUE.CAB")
        .WriteLine (".Set DiskDirectoryTemplate=" & DataDirPath & "\descasque")
        .WriteBlankLines 1
        For i = 0 To UBound(sFiles)
            If sFiles(i) <> "descasque.ddf" Then
                .WriteLine (DataDirPath & "\descasque\" & sFiles(i))
            End If
        Next
        .Close
    End With
    'GRAVA O ARQUIVO EFETIVAMENTE NA PASTA DO BD DESCASQUE
    ShellAndWait "makecab /F " & DataDirPath & "\descasque\descasque.ddf", vbHide
    If sPath <> "" Then
        fs.MoveFile DataDirPath & "\descasque\descasque.cab", sPath
    End If
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ShellAndWait(ByVal program_name As String, ByVal window_style As VbAppWinStyle)
Dim process_id As Long
Dim process_handle As Long

    ' Inicie o programa.
    On Error GoTo ShellError
    process_id = Shell(program_name, window_style)
    On Error GoTo 0

    DoEvents

    ' Aguarde o programa terminar.
    ' pegue o Handle do Processo.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Sub

ShellError:
    MsgBox Err.Description, vbExclamation, "Erro"
    
End Sub

Private Sub vSendMail_Progress(lPercentCompete As Long)

lblProgress = lPercentCompete & "% concluídos"

End Sub

Private Sub vSendMail_SendFailed(Explanation As String)

MsgBox ("Sua tentativa de enviar a mensagem falhou pelo seguinte motivo: " & vbCrLf & Explanation)
lblProgress = ""

End Sub

Private Sub vSendMail_SendSuccesful()

MsgBox "Mensagem enviada com sucesso!"
lblProgress = ""

End Sub

Private Sub vSendMail_Status(Status As String)

lstStatus.AddItem Status
lstStatus.ListIndex = lstStatus.ListCount - 1
lstStatus.ListIndex = -1

End Sub

