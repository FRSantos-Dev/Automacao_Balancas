VERSION 5.00
Begin VB.Form formSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "formSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label lblCopyright 
         Caption         =   "Todos os direitos reservados de acordo com a lei"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   3060
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWarning 
         Caption         =   "Atenção: Proibido a cópia sem a devida autorização do Proprietário"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6045
         TabIndex        =   4
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Microsoft Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3975
         TabIndex        =   5
         Top             =   2340
         Width           =   2880
      End
      Begin VB.Label lblProductName 
         Caption         =   "Accessory Plastic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   2880
         TabIndex        =   6
         Top             =   660
         Width           =   3585
      End
      Begin VB.Image imgLogo 
         Height          =   2460
         Left            =   90
         Picture         =   "formSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licenciado para Accessory Plastic LTDA"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "formSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para manter a janela modal sem ser modal
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Sub Form_Activate()
Dim interv As Date

    Me.MousePointer = vbHourglass
    Me.Refresh
    'Deixa a janela sendo exibida um pouco
'    interv = DateAdd("s", 2, Time)
'    While Not Time > interv
'        DoEvents
'    Wend
    
    Me.MousePointer = vbDefault
    Load MDIForm1
    Unload Me

End Sub

Private Sub Form_Initialize()

    'evita que o aplicativo seja aberto em mais de uma instancia
    If App.PrevInstance Then
        MsgBox ("A aplicação que você está iniciando já está em uso."), vbExclamation, "Aplicação em uso"
        End
    End If

End Sub

Private Sub Form_Load()

    lblVersion = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub
