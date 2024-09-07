VERSION 5.00
Begin VB.Form FormBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caminho Para o Banco de Dados"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "FormBD.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5490
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdProsseguir 
      Caption         =   "&Prosseguir"
      Height          =   315
      Left            =   4050
      TabIndex        =   2
      Top             =   1830
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Estação"
      Height          =   1635
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   5310
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2000
         TabIndex        =   5
         Top             =   1275
         Width           =   3210
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Servidor"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   660
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endereço IP ou Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   375
         TabIndex        =   6
         Top             =   1275
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   420
         TabIndex        =   1
         Top             =   1020
         Width           =   2625
      End
   End
End
Attribute VB_Name = "FormBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProsseguir_Click()

    If Label1 = "" And Option1(0).Value = True Then
        MsgBox "Selecione o caminho para o BD antes de prosseguir!"
        Exit Sub
    End If
    
    If Option1(0).Value = True Then
        SaveSetting "Descasque", "BaseDados", "PathMySQL", txtEndereco
    Else
        SaveSetting "Descasque", "BaseDados", "PathMySQL", "127.0.0.1"
    End If
    Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode <> 1 Then End

End Sub

Private Sub Option1_Click(Index As Integer)

    Select Case Index
    Case 0
        txtEndereco = ""
        txtEndereco.Locked = False
    Case 1
        txtEndereco = "127.0.0.1"
        txtEndereco.Locked = True
    End Select

End Sub
