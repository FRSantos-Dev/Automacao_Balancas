VERSION 5.00
Begin VB.Form formPortas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Portas de Comunicação"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   915
      Left            =   2820
      Picture         =   "formPortas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   300
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Portas"
      Height          =   1095
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2235
      Begin VB.OptionButton Option1 
         Caption         =   "COM2"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "COM1"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "formPortas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If Option1(0).Value = True Then
        SaveSetting "Descasque", "Portas", "Configuracao", "1"
    Else
        SaveSetting "Descasque", "Portas", "Configuracao", "2"
    End If
    Unload Me

End Sub

Private Sub Form_Load()

    If GetSetting("Descasque", "Portas", "Configuracao", "2") = "1" Then
        Option1(0).Value = True
    Else
        Option1(1).Value = True
    End If
    
End Sub
