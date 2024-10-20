VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form formFichas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de controle de produção"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "<<< Cancelar"
      Height          =   315
      Left            =   60
      TabIndex        =   83
      Top             =   5820
      Width           =   2415
   End
   Begin VB.CommandButton cmdAvancar 
      Caption         =   "Avançar >>>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   82
      Top             =   5820
      Width           =   2415
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Temperatura das zonas"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   5220
      TabIndex        =   75
      Top             =   3660
      Width           =   2295
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   33
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   38
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   32
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   37
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   31
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   36
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   30
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   35
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   29
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   34
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   28
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   33
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   81
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   80
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   79
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   78
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   77
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bico"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   76
         Top             =   2100
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Pressão"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   60
      TabIndex        =   64
      Top             =   3660
      Width           =   4995
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   27
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   32
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   26
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   31
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   25
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   30
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   29
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   23
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   28
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   27
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   21
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   26
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   25
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   24
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   18
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   23
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contra Pressão"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   74
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extrator Recuo"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   73
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extrator Avanço"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   72
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amort. Fechamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   71
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prot. do Molde"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   70
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecham. Rápido"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecham. Lento"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   68
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amort. Abertura"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   67
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abert. Rápida"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   66
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aber. Lenta"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Pressão de Injeção"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   5220
      TabIndex        =   57
      Top             =   1080
      Width           =   2295
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   22
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   21
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   20
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   18
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   17
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   63
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   62
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   61
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   60
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   59
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   58
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Velocidades"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   60
      TabIndex        =   43
      Top             =   1080
      Width           =   4995
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   16
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   12
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   11
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   10
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   6
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   5
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Injeção"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   56
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dosagem"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   55
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortecimento"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   54
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechamento 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   53
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prot. do Molde"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   52
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recalque"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2460
         TabIndex        =   51
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechamento 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extrator Recuo"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extrator Avanço"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amort. Abertura"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abert. Rápida"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aber. Lenta"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      TabIndex        =   39
      Top             =   60
      Width           =   7455
      Begin VB.CommandButton cmdImprimir 
         Height          =   555
         Left            =   6540
         Picture         =   "formFichas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Ficha"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdLocalizar 
         Height          =   555
         Left            =   5760
         Picture         =   "formFichas.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Localizar Ficha"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtCavidades 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   98
         Top             =   180
         Width           =   495
      End
      Begin VB.ComboBox cboMoldes 
         Height          =   315
         ItemData        =   "formFichas.frx":314C
         Left            =   2520
         List            =   "formFichas.frx":314E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1515
      End
      Begin VB.ComboBox cboMP 
         Height          =   315
         ItemData        =   "formFichas.frx":3150
         Left            =   1260
         List            =   "formFichas.frx":3152
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   4275
      End
      Begin VB.TextBox txtMaquina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "01"
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Matéria Prima"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   44
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Máquina"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cavidades"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4140
         TabIndex        =   41
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Molde"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   40
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   60
      TabIndex        =   84
      Top             =   1080
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton cmdObs 
         BackColor       =   &H0080C0FF&
         Caption         =   "Observações"
         Height          =   315
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   4260
         Width           =   2415
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         Caption         =   "Observações"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         TabIndex        =   114
         Top             =   2760
         Width           =   4335
         Begin VB.TextBox txtObs 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   107
            Top             =   660
            Width           =   4095
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   1320
            TabIndex        =   106
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label Label45 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Montador"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   115
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         Caption         =   "Pressão de Recalque"
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   2520
         TabIndex        =   99
         Top             =   180
         Width           =   1815
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   840
            MaxLength       =   7
            TabIndex        =   105
            Top             =   2100
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   840
            MaxLength       =   7
            TabIndex        =   104
            Top             =   1740
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   840
            MaxLength       =   7
            TabIndex        =   103
            Top             =   1380
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   840
            MaxLength       =   7
            TabIndex        =   102
            Top             =   1020
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   840
            MaxLength       =   7
            TabIndex        =   101
            Top             =   660
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   840
            MaxLength       =   7
            TabIndex        =   100
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "6"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   113
            Top             =   2100
            Width           =   735
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   112
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   111
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   110
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   109
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   108
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         Caption         =   "Tempos"
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   120
         TabIndex        =   87
         Top             =   180
         Width           =   2115
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   97
            Top             =   1740
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   96
            Top             =   1380
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   95
            Top             =   1020
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   94
            Top             =   660
            Width           =   675
         End
         Begin VB.TextBox txtCampo2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   88
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label51 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pausa"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   93
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label50 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Injeção"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   92
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label49 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Resfriamento"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   91
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label48 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total de ciclo"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   90
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label47 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Alarme de ciclo"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   89
            Top             =   1740
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Situação"
         ForeColor       =   &H80000008&
         Height          =   3555
         Left            =   4560
         TabIndex        =   85
         Top             =   180
         Width           =   2775
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Contador zerado"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   125
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Peças perfeitas"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   124
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Regulagem da máquina OK"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   123
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Máquina com defeito"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   122
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Molde gelando bem"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   121
            Top             =   1680
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Molde com defeito"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   120
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Molde está com proteção"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   119
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Molde c\ água dos 02 lados"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   118
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Bucha de injeção OK"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   2415
         End
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
         Left            =   5880
         TabIndex        =   117
         Top             =   3840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label46 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         TabIndex        =   116
         Top             =   3840
         Width           =   1230
      End
   End
End
Attribute VB_Name = "formFichas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ficha As String
Dim Passo As Integer
Dim Campo(12) As Double
Dim Campo2(10) As Double
Dim dData As Date
'Para manter a janela modal sem ser modal
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Dim BtRel As Boolean

Private Sub cboMoldes_Click()
    
    If cboMoldes.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenForwardOnly
        RS.Open "SELECT cavidades FROM TBL_PRECOS WHERE referencia = '" & cboMoldes.Text & "'", con
        txtCavidades = RS!Cavidades & ""
        RS.Close
    End If
    
End Sub

Private Sub cmdAvancar_Click()
Dim f As Integer
    
    Select Case Passo
    'Primeira tela de preenchimento
    Case 0
        Frame6.Visible = True
        Frame5.Visible = False
        Frame4.Visible = False
        Frame3.Visible = False
        Frame2.Visible = False
        SaveStatus (1)
        Passo = 1
        txtCampo2(0).SetFocus
    Case 1
        'Segunda tela de preenchimento
        SaveStatus (2)
        Unload Me
        Passo = 0
    End Select

End Sub

Private Sub cmdCancelar_Click()

    Select Case Passo
    Case 0
        Unload Me
    Case 1
        Frame6.Visible = False
        Frame5.Visible = True
        Frame4.Visible = True
        Frame3.Visible = True
        Frame2.Visible = True
        Passo = 0
    Case 2
    End Select
    
End Sub

Private Sub cmdImprimir_Click()
                
    If txtMaquina = "" Or cboMoldes.Text = "" Or cboMP.Text = "" Then
        MsgBox "Preencha os campos Máquina, Molde e Matéria-Prima antes de prosseguir!", vbCritical
        Exit Sub
    End If
    BtRel = True
    Maq = txtMaquina
    Cav = txtCavidades & ""
    Molde = cboMoldes.Text
    Mat = Right(cboMP.Text, Len(cboMP.Text) - 6)
    ID_MP = Left(cboMP.Text, 3)
    Passo = 0
    Me.Hide
    formImprFicha.Config
    BtRel = False
    Me.Show
    

End Sub

Private Sub cmdLocalizar_Click()
    
    'Verifica se todos os campos foram preenchidos
    If cboMoldes = "" Then
        MsgBox "Selecione um Molde!", vbCritical
        Exit Sub
    End If
    If txtMaquina = "" Then
        MsgBox "Preencha o campo Máquina!", vbCritical
        Exit Sub
    End If
    If cboMP = "" Then
        MsgBox "Selecione uma Matéria Prima!", vbCritical
        Exit Sub
    End If

    'Preenche os dados na tela
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT * FROM TBL_FICHAS WHERE maquina = " & txtMaquina & " AND molde = '" & cboMoldes.Text & "' AND MP = " & CLng(Left(cboMP, 3)) & "", con
    'Verifica se já existe esta ficha cadastrada
    If RS.EOF Then
        If MsgBox("Esta ficha ainda não foi criada. Deseja criá-la agora?", vbYesNo + vbQuestion) = vbYes Then
            Frame2.Enabled = True
            Frame3.Enabled = True
            Frame4.Enabled = True
            Frame5.Enabled = True
            Frame6.Enabled = True
            cmdAvancar.Enabled = True
            ClearAll
            con.Execute "INSERT INTO TBL_FICHAS (MAQUINA, MOLDE, MP) VALUES('" & txtMaquina & "', '" & cboMoldes.Text & "', '" & CLng(Left(cboMP, 3)) & "')"
        End If
    Else
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Frame5.Enabled = True
        Frame6.Enabled = True
        cmdAvancar.Enabled = True
        GetStatus
    End If
    RS.Close
    
End Sub

Private Sub cmdObs_Click()
    
    Me.Tag = txtMaquina & "|" & cboMoldes.Text & "|" & Left(cboMP, 3)
    Me.Hide
    formObs.Show 1
    Me.Show
    
End Sub

Private Sub Form_Load()
    
    'Para manter a janela modal sem ser modal
    'Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    'Alimenta cx combo
    Carrega_MP
    
    'Alimenta referencias
    PreencheMoldes
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If BtRel = False Then
        Me.Hide
        formProducao.Show
    End If

End Sub

Private Sub txtCampo_GotFocus(Index As Integer)
    
    SendKeys "{Home}+{End}"
    
End Sub

Private Sub txtCampo2_GotFocus(Index As Integer)
    
    SendKeys "{Home}+{End}"
    
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtCampo(Index), KeyAscii)

End Sub

Private Sub txtCampo2_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 11 Then Exit Sub
    KeyAscii = TypeCurrency(txtCampo2(Index), KeyAscii)
    
End Sub

Private Sub Carrega_MP()
    
    'Carrega os dados do combo com as matérias-primas
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT ID, NOME FROM TBL_MP", con
    If RS.EOF = False Then
        While Not RS.EOF
            cboMP.AddItem Format(RS!ID, "000") & " - " & RS!nome
            RS.MoveNext
        Wend
    End If
    RS.Close

End Sub

Private Sub SaveStatus(ByRef CampoTexto As Integer)
Dim i As Integer
Dim k As String
Dim k2 As String

    Select Case CampoTexto
    Case 1
        'Faco um loop nos textbox, pego seus valores e gravo em um array. Depois gravo no registro
        For i = 0 To 33
            k = k & txtCampo(i).Text & "|"
            DoEvents
        Next
        'Atualiza os dados no BD
        con.Execute "UPDATE TBL_FICHAS SET CAMPOS1 = '" & k & "' WHERE maquina = " & txtMaquina & " AND molde = '" & cboMoldes.Text & "' AND MP = " & CLng(Left(cboMP, 3)) & ""
    Case 2
        'Faco um loop nos textbox, pego seus valores e gravo em um array. Depois gravo no registro
        For i = 0 To 11
            k = k & txtCampo2(i).Text & "|"
            DoEvents
        Next
        
        'Atualiza os Checkbox do "Status"
        For i = 0 To Check1.UBound
            k2 = CStr(k2 & CInt(Check1(i).Value) & "|")
            DoEvents
        Next
            
        'Atualiza os dados no BD
        If IsDate(txtData) Then dData = txtData
        If IsDate(txtData) Then
            con.Execute "UPDATE TBL_FICHAS SET CAMPOS2 = '" & k & "', CHECKS = '" & k2 & "', OBS = '" & txtObs & "" & "', DATA = '" & Format(dData, "yyyy/mm/dd") & "' WHERE maquina = " & txtMaquina & " AND molde = '" & cboMoldes.Text & "' AND MP = " & CLng(Left(cboMP, 3)) & ""
        Else
            con.Execute "UPDATE TBL_FICHAS SET CAMPOS2 = '" & k & "', CHECKS = '" & k2 & "', OBS = '" & txtObs & "" & "' WHERE maquina = " & txtMaquina & " AND molde = '" & cboMoldes.Text & "' AND MP = " & CLng(Left(cboMP, 3)) & ""
        End If
        con.Execute "INSERT INTO TBL_OBS (OBS, DATA, MONTADOR, MAQUINA, REFERENCIA, MP) values('" & txtObs & "', '" & Format(txtData, "yyyy/mm/dd") & "', '" & txtCampo2(11) & "', " & txtMaquina & ", '" & cboMoldes.Text & "', " & CLng(Left(cboMP, 3)) & ")"
    End Select
    
End Sub

Private Sub GetStatus()
Dim X() As String
Dim val As String

    'Estou reutilizando parte do código anterior.
    'Para evitar refazer a janela, continuarei dividindo os textboxes matrizes em dois arrays.
    
    'recupero os dados do BD e exibo na inicializacao do form (primeira tela)
    val = RS!campos1 & ""
    If val <> "" Then
        X = Split(val, "|")
        For i = 0 To 33
            txtCampo(i) = X(i)
            DoEvents
        Next
    End If

    'recupero os dados do BD e exibo na inicializacao do form (Segunda tela)
    val = RS!campos2 & ""
    If val <> "" Then
        X = Split(val, "|")
        For i = 0 To 11
            txtCampo2(i) = X(i)
            DoEvents
        Next
    End If
    
    'recupero os dados do BD e atualizo os Checkbox do "Status"
    val = RS!checks & ""
    If val <> "" Then
        X = Split(val, "|")
        For i = 0 To Check1.UBound
            Check1(i) = CInt(X(i))
            DoEvents
        Next
    End If
    txtObs = RS!obs & ""
    If Not IsNull(RS!Data) Then txtData = RS!Data
    
End Sub

Private Sub PreencheMoldes()

    'Carrega o RS de moldes
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT referencia FROM TBL_PRECOS ORDER BY REFERENCIA", con
    While Not RS.EOF
        cboMoldes.AddItem RS!REFERENCIA
        RS.MoveNext
    Wend
    RS.Close
    
End Sub

Private Sub ClearAll()
Dim X As Object

    For Each X In formFichas
        If TypeOf X Is TextBox And X.Name <> "txtMaquina" And X.Name <> "txtCavidades" Then
            
            X.Text = ""
        End If
        If TypeOf X Is CheckBox Then
            X.Value = 0
        End If
    Next
    
    txtData.Text = "__/__/____"
    txtObs.Text = ""

End Sub
