VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form formGrafico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gráfico 1"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "formGrafico.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "formGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    RS.Source = "SELECT DATA, SUM(MP_CONSUMIDA) AS MP FROM TBL_DADOS GROUP BY DATA"
    Set RS.ActiveConnection = cn
    RS.Open
    With MSChart1
      '.ShowLegend = True
      Set .DataSource = RS
   End With

End Sub
