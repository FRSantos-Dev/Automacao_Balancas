VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formEstoque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "formEstoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   795
      Left            =   4980
      Picture         =   "formEstoque.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Incluir Matéria Prima no estoque"
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   795
      Left            =   4980
      Picture         =   "formEstoque.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Para excluir Matéria Prima, selecione um ítem no Grid"
      Top             =   1020
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   795
      Left            =   4980
      Picture         =   "formEstoque.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sair"
      Top             =   1860
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Cadastro de Matéria Prima"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cboMP 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtPeso 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   60
         TabIndex        =   1
         Top             =   780
         Width           =   1035
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
         Left            =   3420
         TabIndex        =   2
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Matéria Prima"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1050
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1875
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3307
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      HighLight       =   2
      Appearance      =   0
      FormatString    =   "|Matéria Prima                                     |Peso          |Data             "
   End
End
Attribute VB_Name = "formEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I've created a variable to inform the earlyer form if the table was updated or not
Dim Alterou As Boolean
Dim Limite As String

Private Sub cmdCadastrar_Click()

    'INSERE UMA NOVA MP NO ESTOQUE.
    If cboMP.ListIndex <> -1 And txtPeso <> "" And txtData <> "__/__/____" Then
        con.Execute "INSERT INTO TBL_ENTRADA_MP (COD_MP, peso, data) SELECT ID, '" & FormatFloatForDB(txtPeso) & "', '" & Format(txtData, "YYYY/MM/DD") & "' FROM TBL_MP WHERE TBL_MP.NOME =  '" & cboMP.Text & "'"
        con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE + '" & FormatFloatForDB(txtPeso) & "') WHERE NOME = '" & cboMP.Text & "'"
        Grid1.Rows = 1
        PreencheGrid
        'Atualiza a janela de Cadastro de Matérias-Prima
        formMP.Grid1.Rows = 1
        Alterou = True
        txtPeso = ""
    Else
        MsgBox "É necessário preencher todos os campos!", vbExclamation
        Exit Sub
    End If
    cboMP.ListIndex = -1
    cboMP.SetFocus

End Sub

Private Sub cmdExcluir_Click()

    If Grid1.Text <> "" And Grid1.Row <> 0 Then
        If MsgBox("Você tem certeza de que deseja excluir esta entrada no estoque?", vbExclamation + vbYesNo) = vbYes Then
            con.Execute "DELETE FROM TBL_ENTRADA_MP WHERE ID = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
            con.Execute "UPDATE TBL_MP SET ESTOQUE = (ESTOQUE - '" & FormatFloatForDB(Grid1.TextMatrix(Grid1.Row, 2)) & "') WHERE NOME = '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
            If Grid1.Rows = 2 Then
                Grid1.Rows = 1
            Else
                Grid1.RemoveItem Grid1.RowSel
            End If
            'Atualiza a janela de Cadastro de Matérias-Prima
            formMP.Grid1.Rows = 1
            Alterou = True
        End If
    Else
        MsgBox "Selecione uma referência válida!"
    End If

End Sub

Private Sub cmdSair_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Alterou = False
    Limite = IIf(GetSetting("Descasque", "Estoque", "Show Records") = "", 100, GetSetting("Descasque", "Estoque", "Show Records"))
    Me.MousePointer = vbHourglass
    txtData = Date
    Set RS = con.Execute("SELECT NOME FROM TBL_MP")
    If RS.EOF = False Then
        While Not RS.EOF
            cboMP.AddItem RS!nome
            RS.MoveNext
        Wend
    End If
    RS.Close
    PreencheGrid
    Me.MousePointer = vbDefault
    
End Sub

Private Sub PreencheGrid()

    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.Open "SELECT TBL_ENTRADA_MP.ID, NOME, peso, DATA FROM TBL_MP INNER JOIN TBL_ENTRADA_MP ON TBL_MP.ID=TBL_ENTRADA_MP.COD_MP ORDER BY DATA DESC LIMIT 0, " & Limite & "", con
    If RS.EOF = False Then
        While Not RS.EOF
            Grid1.AddItem RS!ID & vbTab & RS!nome & vbTab & RS!Peso & vbTab & RS!Data
            RS.MoveNext
        Wend
    End If
    RS.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Alterou = True Then
        'Para fazer uma chamada a uma sub ou function em outro formulario, ela deve ser do tipo Public
        formMP.PreencheGrid
    End If
    
End Sub

Private Sub Grid1_Click()
    
    If Grid1.Text = "" Then Exit Sub
    Grid1.CellBackColor = vbBlue
    Grid1.CellForeColor = vbWhite

End Sub

Private Sub Grid1_LeaveCell()

    
    Grid1.CellBackColor = &H80000005
    Grid1.CellForeColor = &H80000012
    
End Sub

Private Sub txtData_LostFocus()

    CheckIfIsDate txtData

End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)

    KeyAscii = TypeCurrency(txtPeso, KeyAscii)
    
End Sub
