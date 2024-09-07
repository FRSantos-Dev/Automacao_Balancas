Attribute VB_Name = "Module1"
Option Explicit
'--------------------------------------------------------------
'Declara��o das fun��es da balanca (tem que ser no m�dulo)
Declare Function AttachScale Lib "c:\windows\system\PcScale.dll" (ByRef Handle As Long, ByVal COM As Long, ByVal BaudRate As Long, ByVal ByteSize As Long, ByVal StopBits As Long, ByVal Parity As Long) As Long
Declare Function ReadScale Lib "c:\windows\system\PcScale.dll" (ByRef Handle As Long, ByVal Indicator As Long, ByVal Gross As String, ByVal Tare As String, ByVal Net As String, ByVal Unitd As String) As Long
Declare Sub DettachScale Lib "c:\windows\system\PcScale.dll" (ByRef Handle As Long)
'--------------------------------------------------------------

Public Resultado As ADODB.Recordset
Public Resposta As Integer
Public pDataIni As Date
Public pDataFim As Date
Public DataIni As Date 'Formularios de produ��o
Public DataFim As Date 'Formularios de produ��o
Public Titulo As String, Mensagem As String, Caixa As String, NomeDescascador As String, DBPath As String
Public con As ADODB.Connection 'adicionado para o MYsql++++++++++++++++++++++++++++++++++++++++++++
Public RS As ADODB.Recordset
Public RS2 As ADODB.Recordset
Public RSReferencias As ADODB.Recordset
Public CRITERIO As String 'Uso para criterios de consultas
Public DataDirPath As String

'Variaveis do ImprFicha
Public Maq As String, Molde As String, Cav As String, Mat As String, ID_MP As Integer


Public Sub AbrirConexao() 'adicionado para o MYsql++++++++++++++++++++++++++++++++++++++++++++
Dim con_str As String
Dim BD As String, Server As String, User As String, Pass As String
Dim iErr As Byte
Dim RSBD As ADODB.Recordset
Dim xx As Byte

On Error GoTo ErroHandler

    Set con = New ADODB.Connection
    

1   BD = "descasque"
    Server = GetSetting("Descasque", "BaseDados", "PathMySQL")
    User = "root"
    Pass = ""
    'Para mudar a senha no root no MySql
    'SET PASSWORD FOR root@localhost=PASSWORD('new_password');
    
    con_str = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Server & ";DATABASE=" & BD & ";UID=" & User & ";PWD=" & Pass & ";OPTION=35"
    
    With con
        .ConnectionString = con_str
        .Open con_str
    End With
    
    Set RSReferencias = New ADODB.Recordset
    RSReferencias.CursorLocation = adUseClient
    RSReferencias.CursorType = adOpenKeyset
    
    'Como a tabela de referencias � usada o tempo todo, mantenho um recordset aberto
    RSReferencias.Open "SELECT * FROM tbl_precos ORDER BY referencia", con
    
    'Captura o caminho do mysql, consultando diretamente a base de dados.
    Set RSBD = con.Execute("SHOW VARIABLES")
    
    While Not DataDirPath <> ""
        If RSBD!variable_name = "datadir" Then
            DataDirPath = RSBD!Value
        End If
        RSBD.MoveNext
    Wend
    
    RSBD.Close
    Set RSBD = Nothing
    
Exit Sub
ErroHandler:
    'Se chegar a 3 erros, orienta a entrar em contato com o suptec
    iErr = iErr + 1
    If iErr = 3 Then
        Unload formSplash
        MsgBox "N�o foi poss�vel localizar o Banco de dados. Entre em contato com o Suporte T�cnico.", vbCritical
        End
    End If
    If Err = 3024 Or Err = 3044 Or Err = -2147467259 Then
        Unload formSplash
        FormBD.Show 1
        Server = GetSetting("Descasque", "BaseDados", "PathMySQL")
        GoTo 1
    Else
        Unload formSplash
        MsgBox Err.Number & ": " & Err.Description
        End
    End If
    
End Sub

' Sub para apresentar mensagens de erro para o Visual ReportX
Public Sub Rpx_MsgErro(Numero As Long)

    Dim MSG$
    
    If Numero < 0 Then
    
        ' Mensagens de erro previstas
        Select Case Numero - vbObjectError
            Case 1001: MSG = "� necess�rio existir uma impressora instalada no Windows"
            Case 1002: MSG = "N�o h� registros a imprimir"
            Case 1003, 1004, 1005: MSG = "Erro na configura��o interna do relat�rio"
            Case 1006: MSG = "A p�gina configurada para o relat�rio n�o possu� espa�o suficiente para a impress�o"
            Case 1007: MSG = "J� existe um relat�rio em andamento"
        End Select
        
        MsgBox MSG, vbInformation, "Impress�o"
        
    Else
        
        ' Mensagens n�o previstas. Isso pode significar um erro
        ' interno no ReportX. Se isso acontecer, por favor reporte isso
        ' atrav�s de e-mail para ser corrigido.
        If Numero = 401 Then
            MsgBox "N�o chame o relat�rio a partir de um formul�rio MODAL", vbInformation, "Erro de programa��o"
            Exit Sub
        End If
        MsgBox "Erro n�o previsto:" & Numero & vbCrLf & Error(Numero) & _
            IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical, "Impress�o"
        
    End If
    
End Sub


