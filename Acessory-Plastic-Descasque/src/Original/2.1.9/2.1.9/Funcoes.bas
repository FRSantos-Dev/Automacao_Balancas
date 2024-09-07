Attribute VB_Name = "Funcoes"
Option Explicit

' variáveis utilizadas pela função ValExtenso
Dim AUnidades As Variant
Dim APluridades As Variant
Dim ACentenas As Variant
Dim ADezenas As Variant
Dim ALhoes As Variant
Dim ALhao As Variant
Dim Digitos As Integer


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX INÍCIO DO CÓDIGO DE EXTENSO XXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'Use da seguinte maneira:
'Text1.text = ValExtenso(Text2.text)

Function ValExtenso(Numero As String) As String
'* parametros - entrada - numero
'*            - saida   - por extenso

Static N, ND, NC, ind, Val_Txt, Comprimento


AUnidades = Array("zero", "um", "dois", "tres", "quatro", "cinco", "seis", "sete", "oito", "nove")
APluridades = Array("dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove")
ACentenas = Array("duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")
ADezenas = Array("vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa")
ALhoes = Array(" quatrilhoes", " trilhoes", " bilhoes", " milhoes", " mil", "")
ALhao = Array(" quatrilhao", " trilhao", " bilhao", " milhao", " mil", "")

Digitos = 21


Comprimento = Len(Numero)

Val_Txt = ""

 Select Case InStr(1, Numero, ",")
    Case 0
        Numero = Numero & ",00"
    Case (Comprimento - 1)
        Numero = Numero & "0"
    Case Else
        Numero = Numero
End Select

If Numero = "0,00" Then
    ValExtenso = "zero"
End If

'Numero = CStr(P_NUM)
Numero = Space(21 - Len(Numero)) & Numero

If Numero = "*********************" Then
    MsgBox "Valor excede 8 dígitos"
    ValExtenso = Numero
End If
For ind = 0 To 5
    NC = Mid(Numero, (ind) * 3 + 1, 3)
    N = val(Mid(NC, 3, 1))
    ND = val(Mid(NC, 2, 1))
    NC = val(Mid(NC, 1, 1))
    
    If N + ND + NC > 0 Then
       If Len(Val_Txt) > 0 Then
          Val_Txt = Val_Txt + " e "
       End If
       Val_Txt = Val_Txt + Centena(NC, ND, N) + _
                  IIf(N = 1 And ND + NC = 0, ALhao(ind), ALhoes(ind))
    End If
Next ind
    If Val_Txt <> "" Then
        If UCase(Val_Txt) = "UM" Then
            Val_Txt = Val_Txt & " REAL"
        Else
            Val_Txt = Val_Txt & " REAIS"
        End If
    End If
    
    NC = Mid(Numero, 19, 3)
    N = val(Mid(NC, 3, 1))
    ND = val(Mid(NC, 2, 1))
    NC = 0
    If N + ND + NC > 0 Then
       If Len(Val_Txt) > 0 Then
          Val_Txt = Val_Txt + " E "
       End If
       Val_Txt = Val_Txt + Centena(NC, ND, N) + " CENTAVOS"
    End If


ValExtenso = UCase(Val_Txt)

End Function

Function Centena(NC, ND, N)
ACentenas = Array("duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")

    If NC = 1 Then
        If ND + N = 0 Then
            Centena = "cem"
        Else
            Centena = "cento e " + Dezena(ND, N)
        End If
        Exit Function
    End If
    
    If NC = 0 Then
        Centena = Dezena(ND, N)
        Exit Function
    End If
        
    If ND + N <> 0 Then
        Centena = ACentenas(NC - 2) + " e " + Dezena(ND, N)
    Else
        Centena = ACentenas(NC - 2)
    End If
    
    
End Function
    
Static Function Dezena(ND, N)
    
  If ND = 0 Then
    Dezena = Unidade(N)
    Exit Function
  End If
    
  If ND = 1 Then
    Dezena = APluridades(N)
    Exit Function
  End If
    
  If N = 0 Then
    Dezena = ADezenas(ND - 2)
    Exit Function
  End If
  
  Dezena = ADezenas(ND - 2) + " e " + Unidade(N)

End Function

Function Unidade(N) As String
'* subrotina de unidades
    Unidade = AUnidades(N)
End Function
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX FIM DO CÓDIGO DE EXTENSO XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



Public Sub CheckIfIsDate(Controle As Object)
'Funcao que verifica se o valor passado é uma data válida ou não
'Usar no evento KeyPress

    'Verifica se a data é válida, e se o usuario digitou corretamente o campo
    If Not IsDate(Controle.Text) Or InStr(1, Controle.Text, "_") Then
        MsgBox "Data Inválida! Digite novamente"
        With Controle 'controle.text
            .SetFocus
            .SelStart = .SelLength
            Exit Sub
        End With
    End If
    
    'Verifica se o digitador colocou uma data anterior a data atual.
    If CDate(Controle.Text) < Date Then
        Resposta = MsgBox("A data de saida é anterior a data atual! Tem certeza de que deseja continuar?", vbYesNo)
        If Resposta = vbNo Then
            With Controle 'controle.text
                .SetFocus
                .SelStart = .SelLength
                Exit Sub
            End With
        End If
    End If

End Sub

Public Function TypeCurrency(Controle As TextBox, Tecla As Integer) As Integer
'Verfifica se o valor digitado está em formato moeda (0.00)
'Usar no evento KeyPress
Dim a As Boolean
    
    'Verifica se já existe um ponto decimal no texto
    If Tecla = 44 Or Tecla = 46 Then
        a = InStr(1, Controle, ",")
        If a = True Then
            TypeCurrency = 0
            Exit Function
        End If
    End If

    'testa as teclas. Somente aceita numerico e backspace e ponto
    Select Case Tecla
        Case 44 'Ponto
        Case 46
            TypeCurrency = 44
            Exit Function
        Case 8    'backspace
        Case 48 To 57 'numeros
        Case Else
            TypeCurrency = 0  'nada
            Exit Function
    End Select
        TypeCurrency = Tecla
    
End Function

Public Function OnlyNumbers(Controle As TextBox, Tecla As Integer) As Integer
'Verfifica se o valor digitado é somente numeros
'Usar no evento KeyPress

    'testa as teclas. Somente aceita numerico e backspace
    Select Case Tecla
        Case 44
            OnlyNumbers = 46
            Exit Function
        Case 8    'backspace
        Case 48 To 57 'numeros
        Case Else
            OnlyNumbers = 0  'nada
            Exit Function
    End Select
        OnlyNumbers = Tecla
    
End Function

Public Sub DoDates(X As Object, Y As Object)
        
    'Preenche os campos datainicial e final, tendo
    'como ponto de partida a segunda-feira e ponto final o domingo
    Select Case Weekday(Date)
    Case 1 'Domingo
        X = DateAdd("d", -6, Date)
        Y = Date
    Case 2 'Segunda
        X = Date
        Y = DateAdd("d", 6, Date)
    Case 3 'Terça
        X = DateAdd("d", -1, Date)
        Y = DateAdd("d", 5, Date)
    Case 4 'Quarta
        X = DateAdd("d", -2, Date)
        Y = DateAdd("d", 4, Date)
    Case 5 'Quinta
        X = DateAdd("d", -3, Date)
        Y = DateAdd("d", 3, Date)
    Case 6 'Sexta
        X = DateAdd("d", -4, Date)
        Y = DateAdd("d", 2, Date)
    Case 7 'Sábado
        X = DateAdd("d", -5, Date)
        Y = DateAdd("d", 1, Date)
    End Select

End Sub

Public Function FormatFloatForDB(ByVal Valor As String) As String
'PARA INSERIR VALORES MONETARIOS NO MYSQL, USA-SE O PONTO, AO INVÉS DA VÍRGULA
'COMO O SEPARADOR DE DECIMAIS
    
    'PARA EVITAR PROBLEMAS COM O PONTO NA INSERÇÃO DO MYSQL, RETIRO TODOS ANTES DE VOLTAR O VALOR
    Valor = Replace(Valor, ".", "")
    'DEPOIS MUDO A VIRGULA PARA PONTO, PARA O SEPARADOR DE DECIMAIS
    FormatFloatForDB = Replace(Valor, ",", ".")
    
End Function

Public Function AllFiles(ByVal FullPath As String) As String()
'***************************************************
'PURPOSE: Returns all files in a folder using
'the FileSystemObject

'PARAMETER: FullPath = FullPath to folder for
'which you want all files

'RETURN VALUE: An array containing a list of
'all file names in FullPath, or a 1-element
'array with an empty string if FullPath
'does not exist or it has no files

'REQUIRES: Reference to Micrsoft Scripting
'          Runtime

'EXAMPLE:

'Dim sFiles() as string
'dim lCtr as long
'sFiles = AllFiles("C:\Windows\System")
'For lCtr = 0 to Ubound(sFiles)
'  Debug.Print sfiles(lctr)
'Next

'REMARKS:  The FileSystemObject does not
'Allow for the use of wild cards (e.g.,
'*.txt.)  If this is what you need, see
'http://wwww.freevbcode.com/ShowCode.asp?ID=1331
'************************************************

Dim oFs As New FileSystemObject
Dim sAns() As String
Dim oFolder As Folder
Dim oFile As File
Dim lElement As Long

ReDim sAns(0) As String
If oFs.FolderExists(FullPath) Then
    Set oFolder = oFs.GetFolder(FullPath)
 
    For Each oFile In oFolder.Files
      lElement = IIf(sAns(0) = "", 0, lElement + 1)
      ReDim Preserve sAns(lElement) As String
      sAns(lElement) = oFile.Name
    Next
End If

AllFiles = sAns
ErrHandler:
    Set oFs = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
End Function
