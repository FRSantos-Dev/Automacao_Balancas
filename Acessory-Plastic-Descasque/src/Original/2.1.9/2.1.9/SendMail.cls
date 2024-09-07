VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private psTo                As String       ' 'To:' email address, delimited with ';'
Private psToDisplay         As String       ' 'To:' display name, delimited with ';'
Private psCc                As String       ' 'Cc:' email address, delimited with ';'
Private psCcDisplay         As String       ' 'Cc:' display name, delimited with ';'
Private psBcc               As String       ' 'Bcc:' email address, delimited with ';'
Private psFrom              As String       ' 'From:' email address
Private psFromDisplay       As String       ' 'From:' display name
Private psReplyTo           As String       ' 'ReplyTo:' email address
Private psSubject           As String       ' subject
Private psMailMessage       As String       ' the email message body
Private psSMTPHost          As String       ' remote host name or IP number
Private plSMTPPort          As Long         ' remote host port number
Private psAttachment        As String       ' filename for attachment
Private pColErrors          As Collection   ' errors collection
Private pbExitImmediately   As Boolean      ' flag - unrecoverable error, exit sub 'Immediately'
Private pEncodeType         As Integer      ' MIME / UUEncode flag
Private pbUseAuthentication As Boolean      ' flag, use login authentication with host
Private pbHtmlText          As Integer      ' flag, send plain text / html text
Private pb8BitMail          As Boolean
Private psContentBase       As String       ' Content base for HTML text

' error strings used with 'pColErrors' collection to report errors to the user
Private Const ERR_INVALID_HOST = "Invalid Host Name"
Private Const ERR_INVALID_PORT = "Invalid Remote Port"
Private Const ERR_INVALID_REC_EMAIL = "Missing or Invalid Recipient E-mail Address"
Private Const ERR_NO_REC_EMAIL = "No Recipient E-mail Address Specified"
Private Const ERR_INVALID_CC_EMAIL = "Invalid Cc: Recipient E-mail Address"
Private Const ERR_INVALID_BCC_EMAIL = "Invalid Bcc: Recipient E-mail Address"
Private Const ERR_INVALID_SND_EMAIL = "Missing or Invalid Sender E-mail Address"
Private Const ERR_TIMEOUT = "Timeout occurred: The SMTP Host did not respond to the request"
Private Const ERR_FILE_NOT_EXIST = "The file you tried to attach does not exist"
Private Const ERR_RECIPIENT_COUNT = "Too many recipients"
Private Const ERR_HTML_REQUIRES_MIME = "Sending HTML requires MIME encoding"


' list of top level Domains, obtained from www.IANA.com.  Can and will change....
' used in syntax checking for email address & SMTP host names
Private Const TOP_DOMAINS = "COM ORG NET EDU GOV MIL INT AF AL DZ AS AD AO AI AQ AG AR " & _
                            "AM AW AC AU AT AZ BS BH BD BB BY BZ BT BJ BE BM BO BA BW BV " & _
                            "BR IO BN BG BF BI KH CM CA CV KY CF TD CL CN CX CC CO KM CD " & _
                            "CG CK CR CI HR CU CY CZ DK DJ DM DO TP EC EG SV GQ ER EE ET " & _
                            "FK FO FJ FI FR GF PF TF GA GM GE DE GH GI GR GL GD GP GU GT " & _
                            "GG GN GW GY HT HM VA HN HK HU IS IN ID IR IQ IE IM IL IT JM " & _
                            "JP JE JO KZ KE KI KP KR KW KG LA LV LB LS LR LY LI LT LU MO " & _
                            "MK MG MW MY MV ML MT MH MQ MR MU YT MX FM MD MC MN MS MA MZ " & _
                            "MM NA NR NP NL AN NC NZ NI NE NG NU NF MP NO OM PK PW PA PG " & _
                            "PY PE PH PN PL PT PR QA RE RO RU RW KN LC VC WS SM ST SA SN " & _
                            "SC SL SG SK SI SB SO ZA GS ES LK SH PM SD SR SJ SZ SE CH SY " & _
                            "TW TJ TZ TH TG TK TO TT TN TR TM TC TV UG UA AE GB US UM UY " & _
                            "UZ VU VE VN VG VI WF EH YE YU ZR ZM ZW UK"

' constant & arrays used in Base64 & UUEncode encode/decode functions
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63) As String
Private pbBase64Byt(0 To 63) As Byte
Private psUUEncodeChr(0 To 63) As String

' misc defaults
Private Const CONNECT_TIMEOUT = 10          ' seconds to wait before giving up
Private Const CONNECT_RETRY = 5             ' number of times to try before giving up
Private Const MSG_TIMEOUT = 60              ' seconds before timing out on message transmission
Private Const MSG_ATTACH_TIMEOUT = 600      ' seconds before timing out on attachment transmission
Private Const SETTINGS_KEY = "Settings"     ' registry sub key
Private Const DEFAULT_PORT As Long = 25     ' default socket port for SMTP

' maximums per RFC 821                      ' not all of these are currently used...
Private Const MAX_USER_LEN = 64             ' maximum total lenght of a user name
Private Const MAX_DOMAIN_LEN = 64           ' maximum total lenght of a domain name
Private Const MAX_PATH_LEN = 256            ' maximum total lenght of a forward-path or reverse-path
Private Const MAX_COMMAND_LEN = 512         ' maximum total lenght of a command line
Private Const MAX_REPLY_LEN = 512           ' maximum total lenght of a reply line
Private Const MAX_TEXTLINE_LEN = 1000       ' maximum total lenght of a text line
Private Const MAX_RECIPIENTS = 100          ' maximum number of recipients that must be buffered

' enums for host string validation
Public Enum VALIDATE_HOST_METHOD
    VALIDATE_HOST_NONE = 0
    VALIDATE_HOST_SYNTAX = 1
    VALIDATE_HOST_PING = 2
    VALIDATE_HOST_DNS = 3
End Enum
    
' enums for email string validation
Public Enum VALIDATE_METHOD
    VALIDATE_NONE = 0
    VALIDATE_SYNTAX = 1
End Enum

' enums for email encoding
Public Enum ENCODE_METHOD
    MIME_ENCODE = 0
    UU_ENCODE = 1
End Enum

Private plConnectTimeout        As Long
Private plCommandTimeout        As Long
Private plMessageTimeOut        As Long
Private plConnectRetry          As Long
Private plPersistentSettings    As Long
Private psRegKey                As String
Private etEmailValidation       As VALIDATE_METHOD
Private etSMTPHostValidation    As VALIDATE_METHOD

' name for instance of the socket form
Private frmSocket As frmSckHost

Public Event SendSuccesful()
Public Event SendFailed(Explanation As String)
Public Event Status(Status As String)
Public Event Progress(PercentComplete As Long)

Private Sub Class_Initialize()

Dim iPtr    As Integer

Set pColErrors = New Collection

' registry key
psRegKey = App.EXEName

' setup default values...
plPersistentSettings = CLng(GetSetting(psRegKey, SETTINGS_KEY, "PersistentSettings", -1))
If plPersistentSettings Then
    From = GetSetting(psRegKey, SETTINGS_KEY, "From", "")
    FromDisplayName = GetSetting(psRegKey, SETTINGS_KEY, "FromDisplayName", "")
    SMTPHost = GetSetting(psRegKey, SETTINGS_KEY, "RemoteHost", "")
    SMTPPort = CLng(GetSetting(psRegKey, SETTINGS_KEY, "RemotePort", DEFAULT_PORT))
    SMTPHostValidation = GetSetting(psRegKey, SETTINGS_KEY, "SMTPHostValidation", VALIDATE_HOST_DNS)
    EmailAddressValidation = CLng(GetSetting(psRegKey, SETTINGS_KEY, "EmailValidation", VALIDATE_SYNTAX))
    ConnectTimeout = CLng(GetSetting(psRegKey, SETTINGS_KEY, "ConnectTimeout", CONNECT_TIMEOUT))
    MessageTimeout = CLng(GetSetting(psRegKey, SETTINGS_KEY, "MessageTimeout", MSG_TIMEOUT))
    ConnectRetry = CLng(GetSetting(psRegKey, SETTINGS_KEY, "ConnectRetry", CONNECT_RETRY))
    EncodeType = GetSetting(psRegKey, SETTINGS_KEY, "EncodeType", MIME_ENCODE)
    Username = GetSetting(psRegKey, SETTINGS_KEY, "Username", "")
    UseAuthentication = GetSetting(psRegKey, SETTINGS_KEY, "UseAuthentication", False)
Else
    SMTPPort = DEFAULT_PORT                     ' Standard SMTP port (25)
    SMTPHostValidation = VALIDATE_HOST_DNS      ' VALIDATE_HOST_NONE,VALIDATE_HOST_SYNTAX,VALIDATE_HOST_PING,VALIDATE_HOST_DNS
    EmailAddressValidation = VALIDATE_SYNTAX    ' VALIDATE_NONE,VALIDATE_SYNTAX
    ConnectTimeout = CONNECT_TIMEOUT
    MessageTimeout = MSG_TIMEOUT
    ConnectRetry = CONNECT_RETRY
    EncodeType = MIME_ENCODE
End If

' initialize arrays for base64 & uuencode encoders
For iPtr = 0 To 63
    pbBase64Byt(iPtr) = Asc(Mid$(BASE64CHR, iPtr + 1, 1))   ' byte array value
    psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)        ' string array value
    psUUEncodeChr(iPtr) = Chr$(iPtr + &H20)
Next
psUUEncodeChr(0) = Chr$(&H60)

End Sub

Private Sub Class_Terminate()

On Local Error Resume Next

' just in case things were accidently left open...
frmSocket.sckMail.Close
Set frmSocket.EmailClient = Nothing
Unload frmSocket
Set frmSocket = Nothing
Set pColErrors = Nothing

End Sub

Public Sub Shutdown()

' stub function, here to maintain binary
' compatibility with previous versions.

End Sub

Public Property Get Subject() As String

Subject = psSubject

End Property

Public Property Let Subject(ByVal NewValue As String)

psSubject = NewValue

End Property

Public Property Get From() As String
    
From = psFrom

End Property


Public Property Let From(ByVal NewValue As String)
    
Dim bValid As Boolean

' Appropriate validation methods for 'From' are:
' VALIDATE_NONE
' VALIDATE_SYNTAX

If etEmailValidation = VALIDATE_SYNTAX Then
    bValid = IsValidEmailAddress(NewValue)
Else
    bValid = True
End If

If bValid Then
    psFrom = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "From", NewValue
    RemoveError ERR_INVALID_SND_EMAIL
Else
    AddError ERR_INVALID_SND_EMAIL
End If
    
End Property

Public Property Let FromDisplayName(ByVal NewValue As String)

psFromDisplay = NewValue
If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "FromDisplayName", NewValue

End Property

Public Property Get FromDisplayName() As String
    
FromDisplayName = psFromDisplay

End Property

Public Property Let CcDisplayName(ByVal NewValue As String)

psCcDisplay = Delimiter(NewValue)

End Property

Public Property Get CcDisplayName() As String
    
CcDisplayName = psCcDisplay

End Property

Public Property Get CcRecipient() As String

CcRecipient = psCc

End Property

Public Property Let CcRecipient(ByVal NewValue As String)

Dim bValid      As Boolean
Dim vTmp        As Variant
Dim iPtr        As Integer

' Appropriate validation methods for 'CcRecipient' are:
' VALIDATE_NONE
' VALIDATE_SYNTAX

If NewValue = "" Then
    psCc = NewValue
    RemoveError ERR_INVALID_CC_EMAIL

Else
    NewValue = Delimiter(NewValue)
    bValid = True
    If etEmailValidation = VALIDATE_SYNTAX Then
        vTmp = Split(NewValue, ";")
        For iPtr = LBound(vTmp) To UBound(vTmp)
            bValid = IsValidEmailAddress(CStr(vTmp(iPtr)))
            If Not bValid Then Exit For
        Next
    End If
    
    If bValid Then
        psCc = NewValue
        RemoveError ERR_INVALID_CC_EMAIL
    Else
        AddError ERR_INVALID_CC_EMAIL
    End If

End If
    
End Property

Public Property Get BccRecipient() As String

BccRecipient = psBcc

End Property

Public Property Let BccRecipient(ByVal NewValue As String)

Dim bValid      As Boolean
Dim vTmp        As Variant
Dim iPtr        As Integer

' Appropriate validation methods for 'BccRecipient' are:
' VALIDATE_NONE
' VALIDATE_SYNTAX

If NewValue = "" Then
    psBcc = NewValue
    RemoveError ERR_INVALID_BCC_EMAIL

Else
    NewValue = Delimiter(NewValue)
    bValid = True
    If etEmailValidation = VALIDATE_SYNTAX Then
        vTmp = Split(NewValue, ";")
        For iPtr = LBound(vTmp) To UBound(vTmp)
            bValid = IsValidEmailAddress(CStr(vTmp(iPtr)))
            If Not bValid Then Exit For
        Next
    End If
    
    If bValid Then
        psBcc = NewValue
        RemoveError ERR_INVALID_BCC_EMAIL
    Else
        AddError ERR_INVALID_BCC_EMAIL
    End If

End If
    
End Property


Public Property Let EncodeType(ByVal NewValue As ENCODE_METHOD)

If NewValue = MIME_ENCODE Or NewValue = UU_ENCODE Then
    pEncodeType = NewValue
    SaveSetting psRegKey, SETTINGS_KEY, "EncodeType", NewValue
End If

End Property

Public Property Get EncodeType() As ENCODE_METHOD

EncodeType = pEncodeType

End Property

Public Property Get Recipient() As String

Recipient = psTo

End Property

Public Property Let Recipient(ByVal NewValue As String)

Dim bValid      As Boolean
Dim iPtr        As Integer
Dim vTmp        As Variant

' Appropriate validation methods for 'Recipient' are:
' VALIDATE_NONE
' VALIDATE_SYNTAX

NewValue = Delimiter(NewValue)
bValid = True
If etEmailValidation = VALIDATE_SYNTAX Then
    vTmp = Split(NewValue, ";")
    For iPtr = LBound(vTmp) To UBound(vTmp)
        bValid = IsValidEmailAddress(CStr(vTmp(iPtr)))
        If Not bValid Then Exit For
    Next
End If
 
If bValid Then
    psTo = NewValue
    RemoveError ERR_INVALID_REC_EMAIL
Else
    AddError ERR_INVALID_REC_EMAIL
End If
    
End Property
Public Property Get RecipientDisplayName() As String
    
RecipientDisplayName = psToDisplay

End Property

Public Property Let RecipientDisplayName(ByVal NewValue As String)

psToDisplay = Delimiter(NewValue)

End Property

Public Property Get ReplyToAddress() As String

ReplyToAddress = psReplyTo

End Property

Public Property Let ReplyToAddress(ByVal NewValue As String)

psReplyTo = "<" & NewValue & ">"

End Property

Public Property Get Message() As String

Message = psMailMessage

End Property

Public Property Let Message(ByVal NewValue As String)

Dim lPtr    As Long

psMailMessage = NewValue

' check for any 8 bit characters
pb8BitMail = False
For lPtr = 1 To Len(psMailMessage)
    If Asc(Mid$(psMailMessage, lPtr, 1)) > 127 Then
        pb8BitMail = True
        Exit For
    End If
Next

End Property

Public Property Get SMTPHost() As String

SMTPHost = psSMTPHost

End Property

Public Property Let SMTPHost(NewValue As String)

Dim bValid      As Boolean

' Appropriate validation methods for 'SMTPHost' are:
' VALIDATE_NONE
' VALIDATE_SYNTAX
' VALIDATE_PING
' VALIDATE_DNS

' Verify the new host name...
Select Case etSMTPHostValidation
    Case VALIDATE_HOST_SYNTAX: bValid = IsValidIPHost(NewValue)
    Case VALIDATE_HOST_PING: bValid = Ping(NewValue)
    Case VALIDATE_HOST_DNS: If GetIPAddress(NewValue) <> "" Then bValid = True
    Case Else: bValid = True
End Select

If bValid Then
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "RemoteHost", NewValue
    RemoveError ERR_INVALID_HOST
    psSMTPHost = NewValue
Else
    AddError ERR_INVALID_HOST
End If

End Property

Public Property Get SMTPPort() As Long

SMTPPort = plSMTPPort

End Property

Public Property Let SMTPPort(ByVal NewValue As Long)

If NewValue < 1 Or NewValue > 65535 Then
    AddError ERR_INVALID_PORT
Else
    plSMTPPort = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "RemotePort", NewValue
    RemoveError ERR_INVALID_PORT
End If

End Property

Friend Sub AddError(ErrStr As String)

On Local Error Resume Next
pColErrors.Add ErrStr, ErrStr

End Sub

Private Sub RemoveError(ErrStr As String)

Dim i   As Long

On Local Error Resume Next

For i = 1 To pColErrors.Count
    If pColErrors(i) = ErrStr Then pColErrors.Remove ErrStr
Next

End Sub


Public Sub Send()

Dim sSenderName         As String
Dim sToHeader           As String
Dim sCcHeader           As String
Dim sAttachFileName     As String
Dim sSplit()            As String
Dim lMessageTimeOut     As Long
Dim sAttachArray()      As String
Dim iAttach             As Integer
Dim iCtr                As Integer
Dim sToAddr()           As String
Dim sToName()           As String
Dim sCcAddr()           As String
Dim sCcName()           As String
Dim sBccAddr()          As String
Dim sHello              As String
Dim sAuth               As String
Dim sTxt                As String
Dim strBoundry          As String
Dim bMimeMultiPart      As Boolean
Dim Qt$

Qt$ = Chr$(34)

' general catch all error handler only
' works when running in stand alone EXE
If bInEXE Then On Local Error GoTo Err_Send

' initialize the socket control form
Set frmSocket = New frmSckHost
Set frmSocket.EmailClient = Me
frmSocket.sckMail.RemoteHost = psSMTPHost
frmSocket.sckMail.RemotePort = plSMTPPort

' Prepair to send mail...
lMessageTimeOut = IIf(psAttachment = "", plMessageTimeOut, MSG_ATTACH_TIMEOUT)
If psTo = "" Then AddError ERR_NO_REC_EMAIL

' HTML & UU Encode are mutually exclusive
If pbHtmlText = True And pEncodeType = UU_ENCODE Then
    AddError ERR_HTML_REQUIRES_MIME
End If

' build recipient & name arrays
sToAddr = Split(psTo, ";")
sToName = Split(psToDisplay, ";")
sCcAddr = Split(psCc, ";")
sCcName = Split(psCcDisplay, ";")
sBccAddr = Split(psBcc, ";")

' check recipient count
If UBound(sToAddr) + UBound(sCcAddr) + UBound(sBccAddr) + 3 > MAX_RECIPIENTS Then AddError ERR_RECIPIENT_COUNT

' If there's already an error, we won't try to send
If pColErrors.Count > 0 Then
    SendFail
    Exit Sub
End If

' Begin here...
ResetGlobals
glMailSize = EstimateMailSize
If pEncodeType = MIME_ENCODE And Len(psAttachment) > 0 Then
    bMimeMultiPart = True
Else
    bMimeMultiPart = False
End If

' this flag gets set when a socket error occurs or the host cannot process an
' input command, see 'SendFail', 'sckMail_DataArrival' & 'WaitUntilTrue' Subs
pbExitImmediately = False

With frmSocket.sckMail
    ' start an SMTP session...
    ' open a connection with the remote host
    ' try 'plConnectRetry' times before giving up
    ' Notify the user that we're starting
    RaiseEvent Status("Connecting to Server...")
    For iCtr = 1 To plConnectRetry
        If .State <> sckConnected Then
            gbConnectOK = False
            'could be error if .state = sckConnecting
            If .State <> sckConnecting Then .Connect
            WaitUntilTrue gbConnectOK, plConnectTimeout
            If gbConnectOK Then Exit For
            If pbExitImmediately Then Exit Sub
        Else
            Exit For
        End If
    Next
    If Not gbConnectOK Then
        TimeOut
        Exit Sub
    End If
    
    ' once a connection is established, say 'hello'
    RaiseEvent Status("Initializing Communications...")
    gbRequestAccepted = False
    ' EHLO is the extended (ESMTP) hello command, HELO is the standard hello command
    If pbUseAuthentication Then sHello = "EHLO " Else sHello = "HELO "
    .SendData sHello & Mid$(psFrom, InStr(psFrom, "@") + 1) & vbCrLf
    WaitUntilTrue gbRequestAccepted, plConnectTimeout
    If pbExitImmediately Then Exit Sub
    If Not gbRequestAccepted Then
        TimeOut
        Exit Sub
    End If
    
    ' Login Authentication ...
    ' the 'EHLO" command will cause the host to send a list of supported extensions
    ' via a series of 250 replies, wait to see if 'Auth Logon' is listed. The Sub
    ' sckMail_DataArrival will set pbUseAuthentication = True if Auth Login is
    ' supported by the remote host. If it is supported, Sub sckMail_DataArrival will
    ' respond to the host's Username & Password requests (gsUsername, gsPassword).
    If pbUseAuthentication = True Then
        WaitUntilTrue gbAuthLoginSupported, 5
        If gbAuthLoginSupported Then
            RaiseEvent Status("Sending Login Authentication...")
            .SendData "AUTH Login" & vbCrLf
            WaitUntilTrue gbAuthLoginSuccess, 5
            If gbAuthLoginSuccess Then
                RaiseEvent Status("Host Login OK!")
            Else
                RaiseEvent Status("Host Login Failed!")
            End If
            If pbExitImmediately Then Exit Sub
        Else
            RaiseEvent Status("Login Not Supported by Host, Continuing...")
        End If
    End If

    ' tell the host who the mail is 'From'
    RaiseEvent Status("Sending Sender Information...")
    gbRequestAccepted = False
    If gbAuthMailFromOK Then sAuth = " AUTH=" & psFrom Else sAuth = vbNullString
    .SendData "MAIL FROM: <" & psFrom & ">" & sAuth & vbCrLf
    WaitUntilTrue gbRequestAccepted, plConnectTimeout
    If pbExitImmediately Then Exit Sub
    If Not gbRequestAccepted Then
        TimeOut
        Exit Sub
    End If
    
    ' tell the host who the recipients are
    ' build the 'To:' header string 'sToHeader' too
    RaiseEvent Status("Sending Recipient Information...")
    For iCtr = LBound(sToAddr) To UBound(sToAddr)
        ' send the recipient address & wait for a reply
        gbRequestAccepted = False
        .SendData "RCPT TO: <" & Trim$(sToAddr(iCtr)) & ">" & vbCrLf
        WaitUntilTrue gbRequestAccepted, plConnectTimeout
        If pbExitImmediately Then Exit Sub
        If Not gbRequestAccepted Then
            TimeOut
            Exit Sub
        End If
        ' build the 'To:' header string
        If iCtr <= UBound(sToName) Then
            If Len(sToName(iCtr)) Then sToHeader = sToHeader & Qt$ & sToName(iCtr) & Qt$
            If Len(sToHeader) Then sToHeader = sToHeader & " "
        End If
        sToHeader = sToHeader & "<" & sToAddr(iCtr) & ">"
        If iCtr < UBound(sToAddr) Then sToHeader = sToHeader & ", " & vbCrLf & "    "
    Next
    
    ' send Cc: recipient addresses (just more 'RCPT TO' addresses)
    ' build the 'Cc:' header string too
    For iCtr = LBound(sCcAddr) To UBound(sCcAddr)
        ' send the recipient address & wait for a reply
        gbRequestAccepted = False
        .SendData "RCPT TO: <" & Trim$(sCcAddr(iCtr)) & ">" & vbCrLf
        WaitUntilTrue gbRequestAccepted, plConnectTimeout
        If pbExitImmediately Then Exit Sub
        If Not gbRequestAccepted Then
            TimeOut
            Exit Sub
        End If
        ' build the 'Cc:' header string
        If iCtr <= UBound(sCcName) Then
            If Len(sCcName(iCtr)) Then sCcHeader = sCcHeader & Qt$ & sCcName(iCtr) & Qt$
            If Len(sCcHeader) Then sCcHeader = sCcHeader & " "
        End If
        sCcHeader = sCcHeader & "<" & sCcAddr(iCtr) & ">"
        If iCtr < UBound(sCcAddr) Then sCcHeader = sCcHeader & ", " & vbCrLf & "    "
    Next
       
    ' send Bcc: recipient addresses (more of the same)
    ' no display headers here, these are blind
    For iCtr = LBound(sBccAddr) To UBound(sBccAddr)
        ' send the recipient address & wait for a reply
        gbRequestAccepted = False
        .SendData "RCPT TO: <" & Trim$(sBccAddr(iCtr)) & ">" & vbCrLf
        WaitUntilTrue gbRequestAccepted, plConnectTimeout
        If pbExitImmediately Then Exit Sub
        If Not gbRequestAccepted Then
            TimeOut
            Exit Sub
        End If
    Next
       
    ' tell the remote host we're ready to send data
    RaiseEvent Status("Sending Message...")
    gbDataOK = False
    .SendData "DATA" & vbCrLf
    WaitUntilTrue gbDataOK, plConnectTimeout
    If pbExitImmediately Then Exit Sub
    If Not gbDataOK Then
        TimeOut
        Exit Sub
    End If
    
    ' OK, the host is ready for data, this is where the mail message starts
    ' Send the mail headers, what is displayed on the target mail client
    gbRequestAccepted = False
    .SendData "Date: " & Format(Now, "dddd dd-mmm-yyyy hh:nn:ss") & vbCrLf
    If Len(psFromDisplay) Then sSenderName = Qt$ & psFromDisplay & Qt$ & " "
    sSenderName = sSenderName & "<" & psFrom & ">"
    .SendData "From: " & sSenderName & vbCrLf
    .SendData "To: " & sToHeader & vbCrLf
    If Len(sCcHeader) Then .SendData "Cc: " & sCcHeader & vbCrLf
    If Len(psReplyTo) Then .SendData "Reply-to: " & psReplyTo & vbCrLf
    .SendData "Subject: " & psSubject & vbCrLf
    
    If pEncodeType = MIME_ENCODE Then
        ' create a Unique-Boundary string for multi-part MIME encoding
        strBoundry = "----_=_NextPart_000_" & Hex$(GetTickCount Or &H80000000) & _
                     "." & Hex$((GetTickCount Xor &HFFFFFFFF) Or &H80000000)

        ' send the MIME start headers
        .SendData "MIME-Version: 1.0" & vbCrLf
        If bMimeMultiPart Then
            .SendData "Content-Type: multipart/mixed;" & vbCrLf
            .SendData vbTab & "boundary=" & Qt$ & strBoundry & Qt$ & vbCrLf & vbCrLf
            .SendData "This is a multi-part message in MIME format. Since" & vbCrLf
            .SendData "your mail reader does not understand this format," & vbCrLf
            .SendData "some or all of this message may not be legible." & vbCrLf & vbCrLf
            ' send the MIME boundry and content headers for the message body
            .SendData "--" & strBoundry & vbCrLf
        End If
        ' plain or html text...
        If pbHtmlText Then sTxt = "text/html;" Else sTxt = "text/plain;"
        .SendData "Content-Type: " & sTxt & vbCrLf
        .SendData vbTab & "charset=" & Qt$ & "iso-8859-1" & Qt$ & vbCrLf ' "us-ascii", "Windows-1252"
        If pb8BitMail Then sTxt = "8bit" Else sTxt = "7bit"
        .SendData "Content-Transfer-Encoding: " & sTxt & vbCrLf
        ' if we're sending html & the user supplied the content base then send it too
        If pbHtmlText Then If Len(psContentBase) Then .SendData "Content-Base: " & Qt$ & psContentBase & Qt$ & vbCrLf
    End If
    
    .SendData vbCrLf
        
    ' Send the message body
    .SendData FormatMail(psMailMessage) & vbCrLf & vbCrLf
         
    ' Send attachments, if any...
    sAttachArray = Split(psAttachment, ";")
    For iAttach = LBound(sAttachArray) To UBound(sAttachArray)
        If sAttachArray(iAttach) <> "" Then
            ' strip the path information off the filename
            sSplit = Split(sAttachArray(iAttach), "\")
            sAttachFileName = sSplit(UBound(sSplit))
            RaiseEvent Status("Sending Attachment, " & sAttachFileName & "...")
            If pEncodeType = MIME_ENCODE Then
                ' send the next MIME boundry & content headers
                .SendData "--" & strBoundry & vbCrLf
                .SendData "Content-Type: " & GetContentType(sAttachFileName) & ";" & vbCrLf
                .SendData vbTab & "name=" & Qt$ & sAttachFileName & Qt$ & vbCrLf
                .SendData "Content-Transfer-Encoding: base64" & vbCrLf
                .SendData "Content-Disposition: attachment;" & vbCrLf
                .SendData vbTab & "filename=" & Qt$ & sAttachFileName & Qt$ & vbCrLf & vbCrLf
                ' send the encoded file
                EncodeAndSendFile sAttachArray(iAttach), MIME_ENCODE
                If pbExitImmediately Then Exit Sub
                .SendData vbCrLf
            Else
                ' start a UUEncode session
                .SendData "begin 600 " & sAttachFileName & vbCrLf
                ' send the encoded file
                EncodeAndSendFile sAttachArray(iAttach), UU_ENCODE
                If pbExitImmediately Then Exit Sub
                ' send the ending sequence
                .SendData "end" & vbCrLf
            End If
        End If
    Next iAttach
                
    If bMimeMultiPart = True Then
        ' send the MIME closing boundry header
        .SendData "--" & strBoundry & "--" & vbCrLf
    End If
    
    ' Send the 'end of mail' string
    gbRequestAccepted = False
    .SendData "." & vbCrLf
    WaitUntilTrue gbRequestAccepted, lMessageTimeOut
    If pbExitImmediately Then Exit Sub
    If Not gbRequestAccepted Then
        TimeOut
        Exit Sub
    End If
    
    ' Be polite, tell the host we're closing the connection...
    RaiseEvent Status("Closing Connection...")
    .SendData "QUIT" & vbCrLf
    gbRequestAccepted = False
    WaitUntilTrue gbRequestAccepted, 2
    RaiseEvent Progress(100)
    RaiseEvent SendSuccesful
    .Close

End With

' release the socket form
Set frmSocket.EmailClient = Nothing
Unload frmSocket
Set frmSocket = Nothing

Exit Sub

Err_Send:

AddError Err.Description
SendFail

End Sub

Friend Sub SendFail()

Dim iCtr            As Integer
Dim sErrorString    As String

If bInEXE Then On Local Error Resume Next

' report error to the user
For iCtr = 1 To pColErrors.Count
    sErrorString = sErrorString & pColErrors(iCtr) & vbCrLf
Next
RaiseEvent SendFailed(sErrorString)

' close the connection with the remote host
With frmSocket.sckMail
    If Not .State = sckClosed Then
        ' be polite, if the port is open
        ' tell the remote host we're closing
        If .State = sckOpen Or .State = sckConnected Then
            RaiseEvent Status("Error Occurred, Closing Connection...")
            .SendData "QUIT" & vbCrLf
            ' allow time for the message to get out
            gbRequestAccepted = False
            WaitUntilTrue gbRequestAccepted, 2
        End If
        ' close the connection
        .Close
        Do Until .State = sckClosed
            DoEvents
        Loop
    End If
End With

' set flag to exit 'Send' Sub without further processing
pbExitImmediately = True

' release the socket form
Set frmSocket.EmailClient = Nothing
Unload frmSocket
Set frmSocket = Nothing

'clear all errors
Set pColErrors = New Collection

End Sub


Public Function IsValidEmailAddress(AddressString As String)

Dim strTmp      As String
Dim iPos        As Integer

' assume failure
IsValidEmailAddress = False

' len cannot be zero
strTmp = Trim$(AddressString)
If Len(strTmp) = 0 Then Exit Function

' must have one "@" & it can't be the first character
iPos = InStr(strTmp, "@")
If iPos <= 1 Then Exit Function

'can't have multiple "@" chars in the string
If InStr(iPos + 1, strTmp, "@") Then Exit Function

IsValidEmailAddress = IsValidIPHost(strTmp)

End Function

Private Function IsValidIPHost(HostString As String) As Boolean

Dim sHost               As String
Dim sSplit()            As String

sHost = UCase$(Trim$(HostString))

If IsDottedQuad(sHost) Then
    IsValidIPHost = True
    Exit Function
End If

sSplit = Split(sHost, ".")

' it's not numeric, top level domain?
If UBound(sSplit) > 0 And InStr(TOP_DOMAINS, sSplit(UBound(sSplit))) > 0 Then
    IsValidIPHost = True
    Exit Function
End If

End Function

Private Sub TimeOut()

AddError ERR_TIMEOUT
SendFail

End Sub

Private Sub WaitUntilTrue(Flag As Boolean, TimeToWait As Long)

Dim lStart              As Long
Dim lTimetoQuit         As Long

lStart = CLng(Timer)

If TimeToWait < 1 Then Exit Sub

' Deal with timer being reset at Midnight
If lStart + TimeToWait < 86400 Then
    lTimetoQuit = lStart + TimeToWait
Else
    lTimetoQuit = (lStart - 86400) + TimeToWait
End If

Do Until Flag = True Or Timer >= lTimetoQuit Or pbExitImmediately = True
    DoEvents
    Sleep (10)      ' added to reduce CPU load during wait periods
Loop
   
End Sub


Private Sub EncodeAndSendFile(ByVal strFile As String, Encode As ENCODE_METHOD)

'******************************************************************************
'
' Synopsis:     Send a file attachment via an open socket
'
' Parameters:   strFile  - The input file name
'               Encode   -  type of encoding to use; MIME or UU
'
' Return:       nothing
'
' Description:
'   Open the file & read characters in. Send the characters through the appropriate
'   encoder, either MIME (Base64) or UUEncode, before tranmission via an open socket.
'
'******************************************************************************

Dim hFile               As Integer      ' file handle
Dim sValue              As String       ' temp string buffer
Dim bInFile()           As Byte         ' byte array file buffer
Dim lEventCtr           As Long         ' counter
Dim lChunkSize          As Long         ' number of bytes to get
Dim lNumBytes           As Long         ' file pointer

If bInEXE Then On Local Error GoTo File_Error

hFile = FreeFile
Open strFile For Binary Access Read As #hFile
lNumBytes = LOF(hFile)

If Encode = MIME_ENCODE Then

    Do While lNumBytes
        ' set input buffer size, MUST be a multiple of 57
        lChunkSize = IIf(lNumBytes > 2850, 2850, lNumBytes)
        
        ' read & Base 64 encode a group of characters
        bInFile = InputB(lChunkSize, #hFile)
        frmSocket.sckMail.SendData EncodeBase64Byte(bInFile)
        
        ' adjust file pointer
        lNumBytes = lNumBytes - lChunkSize
        
        DoEvents
    Loop

ElseIf Encode = UU_ENCODE Then

    Do While Not EOF(hFile)
        ' read & UU encode a line of characters
        sValue = Input(45, #hFile)
        frmSocket.sckMail.SendData UUEncodeString(sValue) & vbCrLf
    
        ' DoEvents (occasionally)
        lEventCtr = lEventCtr + 1
        If lEventCtr Mod 50 = 0 Then DoEvents
    Loop

End If

File_Done:

Close #hFile

Exit Sub

File_Error:

    AddError Err.Description
    SendFail
    pbExitImmediately = True
    Resume File_Done

End Sub

Private Function EncodeBase64Byte(InArray() As Byte) As Byte()

'******************************************************************************
'
' Synopsis:     Base 64 encode a byte array
'
' Parameters:   InArray  - The input byte array
'
' Return:       encoded byte array
'
' Description:
'   Convert a byte array to a Base 64 encoded byte array. Coerce 3 bytes into 4 by
'   converting 3, 8 bit bytes into 4, 6 bit values. Each 6 bit value (0 to 63) is
'   then used as a pointer into a base64 byte array to derive a character. This
'   rountine performs the same task as the EncodeBase64String function that follows it
'   except that the input and return data types are byte arrays. This is an important
'   distinction, as this routine is compatible with double byte character sets (DBCS)
'   like Chinese while the EncodeBase64String function is not.
'
'******************************************************************************

Dim lInPtr              As Long         ' pointer into input array
Dim lOutPtr             As Long         ' pointer into output array
Dim OutArray()          As Byte         ' output byte array buffer
Dim lLen                As Long         ' number of extra bytes past 3 byte boundry
Dim iNewLine            As Long         ' line counter

' if size of input array is not a multiple of 3,
' increase it to the next multiple of 3
lLen = (UBound(InArray) - LBound(InArray) + 1) Mod 3
If lLen Then
    lLen = 3 - lLen
    ReDim Preserve InArray(UBound(InArray) + lLen)
End If

' create an output buffer
ReDim OutArray(UBound(InArray) * 2 + 100)

' step through the input array, 3 bytes at a time
For lInPtr = 0 To UBound(InArray) Step 3
    
    ' add CrLf as required
    If iNewLine = 19 Then
        OutArray(lOutPtr) = 13
        OutArray(lOutPtr + 1) = 10
        lOutPtr = lOutPtr + 2
        iNewLine = 0
    End If
    
    ' convert 3 bytes into 4 base 64 encoded bytes
    OutArray(lOutPtr) = pbBase64Byt((InArray(lInPtr) And &HFC) \ 4)
    OutArray(lOutPtr + 1) = pbBase64Byt((InArray(lInPtr) And &H3) * &H10 + (InArray(lInPtr + 1) And &HF0) \ &H10)
    OutArray(lOutPtr + 2) = pbBase64Byt((InArray(lInPtr + 1) And &HF) * 4 + (InArray(lInPtr + 2) And &HC0) \ &H40)
    OutArray(lOutPtr + 3) = pbBase64Byt(InArray(lInPtr + 2) And &H3F)
    
    ' update pointers
    lOutPtr = lOutPtr + 4
    iNewLine = iNewLine + 1
Next

' add terminator '=' as required
Select Case lLen
    Case 1: OutArray(lOutPtr - 1) = 61
    Case 2: OutArray(lOutPtr - 1) = 61: OutArray(lOutPtr - 2) = 61
End Select

' add CrLf if not already there
If OutArray(lOutPtr - 2) <> 13 Then
    OutArray(lOutPtr) = 13
    OutArray(lOutPtr + 1) = 10
    lOutPtr = lOutPtr + 2
End If

' resize output buffer and return
ReDim Preserve OutArray(lOutPtr - 1)
EncodeBase64Byte = OutArray

End Function


Friend Function EncodeBase64String(str2Encode As String) As String

'******************************************************************************
'
' Synopsis:     Base 64 encode a string
'
' Parameters:   str2Encode  - The input string
'
' Return:       encoded string
'
' Description:
'   Convert a string to a Base 64 encoded string. Coerce 3 bytes into 4 by converting
'   3, 8 bit bytes into 4, 6 bit values. Each 6 bit value (0 to 63) is then used as
'   a pointer into the base64 string array to derive the character. The string will be
'   a multiple of 4 bytes in lenght after conversion, padded with '=' as required.
'
'******************************************************************************

Dim lCtr                As Long
Dim lPtr                As Long
Dim lLen                As Long
Dim sValue              As String
Dim sEncoded            As String
Dim Bits8(1 To 3)       As Byte
Dim Bits6(1 To 4)       As Byte

For lCtr = 1 To Len(str2Encode) Step 3
    
    ' Get 3 characters
    sValue = Mid$(str2Encode, lCtr, 3)
    lLen = Len(sValue)
    
    ' Move string data into a byte array, then swap
    ' bits to create 4, 6 bit base 64 values (0-63)
    If lLen < 3 Then Erase Bits8
    CopyMemory Bits8(1), sValue, lLen
    Bits6(1) = (Bits8(1) And &HFC) \ 4
    Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
    Bits6(3) = (Bits8(2) And &HF) * 4 + (Bits8(3) And &HC0) \ &H40
    Bits6(4) = Bits8(3) And &H3F

    ' add the new characters to the ouput buffer
    For lPtr = 1 To lLen + 1
       sEncoded = sEncoded & psBase64Chr(Bits6(lPtr))
    Next

Next

' The decoder expects the size to be a multiple of 4 bytes.
' Possible sizes for the last packet are: 2, 3 & 4.
Select Case lLen + 1
    Case 2: sEncoded = sEncoded & "=="      ' send two pad characters
    Case 3: sEncoded = sEncoded & "="       ' send one pad character
    Case 4:                                 ' we're done
End Select

EncodeBase64String = sEncoded

End Function


Friend Function DecodeBase64String(str2Decode As String) As String

'******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
'   Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
'   values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
'   ascii character equivalent. Stop converting at the end of the input string
'   or when the first '=' (equal sign) is encountered.
'
'******************************************************************************

Dim lPtr            As Long
Dim iValue          As Integer
Dim iLen            As Integer
Dim iCtr            As Integer
Dim Bits(1 To 4)    As Byte
Dim strDecode       As String

' for each 4 character group....
For lPtr = 1 To Len(str2Decode) Step 4
    iLen = 4
    For iCtr = 0 To 3
        ' retrive the base 64 value, 4 at a time
        iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
        Select Case iValue
            ' A~Za~z0~9+/
            Case 1 To 64: Bits(iCtr + 1) = iValue - 1
            ' =
            Case 65
                iLen = iCtr
                Exit For
            ' not found
            Case 0: Exit Function
        End Select
    Next
    
    ' convert the 4, 6 bit values into 3, 8 bit values
    Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
    Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
    Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)
    
    ' add the three new characters to the output string
    For iCtr = 1 To iLen - 1
        strDecode = strDecode & Chr$(Bits(iCtr))
    Next

Next

DecodeBase64String = strDecode

End Function


Private Function UUEncodeString(str2UUEncode As String) As String
'******************************************************************************
'
' Synopsis:     UUEncode a string
'
' Parameters:   str2UUEncode  - The input string
'
' Return:       encoded string
'
' Description:
'   UU Encode a string. Coerce 3 bytes into 4 by converting 3, 8 bit bytes into
'   4, 6 bit values. Each 6 bit value (0 to 63) is then used as a pointer into the
'   UUEncode string array to derive the correct character. The string will be a multiple
'   of 4 bytes in lenght after conversion, padded with '=' as required. The line length
'   will be encoded as a leading character (same 0 to 63 encoding) in the return string.
'
'******************************************************************************

Dim sValue              As String
Dim lPtr                As Long
Dim lCtr                As Long
Dim lLen                As Long
Dim lLineLen            As Long
Dim sEncoded            As String
Dim Bits8(1 To 3)       As Byte
Dim Bits6(1 To 4)       As Byte

lLineLen = Len(str2UUEncode)
' lines are limited to 63
If lLineLen > 63 Then Exit Function

For lCtr = 1 To Len(str2UUEncode) Step 3
    
     ' Get 3 characters
    sValue = Mid$(str2UUEncode, lCtr, 3)
    lLen = Len(sValue)
    
    ' Move string data into a byte array, then
    ' swap bits to create 4, 6 bit values (0-63)
    If lLen < 3 Then Erase Bits8
    CopyMemory Bits8(1), sValue, lLen
    Bits6(1) = (Bits8(1) And &HFC) \ &H4
    Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
    Bits6(3) = (Bits8(2) And &HF) * &H4 + (Bits8(3) And &HC0) \ &H40
    Bits6(4) = Bits8(3) And &H3F

    ' Encode new 4 byte string by selecting a character from
    ' the array. Length is determined by 'lLen' to make sure
    ' the file attachment is the right length
    For lPtr = 1 To lLen + 1
        sEncoded = sEncoded & psUUEncodeChr(Bits6(lPtr))
    Next

Next

' add the line length character
sEncoded = psUUEncodeChr(lLineLen) & sEncoded

' The decoder expects the size to be a multiple of 4 bytes.
' Possible sizes for the last packet are: 2, 3 & 4.
Select Case lLen + 1
    Case 2: sEncoded = sEncoded & "=="      ' send two pad characters
    Case 3: sEncoded = sEncoded & "="       ' send one pad character
    Case 4:                                 ' no pad characers needed
End Select

UUEncodeString = sEncoded

End Function


Private Function IsDottedQuad(HostString As String) As Boolean

Dim sSplit()        As String
Dim iCtr            As Integer
Dim bAns            As Boolean

IsDottedQuad = False
sSplit = Split(HostString, ".")

If UBound(sSplit) = 3 Then
    For iCtr = 0 To 3
        If Not IsNumeric(sSplit(iCtr)) Then Exit Function
    Next
    bAns = True
    For iCtr = 0 To 3
        If iCtr = 0 Then
            bAns = val(sSplit(iCtr)) <= 239
            If bAns = False Then Exit For
        Else
            bAns = val(sSplit(iCtr)) <= 255
            If bAns = False Then Exit For
        End If
    Next
    IsDottedQuad = bAns
Else
    Exit Function
End If

End Function

Public Function Ping(Address As String, _
                    Optional RoundTripTime As String = "", _
                    Optional DataSize As String = "", _
                    Optional DataMatch As Boolean = False) As Boolean

Dim ECHO            As ICMP_ECHO_REPLY
Dim iPtr            As Integer
Dim Dt              As String
Dim sAddress        As String

'THIS CODE IS BASED ON FUNCTIONS
'WITHIN RICHARD DEEMING'S IP UTILITIES:
'http://www.freevbcode.com/showcode.asp?ID=199

If bInEXE Then On Local Error GoTo DPErr

' if passed a name, get the IP address
If Not IsDottedQuad(Address) Then
    sAddress = GetIPAddress(Address)
Else
    sAddress = Address
End If
    
If sAddress = "" Then Exit Function
    
If SocketsInitialize Then
       
    ' build string of random characters
    For iPtr = 1 To DATA_SIZE
        Dt = Dt & Chr$(Rnd() * 254 + 1)
    Next
        
    'ping an ip address, passing the
    'address and the ECHO structure
    Ping = DoPing(sAddress, Dt, ECHO)
        
    'get the results from the ECHO structure
    RoundTripTime = ECHO.RoundTripTime
    Address = (ECHO.Address And &HFF) & "." & _
              ((ECHO.Address \ &H100) And &HFF) & "." & _
              ((ECHO.Address \ &H10000) And &HFF) & "." & _
              ((ECHO.Address \ &H1000000) And &HFF)
 
    DataSize = ECHO.DataSize & " bytes"
      
    iPtr = InStr(ECHO.Data, Chr$(0))
    If iPtr > 1 Then DataMatch = (Left$(ECHO.Data, iPtr - 1) = Dt)
   
    SocketsCleanup
Else
    Ping = IP_GENERAL_FAILURE
    
End If

Exit Function

DPErr:
    Ping = IP_GENERAL_FAILURE

End Function

Private Function DoPing(szAddress As String, sDataToSend As String, _
                        ECHO As ICMP_ECHO_REPLY) As Boolean

Dim hPort           As Long
Dim dwAddress       As Long
Dim lAns            As Long

dwAddress = AddressStringToLong(szAddress)
hPort = IcmpCreateFile()
    
If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, _
    ECHO, Len(ECHO), PING_TIMEOUT) Then
    'the ping succeeded,
    '.Status will be 0
    '.RoundTripTime is the time in ms for the ping to complete,
    '.Data is the data returned (NULL terminated)
    '.Address is the Ip address that actually replied
    '.DataSize is the size of the string in .Data
    lAns = IP_SUCCESS
Else
    If ECHO.Status = 0 Then
        lAns = -1
    Else
        lAns = ECHO.Status * -1
    End If
End If
                       
Call IcmpCloseHandle(hPort)
DoPing = lAns = IP_SUCCESS

End Function

Public Function GetIPHost() As String

'Resolves the local hosts name
Dim sHostName   As String
Dim iPtr        As Integer

sHostName = String$(256, Chr$(0))

If Not SocketsInitialize() Then Exit Function

If gethostname(sHostName, Len(sHostName)) = ERROR_SUCCESS Then
    iPtr = InStr(sHostName, Chr$(0))
    If iPtr > 1 Then GetIPHost = Mid$(sHostName, 1, iPtr - 1)
End If

SocketsCleanup

End Function


Public Function GetIPAddress(sHostName As String) As String

'Resolves the host-name to an IP address (DNS)
Dim lpHost          As Long
Dim HOST            As HOSTENT
Dim dwIPAddr        As Long
Dim tmpIPAddr()     As Byte
Dim i               As Integer
Dim sIPAddr         As String
Dim werr            As Long

If Not SocketsInitialize() Then
    GetIPAddress = ""
    Exit Function
End If
    
If sHostName = "" Then sHostName = GetIPHost
sHostName = Trim$(sHostName) & Chr$(0)
lpHost = gethostbyname(sHostName)

If lpHost = 0 Then
    werr = WSAGetLastError()
    GetIPAddress = ""
    SocketsCleanup
    Exit Function
End If

CopyMemory HOST, lpHost, Len(HOST)
CopyMemory dwIPAddr, HOST.hAddrList, 4
ReDim tmpIPAddr(1 To HOST.hLen)
CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen

For i = 1 To HOST.hLen
    sIPAddr = sIPAddr & tmpIPAddr(i) & "."
Next

GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
SocketsCleanup

End Function

Private Function AddressStringToLong(ByVal tmp As String) As Long

Dim i           As Integer
Dim vParts      As Variant
  
' we have to extract each part of the
' address string, ex: 123.456.789.123
     
vParts = Split(tmp, ".")

If UBound(vParts) <> 3 Then
    AddressStringToLong = 0
    Exit Function
End If
   
' build the long value out of the
' hex of the extracted strings
    
AddressStringToLong = val("&H" & Right("00" & Hex(vParts(3)), 2) & _
                                 Right("00" & Hex(vParts(2)), 2) & _
                                 Right("00" & Hex(vParts(1)), 2) & _
                                 Right("00" & Hex(vParts(0)), 2))
   
End Function

Public Property Get EmailAddressValidation() As VALIDATE_METHOD

EmailAddressValidation = etEmailValidation

End Property

Public Property Let EmailAddressValidation(ByVal NewValue As VALIDATE_METHOD)

If NewValue >= VALIDATE_NONE And NewValue <= VALIDATE_SYNTAX Then
    etEmailValidation = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "EmailValidation", NewValue
End If

End Property

Public Property Get SMTPHostValidation() As VALIDATE_HOST_METHOD

SMTPHostValidation = etSMTPHostValidation

End Property

Public Property Let SMTPHostValidation(ByVal NewValue As VALIDATE_HOST_METHOD)

If NewValue >= VALIDATE_HOST_NONE And NewValue <= VALIDATE_HOST_DNS Then
    etSMTPHostValidation = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "SMTPHostValidation", NewValue
End If

'in case this is set after the host value is set
If psSMTPHost <> "" Then SMTPHost = psSMTPHost

End Property

Public Property Let Attachment(ByVal NewValue As String)

Dim sAttachName()   As String
Dim bNamesOK        As Boolean
Dim iAttach         As Integer
Dim lPtr            As Long

If Trim(NewValue) <> "" Then
 
    NewValue = Delimiter(NewValue)

    sAttachName = Split(Trim(NewValue), ";")
    bNamesOK = True
    
    For iAttach = 0 To UBound(sAttachName)
        If Dir(sAttachName(iAttach)) = "" Then bNamesOK = False
    Next iAttach
    
    If bNamesOK = True Then
        psAttachment = NewValue
        RemoveError ERR_FILE_NOT_EXIST
    Else
        AddError ERR_FILE_NOT_EXIST
    End If

Else
    psAttachment = NewValue
    RemoveError ERR_FILE_NOT_EXIST
End If

End Property

Public Property Get Attachment() As String

Attachment = psAttachment

End Property


Public Property Get ConnectTimeout() As Long

ConnectTimeout = plConnectTimeout

End Property

Public Property Let ConnectTimeout(ByVal NewValue As Long)

If NewValue > 0 And NewValue <= 120 Then
    plConnectTimeout = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "ConnectTimeout", NewValue
End If

End Property

Public Property Get MessageTimeout() As Long

MessageTimeout = plMessageTimeOut

End Property

Public Property Let MessageTimeout(ByVal NewValue As Long)

If NewValue > 0 And NewValue <= 3600 Then
    plMessageTimeOut = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "MessageTimeout", NewValue
End If

End Property

Public Property Get ConnectRetry() As Long

ConnectRetry = plConnectRetry

End Property

Public Property Let ConnectRetry(ByVal NewValue As Long)

If NewValue > 0 And NewValue <= 20 Then
    plConnectRetry = NewValue
    If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "ConnectRetry", NewValue
End If

End Property

Public Property Get PersistentSettings() As Boolean

PersistentSettings = plPersistentSettings

End Property

Public Property Let PersistentSettings(ByVal NewValue As Boolean)

plPersistentSettings = NewValue
If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "PersistentSettings", CStr(CLng(NewValue))

End Property

Public Property Let Username(ByVal NewValue As String)

gsUserName = NewValue
If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "Username", NewValue

End Property

Public Property Get Username() As String

Username = gsUserName

End Property

Public Property Let Password(ByVal NewValue As String)

gsPassword = NewValue

End Property

Public Property Get Password() As String

Password = gsPassword

End Property

Public Property Let UseAuthentication(ByVal NewValue As Boolean)

pbUseAuthentication = NewValue
If plPersistentSettings Then SaveSetting psRegKey, SETTINGS_KEY, "UseAuthentication", CStr(CLng(NewValue))

End Property

Public Property Get UseAuthentication() As Boolean

UseAuthentication = pbUseAuthentication

End Property

Public Property Let ContentBase(ByVal NewValue As String)

' fix some common mistakes...
NewValue = TrimWithoutPrejudice(NewValue)
If Len(NewValue) Then
    Replace NewValue, "\", "/"
    If InStr(1, NewValue, "http://", vbTextCompare) = 0 Then NewValue = "http://" & NewValue
    If Right$(NewValue, 1) <> "/" Then NewValue = NewValue & "/"
End If
psContentBase = NewValue

End Property

Public Property Get ContentBase() As String

ContentBase = psContentBase

End Property

Public Property Let AsHTML(ByVal NewValue As Boolean)

pbHtmlText = NewValue

End Property

Public Property Get AsHTML() As Boolean

AsHTML = pbHtmlText

End Property


Private Function FormatMail(ByVal strIn As String) As String

'******************************************************************************
'
' Synopsis:     Re-format text lines per RFC 821
'
' Parameters:   strIn   - The input string to be formated
'
' Return:       re-formated string
'
' Description:
'   RFC 821 places the following restrictions on user text:
'       1) Before sending a line of mail text begining with a '.'
'          the sender will add an additional '.'
'
'       2) The receiver checks each line of mail text, if a line is single '.'
'          it is the end of the mail message. If the first character is
'          a '.' and there are other characters on the line, the first '.'
'          is deleted.
'
'       3) The maximum line lenght will not exceed 1000 characters
'
'******************************************************************************

Dim sTextLine()     As String
Dim sRemainder      As String
Dim sNewLine        As String
Dim sDelimiter      As String
Dim lPtr            As Long
Dim lStrPos         As Long
Dim lSplit          As Long

If Len(strIn) = 0 Then Exit Function

' Select the correct delimiter character
If InStr(strIn, vbCrLf) Then
    sDelimiter = vbCrLf
ElseIf InStr(strIn, vbCr) Then
    sDelimiter = vbCr
Else
    sDelimiter = vbNullString
End If

' split the text into seperate lines
sTextLine() = Split(strIn, sDelimiter)

' process each line
For lPtr = LBound(sTextLine) To UBound(sTextLine)
    ' check for lines starting with a '.'
    ' when found, add a second '.'
    If Left$(sTextLine(lPtr), 1) = "." Then sTextLine(lPtr) = "." & sTextLine(lPtr)
    
    ' check that the line is not too long (account for 2 extra characters - vbCrLf)
    ' break into smaller elements as required
    If Len(sTextLine(lPtr)) > MAX_TEXTLINE_LEN - 2 Then
        sRemainder = sTextLine(lPtr)
        sNewLine = vbNullString
        If sDelimiter = vbNullString Then sDelimiter = vbCrLf
        Do While Len(sRemainder) > MAX_TEXTLINE_LEN - 2
            ' try to split at a space character, if not then split at MAX_TEXTLINE_LEN - 2
            lSplit = InStrRev(sRemainder, " ", MAX_TEXTLINE_LEN - 2)
            If lSplit = 0 Then lSplit = MAX_TEXTLINE_LEN - 2
            ' insert a vbCrLf at the split point
            sNewLine = sNewLine & Mid$(sRemainder, 1, lSplit) & sDelimiter
            sRemainder = Mid$(sRemainder, lSplit + 1)
        Loop
        sTextLine(lPtr) = sNewLine & sRemainder
    End If
Next

FormatMail = Join(sTextLine, sDelimiter)

End Function


Private Function Delimiter(ByVal strIn As String) As String

Dim lPtr        As Long

' allow both comma and semicolon as delimiters
' the rest of the app will use semicolons so
' convert any commas to semicolons here

lPtr = InStr(strIn, ",")
Do While lPtr
    Mid$(strIn, lPtr, 1) = ";"
    lPtr = InStr(lPtr + 1, strIn, ",")
Loop

Delimiter = strIn

End Function

Private Function EstimateMailSize() As Long

'******************************************************************************
'
' Synopsis:     Estimate the size (number of bytes) of the mail message
'
' Parameters:   none
'
' Return:       long - number of bytes
'
' Description:
'   Estimate the size in bytes of the mail message being sent. Include the
'   message body, headers, attachments, etc. Account for type of encoding.
'   This is only an estimate, it is not exact.  The result is used to
'   calculate the progress while sending....
'
'******************************************************************************

Dim lNumBytes       As Long
Dim lAttachSize     As Long
Dim sAttach()       As String
Dim iCtr            As Integer
Dim iNumRecip       As Integer
Dim iNumAttach      As Integer
Dim sSplit()        As String

If bInEXE Then On Local Error Resume Next

' Helo
lNumBytes = lNumBytes + Len(Mid$(psFrom, InStr(psFrom, "@") + 1)) + 7

' Mail From
lNumBytes = lNumBytes + Len(psFrom) + 12

' login authentication
If pbUseAuthentication Then
    lNumBytes = lNumBytes + 22 + Len(psFrom)
    If Len(gsUserName) > 0 Then lNumBytes = lNumBytes + (Len(gsUserName) * 4 \ 3)
    If Len(gsPassword) > 0 Then lNumBytes = lNumBytes + (Len(gsPassword) * 4 \ 3)
End If

' To: recipients
sSplit() = Split(psTo, ";")
For iCtr = LBound(sSplit) To UBound(sSplit)
    lNumBytes = lNumBytes + Len(sSplit(iCtr)) * 2           ' sent twice, RCPT & 'To:' header
    lNumBytes = lNumBytes + 15                              ' RCPT TO:
    If iCtr > LBound(sSplit) Then lNumBytes = lNumBytes + 6
Next

' To Display
sSplit() = Split(psToDisplay, ";")
For iCtr = LBound(sSplit) To UBound(sSplit)
    lNumBytes = lNumBytes + Len(sSplit(iCtr))
    lNumBytes = lNumBytes + 11
Next

' Cc: recipients
sSplit() = Split(psCc, ";")
For iCtr = LBound(sSplit) To UBound(sSplit)
    lNumBytes = lNumBytes + Len(sSplit(iCtr)) * 2           ' sent twice, RCPT & 'Cc:' header
    lNumBytes = lNumBytes + 15                              ' RCPT TO:
    If iCtr > LBound(sSplit) Then lNumBytes = lNumBytes + 6 ' header
Next

' Bcc: recipients
sSplit() = Split(psBcc, ";")
For iCtr = LBound(sSplit) To UBound(sSplit)
    lNumBytes = lNumBytes + Len(sSplit(iCtr))               ' RCPT & 'Bcc:' header
    lNumBytes = lNumBytes + 15                              ' RCPT TO:
    If iCtr > LBound(sSplit) Then lNumBytes = lNumBytes + 6 ' header
Next

' Cc Display
sSplit() = Split(psCcDisplay, ";")
For iCtr = LBound(sSplit) To UBound(sSplit)
    lNumBytes = lNumBytes + Len(sSplit(iCtr))
    lNumBytes = lNumBytes + 11
Next

' data
lNumBytes = lNumBytes + 6
' From:
If Len(psFromDisplay) Then lNumBytes = lNumBytes + Len(psFromDisplay) + 3
lNumBytes = lNumBytes + Len(psFrom) + 10
' ReplyTo
If Len(psReplyTo) Then lNumBytes = lNumBytes + Len(psReplyTo) + 12
' Subject
lNumBytes = lNumBytes + Len(psSubject) + 11
' Date
lNumBytes = lNumBytes + Len(Format(Now, "dddd dd-mmm-yyyy hh:nn:ss")) + 8

' Message body
lNumBytes = lNumBytes + Len(psMailMessage) + 4

' Attachments and their names
sSplit = Split(psAttachment, ";")
For iCtr = LBound(sSplit) To UBound(sSplit)
    If sSplit(iCtr) <> "" Then
        ' the length of the attached files
        lAttachSize = lAttachSize + FileLen(sSplit(iCtr))
        iNumAttach = iNumAttach + 1
        ' the lenght of tbe file names
        sAttach() = Split(sSplit(iCtr), "\")
        If pEncodeType = MIME_ENCODE Then
            lNumBytes = lNumBytes + (Len(sAttach(UBound(sAttach))) * 2)
        Else
            lNumBytes = lNumBytes + Len(sAttach(UBound(sAttach)))
        End If
    End If
Next

' MIME headers....
If pEncodeType = MIME_ENCODE Then
    lNumBytes = lNumBytes + 102
    If pbHtmlText = True And Len(psContentBase) > 0 Then lNumBytes = lNumBytes + 18 + Len(psContentBase)
    If iNumAttach > 0 Then lNumBytes = lNumBytes + 282
End If

' adjust the size of lAttachSize for the type encoding used
If iNumAttach > 0 Then
    If pEncodeType = MIME_ENCODE Then
        lAttachSize = lAttachSize * 4 \ 3 + 42
        lAttachSize = lAttachSize + (lAttachSize \ 76) * 2
    Else
        lAttachSize = lAttachSize * 4 \ 3
        lAttachSize = lAttachSize + (lAttachSize \ 60) * 3
    End If
End If
lNumBytes = lNumBytes + lAttachSize

' add MIME/UUEncode attachment headers
If pEncodeType = MIME_ENCODE Then
    lNumBytes = lNumBytes + (iNumAttach * 182)
Else
    lNumBytes = lNumBytes + (iNumAttach * 20)
End If

' close
lNumBytes = lNumBytes + 12

EstimateMailSize = lNumBytes

End Function

Friend Sub UpdateProgress(lProgress As Long)

RaiseEvent Progress(lProgress)

End Sub

Public Function GetContentType(ByVal strFile As String, Optional strDefault As String = "application/octet-stream") As String

'******************************************************************************
'
' Synopsis:     Get the Content Type from the Registry.
'
' Parameters:   strFile         - The filename to get the Content Type for
'               strDefault      - The default data to return if nothing is found

' Return:       The Content Type string

' Description:
'   The Content Type string for registered file extensions is located in the
'   system registry, in the root key HKEY_CLASSES_ROOT. Open the registry key
'   for the given file extension and read the 'Content Type' value. If the key
'   and/or value are not found, assign a default value of 'application/octet-stream'
'
'******************************************************************************

Dim hKey                As Long         ' key handle
Dim strBuff             As String       ' buffer for API to write to
Dim lBuffLen            As Long         ' lenght of API return string
Dim lRet                As Long         ' API return code
Dim lValueType          As Long         ' data type for retun value
Dim iPtr                As Integer      ' scratch pointer
Dim strValueName        As String       ' registry 'value name'
Dim strKeyName          As String       ' registry 'key name'

Const HKEY_CLASSES_ROOT = &H80000000
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1&

If bInEXE Then On Local Error Resume Next

' registry value name
strValueName = "Content Type"

' get the passed in key name. We only want
' the file extension here e.g. .exe, .doc, etc.
' if an extension is not found, assign default
' value and return
iPtr = InStrRev(strFile, ".")
If iPtr Then
    strKeyName = Mid$(strFile, iPtr)
Else
    GetContentType = strDefault
    Exit Function
End If

' open the Registry key, if key not found, return the defaut value
lRet = RegOpenKey(HKEY_CLASSES_ROOT, strKeyName, hKey)
If lRet <> ERROR_SUCCESS Then
    GetContentType = strDefault
    Exit Function
End If

' query the key value to get it's data type & length
lRet = RegQueryValueEx(hKey, strValueName, 0&, lValueType, 0&, lBuffLen)

If lValueType = REG_SZ Then
    ' create a buffer & call the API again
    strBuff = String(lBuffLen, " ")
    lRet = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal strBuff, lBuffLen)
    If lRet = ERROR_SUCCESS Then
        ' get the string value, drop the trailing '0'
        GetContentType = Left(strBuff, lBuffLen - 1)
    Else
        GetContentType = strDefault
    End If
Else
    ' value name not found or wrong data type
    GetContentType = strDefault
End If

' close the key
If hKey Then lRet = RegCloseKey(hKey)

End Function

