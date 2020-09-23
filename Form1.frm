VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtPlatformID 
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtType 
      Height          =   3255
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtComment 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtServerName 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtMinorNo 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtMajorNo 
      Height          =   405
      Left            =   5280
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Logon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1050
      Width           =   1335
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   570
      Width           =   1335
   End
   Begin VB.CommandButton Command0 
      Caption         =   "Check Domain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "type"
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Platform ID"
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Comment"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Server Name"
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Minor No"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Major No"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Domain Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SEC_E_OK = &H0
Const SEC_E_INSUFFICIENT_MEMORY = &H80090300
Const SEC_E_INVALID_HANDLE = &H80090301
Const SEC_E_UNSUPPORTED_FUNCTION = &H80090302
Const SEC_E_TARGET_UNKNOWN = &H80090303
Const SEC_E_INTERNAL_ERROR = &H80090304
Const SEC_E_SECPKG_NOT_FOUND = &H80090305
Const SEC_E_NOT_OWNER = &H80090306
Const SEC_E_CANNOT_INSTALL = &H80090307
Const SEC_E_INVALID_TOKEN = &H80090308
Const SEC_E_CANNOT_PACK = &H80090309
Const SEC_E_QOP_NOT_SUPPORTED = &H8009030A
Const SEC_E_NO_IMPERSONATION = &H8009030B
Const SEC_E_LOGON_DENIED = &H8009030C
Const SEC_E_UNKNOWN_CREDENTIALS = &H8009030D
Const SEC_E_NO_CREDENTIALS = &H8009030E
Const SEC_E_MESSAGE_ALTERED = &H8009030F
Const SEC_E_OUT_OF_SEQUENCE = &H80090310
Const SEC_E_NO_AUTHENTICATING_AUTHORITY = &H80090311
Const SEC_I_CONTINUE_NEEDED = &H90312
Const SEC_I_COMPLETE_NEEDED = &H90313
Const SEC_I_COMPLETE_AND_CONTINUE = &H90314
Const SEC_I_LOCAL_LOGON = &H90315
Const SEC_E_BAD_PKGID = &H80090316
Const SEC_E_CONTEXT_EXPIRED = &H80090317
Const SEC_E_INCOMPLETE_MESSAGE = &H80090318
Const SEC_E_INCOMPLETE_CREDENTIALS = &H80090320
Const SEC_E_BUFFER_TOO_SMALL = &H80090321
Const SEC_I_INCOMPLETE_CREDENTIALS = &H90320
Const SEC_I_RENEGOTIATE = &H90321
Const SEC_E_WRONG_PRINCIPAL = &H80090322

Const SECPKG_CRED_OUTBOUND = 2
Const SECPKG_CRED_INBOUND = 1

Const SEC_WINNT_AUTH_IDENTITY_ANSI = 1
Const SEC_WINNT_AUTH_IDENTITY_UNICODE = 2

Const SECURITY_NATIVE_DREP = 16
Const SECURITY_NETWORK_DREP = 0

Const SECBUFFER_TOKEN = 2


'typedef struct _SERVER_INFO {
'    DWORD     sv101_platform_id;
'    LPWSTR    sv101_name;
'    DWORD     sv101_version_major;
'    DWORD     sv101_version_minor;
'    DWORD     sv101_type;
'    LPWSTR    sv101_comment;
'} SERVER_INFO;
 
Private Type SERVER_INFO
    platform_id As String
    name As String
    version_major As String
    version_minor As String
    type As String
    comment As String
End Type
    
    
Private Declare Function NetServerGetInfoNT Lib "netapi32.dll" Alias "NetServerGetInfo" _
    (ByVal ServerName As Long, ByVal level As Long, ByVal bufptr As Long) As Long
Private Declare Function NetServerGetInfo9X Lib "svrapi.dll" Alias "NetServerGetInfo" _
    (ByVal ServerName As String, ByVal level As Integer, ByVal bufptr As Long, ByVal buflen As Integer, ByVal totalavail As Long) As Long

Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal bufptr As Long) As Long

Private Type SecPkgInfo
    fCapabilities As Long   'unsigned long          Capability bitmask
    wVersion As Integer     'unsigned short         Version of driver
    wRPCID As Integer       'unsigned short         ID for RPC Runtime
    cbMaxToken As Long      'unsigned long          Size of authentication token (max)
    name As Long            'SEC_CHAR SEC_FAR *     Text name
    comment As Long         'SEC_CHAR SEC_FAR *     Comment
End Type

Private Type SEC_WINNT_AUTH_IDENTITY
    User As Long            'unsigned char __RPC_FAR *
    UserLength As Long      'unsigned long
    domain As Long          'unsigned char __RPC_FAR *
    DomainLength As Long    'unsigned long
    password As Long        'unsigned char __RPC_FAR *
    PasswordLength As Long  'unsigned long
    Flags As Long           'unsigned long
End Type

Private Type DWORD
    dwLower As Long         'unsigned long
    dwUpper As Long         'unsigned long
End Type

Private Type SecBuffer
    cbBuffer As Long        'unsigned long      Size of the buffer, in bytes
    BufferType As Long      'unsigned long      Type of the buffer (below)
    pvBuffer As Long        'void SEC_FAR *     Pointer to the buffer
End Type

Private Type SecBufferDesc
    ulVersion As Long       'unsigned long      Version number
    cBuffers As Long        'unsigned long      Number of buffers
    pBuffers As Long        'PSecBuffer         Pointer to array of buffers
End Type

Private Declare Function AcquireCredentialsHandleNT Lib "security.dll" _
    Alias "AcquireCredentialsHandleA" ( _
    ByVal pszPrincipal As Long, ByVal pszPackage As String, _
    ByVal fCredentialUse As Long, ByVal pvLogonId As Long, _
    ByVal pAuthData As Long, ByVal pGetKeyFn As Long, _
    ByVal pvGetKeyArgument As Long, ByRef PCredHandle As DWORD, _
    ByRef ptsExpiry As DWORD) As Long
Private Declare Function AcquireCredentialsHandle9X Lib "secur32.dll" _
    Alias "AcquireCredentialsHandleA" ( _
    ByVal pszPrincipal As Long, ByVal pszPackage As String, _
    ByVal fCredentialUse As Long, ByVal pvLogonId As Long, _
    ByVal pAuthData As Long, ByVal pGetKeyFn As Long, _
    ByVal pvGetKeyArgument As Long, ByRef PCredHandle As DWORD, _
    ByRef ptsExpiry As DWORD) As Long

Private Declare Function InitializeSecurityContextNT Lib "security.dll" _
    Alias "InitializeSecurityContextA" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByVal pszTargetName As String, ByVal fContextReq As Long, _
    ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    ByVal pInput As Long, ByVal Reserved2 As Long, _
    ByRef phNewContext As DWORD, ByRef pOutput As SecBufferDesc, _
    ByRef pfContextAttr As Long, ByRef ptsExpiry As DWORD) As Long
Private Declare Function InitializeSecurityContext9X Lib "secur32.dll" _
    Alias "InitializeSecurityContextA" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByVal pszTargetName As String, ByVal fContextReq As Long, _
    ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    ByVal pInput As Long, ByVal Reserved2 As Long, _
    ByRef phNewContext As DWORD, ByRef pOutput As SecBufferDesc, _
    ByRef pfContextAttr As Long, ByRef ptsExpiry As DWORD) As Long

Private Declare Function AcceptSecurityContextNT Lib "security.dll" _
    Alias "AcceptSecurityContext" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByRef pInput As SecBufferDesc, ByVal fContextReq As Long, _
    ByVal TargetDataRep As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
Private Declare Function AcceptSecurityContext9X Lib "secur32.dll" _
    Alias "AcceptSecurityContext" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByRef pInput As SecBufferDesc, ByVal fContextReq As Long, _
    ByVal TargetDataRep As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
   
Private Declare Function CompleteAuthTokenNT Lib "security.dll" _
    Alias "CompleteAuthToken" _
    (ByRef phContext As DWORD, ByRef pToken As SecBufferDesc) As Long
Private Declare Function CompleteAuthToken9X Lib "secur32.dll" _
    Alias "CompleteAuthToken" _
    (ByRef phContext As DWORD, ByRef pToken As SecBufferDesc) As Long
    
Private Declare Function FreeContextBufferNT Lib "security.dll" _
    Alias "FreeContextBuffer" (ByVal pvContextBuffer As Long) As Long
Private Declare Function FreeContextBuffer9X Lib "secur32.dll" _
    Alias "FreeContextBuffer" (ByVal pvContextBuffer As Long) As Long

Private Declare Function FreeCredentialsHandleNT Lib "security.dll" _
    Alias "FreeCredentialsHandle" (ByRef hcred As DWORD) As Long
Private Declare Function FreeCredentialsHandle9X Lib "secur32.dll" _
    Alias "FreeCredentialsHandle" (ByRef hcred As DWORD) As Long

Private Declare Function DeleteSecurityContextNT Lib "security.dll" _
    Alias "DeleteSecurityContext" (ByRef hctxt As DWORD) As Long
Private Declare Function DeleteSecurityContext9X Lib "secur32.dll" _
    Alias "DeleteSecurityContext" (ByRef hctxt As DWORD) As Long

Private Declare Function InitSecurityInterfaceNT Lib "security.dll" _
    Alias "InitSecurityInterfaceA" () As Long
Private Declare Function InitSecurityInterface9X Lib "secur32.dll" _
    Alias "InitSecurityInterfaceA" () As Long

Private Declare Function QuerySecurityPackageInfoNT Lib "security.dll" _
    Alias "QuerySecurityPackageInfoA" _
    (ByVal pszPackageName As String, ByRef ppPackageInfo As Long) As Integer
Private Declare Function QuerySecurityPackageInfo9X Lib "secur32.dll" _
    Alias "QuerySecurityPackageInfoA" _
    (ByVal pszPackageName As String, ByRef ppPackageInfo As Long) As Integer

Private Declare Function ImpersonateSecurityContext Lib "security.dll" _
    (ByRef hctxt As DWORD) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, ByVal Source As Long, ByVal Length As Long)

Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Dim osvi As OSVERSIONINFO

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long


'    SEC_CHAR SEC_FAR * pszPrincipal,    // Name of principal
'    SEC_CHAR SEC_FAR * pszPackage,      // Name of package
'    unsigned long fCredentialUse,       // Flags indicating use
'    void SEC_FAR * pvLogonId,           // Pointer to logon ID
'    void SEC_FAR * pAuthData,           // Package specific data
'    SEC_GET_KEY_FN pGetKeyFn,           // Pointer to GetKey() func
'    void SEC_FAR * pvGetKeyArgument,    // Value to pass to GetKey()
'    PCredHandle phCredential,           // (out) Cred Handle
'    PTimeStamp ptsExpiry                // (out) Lifetime (optional)

Private Function AcquireCredentialsHandle(ByVal pszPrincipal As Long, _
    ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonId As Long, _
    ByVal pAuthData As Long, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
    ByRef PCredHandle As DWORD, ByRef ptsExpiry As DWORD) As Long
    
    If IsNT() Then
        AcquireCredentialsHandle = AcquireCredentialsHandleNT(pszPrincipal, _
            pszPackage, fCredentialUse, pvLogonId, pAuthData, pGetKeyFn, _
            pvGetKeyArgument, PCredHandle, ptsExpiry)
    Else
        AcquireCredentialsHandle = AcquireCredentialsHandle9X(pszPrincipal, _
            pszPackage, fCredentialUse, pvLogonId, pAuthData, pGetKeyFn, _
            pvGetKeyArgument, PCredHandle, ptsExpiry)
    End If
End Function
    
'    PCredHandle phCredential,               // Cred to base context
'    PCtxtHandle phContext,                  // Existing context (OPT)
'    SEC_CHAR SEC_FAR * pszTargetName,       // Name of target
'    unsigned long fContextReq,              // Context Requirements
'    unsigned long Reserved1,                // Reserved, MBZ
'    unsigned long TargetDataRep,            // Data rep of target
'    PSecBufferDesc pInput,                  // Input Buffers
'    unsigned long Reserved2,                // Reserved, MBZ
'    PCtxtHandle phNewContext,               // (out) New Context handle
'    PSecBufferDesc pOutput,                 // (inout) Output Buffers
'    unsigned long SEC_FAR * pfContextAttr,  // (out) Context attrs
'    PTimeStamp ptsExpiry                    // (out) Life span (OPT)

Private Function InitializeSecurityContext(ByRef phCredential As DWORD, _
    ByVal phContext As Long, ByVal pszTargetName As String, _
    ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    ByVal pInput As Long, ByVal Reserved2 As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
    
    If IsNT() Then
        InitializeSecurityContext = InitializeSecurityContextNT(phCredential, _
            phContext, pszTargetName, fContextReq, Reserved1, TargetDataRep, _
            pInput, Reserved2, phNewContext, pOutput, pfContextAttr, ptsExpiry)
    Else
        InitializeSecurityContext = InitializeSecurityContext9X(phCredential, _
            phContext, pszTargetName, fContextReq, Reserved1, TargetDataRep, _
            pInput, Reserved2, phNewContext, pOutput, pfContextAttr, ptsExpiry)
    End If
End Function
    
'    PCredHandle phCredential,               // Cred to base context
'    PCtxtHandle phContext,                  // Existing context (OPT)
'    PSecBufferDesc pInput,                  // Input buffer
'    unsigned long fContextReq,              // Context Requirements
'    unsigned long TargetDataRep,            // Target Data Rep
'    PCtxtHandle phNewContext,               // (out) New context handle
'    PSecBufferDesc pOutput,                 // (inout) Output buffers
'    unsigned long SEC_FAR * pfContextAttr,  // (out) Context attributes
'    PTimeStamp ptsExpiry                    // (out) Life span (OPT)

Private Function AcceptSecurityContext( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByRef pInput As SecBufferDesc, ByVal fContextReq As Long, _
    ByVal TargetDataRep As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
    
    If IsNT() Then
        AcceptSecurityContext = AcceptSecurityContextNT(phCredential, _
            phContext, pInput, fContextReq, TargetDataRep, phNewContext, _
            pOutput, pfContextAttr, ptsExpiry)
    Else
        AcceptSecurityContext = AcceptSecurityContext9X(phCredential, _
            phContext, pInput, fContextReq, TargetDataRep, phNewContext, _
            pOutput, pfContextAttr, ptsExpiry)
    End If
End Function
    
'    PCtxtHandle phContext,              // Context to complete
'    PSecBufferDesc pToken               // Token to complete

Private Function CompleteAuthToken _
    (ByRef phContext As DWORD, ByRef pToken As SecBufferDesc) As Long
    
    If IsNT() Then
        CompleteAuthToken = CompleteAuthTokenNT(phContext, pToken)
    Else
        CompleteAuthToken = CompleteAuthToken9X(phContext, pToken)
    End If
End Function
    
Private Function DeleteSecurityContext(ByRef hctxt As DWORD) As Long
    If IsNT() Then
        DeleteSecurityContext = DeleteSecurityContextNT(hctxt)
    Else
        DeleteSecurityContext = DeleteSecurityContext9X(hctxt)
    End If
End Function

Private Function FreeContextBuffer(ByVal pvContextBuffer As Long) As Long
    If IsNT() Then
        FreeContextBuffer = FreeContextBufferNT(pvContextBuffer)
    Else
        FreeContextBuffer = FreeContextBuffer9X(pvContextBuffer)
    End If
End Function

Private Function FreeCredentialsHandle(ByRef hcred As DWORD) As Long
    If IsNT() Then
        FreeCredentialsHandle = FreeCredentialsHandleNT(hcred)
    Else
        FreeCredentialsHandle = FreeCredentialsHandle9X(hcred)
    End If
End Function

Private Function InitSecurityInterface() As Long
    If IsNT() Then
        InitSecurityInterface = InitSecurityInterfaceNT()
    Else
        InitSecurityInterface = InitSecurityInterface9X()
    End If
End Function

Private Function QuerySecurityPackageInfo( _
    ByVal pszPackageName As String, ByRef ppPackageInfo As Long) As Integer
    If IsNT() Then
        QuerySecurityPackageInfo = _
            QuerySecurityPackageInfoNT(pszPackageName, ppPackageInfo)
    Else
        QuerySecurityPackageInfo = _
            QuerySecurityPackageInfo9X(pszPackageName, ppPackageInfo)
    End If
End Function


Function ByteToStr(b() As Byte, i As Integer) As String
    Dim s As String
    
    Do While b(i) <> 0
        s = s & Chr(b(i))
        i = i + 1
    Loop
    ByteToStr = s
End Function

Function GetServerType(i As Long) As String
    Dim s As String
    Dim nl As String
    
    nl = vbCr & vbLf
    
    If i And &H1 Then s = s & "WORKSTATION" & nl            '#define SV_TYPE_WORKSTATION         0x00000001
    If i And &H2 Then s = s & "SERVER" & nl                 '#define SV_TYPE_SERVER              0x00000002
    If i And &H4 Then s = s & "SQLSERVER" & nl              '#define SV_TYPE_SQLSERVER           0x00000004
    If i And &H8 Then s = s & "DOMAIN_CTRL" & nl            '#define SV_TYPE_DOMAIN_CTRL         0x00000008
    If i And &H10 Then s = s & "DOMAIN_BAKCTRL" & nl        '#define SV_TYPE_DOMAIN_BAKCTRL      0x00000010
    If i And &H20 Then s = s & "TIME_SOURCE" & nl           '#define SV_TYPE_TIME_SOURCE         0x00000020
    If i And &H40 Then s = s & "AFP" & nl                   '#define SV_TYPE_AFP                 0x00000040
    If i And &H80 Then s = s & "NOVELL" & nl                '#define SV_TYPE_NOVELL              0x00000080
    If i And &H100 Then s = s & "DOMAIN_MEMBER" & nl        '#define SV_TYPE_DOMAIN_MEMBER       0x00000100
    If i And &H200 Then s = s & "PRINTQ_SERVER" & nl        '#define SV_TYPE_PRINTQ_SERVER       0x00000200
    If i And &H400 Then s = s & "DIALIN_SERVER" & nl        '#define SV_TYPE_DIALIN_SERVER       0x00000400
    If i And &H800 Then s = s & "UNIX" & nl                 '#define SV_TYPE_SERVER_UNIX         0x00000800
    If i And &H1000 Then s = s & "NT" & nl                  '#define SV_TYPE_NT                  0x00001000
    If i And &H2000 Then s = s & "WFW" & nl                 '#define SV_TYPE_WFW                 0x00002000
    If i And &H4000 Then s = s & "SERVER_MFPN" & nl         '#define SV_TYPE_SERVER_MFPN         0x00004000
    If i And &H8000 Then s = s & "SERVER_NT" & nl           '#define SV_TYPE_SERVER_NT           0x00008000
    If i And &H10000 Then s = s & "POTENTIAL_BROWSER" & nl  '#define SV_TYPE_POTENTIAL_BROWSER   0x00010000
    If i And &H20000 Then s = s & "BACKUP_BROWSER" & nl     '#define SV_TYPE_BACKUP_BROWSER      0x00020000
    If i And &H40000 Then s = s & "MASTER_BROWSER" & nl     '#define SV_TYPE_MASTER_BROWSER      0x00040000
    If i And &H80000 Then s = s & "DOMAIN_MASTER" & nl      '#define SV_TYPE_DOMAIN_MASTER       0x00080000
    If i And &H100000 Then s = s & "SERVER_OSF" & nl        '#define SV_TYPE_SERVER_OSF          0x00100000
    If i And &H200000 Then s = s & "SERVER_VMS" & nl        '#define SV_TYPE_SERVER_VMS          0x00200000
    If i And &H400000 Then s = s & "WINDOWS" & nl           '#define SV_TYPE_WINDOWS             0x00400000  /* Windows95 and above */
    If i And &H800000 Then s = s & "DFS" & nl               '#define SV_TYPE_DFS                 0x00800000  /* Root of a DFS tree */
    If i And &H1000000 Then s = s & "CLUSTER_NT" & nl       '#define SV_TYPE_CLUSTER_NT          0x01000000  /* NT Cluster */
    If i And &H10000000 Then s = s & "DCE" & nl             '#define SV_TYPE_DCE                 0x10000000  /* IBM DSS (Directory and Security Services) or equivalent */
    If i And &H20000000 Then s = s & "ALTERNATE_XPORT" & nl '#define SV_TYPE_ALTERNATE_XPORT     0x20000000  /* return list for alternate transport */
    If i And &H40000000 Then s = s & "LOCAL_LIST_ONLY" & nl '#define SV_TYPE_LOCAL_LIST_ONLY     0x40000000  /* Return local list only */
    If i And &H80000000 Then s = s & "DOMAIN_ENUM" & nl     '#define SV_TYPE_DOMAIN_ENUM         0x80000000

    GetServerType = s
End Function


'5      ERROR_ACCESS_DENIED         ' win95 win98 haven't logon
'53     ERROR_BAD_NETPATH           ' server does not exist
'87     ERROR_INVALID_PARAMETER
'124    ERROR_INVALID_LEVEL
'234    ERROR_MORE_DATA             ' buffer not big enough, correct size returned

'2102   NERR_NetNotStarted
'2113   NERR_BufTooSmall            ' buffer too small for fixed length structure
'2114   NERR_ServerNotStarted
'2351   NERR_InvalidComputer        ' invalid computer name e.g. \\\\abcd

Private Function NetServerGetInfo(ServerName As String, SvrInfo As SERVER_INFO) As Long
    Dim status As Long
    Dim i As Long
    Dim s As String
    
    SvrInfo.platform_id = ""
    SvrInfo.name = ""
    SvrInfo.version_major = ""
    SvrInfo.version_minor = ""
    SvrInfo.type = ""
    SvrInfo.comment = ""
    
    If IsNT() Then
        Dim b As Long
        
        status = NetServerGetInfoNT(StrPtr(ServerName & Chr(0)), 101, VarPtr(b))
        If status = 0 Then
            CopyMemory i, b, 4
            SvrInfo.platform_id = i
            CopyMemory i, b + 8, 4
            SvrInfo.version_major = i
            CopyMemory i, b + 12, 4
            SvrInfo.version_minor = i
            CopyMemory i, b + 16, 4
            SvrInfo.type = GetServerType(i)
            i = NetApiBufferFree(b)

            '??? free buffer here
        End If
    Else
        Dim buf(100) As Byte
        Dim cb As Integer
        
        status = NetServerGetInfo9X("\\" & ServerName, 1, VarPtr(buf(0)), 100, VarPtr(cb))
        If status = 0 Then
            SvrInfo.name = ByteToStr(buf, 0)
            SvrInfo.version_major = buf(16) And &HF
            SvrInfo.version_minor = buf(17)
            CopyMemory i, VarPtr(buf(0)) + 18, 4
            SvrInfo.type = GetServerType(i)
            CopyMemory i, VarPtr(buf(0)) + 22, 4
            SvrInfo.comment = ByteToStr(buf, i - VarPtr(buf(0)))
        End If
    End If
    
    NetServerGetInfo = status
End Function

Private Function IsNT() As Boolean
    IsNT = (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Private Function GetMsg(i As Long) As String
    Select Case i
        Case SEC_E_OK
                GetMsg = "OK"
        Case SEC_E_INSUFFICIENT_MEMORY
                GetMsg = "E: INSUFFICIENT_MEMORY"
        Case SEC_E_INVALID_HANDLE
                GetMsg = "E: INVALID_HANDLE"
        Case SEC_E_UNSUPPORTED_FUNCTION
                GetMsg = "E: UNSUPPORTED_FUNCTION"
        Case SEC_E_TARGET_UNKNOWN
                GetMsg = "E: TARGET_UNKNOWN"
        Case SEC_E_INTERNAL_ERROR
                GetMsg = "E: INTERNAL_ERROR"
        Case SEC_E_SECPKG_NOT_FOUND
                GetMsg = "E: SECPKG_NOT_FOUND"
        Case SEC_E_NOT_OWNER
                GetMsg = "E: NOT_OWNER"
        Case SEC_E_CANNOT_INSTALL
                GetMsg = "E: CANNOT_INSTALL"
        Case SEC_E_INVALID_TOKEN
                GetMsg = "E: INVALID_TOKEN"
        Case SEC_E_CANNOT_PACK
                GetMsg = "E: CANNOT_PACK"
        Case SEC_E_QOP_NOT_SUPPORTED
                GetMsg = "E: QOP_NOT_SUPPORTED"
        Case SEC_E_NO_IMPERSONATION
                GetMsg = "E: NO_IMPERSONATION"
        Case SEC_E_LOGON_DENIED
                GetMsg = "E: LOGON_DENIED"
        Case SEC_E_UNKNOWN_CREDENTIALS
                GetMsg = "E: UNKNOWN_CREDENTIALS"
        Case SEC_E_NO_CREDENTIALS
                GetMsg = "E: NO_CREDENTIALS"
        Case SEC_E_MESSAGE_ALTERED
                GetMsg = "E: MESSAGE_ALTERED"
        Case SEC_E_OUT_OF_SEQUENCE
                GetMsg = "E: OUT_OF_SEQUENCE"
        Case SEC_E_NO_AUTHENTICATING_AUTHORITY
                GetMsg = "E: NO_AUTHENTICATING_AUTHORITY"
        Case SEC_I_CONTINUE_NEEDED
                GetMsg = "I: CONTINUE_NEEDED"
        Case SEC_I_COMPLETE_NEEDED
                GetMsg = "I: COMPLETE_NEEDED"
        Case SEC_I_COMPLETE_AND_CONTINUE
                GetMsg = "I: COMPLETE_AND_CONTINUE"
        Case SEC_I_LOCAL_LOGON
                GetMsg = "I: LOCAL_LOGON"
        Case SEC_E_BAD_PKGID
                GetMsg = "E: BAD_PKGID"
        Case SEC_E_CONTEXT_EXPIRED
                GetMsg = "E: CONTEXT_EXPIRED"
        Case SEC_E_INCOMPLETE_MESSAGE
                GetMsg = "E: INCOMPLETE_MESSAGE"
        Case SEC_E_INCOMPLETE_CREDENTIALS
                GetMsg = "E: INCOMPLETE_CREDENTIALS"
        Case SEC_E_BUFFER_TOO_SMALL
                GetMsg = "E: BUFFER_TOO_SMALL"
        Case SEC_I_INCOMPLETE_CREDENTIALS
                GetMsg = "I: INCOMPLETE_CREDENTIALS"
        Case SEC_I_RENEGOTIATE
                GetMsg = "I: RENEGOTIATE"
        Case SEC_E_WRONG_PRINCIPAL
                GetMsg = "E: WRONG_PRINCIPAL"
        Case Else
                GetMsg = "Unknown Error"
    End Select
End Function

Private Sub StrToByte(s As String, b() As Byte)
    Dim i As Integer
    
    For i = 0 To Len(s) - 1
        b(i) = Asc(Mid(s, i + 1, 1))
    Next i
    b(i) = 0
End Sub

Public Function SSPILogonUser(User As String, password As String, _
                                domain As String, errmsg) As Boolean
    Dim i As Long
    Dim ppkgInfo As Long
    Dim hcred As DWORD
    Dim AuthIdentity As SEC_WINNT_AUTH_IDENTITY
    Dim UserBuf(20) As Byte
    Dim DomainBuf(20) As Byte
    Dim PasswordBuf(20) As Byte
    Dim hctxt As DWORD
    Dim OutBuffDesc As SecBufferDesc
    Dim OutSecBuff As SecBuffer
    Dim ContextAttributes As Long
    Dim LifeTime As DWORD
    Dim cbMaxMessage As Long
  
    AuthIdentity.domain = VarPtr(DomainBuf(0))
    AuthIdentity.DomainLength = Len(domain)
    AuthIdentity.password = VarPtr(PasswordBuf(0))
    AuthIdentity.PasswordLength = Len(password)
    AuthIdentity.User = VarPtr(UserBuf(0))
    AuthIdentity.UserLength = Len(User)
    AuthIdentity.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI
     
    StrToByte domain, DomainBuf
    StrToByte User, UserBuf
    StrToByte password, PasswordBuf
    
    i = InitSecurityInterface
    If i < 0 Then GoTo error
    
    i = QuerySecurityPackageInfo("NTLM", ppkgInfo)
    If i < 0 Then GoTo error
    
    i = FreeContextBuffer(ppkgInfo)
    If i < 0 Then GoTo error
    
    CopyMemory cbMaxMessage, ppkgInfo + 8, 4

    '----------------------------------- negotiate, Client Initialization
    
    ReDim pOut(cbMaxMessage) As Byte
    
    OutSecBuff.cbBuffer = cbMaxMessage
    OutSecBuff.pvBuffer = VarPtr(pOut(0))
    OutSecBuff.BufferType = SECBUFFER_TOKEN
    OutBuffDesc.ulVersion = 0
    OutBuffDesc.cBuffers = 1
    OutBuffDesc.pBuffers = VarPtr(OutSecBuff)

    i = AcquireCredentialsHandle(0, "NTLM", SECPKG_CRED_OUTBOUND, 0, _
            VarPtr(AuthIdentity), 0, 0, hcred, LifeTime)
    If i < 0 Then GoTo error
         
    i = InitializeSecurityContext(hcred, 0, "\\AuthSamp", 0, 0, _
            SECURITY_NATIVE_DREP, 0, 0, hctxt, OutBuffDesc, _
            ContextAttributes, LifeTime)
    If i < 0 Then GoTo error
            
    If i = SEC_I_COMPLETE_NEEDED Or i = SEC_I_COMPLETE_AND_CONTINUE Then
        i = CompleteAuthToken(hctxt, OutBuffDesc)
        MsgBox ("COMPLETE should not be required for NTLM.")
    End If
    
    '----------------------------------- challenge
    
    Dim hCred2 As DWORD
    Dim hctxt2 As DWORD
    Dim InBuffDesc2 As SecBufferDesc
    Dim InSecBuff2 As SecBuffer
    Dim OutBuffDesc2 As SecBufferDesc
    Dim OutSecBuff2 As SecBuffer
    ReDim pout2(cbMaxMessage) As Byte
    
    i = AcquireCredentialsHandle(0, "NTLM", SECPKG_CRED_INBOUND, 0, _
            0, 0, 0, hCred2, LifeTime)  ' Server initialization
    If i < 0 Then GoTo error
         
    InSecBuff2.cbBuffer = OutSecBuff.cbBuffer
    InSecBuff2.pvBuffer = OutSecBuff.pvBuffer
    InSecBuff2.BufferType = SECBUFFER_TOKEN
    InBuffDesc2.ulVersion = 0
    InBuffDesc2.cBuffers = 1
    InBuffDesc2.pBuffers = VarPtr(InSecBuff2)
    
    OutSecBuff2.cbBuffer = cbMaxMessage
    OutSecBuff2.pvBuffer = VarPtr(pout2(0))
    OutSecBuff2.BufferType = SECBUFFER_TOKEN
    OutBuffDesc2.ulVersion = 0
    OutBuffDesc2.cBuffers = 1
    OutBuffDesc2.pBuffers = VarPtr(OutSecBuff2)
    
    'server transmits the output security buffer and length back to the client.
    
    i = AcceptSecurityContext(hCred2, 0, InBuffDesc2, 0, SECURITY_NATIVE_DREP, _
            hctxt2, OutBuffDesc2, ContextAttributes, LifeTime)
    If i < 0 Then GoTo error
    
    '----------------------------------- authenticate

    Dim InSecBuff As SecBuffer
    Dim InBuffDesc As SecBufferDesc
    
    InSecBuff.cbBuffer = OutSecBuff2.cbBuffer
    InSecBuff.pvBuffer = OutSecBuff2.pvBuffer
    InSecBuff.BufferType = SECBUFFER_TOKEN
    InBuffDesc.ulVersion = 0
    InBuffDesc.cBuffers = 1
    InBuffDesc.pBuffers = VarPtr(InSecBuff)
    
    OutSecBuff.cbBuffer = cbMaxMessage
    
    ' the client transmits the output security buffer and buffer length to the server,
    ' as it did after the first call to InitializeSecurityContext.
    ' The client has now finished setting up the security context.
    
    i = InitializeSecurityContext(hcred, VarPtr(hctxt), "\\AuthSamp", 0, 0, _
            SECURITY_NATIVE_DREP, VarPtr(InBuffDesc), 0, hctxt, OutBuffDesc, _
            ContextAttributes, LifeTime)
    If i < 0 Then GoTo error

    '----------------------------------- authenticate
    
    InSecBuff2.cbBuffer = OutSecBuff.cbBuffer
    InSecBuff2.pvBuffer = OutSecBuff.pvBuffer
    
    OutSecBuff2.cbBuffer = cbMaxMessage
    
    ' The server makes the final call to AcceptSecurityContext
    
    i = AcceptSecurityContext(hCred2, VarPtr(hctxt2), InBuffDesc2, 0, _
            SECURITY_NATIVE_DREP, hctxt2, OutBuffDesc2, ContextAttributes, LifeTime)
                                         '^ no output security token, can use null instead of OutSecBuff2
    If i < 0 Then GoTo error
    
    ' the client was successfully authenticated
    
'    i = ImpersonateSecurityContext(hctxt2)
'    If i < 0 Then GoTo error
    
    
    
    i = DeleteSecurityContext(hctxt)
    If i < 0 Then GoTo error
    i = DeleteSecurityContext(hctxt2)
    If i < 0 Then GoTo error
    i = FreeCredentialsHandle(hcred)
    If i < 0 Then GoTo error
    i = FreeCredentialsHandle(hCred2)
    If i < 0 Then GoTo error
    
    SSPILogonUser = True
    Exit Function

error:

    errmsg = GetMsg(i)
    SSPILogonUser = False
    
End Function

Private Sub Command0_Click()
    Dim s As String
    Dim a As Long
    
    Dim SvrInfo As SERVER_INFO
    Dim bufadd As Long
    
    Dim i As Long
    
    
    osvi.dwOSVersionInfoSize = 148
    i = GetVersionEx(osvi)
    
    Screen.MousePointer = vbHourglass
'    DoCmd.Hourglass True
    
    i = NetServerGetInfo(txtInput, SvrInfo)
    
    txtMajorNo = SvrInfo.version_major
    txtMinorNo = SvrInfo.version_minor
    txtServerName = SvrInfo.name
    txtComment = SvrInfo.comment
    txtType = SvrInfo.type
    txtPlatformID = SvrInfo.platform_id
    txtStatus = i
    
    Screen.MousePointer = vbNormal
    
'    DoCmd.Hourglass False
    
    
End Sub


'Q125700

Private Sub Command2_Click()
    Dim errmsg As String
    Dim status As Integer
    Dim i As Integer
    
'    DoCmd.Hourglass True
    
    osvi.dwOSVersionInfoSize = 148
    i = GetVersionEx(osvi)
    
    status = SSPILogonUser(txtName, txtPassword, txtDomain, errmsg)
       
    If status Then
        MsgBox "Logon successful! You can do whatever you like."
    Else
        If errmsg = "E: NO_AUTHENTICATING_AUTHORITY" Then
            MsgBox "No authenticating authority. You need to enable user level access in Win95/Win98. " & _
                "It can be set in control panel/network/access control/" & _
                "user level access control."
            Exit Sub
        End If
        
        If errmsg = "E: LOGON_DENIED" Then
            MsgBox "Incorrect password, user name or domain name! Please re-enter!"
            Exit Sub
        End If
       
        If errmsg = "E: UNSUPPORTED_FUNCTION" Then
            MsgBox "If you are running Win95/Win98 but haven't " & _
            "installed DCOM95/DCOM98, please do so and try again."
            Exit Sub
        End If

        MsgBox "Please request help from the administrator. " & errmsg
    End If
    
'    DoCmd.Hourglass False

End Sub



