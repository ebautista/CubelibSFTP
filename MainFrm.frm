VERSION 5.00
Object = "{E532970A-FEEB-4A38-A1BB-4E462DDCA8B9}#8.0#0"; "SFTPBBoxCli8.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2A486285-0C52-4069-8D0C-4E5EB6433DE0}#8.0#0"; "BaseBBox8.dll"
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sftp Demo"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin SFTPBBoxCli8.ElSimpleSftpClientX ElSimpleSftpClientX 
      Left            =   6480
      Top             =   3240
      NewLineConvention=   $"MainFrm.frx":0000
      ClientUserName  =   ""
      ClientHostName  =   ""
      ForceCompression=   0   'False
      CompressionLevel=   6
      SoftwareName    =   "SecureBlackbox.8"
      Username        =   ""
      Password        =   ""
      Address         =   ""
      Port            =   22
      UseInternalSocket=   -1  'True
      SocketTimeout   =   0
      UseSocks        =   0   'False
      SocksServer     =   ""
      SocksPort       =   1080
      SocksUserCode   =   ""
      SocksPassword   =   ""
      SocksVersion    =   1
      SocksResolveAddress=   0   'False
      SocksAuthentication=   0
      UseWebTunneling =   0   'False
      WebTunnelAddress=   ""
      WebTunnelPort   =   3128
      WebTunnelAuthentication=   0
      WebTunnelUserId =   ""
      WebTunnelPassword=   ""
      SFTPBufferSize  =   131072
      PipelineLength  =   32
      DownloadBlockSize=   8192
      UploadBlockSize =   32768
      CurrentOperationCancel=   0   'False
      ASCIIMode       =   0   'False
      LocalNewLineConvention=   $"MainFrm.frx":0006
      DefaultWindowSize=   2048000
      MinWindowSize   =   2048
      SSHAuthOrder    =   1
      AutoAdjustCiphers=   -1  'True
      AutoAdjustTransferBlock=   -1  'True
      LocalAddress    =   ""
      LocalPort       =   0
      UseUTF8         =   -1  'True
      OperationErrorHandling=   0
      RequestPasswordChange=   0   'False
      AuthAttempts    =   1
      CertAuthMode    =   1
      IncomingSpeedLimit=   0
      OutgoingSpeedLimit=   0
      KeepAlivePeriod =   0
      SocksUseIPv6    =   0   'False
      UseIPv6         =   0   'False
      AdjustFileTimes =   0   'False
      FIPSMode        =   0   'False
      GSSHostName     =   ""
      GSSDelegateCredentials=   0   'False
      UseTruncateFlagOnUpload=   -1  'True
      TreatZeroSizeAsUndefined=   -1  'True
      UseUTF8OnV3     =   0   'False
   End
   Begin VB.Timer TimeOutCounter 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   1200
   End
   Begin BaseBBox8.ElSBLicenseManagerX ElSBLicenseManagerX 
      Left            =   6480
      Top             =   2520
   End
   Begin ComctlLib.ListView LogListView 
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Timestamp"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Event"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.CommandButton btnUpdateInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6120
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Refresh"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnPutFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4920
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Upload"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnGetFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3720
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":0690
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Download selected"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnDelete 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2520
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":09D2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete selected"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnRename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":0D14
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Rename selected"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnMkDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Picture         =   "MainFrm.frx":1056
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Make directory"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Edit4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   7215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Permissions"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox editPath 
      Height          =   300
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "."
      Top             =   1800
      Width           =   7215
   End
   Begin MSComDlg.CommonDialog OpenDialog1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog SaveDialog1 
      Left            =   5640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection properties"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox editPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox editUserName 
         Height          =   300
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox editHost 
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Text            =   "192.168.0.1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "User name"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public mvarCSFTP As CSFTPUploadDownload

Private Const FILE_BLOCK_SIZE = 4096
Private Const STATE_OPEN_DIRECTORY_SENT = 1
Private Const STATE_READ_DIRECTORY_SENT = 2
Private Const STATE_CHANGE_DIR = 3
Private Const STATE_MAKE_DIR = 4
Private Const STATE_RENAME = 5
Private Const STATE_REMOVE = 6
Private Const STATE_DOWNLOAD_OPEN = 7
Private Const STATE_DOWNLOAD_RECEIVE = 8
Private Const STATE_UPLOAD_OPEN = 9
Private Const STATE_UPLOAD_SEND = 10
Private Const STATE_CLOSE_HANDLE = 11

Private Const MESSAGE_PREFIX = "SFTP SERVER: "

Private m_strCurrentHandle As String
Private m_strCurrentDir As String
Private m_strRelDir As String
Private m_strCurrentFile As String

Public m_colCurrentFileList As Collection

Private m_lngState As Long
Private m_lngCurrentFileOffset As Long
Private m_lngCurrentFileSize As Long
Private m_lngCurrentFile As Integer

Private m_blnClientDataAvailable As Boolean
Public m_blnDirectoryReadFinished As Boolean
Public m_blnFileUploaded As Boolean
Public m_blnFileDownloaded As Boolean
Public m_blnFileDeleted As Boolean

Public m_blnHasError As Boolean

Private m_blnAuthenticationFailed As Boolean

Public Sub LoadForm(ByRef CSFTP As CSFTPUploadDownload)
    Load Me
    
    Set mvarCSFTP = CSFTP
End Sub

Private Sub Form_Load()
    
    Set m_colCurrentFileList = New Collection
  
    ElSBLicenseManagerX.SetLicenseKey ("90D40DF1DDFEEC8F659575583B2619AE021FB2D4DCAB1F82E429A554D48A77E8FCD05FBB554D713297DEDEEFE375828F822A11D20B2B7A2671A844123C45D8176FA1898EECA5F4401ACAF8999496A60AD7BCEE80B4C2764E534F7215FF42C83FE42CCA6414394BE80394EFE3A67C6DF36494EB440BC16BB62C32194B4AB2E8FAEAAA11F99004851D8DF675F2C33B3F70C4811A487E59D4023E7C31950A4AC948CAEE628EC8134DAFE72F314D88BDCE932328F8D75AD620E169D90348B51FBFD651779357026431BD0235B2F8FBB10F880EBCEDEEE714E88E644082B878E4297917F1336A5DA52446736870F2AAF8C070EEB5A89519583E9B161D50D21B1B1846")
    
    ElSimpleSftpClientX.EnableVersion SB_SFTP_VERSION_3
    ElSimpleSftpClientX.EnableAuthenticationType SSH_AUTH_TYPE_PASSWORD
    ElSimpleSftpClientX.EnableAuthenticationType SSH_AUTH_TYPE_KEYBOARD
End Sub

Private Sub ElSimpleSftpClientX_OnError(ByVal ErrorCode As Long)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "SSH error " & ErrorCode
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "If you have ensured that all connection parameters are correct and you still can't connect,"
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "please contact CANDS support."
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Remember to provide details about the error that happened."
    
    If Len(ElSimpleSftpClientX.ServerSoftwareName) > 0 Then
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Server software identified itself as: " & ElSimpleSftpClientX.ServerSoftwareName
    End If
    
    Call mvarCSFTP.DisconnectFromServer(False)
End Sub

Private Sub ElSimpleSftpClientX_OnAuthenticationStart(ByVal SupportedAuths As Long)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "TCP connection opened..."
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authenticating Username: " & mvarCSFTP.UserName & ", Password: " & mvarCSFTP.Password & " using SSHClient..."
    
    m_blnAuthenticationFailed = False
End Sub

Private Sub ElSimpleSftpClientX_OnAuthenticationFailed(ByVal AuthenticationType As SSHBBoxCli8.TxSSHAuthenticationType)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authentication type " & AuthenticationType & " failed."
    m_blnAuthenticationFailed = True
    Call mvarCSFTP.DisconnectFromServer(False)
End Sub

Private Sub ElSimpleSftpClientX_OnAuthenticationSuccess()
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authentication succeeded..."
End Sub

Private Sub ElSimpleSftpClientX_OnSend(ByVal Data As Variant)
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Sending Data ..."
End Sub


'Private Sub scktClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & " Number: " & Number & ", Description: " & Description & ", Scode: " & Scode & ", Source: " & Source
'End Sub
'
'Private Sub SSHClient_OnError(ByVal ErrorCode As Long)
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "SSH error " & ErrorCode
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "If you have ensured that all connection parameters are correct and you still can't connect,"
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "please contact CANDS support."
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Remember to provide details about the error that happened."
'
'    If Len(SSHClient.ServerSoftwareName) > 0 Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Server software identified itself as: " & SSHClient.ServerSoftwareName
'    End If
'
'    Call mvarCSFTP.DisconnectFromServer(False)
'End Sub

'Private Sub SSHClient_OnSend(ByVal Data As Variant)
'  If Not m_blnAuthenticationFailed Then
'    Call scktClient.SendData(Data)
'  End If
'End Sub

'Private Sub SSHClient_OnReceive(Data As Variant, ByVal MaxSize As Long)
'    If scktClient.State <> sckConnected Then
'        m_blnClientDataAvailable = False
'        Exit Sub
'    End If
'
'    Call scktClient.GetData(Data, vbArray + vbByte, MaxSize)
'
'    If ArrSize(Data) = 0 Then
'        m_blnClientDataAvailable = False
'    End If
'End Sub

'Private Sub SSHClient_OnOpenConnection()
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "SSH Connection started..."
'End Sub
'
'Private Sub SSHClient_OnCloseConnection()
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "SSH Connection closed..."
'End Sub

'Private Sub SSHClient_OnAuthenticationSuccess()
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authentication succeeded..."
'End Sub

'Private Sub SSHClient_OnAuthenticationFailed(ByVal AuthenticationType As Long)
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authentication type " & AuthenticationType & " failed."
'    m_blnAuthenticationFailed = True
'    Call mvarCSFTP.DisconnectFromServer(False)
'End Sub

'Private Sub SftpClient_OnOpenConnection()
'    Dim strVersion As String
'
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Sftp connection started..."
'
'    If SftpClient.Version = SB_SFTP_VERSION_3 Then
'        strVersion = "3"
'    ElseIf SftpClient.Version = SB_SFTP_VERSION_4 Then
'        strVersion = "4"
'    Else
'        strVersion = "unknown"
'    End If
'
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Sftp version is " & strVersion & "..."
'
'    m_strCurrentDir = "."
'    Call BuildFileList(".")
'End Sub

'Private Sub SftpClient_OnCloseConnection()
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Sftp connection closed..."
'End Sub
'
'Private Sub SftpClient_OnOpenFile(ByVal Handle As String)
'    Select Case m_lngState
'        Case STATE_OPEN_DIRECTORY_SENT
'            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Root directory opened..."
'            m_strCurrentHandle = Handle
'            Call SftpClient.ReadDirectory(m_strCurrentHandle)
'            m_lngState = STATE_READ_DIRECTORY_SENT
'
'        Case STATE_CHANGE_DIR
'            Call SftpClient.CloseHandle(Handle)
'
'        Case STATE_DOWNLOAD_OPEN
'            m_strCurrentHandle = Handle
'            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Start of SFTP Download..."
'            Call SftpClient.Read(Handle, m_lngCurrentFileOffset, 0, FILE_BLOCK_SIZE)
'            m_lngState = STATE_DOWNLOAD_RECEIVE
'
'        Case STATE_UPLOAD_OPEN
'            m_strCurrentHandle = Handle
'            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Start of SFTP Upload..."
'            Call WriteNextBlockToFile
'            m_lngState = STATE_UPLOAD_SEND
'
'    End Select
'End Sub
'
'Private Sub SftpClient_OnError(ByVal ErrorCode As Long, ByVal Comment As String)
'    If (m_lngState = STATE_READ_DIRECTORY_SENT) And (ErrorCode = SFTP_ERROR_EOF) Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Directory List Received..."
'        Call CloseCurrentHandle
'        m_blnDirectoryReadFinished = True
'    ElseIf (m_lngState = STATE_DOWNLOAD_RECEIVE) And (ErrorCode = SFTP_ERROR_EOF) Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Message Received..."
'        Close #m_lngCurrentFile
'        Call CloseCurrentHandle
'        m_blnFileDownloaded = True
'    Else
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Log in SftpClient_OnError: Error Code (" & ErrorCode & "), Description (" & Comment & ")."
'        Call mvarCSFTP.DisconnectFromServer(False)
'    End If
'
'    'mvarCSFTP.TraceText = MESSAGE_PREFIX & "TEST LOGS - " & Comment
'End Sub
'
'Private Sub SftpClient_OnSuccess(ByVal Comment As String)
'    Select Case m_lngState
'        Case STATE_REMOVE
'            m_blnFileDeleted = True
'
'        Case STATE_UPLOAD_SEND
'            Call WriteNextBlockToFile
'
'        Case STATE_CLOSE_HANDLE
'            Close #m_lngCurrentFile
'            Call BuildFileList(m_strCurrentDir)
'
'    End Select
'
'    If Len(Comment) > 0 Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "TEST LOGS - " & Comment
'    End If
'End Sub
'
'Private Sub SftpClient_OnDirectoryListing(ByVal Listing As Variant)
'    Dim lngIdx As Long
'    Dim FileInfo1 As IElSftpFileInfoX
'    Dim FI2 As IElSftpFileInfoX
'
'    If m_lngState = STATE_READ_DIRECTORY_SENT Then
'        For lngIdx = LBound(Listing) To UBound(Listing)
'            Set FileInfo1 = New ElSftpFileInfoX
'
'            On Error Resume Next
'            Set FI2 = Listing(lngIdx)
'            On Error GoTo 0
'
'            Call FI2.CopyTo(FileInfo1)
'
'            Call m_colCurrentFileList.Add(FileInfo1)
'            Set FileInfo1 = Nothing
'        Next
'
'        Call SftpClient.ReadDirectory(m_strCurrentHandle)
'    End If
'End Sub
'
'Private Sub SftpClient_OnData(ByVal Buffer As Variant)
'  Dim Size As Long
'
'    Size = ArrSize(Buffer)
'
'    If m_lngState = STATE_DOWNLOAD_RECEIVE Then
'        Call BlockWrite(m_lngCurrentFile, Buffer)
'        m_lngCurrentFileOffset = m_lngCurrentFileOffset + Size
'
'        If m_lngCurrentFileOffset >= m_lngCurrentFileSize Then
'            mvarCSFTP.TraceText = MESSAGE_PREFIX & "File received..."
'            Close #m_lngCurrentFile
'
'            Call CloseCurrentHandle
'            m_blnFileDownloaded = True
'        Else
'            Call SftpClient.Read(m_strCurrentHandle, m_lngCurrentFileOffset, 0, FILE_BLOCK_SIZE)
'        End If
'    End If
'End Sub
'
'Private Sub SftpClient_OnAbsolutePath(ByVal Path As String)
'    m_strCurrentDir = Path
'    Call BuildFileList(m_strCurrentDir)
'    editPath.Text = Path
'End Sub
'
'Private Sub SftpClient_OnFileAttributes(ByVal Attributes As IElSftpFileAttributesX)
'    Dim info As IElSftpFileInfoX
'    Dim Row As Integer
'    Dim SizeLo As Long, SizeHi As Long
'    Dim Attr As IElSftpFileAttributesX
'
'    If ListView1.SelectedItem Is Nothing Then
'        Exit Sub
'    End If
'
'    Row = ListView1.SelectedItem.Index
'    Set Attr = Attributes
'
'    Set info = m_colCurrentFileList(Row)
'
'    If Attr.IsAttributeIncluded(SB_SFTP_ATTR_SIZE) Then
'        Call info.Attributes.GetSize(SizeLo, SizeHi)
'        Call Attr.SetSize(SizeLo, SizeHi)
'    End If
'
'    If Attr.IsAttributeIncluded(SB_SFTP_ATTR_UID) Then
'        info.Attributes.UID = Attr.UID
'    End If
'
'    If Attr.IsAttributeIncluded(SB_SFTP_ATTR_GID) Then
'        info.Attributes.GID = Attr.GID
'    End If
'
'    If Attr.IsAttributeIncluded(SB_SFTP_ATTR_ATIME) Then
'        info.Attributes.ATime = Attr.ATime
'    End If
'
'    If Attr.IsAttributeIncluded(SB_SFTP_ATTR_MTIME) Then
'        info.Attributes.MTime = Attr.MTime
'    End If
'
'    If Attr.IsAttributeIncluded(SB_SFTP_ATTR_EXTENDEDCOUNT) Then
'        info.Attributes.ExtendedCount = Attr.ExtendedCount
'    End If
'
'    info.Attributes.Directory = Attr.Directory
'
'    If (Attr.IsAttributeIncluded(SB_SFTP_ATTR_PERMISSIONS)) Then
'        info.Attributes.UserRead = Attr.UserRead
'        info.Attributes.UserWrite = Attr.UserWrite
'        info.Attributes.UserExecute = Attr.UserExecute
'
'        info.Attributes.GroupRead = Attr.GroupRead
'        info.Attributes.GroupWrite = Attr.GroupWrite
'        info.Attributes.GroupExecute = Attr.GroupExecute
'
'        info.Attributes.OtherRead = Attr.OtherRead
'        info.Attributes.OtherWrite = Attr.OtherWrite
'        info.Attributes.OtherExecute = Attr.OtherExecute
'    End If
'
'    Call SetCellInfo(Row, info)
'End Sub
'
'Private Sub scktClient_Connect()
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "TCP connection opened..."
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Authenticating Username: " & mvarCSFTP.UserName & ", Password: " & mvarCSFTP.Password & " using SSHClient..."
'
'    m_blnAuthenticationFailed = False
'
'    SSHClient.UserName = mvarCSFTP.UserName
'    SSHClient.Password = mvarCSFTP.Password
'    Call SSHClient.Open
'End Sub
'
'Private Sub scktClient_DataArrival(ByVal bytesTotal As Long)
'    m_blnClientDataAvailable = True
'
'    While m_blnClientDataAvailable
'        SSHClient.DataAvailable
'    Wend
'End Sub
'
'
'Private Sub scktClient_Close()
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "TCP connection closed..."
'End Sub

'Public Sub BuildFileList(ByVal Path As String)
'    If scktClient.State <> sckConnected Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "BuildFileList Error: not connected..."
'        Exit Sub
'    End If
'
'    While m_colCurrentFileList.Count > 0
'        Call m_colCurrentFileList.Remove(1)
'    Wend
'
'    If Path = "." Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Opening remote directory: Root Path..."
'    Else
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Opening remote directory: " & Path & " ..."
'    End If
'
'    Call SftpClient.OpenDirectory(Path)
'    m_lngState = STATE_OPEN_DIRECTORY_SENT
'End Sub
'
'Private Sub CloseCurrentHandle()
'    If scktClient.State <> sckConnected Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "CloseCurrentHandle Error: not connected..."
'        Exit Sub
'    End If
'
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Closing active handle..."
'    Call SftpClient.CloseHandle(m_strCurrentHandle)
'End Sub

'Private Function WritePermissions(ByVal Attributes As IElSftpFileAttributesX) As String
'    Dim Result As String
'
'    Result = vbNullString
'
'    If Attributes.Directory Then
'        Result = Result + "d"
'    End If
'
'    If Attributes.UserRead Then
'        Result = Result + "r"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.UserWrite Then
'        Result = Result + "w"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.UserExecute Then
'        Result = Result + "x"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.GroupRead Then
'        Result = Result + "r"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.GroupWrite Then
'        Result = Result + "w"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.GroupExecute Then
'        Result = Result + "x"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.OtherRead Then
'        Result = Result + "r"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.OtherWrite Then
'        Result = Result + "w"
'    Else
'        Result = Result + "-"
'    End If
'
'    If Attributes.OtherExecute Then
'        Result = Result + "x"
'    Else
'        Result = Result + "-"
'    End If
'
'    WritePermissions = Result
'End Function

'Private Sub OutputFileList()
'    Dim lngIdx As Long
'    Dim info As IElSftpFileInfoX
'
'    For lngIdx = 1 To m_colCurrentFileList.Count
'        Set info = m_colCurrentFileList(lngIdx)
'
'        If mvarCSFTP.HasTimeOut = False Then
'            If Right(Trim$(info.Name), 4) = ".rcv" Then
'                m_strCurrentFile = info.Name
'
'                DownloadFile info
'
'                Exit For
'            End If
'        Else
'            If scktClient.State <> sckConnected Then
'                scktClient.Close
'            End If
'
'            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Connection timed out. End receiving of messages from server..."
'            Exit For
'        End If
'    Next
'
'End Sub

Public Sub DeleteFile(ByVal Name As String)
    m_blnHasError = False
    
    If Not ElSimpleSftpClientX.Active Then
        m_blnHasError = True
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "Delete File Error: not connected..."
        Exit Sub
    End If
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Removing File " & Name
    Call ElSimpleSftpClientX.RemoveFile(m_strCurrentDir & "/" & Name)
    m_blnFileDeleted = True
End Sub

Public Sub DownloadFile(ByVal info As IElSftpFileInfoX)
    m_blnHasError = False
    
    If Not ElSimpleSftpClientX.Active Then
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "DownloadFile Error: not connected..."
        Exit Sub
    End If
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Starting file download, " & info.Name
     
    On Error GoTo ErrHandler
    Call ElSimpleSftpClientX.DownloadFile(m_strCurrentDir & "/" & info.Name, mvarCSFTP.MdbPath + "\" + info.Name)
    m_blnFileDownloaded = True
    
    Exit Sub
    
ErrHandler:
    If Err.Number > 0 Then
        m_blnHasError = True
    End If
    
End Sub

Public Sub UploadFile(ByVal LocalFile As String)
    Dim FName As String
    
    If Not ElSimpleSftpClientX.Active Then
        mvarCSFTP.TraceText = MESSAGE_PREFIX & "UploadFile Error: not connected..."
        Exit Sub
    End If
    
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Starting file upload, " & LocalFile
    
    FName = ExtractFileName(LocalFile)
    Dim RemoteName As String
    RemoteName = m_strCurrentDir & "/" & FName
    
    On Error GoTo ErrHandler
    m_blnFileUploaded = False
    Call ElSimpleSftpClientX.UploadFile(FName, RemoteName)
    m_blnFileUploaded = True
        
    Exit Sub
    
ErrHandler:
    If Err.Number > 0 Then
        m_blnHasError = True
    End If
    
    'Call OpenFileForRead(m_lngCurrentFile, LocalFile)
    'm_lngCurrentFileOffset = 0
    'm_lngCurrentFileSize = LOF(m_lngCurrentFile)
    'Call SftpClient.CreateFile(m_strCurrentDir & "/" & FName)
    'm_lngState = STATE_UPLOAD_OPEN
End Sub

'Private Sub WriteNextBlockToFile()
'    Dim Buf As Variant
'    Dim Transferred As Long
'
'    If scktClient.State <> sckConnected Then
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "WriteNextBlockToFile Error: not connected..."
'        Exit Sub
'    End If
'
'    mvarCSFTP.TraceText = MESSAGE_PREFIX & "WriteNextBlockToFile : FileOffset - " & m_lngCurrentFileOffset & " // FileSize - " & m_lngCurrentFileSize
'
'    If m_lngCurrentFileOffset >= m_lngCurrentFileSize Then
'        m_lngState = STATE_CLOSE_HANDLE
'        mvarCSFTP.TraceText = MESSAGE_PREFIX & "WriteNextBlockToFile Success: Upload is finished."
'        Close #m_lngCurrentFile
'        Call CloseCurrentHandle
'        m_blnFileUploaded = True
'        Exit Sub
'    End If
'
'    Call BlockRead(m_lngCurrentFile, Buf, FILE_BLOCK_SIZE)
'    Call SftpClient.Write(m_strCurrentHandle, m_lngCurrentFileOffset, 0, Buf)
'    Transferred = ArrSize(Buf)
'    m_lngCurrentFileOffset = m_lngCurrentFileOffset + Transferred
'End Sub

'Private Sub RequestAbsolutePath(ByVal Path As String)
'  Call SftpClient.RequestAbsolutePath(Path)
'End Sub

'Private Sub SetCellInfo(ByVal Index As Long, ByVal info As IElSftpFileInfoX)
'    Dim SizeLo As Long
'    Dim SizeHi As Long
'
'    Call info.Attributes.GetSize(SizeLo, SizeHi)
'
'    ListView1.ListItems(Index).Text = info.Name
'    ListView1.ListItems(Index).SubItems(1) = Str(SizeLo)
'    ListView1.ListItems(Index).SubItems(2) = WritePermissions(info.Attributes)
'End Sub

Private Function Str2ByteArr(ByVal S As String) As Variant
    Dim i As Integer
    Dim arr() As Byte
    
    ReDim arr(0 To Len(S) - 1)
    
    For i = 0 To Len(S) - 1
        arr(i) = Asc(Mid(S, i + 1, 1))
    Next
    
    Str2ByteArr = arr
End Function

Public Function ArrSize(ByRef v As Variant) As Integer
    If (VarType(v) And vbArray) <> vbArray Then
        ArrSize = 0
        Exit Function
    End If
    
    ArrSize = UBound(v) - LBound(v) + 1
End Function

Function OpenFileForRead(ByRef File As Integer, ByVal FileName As String) As Boolean
    File = FreeFile()
    Open FileName For Binary Access Read As #File
End Function

Function OpenFileForWrite(ByRef File As Integer, ByVal FileName As String) As Boolean
    File = FreeFile()
    Open FileName For Output Access Write As #File
    
    Close File
    Open FileName For Binary Access Write As #File
End Function

'Sub BlockRead(ByVal File As Integer, _
'    ByRef Buffer As Variant, _
'    ByVal Count As Long)
'
'    Dim S As String
'
'    S = Input(Count, File)
'
'    If Len(S) = 0 Then
'        Buffer = Empty
'    Else
'        Buffer = Str2ByteArr(S)
'    End If
'End Sub
'
'Sub BlockWrite(ByVal File As Integer, ByRef Buffer As Variant)
'    Dim lngIdx As Long
'    Dim bytData As Byte
'
'    Select Case VarType(Buffer)
'        Case vbArray + vbByte
'            For lngIdx = LBound(Buffer) To UBound(Buffer)
'                bytData = Buffer(lngIdx)
'                Put File, , bytData
'            Next
'
'        Case vbString
'            For lngIdx = 1 To Len(Buffer)
'                Put File, , Mid(Buffer, lngIdx, 1)
'            Next
'
'        Case Else
'            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Invalid data type..."
'            Call mvarCSFTP.DisconnectFromServer(False)
'
'    End Select
'End Sub

Function ExtractFileName(ByVal FileName As String) As String
    Dim ch As String
    Dim i As Integer, Idx As Integer
    
    Idx = 0
    For i = Len(FileName) To 1 Step -1
        ch = Mid(FileName, i, 1)
        
        If (ch = ":") Or (ch = "\") Or (ch = "/") Then
            Idx = i
            Exit For
        End If
    Next
    
    ExtractFileName = Mid(FileName, Idx + 1)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim FileInfo1 As IElSftpFileInfoX
    
    TimeOutCounter.Enabled = False
    
    'Call UnloadControls(Me)
    
    While m_colCurrentFileList.Count > 0
        Set FileInfo1 = m_colCurrentFileList(1)
        Set FileInfo1 = Nothing
        Call m_colCurrentFileList.Remove(1)
    Wend
    
    Set m_colCurrentFileList = Nothing
    
    Unload Me
    
    Set mvarCSFTP = Nothing
    Set MainForm = Nothing
End Sub

Private Sub TimeOutCounter_Timer()
    mvarCSFTP.HasTimeOut = True
End Sub


Public Sub ConnectSFTP()
    If ElSimpleSftpClientX.Active Then
        Call ElSimpleSftpClientX.Close
    End If
    
    ElSimpleSftpClientX.UserName = mvarCSFTP.UserName
    ElSimpleSftpClientX.Password = mvarCSFTP.Password
    ElSimpleSftpClientX.EnableVersion SB_SFTP_VERSION_3
    ElSimpleSftpClientX.Address = mvarCSFTP.HostName
    ElSimpleSftpClientX.Port = mvarCSFTP.PortNumber
    
    mvarCSFTP.TraceText = "Connecting to Hostname: " & mvarCSFTP.HostName & ", PortNumber: " & mvarCSFTP.PortNumber & "..."
    
    On Error GoTo ErrHandler
    Call ElSimpleSftpClientX.Open

ErrHandler:
    Select Case Err.Number
        Case 0
            
        Case Else
            mvarCSFTP.TraceText = MESSAGE_PREFIX & "Connection Error - " & Err.Description & " (" & Err.Number & ") "
            
    End Select
    
End Sub


Public Sub DisconnectSFTP()
    If ElSimpleSftpClientX.Active Then
        Call ElSimpleSftpClientX.Close
    End If
    
    m_blnDirectoryReadFinished = False
    m_blnFileDownloaded = False
    m_blnFileDeleted = False
End Sub


Public Sub RefreshRootDirectoryList()
    Dim Listing As Variant, i As Long
    Dim info As IElSftpFileInfoX
    Dim info_copy As IElSftpFileInfoX
    Dim item As ListItem
    Dim a() As IElSftpFileInfoX

    m_blnHasError = False

    If Not ElSimpleSftpClientX.Active Then
        Exit Sub
    End If

    'Clearing old data
    While m_colCurrentFileList.Count > 0
        m_colCurrentFileList.Remove (1)
    Wend

    On Error GoTo HandleErr
    m_strCurrentDir = vbNullString
    m_strCurrentDir = ElSimpleSftpClientX.RequestAbsolutePath(m_strCurrentDir)

    'Retrieving directory contents
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Retrieving file list..."
    Call ElSimpleSftpClientX.ListDirectory(m_strCurrentDir, Listing)
    For i = LBound(Listing) To UBound(Listing)
        Set info_copy = New ElSftpFileInfoX

        On Error Resume Next
        Set info = Listing(i)
        On Error GoTo 0

        Call info.CopyTo(info_copy)

        If Not info_copy.Attributes.Directory Then
            Call m_colCurrentFileList.Add(info_copy)
        End If
    Next

    m_blnDirectoryReadFinished = True
    Exit Sub

HandleErr:
    m_blnDirectoryReadFinished = True
    m_blnHasError = True
    mvarCSFTP.TraceText = MESSAGE_PREFIX & "Refresh Root Directory List Error - " & Err.Description & " (" & Err.Number & ") "

End Sub
