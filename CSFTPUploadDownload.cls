VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSFTPUploadDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MESSAGE_PREFIX = "SFTP SERVER: "

Private mvarMainSFTP As MainForm

Private mvarHostName As String
Private mvarUserName As String
Private mvarPassword As String
Private mvarTimeoutInterval As String
Private mvarMdbPath As String

Private mvarPortNumber As Long

Private mvarForPLDA As Boolean
Private mvarForSending As Boolean
Private mvarTraceOn As Boolean
Private mvarHasTimeOut As Boolean

Public Event TraceMessage(ByVal TraceText As String, ByVal ForSending As Boolean)
Public Event MessageReceived(ByVal Data As String, ByVal PullQueue As String)
Public Event PullExit(ByVal QueuesEmptied As Boolean)
Public Event SendExit(ByVal MessageSent As Boolean)

Public Property Let MdbPath(ByVal vData As String)
    mvarMdbPath = vData
End Property

Public Property Get MdbPath() As String
    MdbPath = mvarMdbPath
End Property

Public Property Let HasTimeOut(ByVal vData As Boolean)
    mvarHasTimeOut = vData
End Property

Public Property Get HasTimeOut() As Boolean
    HasTimeOut = mvarHasTimeOut
End Property

Public Property Let TimeoutInterval(ByVal vData As String)
    mvarTimeoutInterval = vData
End Property

Public Property Get TimeoutInterval() As String
    TimeoutInterval = mvarTimeoutInterval
End Property

Public Property Let TraceOn(ByVal vData As Boolean)
    mvarTraceOn = vData
End Property

Public Property Get TraceOn() As Boolean
    TraceOn = mvarTraceOn
End Property

Public Property Let ForSending(ByVal vData As Boolean)
    mvarForSending = vData
End Property

Public Property Get ForSending() As Boolean
    ForSending = mvarForSending
End Property

Public Property Let ForPLDA(ByVal vData As Boolean)
    mvarForPLDA = vData
End Property

Public Property Get ForPLDA() As Boolean
    ForPLDA = mvarForPLDA
End Property

Public Property Let PortNumber(ByVal vData As Long)
    mvarPortNumber = vData
End Property

Public Property Get PortNumber() As Long
    PortNumber = mvarPortNumber
End Property

Public Property Let Password(ByVal vData As String)
    mvarPassword = vData
End Property

Public Property Get Password() As String
    Password = mvarPassword
End Property

Public Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property

Public Property Get UserName() As String
    UserName = mvarUserName
End Property

Public Property Let HostName(ByVal vData As String)
    mvarHostName = vData
End Property

Public Property Get HostName() As String
    HostName = mvarHostName
End Property

Private Sub Class_Initialize()
    Set mvarMainSFTP = New MainForm
        
    mvarMainSFTP.LoadForm Me
End Sub

Public Sub TerminateClass()
    On Error Resume Next
    Unload mvarMainSFTP
    On Error GoTo 0
    
    On Error Resume Next
    Set mvarMainSFTP = Nothing
    On Error GoTo 0
End Sub

Public Property Let TraceText(ByVal vNewValue As String)
    'Everytime TraceText is set, it will raise an event in the main application which will trigger AddToTrace function to collect SFTP logs
    If TraceOn = True Then RaiseEvent TraceMessage(vNewValue, ForSending)
End Property

Public Sub PullMessages()
    Dim lngIdx As Long
    Dim info As IElSftpFileInfoX
    Dim strMessage As String
    
    On Error GoTo ErrorHandler
    
    HasTimeOut = False
    
    With mvarMainSFTP
        .ConnectSFTP
            
        .TimeOutCounter.Interval = 1000 * TimeoutInterval
        .TimeOutCounter.Enabled = True
            
        '************************************************************************************************************************************************
        'START REFRESH DIRECTORY LIST
        '************************************************************************************************************************************************
        .RefreshRootDirectoryList
        
        Do While True
            DoEvents
            
            If .m_blnDirectoryReadFinished = True Then Exit Do
            
            'If timeout time is reached or there was an error exit pullmessages
            If (HasTimeOut = True) Or (.m_blnHasError = True) Then
                If .m_blnHasError Then
                    Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to refresh directory to SFTP server..."
                Else
                    Me.TraceText = MESSAGE_PREFIX & "Queue has timeout while trying to connect to SFTP server..."
                End If
                
                Call DisconnectFromServer(False)
                Exit Sub
            End If
        Loop
        '************************************************************************************************************************************************
        
        For lngIdx = 1 To .m_colCurrentFileList.Count
            Set info = .m_colCurrentFileList(lngIdx)
    
            If HasTimeOut = False Then
                If (Right(Trim$(info.Name), 4) = ".rcv") And _
                   ((Left(info.Name, 4) = "PLDA" And ForPLDA) Or _
                   (Left(info.Name, 4) = "NCTS" And Not ForPLDA)) Then
                    '************************************************************************************************************************************************
                    'START FILE DOWNLOAD
                    '************************************************************************************************************************************************
                    'Delete the incomplete download
                    'On Error Resume Next
                    'Kill MdbPath + "\" + Info.Name
                    'On Error GoTo 0
                    
                    .m_blnFileDownloaded = False
                    .DownloadFile info
                    
                    Do While True
                        DoEvents
                        
                        If .m_blnFileDownloaded = True Then Exit Do
                        
                        If .m_blnHasError = True Then
                            Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to download customs messages from server..."
                            
                            Call DisconnectFromServer(False)
                            Exit Sub
                        End If
                    Loop
                    '************************************************************************************************************************************************
                    
                    '************************************************************************************************************************************************
                    'START FILE DELETE FROM SERVER AFTER SUCCESSFUL DOWNLOAD
                    '************************************************************************************************************************************************
                    .m_blnFileDeleted = False
                    .DeleteFile info.Name
                    
                    Do While True
                        DoEvents
                        
                        If .m_blnFileDeleted = True Then Exit Do
                        
                        If .m_blnHasError = True Then
                            Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to delete message from server..."
                            
                            Call DisconnectFromServer(False)
                            Exit Sub
                        End If
                    Loop
                    '************************************************************************************************************************************************
                    
                    '************************************************************************************************************************************************
                    'PROCESS MESSAGE RECEIVED
                    '************************************************************************************************************************************************
                    Me.TraceText = MESSAGE_PREFIX & "Start Processing Message..."
                    
                    If Left(info.Name, 4) = "PLDA" Then ForPLDA = True
                    
                    strMessage = ReadMessage(info.Name)
                    
                    Me.TraceText = MESSAGE_PREFIX & strMessage
                    
                    RaiseEvent MessageReceived(strMessage, HostName) 'second argument not yet finalized
                    
                    Me.TraceText = MESSAGE_PREFIX & "End Processing Message..."
                    '************************************************************************************************************************************************
                End If
            Else
                Me.TraceText = MESSAGE_PREFIX & "Connection timed out. End Receiving of Messages from Server..."
                Call DisconnectFromServer(False)
                Exit For
            End If
            
            .m_blnFileDownloaded = False
            .m_blnFileDeleted = False
        Next
    End With
    
    Call DisconnectFromServer(True)
    
ErrorHandler:
    Select Case Err.Number
        Case 0
        
        Case Else
            Call DisconnectFromServer(False)
            Me.TraceText = MESSAGE_PREFIX & "Error in CSFTPUploadDownload.PullMessages, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
            
    End Select
    
End Sub


Private Function ReadMessage(ByVal FileName As String) As String
    
    Dim intFreeFile As Integer
    Dim strLine As String
    
    intFreeFile = FreeFile()
                                
    Open MdbPath & "\" & FileName For Input As #intFreeFile
    
    'Read FlatFile to get ReadMessage
    Do Until EOF(intFreeFile)
        Line Input #intFreeFile, strLine
        
        ReadMessage = ReadMessage & strLine & vbCrLf
    Loop
    
    Close #intFreeFile
    
    'Delete flatfile
    On Error Resume Next
    Kill MdbPath & "\" & FileName
    On Error GoTo 0
    
End Function


Private Sub WriteMessage(ByVal Message As String, ByVal FileName As String)
    
    Dim intFreeFile As Integer
    
    intFreeFile = FreeFile()
    
    Open MdbPath & "\" & FileName For Output As #intFreeFile
    
    Print #intFreeFile, Message
    
    Close #intFreeFile
    
End Sub

Public Sub DisconnectFromServer(ByVal Successful As Boolean)
    Call mvarMainSFTP.DisconnectSFTP
    
    If ForSending = True Then RaiseEvent SendExit(Successful)
    
    RaiseEvent PullExit(True)
End Sub


Public Sub SendMessage(ByVal Message As String, ByVal Code As String, ByVal DocumentType As String)
    
    Dim strFileName As String
    Dim lngIdx As Long
    Dim info As IElSftpFileInfoX
    Dim sngDelay As Single
    
    With mvarMainSFTP
        .ConnectSFTP
            
        .TimeOutCounter.Interval = 1000 * TimeoutInterval
        .TimeOutCounter.Enabled = True
        
        
        '************************************************************************************************************************************************
        'START REFRESH DIRECTORY LIST
        '************************************************************************************************************************************************
        .RefreshRootDirectoryList
        
        Do While True
            DoEvents
            
            If .m_blnDirectoryReadFinished = True Then Exit Do
            
            'If timeout time is reached or there was an error exit pullmessages
            If (HasTimeOut = True) Or (.m_blnHasError = True) Then
                If .m_blnHasError Then
                    Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to refresh directory to SFTP server..."
                Else
                    Me.TraceText = MESSAGE_PREFIX & "Queue has timeout while trying to connect to SFTP server..."
                End If
                
                Call DisconnectFromServer(False)
                Exit Sub
            End If
        Loop
        '************************************************************************************************************************************************
            
        'Write Message to file for sending
        strFileName = CreateFileName(DocumentType, Code)
        Call WriteMessage(Message, strFileName)
        
        .UploadFile MdbPath & "\" & strFileName
        
        Do While True
            DoEvents
            If .m_blnFileUploaded = True Then Exit Do
            
            If .m_blnHasError = True Then
                Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to send message to server..."
                
                Call DisconnectFromServer(False)
                Exit Sub
            End If
        Loop
        
        'Cleanup sent files
        On Error Resume Next
        Kill MdbPath & "\" & "*.snd"
        On Error GoTo 0
        
        '*******************************************************************************************************************************
        'Workaround on intermittent problem
        '*******************************************************************************************************************************
        sngDelay = Timer + 1
        
        Do
            'Create a 0.5 sec delay
        Loop Until Timer >= sngDelay
        
        '************************************************************************************************************************************************
        'START REFRESH DIRECTORY LIST
        '************************************************************************************************************************************************
        .RefreshRootDirectoryList
        
        Do While True
            DoEvents
            
            If .m_blnDirectoryReadFinished = True Then Exit Do
            
            'If timeout time is reached or there was an error exit pullmessages
            If (HasTimeOut = True) Or (.m_blnHasError = True) Then
                If .m_blnHasError Then
                    Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to refresh directory to SFTP server..."
                Else
                    Me.TraceText = MESSAGE_PREFIX & "Queue has timeout while trying to connect to SFTP server..."
                End If
                
                Call DisconnectFromServer(False)
                Exit Sub
            End If
        Loop
        '************************************************************************************************************************************************
        
        For lngIdx = 1 To .m_colCurrentFileList.Count
            Set info = .m_colCurrentFileList(lngIdx)
            
            If info.Name = strFileName Then
                .m_blnFileDeleted = False
                .DeleteFile info.Name
                
                Do While True
                    DoEvents
                    
                    If .m_blnFileDeleted = True Then Exit Do
                    
                    If .m_blnHasError = True Then
                        Me.TraceText = MESSAGE_PREFIX & "Error has occured while trying to delete message from server..."
                        
                        Call DisconnectFromServer(False)
                        Exit Sub
                    End If
                Loop
                    
                Me.TraceText = MESSAGE_PREFIX & "Communication Error, file sent to SFTP is corrupted, file will be resend on next queue..."
                Call DisconnectFromServer(False)
                Exit Sub
            End If
        Next
        '*******************************************************************************************************************************
    End With
    
    Me.TraceText = MESSAGE_PREFIX & "End of sending messages to SFTP server..."
    Call DisconnectFromServer(True)
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 0
        
        Case Else
            Call DisconnectFromServer(False)
            Me.TraceText = MESSAGE_PREFIX & "Error in CSFTPUploadDownload.SendMessage, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
            
    End Select
    
End Sub


Private Function CreateFileName(ByVal DType As String, ByVal Code As String) As String
    
    Select Case DType
        Case 14, 18
            CreateFileName = "PLDA" & Code & "_" & Format(Now, "ddMMyyyyhhmmss") & ".snd"
            
        Case Else
            CreateFileName = "NCTS" & Code & "_" & Format(Now, "ddMMyyyyhhmmss") & ".snd"
            
    End Select
    
End Function
