VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   Caption         =   "SpeakFunSend"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   330
      Left            =   6210
      TabIndex        =   10
      ToolTipText     =   "Click here to stop SpeakFun on the victim's PC."
      Top             =   2010
      Width           =   780
   End
   Begin VB.TextBox txtInternetAddress 
      Enabled         =   0   'False
      Height          =   285
      Left            =   825
      TabIndex        =   9
      Top             =   2010
      Width           =   2415
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "Other"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   8
      ToolTipText     =   "Cllick here to allow you to enter the name or IP address of the victims PC (if not on the LAN list above)."
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "LAN"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   7
      ToolTipText     =   "Cllick here to select the victim from a list of PCs on your Local Area Network (LAN)."
      Top             =   1710
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.ComboBox cmbVoiceType 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select the gender and accent of the spoken text."
      Top             =   2010
      Width           =   2115
   End
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   2400
      Top             =   1920
   End
   Begin VB.TextBox txtMessage 
      Height          =   1485
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmMain.frx":0442
      ToolTipText     =   "Type in the text you want spoken on the victim's machine. Can be multiple lines, so cut & paste away!"
      Top             =   75
      Width           =   7785
   End
   Begin VB.ComboBox cmbNetwork 
      Height          =   315
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select the victim from this list of PCs on your Local Area Network (LAN). They must have SpeakFun installed and running!"
      Top             =   1635
      Visible         =   0   'False
      Width           =   4890
   End
   Begin VB.CommandButton butTest 
      Caption         =   "Test"
      Height          =   330
      Left            =   7095
      TabIndex        =   1
      ToolTipText     =   "Click here to test what the text sounds like on you PC."
      Top             =   1620
      Width           =   780
   End
   Begin VB.CommandButton butSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   330
      Left            =   7095
      TabIndex        =   0
      ToolTipText     =   "Click here to attempt to send the spoken text to the victim's PC"
      Top             =   2010
      Width           =   780
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3105
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrConnection 
      Enabled         =   0   'False
      Left            =   2655
      Top             =   1920
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   75
      Left            =   2925
      TabIndex        =   5
      Top             =   2130
      Width           =   120
      ExtentX         =   212
      ExtentY         =   132
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
   Begin VB.Label labScanning 
      Alignment       =   2  'Center
      Caption         =   "Scanning LAN - please wait..."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1695
      Width           =   5505
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private remote_stop As Boolean
Private remote_unlink As Boolean

Private Sub butSend_Click()
   
   On Error GoTo CONNECT_ERROR
   
   ' Disable releavnt buttons
   butSend.Enabled = False
   butStop.Enabled = False
   
   frmMain.MousePointer = vbArrowHourglass
   
   ' Acquire IP address of either LAN PC or user-defined address
   If optDestination(0).Value = True Then
      Winsock1.RemoteHost = GetIPFromHostName(Mid$(Trim(cmbNetwork.List(cmbNetwork.ListIndex)), 3))
   Else
      Winsock1.RemoteHost = GetIPFromHostName(Trim(txtInternetAddress.Text))
   End If
   
   If Winsock1.RemoteHost = "" Then
      ' Unable to resolve IP address!
      MsgBox "Can't resolve!", vbApplicationModal + vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "SpeakFun error"
      butSend.Enabled = True
      butStop.Enabled = True
   Else
      ' Have resolved address - if resolved from empty user-defined entry, then display the IP address - it's this PCs!
      If optDestination(1).Value = True Then
         If Trim(txtInternetAddress.Text) = "" Then
            txtInternetAddress.Text = Winsock1.RemoteHost
         End If
      End If
   End If
   
   ' Trigger timeout (in case connection fails).
   tmrConnection.Interval = 5000
   tmrConnection.Enabled = True
   
   ' Initialise the socket - it tries to connect to the same port number as SpeakFunSend is listening
   Winsock1.RemotePort = 1302
   Winsock1.Connect
      
   Exit Sub
   
CONNECT_ERROR:
   frmMain.MousePointer = vbDefault
   If Err.Number = 10049 Then
      tmrConnection.Enabled = False
      If Winsock1.State <> sckClosed Then
         Winsock1.Close
      End If
      butSend.Enabled = True
      butStop.Enabled = True
      MsgBox Err.Description, vbExclamation + vbOKOnly, "IP Address Error"
   End If

End Sub

Private Sub butStop_Click()
   
   On Error Resume Next
   
   Select Case MsgBox("Do you want to prevent SpeakFun from running" & vbCrLf & "when the victim's PC is rebooted?", vbApplicationModal + vbYesNoCancel + vbMsgBoxSetForeground + vbQuestion, "SpeakFun stop request")
      Case vbCancel
         ' Do nothing
         Exit Sub
      
      Case vbYes
         remote_stop = True
         remote_unlink = True
      
      Case vbNo
         remote_stop = True
         remote_unlink = False
   End Select

   butSend.Enabled = False
   butStop.Enabled = False
   butSend_Click
End Sub

Private Sub butTest_Click()
   WebBrowser1.Navigate GetURL(cmbVoiceType.ListIndex)
End Sub

Private Sub Form_Load()
   ' Reset variabales used to stop the remote SpeakFun application
   remote_stop = False
   remote_unlink = False
   
   ' Populate the vioce type combobox
   cmbVoiceType.AddItem "American Female 1"
   cmbVoiceType.AddItem "American Female 2"
   cmbVoiceType.AddItem "American Female 3"
   cmbVoiceType.AddItem "American Female 4"
   cmbVoiceType.AddItem "American Female 5"
   cmbVoiceType.AddItem "American Male 1"
   cmbVoiceType.AddItem "American Male 2"
   cmbVoiceType.AddItem "American Male 3"
   cmbVoiceType.AddItem "British Female"
   cmbVoiceType.AddItem "British Male 1"
   cmbVoiceType.AddItem "British Male 2"
   cmbVoiceType.AddItem "Scottish Female"
   cmbVoiceType.AddItem "Scottish Male"
   cmbVoiceType.AddItem "Australian Female"
   cmbVoiceType.ListIndex = 8 ' default to British Female voice
End Sub

Private Sub tmrConnection_Timer()
   
   On Error Resume Next
         
   tmrConnection.Enabled = False
   
   ' Timeout trying to send speech request, so close socket (if applicable)
   If Winsock1.State <> sckClosed Then
       Winsock1.Close
   End If
         
   remote_stop = False
   remote_unlink = False
         
   butSend.Enabled = True
   butStop.Enabled = True
   frmMain.MousePointer = vbDefault
   
   MsgBox "No response!", vbApplicationModal + vbExclamation + vbOKOnly + vbMsgBoxSetForeground, "SpeakFun send"
End Sub


Private Sub tmrStart_Timer()
   
   On Error Resume Next
   
   ' Single-shot  process on startup, so stop this timer for all time.
   tmrStart.Enabled = False
   
   ' Get list of all other PCs on LAN
   PopulateNetworkCombobox cmbNetwork, 0
   
   If cmbNetwork.ListCount > 0 Then
      cmbNetwork.Visible = True
      optDestination(0).Enabled = True
      optDestination(1).Enabled = True
      txtInternetAddress.Enabled = True
   Else
      optDestination(1).Enabled = True
      optDestination(1).Value = True
      txtInternetAddress.Enabled = True
   End If
   butSend.Enabled = True
   butStop.Enabled = True

End Sub

Private Sub Winsock1_Close()
   
   On Error Resume Next
   
   butSend.Enabled = True
   butStop.Enabled = True
   frmMain.MousePointer = vbDefault
   Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
   
   On Error Resume Next
   
   tmrConnection.Enabled = False
   frmMain.MousePointer = vbDefault
   
   If remote_stop = True Then
      ' If stopping the SpeakFun application, then send relevant command string.
      If remote_unlink = False Then
         Winsock1.SendData "@@@STOP@@@"
      Else
         Winsock1.SendData "@@@UNLINK@@@"
      End If
   Else
      ' Send URL containong speech request
      Winsock1.SendData GetURL(cmbVoiceType.ListIndex)
   End If

   remote_stop = False
   remote_unlink = False
End Sub

Private Function GetURL(ByVal voice_type_index&) As String
   Dim text_string$, url_string$, text_index&, char_code%
   Dim voice_type_array As Variant

   ' This routine gets tezt string and converts it into a URL-friendly format
   ' It is then inserted into relevant URL to Rhetorical Systems rVioce Demo website.
   ' When navigated to on a target PC running SpeakFun, the voice will be herad there!
   
   voice_type_array = Array("ga_f05", "ga_f01", "ga_f02", "ga_f03", "ga_f04", _
                            "ga_m01", "ga_m02", "ga_m03", _
                            "rp_f01", _
                            "rp_m01", "rp_m02", _
                            "sc_f01", _
                            "sc_m01", _
                            "au_f01")

   text_string = Trim(txtMessage.Text)
   url_string = ""
   For text_index = 1 To Len(text_string)
      char_code = Asc(Mid$(text_string, text_index, 1))
      If char_code = 32 Then
         url_string = url_string & "+"
      Else
         url_string = url_string & "%" & UCase(Right$("00" & Hex$(char_code), 2))
      End If
   Next

   GetURL = "http://www.rhetorical.com/cgi-bin/demo.cgi?text=" & url_string & "&language=en&rate=16000&media=1&gender=male&accent=American&voice=en_" & CStr(voice_type_array(voice_type_index))
   
End Function
