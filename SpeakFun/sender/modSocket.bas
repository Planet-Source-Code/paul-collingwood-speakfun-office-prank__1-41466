Attribute VB_Name = "modSocket"
Option Explicit

Public Const MAX_WSADescription& = 256
Public Const MAX_WSASYSStatus& = 128

Public Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Public Const IP_SUCCESS& = 0

Public Const WS_VERSION_REQD& = &H101

Private Declare Function gethostbyname& Lib "wsock32" (ByVal hostname$)
  
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes&)

Private Declare Function lstrlenA& Lib "kernel32" (lpString As Any)

Public Declare Function WSAStartup& Lib "wsock32" (ByVal wVersionRequired&, lpWSADATA As WSADATA)
    
Public Declare Function WSACleanup& Lib "wsock32" ()

Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub


Public Function GetIPFromHostName$(ByVal host_name$)
   Dim hostent_pointer&, address_array_offset&, address_array_pointer&, address_pointer&, address_buffer$
   Dim wsa_data As WSADATA
   
   On Error Resume Next
   
   GetIPFromHostName = ""
   
   If WSAStartup(WS_VERSION_REQD, wsa_data) = IP_SUCCESS Then
   
   
      hostent_pointer = gethostbyname(host_name & vbNullChar)
   
      If hostent_pointer <> 0 Then
   
         address_array_offset = hostent_pointer + 12
         
         CopyMemory address_array_pointer, ByVal address_array_offset, 4
         
         CopyMemory address_pointer, ByVal address_array_pointer, 4
         
         address_buffer = Space$(4)
            
         ' Got the first IP address from the table
         CopyMemory ByVal address_buffer, ByVal address_pointer, 4
   
         ' Convert it to xxx.xxx.xxx.xxx format string
         GetIPFromHostName = CStr(Asc(address_buffer)) & "." & _
              CStr(Asc(Mid$(address_buffer, 2, 1))) & "." & _
              CStr(Asc(Mid$(address_buffer, 3, 1))) & "." & _
              CStr(Asc(Mid$(address_buffer, 4, 1)))
   
      End If
      
      WSACleanup
   End If
End Function


