Attribute VB_Name = "modNetwork"
Option Explicit

Public Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

Private Declare Function WNetOpenEnum& Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope&, ByVal dwType&, ByVal dwUsage&, lpNetResource As Any, lphEnum&)

Private Declare Function WNetEnumResource& Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum&, lpcCount&, ByVal lpBuffer&, lpBufferSize&)

Private Declare Function WNetCloseEnum& Lib "mpr.dll" (ByVal hEnum&)

Private Const RESOURCE_GLOBALNET = &H2

Private Const RESOURCETYPE_ANY = &H0

Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags&, ByVal dwBytes&)

Private Declare Function GlobalFree& Lib "kernel32" (ByVal hMem&)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy&)
   
Private Declare Function lstrcpy& Lib "kernel32" Alias "lstrcpyA" (ByVal NewString$, ByVal OldString&)

Public Function PopulateNetworkCombobox(combobox_control As ComboBox, ByVal remote_name&) As Boolean
   Dim elem_handle&, elem_buffer&, buffer_size&, num_elems&, elem_pointer&, return_value&, elem_index&
   Dim enum_success As Boolean
   Dim net_resource As NETRESOURCE

   On Error Resume Next
   If Err.Number > 0 Then Exit Function

   ' Remote_name value of zero denotes root search
   If remote_name = 0& Then
      combobox_control.Clear
   End If
   net_resource.lpRemoteName = remote_name
   
   On Error GoTo ErrorHandler

   return_value = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, 0, net_resource, elem_handle)

   If return_value = 0 Then
      ' Allocate a 31 Kbyte buffer
      buffer_size = 1024 * 31
      elem_buffer = GlobalAlloc(GPTR, buffer_size)

      If elem_buffer <> 0 Then
         ' Now populate buffer with array of NETRESOURCE structures - we want all the domains
         num_elems = buffer_size \ LenB(net_resource) ' Request as many entries as possible
         return_value = WNetEnumResource(elem_handle, num_elems, elem_buffer, buffer_size)
         
         If return_value = 0 Then
            enum_success = True
            ' Set up a pointer to the first element.
            elem_pointer = elem_buffer
            ' For eachy element in the array
            For elem_index = 1 To num_elems
               DoEvents
               ' Copy data from the buffer into our local structure
               CopyMemory net_resource, ByVal elem_pointer, LenB(net_resource)
               If net_resource.dwDisplayType = RESOURCEDISPLAYTYPE_DOMAIN Then
                  ' Recursive search through each domain
                  If PopulateNetworkCombobox(combobox_control, net_resource.lpRemoteName) = False Then
                     enum_success = False
                  End If
               ElseIf net_resource.dwDisplayType = RESOURCEDISPLAYTYPE_SERVER Then
                  ' Add PC reference to the combobox list
                  combobox_control.AddItem PointerToString(net_resource.lpRemoteName)
                  If combobox_control.ListCount = 1 Then
                     combobox_control.ListIndex = 0
                  End If
               End If
               ' Move the pointer on
               elem_pointer = elem_pointer + LenB(net_resource)
            Next
      
            PopulateNetworkCombobox = enum_success
         End If
      End If
   End If
   
ErrorHandler:
   On Error Resume Next
   
   If elem_buffer <> 0 Then
      GlobalFree elem_buffer
   End If
   
   If elem_handle <> 0 Then
      WNetCloseEnum elem_handle
   End If

End Function

Private Function PointerToString$(ByVal pointer&)
   Dim string_buffer$
   
   string_buffer = String(65535, Chr$(0))
   lstrcpy string_buffer, pointer
   PointerToString = Left(string_buffer, InStr(string_buffer, Chr$(0)) - 1)
End Function








