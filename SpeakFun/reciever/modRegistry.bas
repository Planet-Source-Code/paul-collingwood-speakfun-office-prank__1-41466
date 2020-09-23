Attribute VB_Name = "modRegistry"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_EXPAND_SZ = 2
Public Const REG_DWORD = 4
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_DWORD_LITTLE_ENDIAN = 4

Public Const ERROR_SUCCESS = 0
Public Const ERROR_REG = 1

Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)

Private Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, ByVal lpSecurityAttributes&, phkResult&, lpdwDisposition&)

Private Declare Function RegQueryValueExString& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, ByVal lpData$, lpcbData&)

Private Declare Function RegQueryValueExLong& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData&, lpcbData&)

Private Declare Function RegQueryValueExNULL& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, ByVal lpData&, lpcbData&)

Private Declare Function RegSetValueExString& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal lpValue$, ByVal cbData&)

Private Declare Function RegSetValueExLong& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, lpValue&, ByVal cbData&)

Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey&, ByVal lpSubKey$)

Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$)

Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)

Private Function GetOSVersion() As Integer
   Dim result&
   Dim OSV As OSVERSIONINFO
   
   OSV.dwOSVersionInfoSize = Len(OSV)
   
   result = GetVersionEx(OSV)
   
   GetOSVersion = OSV.dwPlatformId
   Select Case GetOSVersion
      Case 1
         If OSV.dwMinorVersion = 10 Then GetOSVersion = 3
      Case 2
         If OSV.dwMinorVersion = 51 Then GetOSVersion = 4
   End Select

End Function

Private Function GetRegistryValue(RootKey&, KeyName$, ValueName$) As Variant
   Dim lRtn&, hKey&, lCdata&, lValue&, sValue$, lRtype&
  
  lRtn = RegOpenKeyEx(RootKey, KeyName, 0&, KEY_ALL_ACCESS, hKey)
  
  If lRtn <> ERROR_SUCCESS Then
    RegCloseKey (hKey)
    Exit Function
  End If
  
  lRtn = RegQueryValueExNULL(hKey, ValueName, 0&, lRtype, 0&, lCdata)
  
  If lRtn <> ERROR_SUCCESS Then
    RegCloseKey (hKey)
    Exit Function
  End If
  
  Select Case lRtype
    Case 1
      sValue = String(lCdata, 0)
      lRtn = RegQueryValueExString(hKey, ValueName, 0&, lRtype, sValue, lCdata)
      
      If lRtn = ERROR_SUCCESS Then
        GetRegistryValue = sValue
      Else
        GetRegistryValue = Empty
      End If
    Case 4
      lRtn = RegQueryValueExLong(hKey, ValueName, 0&, lRtype, lValue, lCdata)
            
      If lRtn = ERROR_SUCCESS Then
        GetRegistryValue = lValue
      Else
        GetRegistryValue = Empty
      End If
  End Select
  
  RegCloseKey (hKey)
  
End Function

Private Sub SetRegistryValue(RootKey&, KeyName$, ValueName$, KeyType%, KeyValue As Variant)
   Dim lRtn&, hKey&, lValue&, sValue$, lSize&
  
   lRtn = RegOpenKeyEx(RootKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
  
   If lRtn <> ERROR_SUCCESS Then
      RegCloseKey (hKey)
      Exit Sub
   End If
  
  Select Case KeyType
    Case 1
      sValue = KeyValue
      lSize = Len(sValue)
      
      lRtn = RegSetValueExString(hKey, ValueName, 0&, REG_SZ, sValue, lSize)
      
      If lRtn <> ERROR_SUCCESS Then
        RegCloseKey (hKey)
        Exit Sub
      End If
    Case 4
      lValue = KeyValue
      
      lRtn = RegSetValueExLong(hKey, ValueName, 0&, REG_DWORD, lValue, 4)
      
      If lRtn <> ERROR_SUCCESS Then
        RegCloseKey (hKey)
        Exit Sub
      End If
  End Select

  RegCloseKey (hKey)
  
End Sub

Private Sub DeleteRegistryValue(RootKey&, KeyName$, ValueName$)
  Dim lRtn&, hKey&
  
  lRtn = RegOpenKeyEx(RootKey, KeyName, 0&, KEY_ALL_ACCESS, hKey)
  
  If lRtn <> ERROR_SUCCESS Then
    RegCloseKey (hKey)
    Exit Sub
  End If
  
  lRtn = RegDeleteValue(hKey, ValueName)
  
  If lRtn <> ERROR_SUCCESS Then
    RegCloseKey (hKey)
    Exit Sub
  End If
  
  RegCloseKey (hKey)
  
End Sub

Public Sub AddStartupRegistryEntry(KeyLabel$)
   Dim s1$, s2$

   s1 = LCase(App.Path & "\" & App.EXEName & ".EXE")
   
   Select Case GetOSVersion
   Case 1, 2, 3, 4
      s2 = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", KeyLabel)
   
      If s1 <> s2 Then
         SetRegistryValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", KeyLabel, REG_SZ, s1
      End If
      
   End Select

End Sub

Public Sub RemoveStartupRegistryEntry(KeyLabel$)

   DeleteRegistryValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", KeyLabel

End Sub





