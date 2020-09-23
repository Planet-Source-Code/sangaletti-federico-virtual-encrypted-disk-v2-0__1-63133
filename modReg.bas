Attribute VB_Name = "modReg"
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... modReg
'   Author...... Sangaletti Federico
'   email:...... sangaletti@aliceposta.it
'   License..... FREE (But respect copyright of my work!)
'
'   Decription.. This module contains functions to
'                associate .vdsk files with the application
'-------------------------------------------------------

Option Explicit

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_SET_VALUE = &H2&
Private Const MAX_PATH = 260
Private Const REG_DWORD As Long = 4
Private Const REG_SZ = 1
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL

Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const FILE_TYPE = ".vdsk"
Public Const TYPE_NAME = "VDISK_Image"
Public Const TYPE_DESCRIPTION = "Virtual Encrypted Disk Image"
Public Const ICON_NAME = "vdisk.ico"
Public Const ICON_RES_ID = 501

Public Function WriteFileType() As Boolean
    On Error GoTo hError
    
    Dim ret As Long
    Dim KeyHandle As Long
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, FILE_TYPE, KeyHandle)
    If ret = ERROR_SUCCESS Then
        ret = RegSetValue(KeyHandle, vbNullString, REG_SZ, TYPE_NAME, Len(TYPE_NAME))
        If ret <> ERROR_SUCCESS Then
            WriteFileType = False
            RegCloseKey KeyHandle
            Exit Function
        End If
    Else
        WriteFileType = False
        RegCloseKey KeyHandle
        Exit Function
    End If
    RegCloseKey KeyHandle
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, TYPE_NAME, KeyHandle)
    If ret = ERROR_SUCCESS Then
        ret = RegSetValue(KeyHandle, vbNullString, REG_SZ, TYPE_DESCRIPTION, Len(TYPE_DESCRIPTION))
        If ret <> ERROR_SUCCESS Then
            WriteFileType = False
            RegCloseKey KeyHandle
            Exit Function
        End If
    Else
        WriteFileType = False
        RegCloseKey KeyHandle
        Exit Function
    End If
    RegCloseKey KeyHandle
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, TYPE_NAME & "\DefaultIcon", KeyHandle)
    If ret = ERROR_SUCCESS Then
        ret = RegSetValue(KeyHandle, vbNullString, REG_SZ, GetWinDir & "\" & ICON_NAME, Len(GetWinDir & "\" & ICON_NAME))
        If ret <> ERROR_SUCCESS Then
            WriteFileType = False
            RegCloseKey KeyHandle
            Exit Function
        End If
    Else
        WriteFileType = False
        RegCloseKey KeyHandle
        Exit Function
    End If
    RegCloseKey KeyHandle
    
    WriteFileType = True
    Exit Function
    
hError:
    WriteFileType = False
End Function

Private Function GetWinDir() As String
    Dim Buffer As String * MAX_PATH
    Dim nSize As Long
    
    nSize = GetWindowsDirectory(Buffer, MAX_PATH)
    GetWinDir = Left(Buffer, nSize)
End Function

Public Function CreateIcon() As Boolean
    On Error GoTo hError
    Dim bIcon() As Byte
    
    If Dir(GetWinDir & "\" & ICON_NAME) = vbNullString Then
        bIcon = LoadResData(ICON_RES_ID, "CUSTOM")
        Open GetWinDir & "\" & ICON_NAME For Binary Access Write As #1
            Put #1, , bIcon
        Close #1
        CreateIcon = True
    Else
        CreateIcon = False
    End If
    
    Exit Function

hError:
    CreateIcon = False
End Function
