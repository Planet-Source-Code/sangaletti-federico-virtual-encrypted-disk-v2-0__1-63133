Attribute VB_Name = "modWipeFiles"
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... modWipeFiles
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This module contains functions for a
'                *SECURE* files delete
'-------------------------------------------------------

Option Explicit

Private Const MAX_PATH = 260
Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000

Private Const GENERIC_WRITE = &H40000000
Private Const TRUNCATE_EXISTING = 5
Private Const FILE_SHARE_WRITE = &H2

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function WipeAllFiles(PathName As String) As Boolean
    Dim hFile As Long, hResult As Long
    Dim FindData As WIN32_FIND_DATA
    
    'Go to the specified directory
    If SetCurrentDirectory(PathName) = 0 Then
        WipeAllFiles = False
        Exit Function
    End If
    
    'Point to the first file
    hFile = FindFirstFile("*", FindData)
    If hFile = INVALID_HANDLE_VALUE Then
        WipeAllFiles = False
        Exit Function
    End If
    
    'Wipe all files
    Do
        If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
            If DestroyFile(IIf(Right(PathName, 1) <> "\", PathName & "\", PathName) & FindData.cFileName) = INVALID_HANDLE_VALUE Then
                WipeAllFiles = False
                FindClose hFile
                Exit Function
            End If
        End If
        hResult = FindNextFile(hFile, FindData)
    Loop Until hResult = 0
    
    'All files wiped
    If GetLastError = ERROR_NO_MORE_FILES Then WipeAllFiles = True Else WipeAllFiles = False
    
    FindClose hFile
    SetCurrentDirectory "C:\"
End Function

Private Function DestroyFile(PathName As String) As Long
    Dim hFile As Long
    
    'Open the specified file for writing with no buffering
    hFile = CreateFile(PathName, GENERIC_WRITE, FILE_SHARE_WRITE, 0, TRUNCATE_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_WRITE_THROUGH, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        DestroyFile = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Flush the file data
    If FlushFileBuffers(hFile) = 0 Then
        DestroyFile = INVALID_HANDLE_VALUE
        CloseHandle hFile
        Exit Function
    End If
    
    CloseHandle hFile
    
    If DeleteFile(PathName) = 0 Then DestroyFile = INVALID_HANDLE_VALUE Else DestroyFile = hFile
End Function
