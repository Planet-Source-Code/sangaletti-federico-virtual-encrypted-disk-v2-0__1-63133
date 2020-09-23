Attribute VB_Name = "modDOSDevice"
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... modDOSDevice
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This module contains functions to create
'                delete and get infos of a virtual drive
'-------------------------------------------------------

Option Explicit

Const DDD_EXACT_MATCH_ON_REMOVE = &H4
Const DDD_RAW_TARGET_PATH = &H1
Const DDD_REMOVE_DEFINITION = &H2
Const FIRST_DEVICE_NAME = &H41

Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Public Function CreateDevice(DeviceName As String, PathName As String) As String
    If DefineDosDevice(0, DeviceName, PathName) = 0 Then
        MsgBox "Unable to create device " & DeviceName, vbCritical
        CreateDevice = vbNullString
    Else
        CreateDevice = DeviceName
    End If
End Function

Public Function RemoveDevice(DeviceName As String) As Boolean
    If DefineDosDevice(DDD_REMOVE_DEFINITION, DeviceName, vbNullString) = 0 Then
        MsgBox "Unable to remove device " & DeviceName, vbCritical
        RemoveDevice = False
    Else
        RemoveDevice = True
    End If
End Function

Public Function GetDeviceInfo(DeviceName As String) As String
    Dim Buffer As String * 400
    Dim nChar As Long
    
    nChar = QueryDosDevice(DeviceName, Buffer, 400)
    If nChar = 0 Then
        GetDeviceInfo = "Unable to get device info"
    Else
        GetDeviceInfo = Left(Buffer, nChar - 2)
    End If
End Function

Public Function GetFirstFreeDevice() As String
    Dim Devices As String
    Dim i As Integer
    
    Devices = ListDevices
    For i = 3 To 25
        If InStr(1, Devices, Chr$(FIRST_DEVICE_NAME + i)) = 0 Then
            GetFirstFreeDevice = Chr$(FIRST_DEVICE_NAME + i) & ":"
            Exit Function
        End If
    Next i
End Function

Private Function ListDevices() As String
    Dim Devices As Long, i As Integer
    
    Devices = GetLogicalDrives
    
    For i = 0 To 25
        If (Devices And 2 ^ i) <> 0 Then
            ListDevices = ListDevices & Chr$(FIRST_DEVICE_NAME + i)
        End If
    Next i
End Function

Public Function GetFreeSpace(PathName As String) As Double
    Dim sCluster As Long, bSector As Long, nFreeCluster As Long, nCluster As Long
    Dim FreeSectors As Double
    
    GetDiskFreeSpace PathName, sCluster, bSector, nFreeCluster, nCluster
    
    FreeSectors = sCluster * nFreeCluster
    GetFreeSpace = FreeSectors * bSector
End Function
