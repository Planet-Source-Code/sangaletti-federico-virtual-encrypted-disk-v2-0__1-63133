Attribute VB_Name = "modGlobal"
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... modGlobal
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This is the core of the application
'                it contains all the functions to work
'                with VDISK
'-------------------------------------------------------

Option Explicit

Public Const SIGNATURE_MAGIC = &H6BF0F076

Public Const DISK_TYPE_NORMAL = &HF

Public Const MAX_NAME_LENGHT = 64
Public Const MAX_FILE_NUMBER = 256
Public Const MAX_PASSWORD_LENGHT = 32
Public Const MAX_DISK_NAME_LENGHT = 32
Public Const MAX_COMMENT_LENGHT = 256

Public Const FILL_PATTERN = &HFFFFFFFF

Public Type FILE_DESCRIPTOR
    Index As Long
    fName(MAX_NAME_LENGHT - 1) As Byte
    Size As Long
    CreationDate As Long
    StartOffset As Long
    Comment(MAX_COMMENT_LENGHT - 1) As Byte
    CRC32 As Long
End Type

Public Type DISK_DESCRIPTOR
    Magic As Long
    DiskName(MAX_DISK_NAME_LENGHT - 1) As Byte
    Type As Long
    Password(MAX_PASSWORD_LENGHT - 1) As Byte
    MaxFileCapacity As Long
    MaxNameLenght As Long
    Reserved1 As Long
    Reserved2 As Long
    FilesDescriptorOffset As Long
    StartOfData As Long
End Type

Public Type FILE_DESCRIPTOR_OFFSET
    Index As Long
    Offset As Long
End Type

Public Const INVALID_HANDLE_VALUE = -1
Public Const CREATE_NEW = 1
Public Const OPEN_EXISTING = 3

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2


Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long

Public Enum ICON_INDEX
    ICON_UNKNOWN = 1
    ICON_TEXT = 2
    ICON_IMAGE = 3
    ICON_APPLICATION = 4
    ICON_WEB = 5
    ICON_COMPRESSED = 6
    ICON_AUDIO = 7
    ICON_VIDEO = 8
End Enum

Public Enum EncryptOperation
    EncryptData = True
    DecryptData = False
End Enum

Public Enum pMode
    SimpleMode = False
    AdvancedMode = True
End Enum

Private EngineCRC32 As clsCRC
Public EngineENCRYPT As clsCryptAPI

Private vdskPassword As String

Global ProgramMode As pMode
Global DeviceID As String

Public Function MountVirtualDisk(Path As String, DiskPassword As String) As Long
    Dim hVDisk As Long
    Dim DecryptBuffer As String
    
    'Open an existing VDISK
    hVDisk = CreateFile(Path, ByVal (GENERIC_READ Or GENERIC_WRITE), ByVal (FILE_SHARE_READ Or FILE_SHARE_WRITE), ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
    If hVDisk = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to open Virtual Disk", vbCritical
        MountVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Check the signature
    If GetDiskDescriptor(hVDisk).Magic <> SIGNATURE_MAGIC Then
        MsgBox "Invalid Virtual Disk image" & vbCrLf & "Invalid signature", vbCritical
        CloseHandle hVDisk
        MountVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Check compatibility
    If GetDiskDescriptor(hVDisk).MaxFileCapacity <> MAX_FILE_NUMBER Or GetDiskDescriptor(hVDisk).MaxNameLenght <> MAX_NAME_LENGHT Then
        MsgBox "Invalid Virtual Disk image" & vbCrLf & "Incompatible image format", vbCritical
        CloseHandle hVDisk
        MountVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Check the password
    DecryptBuffer = ArrayToString(GetDiskDescriptor(hVDisk).Password)
    If EngineENCRYPT.DecryptString(DecryptBuffer, DiskPassword) <> DiskPassword Then
        MsgBox "Invalid password" & vbCrLf & "Unable to mount Virtual Disk" & vbCrLf & vbCrLf & "The password is case-sensitive, check for CAPS-LOCK", vbCritical
        CloseHandle hVDisk
        MountVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Store the password for other operations
    vdskPassword = DiskPassword
    MountVirtualDisk = hVDisk
End Function

Public Function CreateVirtualDisk(Path As String, DiskName As String, DiskType As Long, DiskPassword() As Byte) As Long
    Dim hVDisk As Long, bRW As Long, rwValue As Long, i As Integer, SizeOfFileDescr As Long
    Dim vDisk As DISK_DESCRIPTOR
    Dim Files(MAX_FILE_NUMBER) As FILE_DESCRIPTOR
    Dim VolumeInfo() As Byte
    Dim CryptBuffer() As Byte
    
    'Check if already exist
    If Dir(Path) <> vbNullString Then
        MsgBox "Virtual disk called '" & Path & "' already existing" & vbCrLf & "Please delete it before creating new", vbExclamation
        CreateVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Create a new VDISK
    hVDisk = CreateFile(Path, ByVal (GENERIC_READ Or GENERIC_WRITE), ByVal (FILE_SHARE_READ Or FILE_SHARE_WRITE), ByVal 0&, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, ByVal 0)
    If hVDisk = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to create Virtual Disk", vbCritical
        CreateVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Fill the disk descriptor
    With vDisk
        .Magic = SIGNATURE_MAGIC
        StringToArray DiskName, .DiskName
        .Type = DiskType
        
        ReDim CryptBuffer(UBound(DiskPassword))
        StringToArray EngineENCRYPT.EncryptString(ArrayToString(DiskPassword), ArrayToString(DiskPassword)), CryptBuffer
        CopyMemory .Password(0), CryptBuffer(0), MAX_PASSWORD_LENGHT
        
        .MaxFileCapacity = MAX_FILE_NUMBER
        .MaxNameLenght = MAX_NAME_LENGHT
        .Reserved1 = 0
        .Reserved2 = 0
        .FilesDescriptorOffset = (Len(vDisk))
        
        SizeOfFileDescr = Len(Files(0))
        .StartOfData = .FilesDescriptorOffset + (SizeOfFileDescr * MAX_FILE_NUMBER)
    End With
    
    'Write the file descriptor al VDISK beginning
    SetFilePointer hVDisk, 0, 0, FILE_BEGIN
    rwValue = WriteFile(hVDisk, ByVal vDisk, Len(vDisk), bRW, ByVal 0&)
    If rwValue = 0 Then
        MsgBox "Unable to write disk decriptor", vbCritical
        CloseHandle hVDisk
        CreateVirtualDisk = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Process all file descriptors
    For i = 0 To MAX_FILE_NUMBER - 1
        'Zero file descriptor
        With Files(i)
            .Index = CLng(i + 1)
            'ZeroArray (.fName)
            .CreationDate = FILL_PATTERN
            .Size = FILL_PATTERN
            .StartOffset = FILL_PATTERN
            'ZeroArray (.Comment)
            .CRC32 = FILL_PATTERN
        End With
        
        'Write descriptor to VDISK
        rwValue = WriteFile(hVDisk, ByVal Files(i), Len(Files(i)), bRW, ByVal 0&)
        If rwValue = 0 Then
            MsgBox "Unable to write file decriptors", vbCritical
            CloseHandle hVDisk
            CreateVirtualDisk = INVALID_HANDLE_VALUE
            Exit Function
        End If
    Next i
    
    CloseHandle hVDisk
    CreateVirtualDisk = 1
End Function

Public Function AddFile(vDiskHandle As Long, Name As String, Size As Long, FileData() As Byte) As Boolean
    Dim ExtendedSize As Long, bRW As Long, rwValue As Long
    Dim FileDescr As FILE_DESCRIPTOR
    Dim FileDescrOffset As FILE_DESCRIPTOR_OFFSET
    
    Set EngineCRC32 = New clsCRC
    
    'Check if file can stored
    FileDescrOffset = GetFreeFileDescriptor(vDiskHandle)
    If FileDescrOffset.Offset = 0 Then
        MsgBox "This Virtual Disk can store maximum " & MAX_FILE_NUMBER & " files" & vbCrLf & "Not enough space for descriptor", vbExclamation
        AddFile = False
        Set EngineCRC32 = Nothing
        Exit Function
    End If
    
    'Point to the first free descriptor
    SetFilePointer vDiskHandle, FileDescrOffset.Offset, 0, 0
    
    'Fill the descriptor with file info
    With FileDescr
        .Index = FileDescrOffset.Index
        StringToArray Name, .fName
        .Size = Size
        .CreationDate = CLng(Date)
        .StartOffset = GetFileSize(vDiskHandle, ExtendedSize)
        If MsgBox("Do you want to add a comment for this file?", vbQuestion + vbYesNo) = vbYes Then StringToArray frmInput.GetInput("Please enter a comment for this file", MAX_COMMENT_LENGHT), .Comment
        .CRC32 = EngineCRC32.CalculateBytes(FileData)
    End With
    
    'Encrypt the decriptor
    EncryptFileDescriptor FileDescr, EncryptData
    
    'Write the new decriptor to VDISK
    rwValue = WriteFile(vDiskHandle, ByVal FileDescr, Len(FileDescr), bRW, ByVal 0&)
    If rwValue = 0 Then GoTo WriteError
    
    'Point to RAW data section
    SetFilePointer vDiskHandle, GetFileSize(vDiskHandle, ExtendedSize), 0, 0
    
    'Encrypt RAW file data and write to VDISK
    EngineENCRYPT.EncryptByte FileData, vdskPassword
    rwValue = WriteFile(vDiskHandle, FileData(0), UBound(FileData) + 1, bRW, ByVal 0&)
    If rwValue = 0 Then GoTo WriteError
    
    Set EngineCRC32 = Nothing
    
    AddFile = True
    Exit Function
    
WriteError:
    MsgBox "Unable to write file to Virtual Disk.", vbExclamation
    AddFile = False
    Set EngineCRC32 = Nothing
End Function

Public Function DeleteFile(vDiskHandle As Long, FileIndex As Long) As Boolean
    Dim Descr As FILE_DESCRIPTOR
    Dim Buffer() As Byte
    Dim rwResult As Long, bRW As Long
    
    'Get the current descriptor
    Descr = GetFileDescriptor(vDiskHandle, FileIndex)
    
    ReDim Buffer(Descr.Size - 1)
    
    'Point to file's RAW data
    SetFilePointer vDiskHandle, Descr.StartOffset, 0, 0
    
    'Overwrite RAW data with null
    rwResult = WriteFile(vDiskHandle, Buffer(0), Descr.Size, bRW, ByVal 0&)
    If rwResult = 0 Then
        MsgBox "Unable to remove file data", vbCritical
        DeleteFile = False
        Exit Function
    End If
    
    'Overwrite the decriptor with null
    If DeleteFileDescriptor(vDiskHandle, FileIndex) Then
        DeleteFile = True
    Else
        DeleteFile = False
    End If
End Function

Public Function GetFreeFileDescriptor(vDiskHandle As Long) As FILE_DESCRIPTOR_OFFSET
    Dim bRW As Long
    Dim FileDescr As FILE_DESCRIPTOR
    
    'Point to first descriptor
    SetFilePointer vDiskHandle, GetDiskDescriptor(vDiskHandle).FilesDescriptorOffset, 0, 0
    
    'Find the first free descriptor
    Do
        ReadFile vDiskHandle, FileDescr, Len(FileDescr), bRW, ByVal 0&
    Loop Until FileDescr.fName(0) = 0 Or FileDescr.Index > MAX_FILE_NUMBER
    
    If FileDescr.Index > MAX_FILE_NUMBER Or FileDescr.fName(0) <> 0 Then
        GetFreeFileDescriptor.Offset = 0
        GetFreeFileDescriptor.Index = 0
    Else
        GetFreeFileDescriptor.Offset = SetFilePointer(vDiskHandle, 0, 0, FILE_CURRENT) - Len(FileDescr)
        GetFreeFileDescriptor.Index = FileDescr.Index
    End If
End Function

Public Function GetDiskDescriptor(vDiskHandle As Long) As DISK_DESCRIPTOR
    Dim bRW As Long
    
    'Point to VDISK beginning
    SetFilePointer vDiskHandle, 0, 0, FILE_BEGIN
    
    'Read the disk descriptor
    ReadFile vDiskHandle, GetDiskDescriptor, Len(GetDiskDescriptor), bRW, ByVal 0&
End Function

Public Function GetFileDescriptor(vDiskHandle As Long, FileIndex As Long) As FILE_DESCRIPTOR
    Dim bRW As Long, i As Long
    
    i = 0
    'Point to first file descriptor
    SetFilePointer vDiskHandle, GetDiskDescriptor(vDiskHandle).FilesDescriptorOffset, 0, 0
    Do
        'Find the specified descriptor
        ReadFile vDiskHandle, GetFileDescriptor, Len(GetFileDescriptor), bRW, ByVal 0&
        i = i + 1
    Loop Until GetFileDescriptor.Index = FileIndex Or i > MAX_FILE_NUMBER
    
    'Decrypt the descriptor
    EncryptFileDescriptor GetFileDescriptor, DecryptData
End Function

Public Function DeleteFileDescriptor(vDiskHandle As Long, FileIndex As Long) As Boolean
    Dim bRW As Long, SizeOfFileDescr As Long, rwResult As Long
    Dim Descr As FILE_DESCRIPTOR
    
    SizeOfFileDescr = Len(Descr)
    'Point to specified descriptor
    SetFilePointer vDiskHandle, GetDiskDescriptor(vDiskHandle).FilesDescriptorOffset + ((FileIndex - 1) * SizeOfFileDescr), 0, 0
    
    'Fill the new descriptor with null
    With Descr
        .fName(0) = 0 'ZeroArray (.fName)
        .Size = FILL_PATTERN
        .CreationDate = FILL_PATTERN
        .StartOffset = FILL_PATTERN
        'ZeroArray (.Comment)
        .CRC32 = FILL_PATTERN
    End With
   
    'Write the new descriptor
    rwResult = WriteFile(vDiskHandle, Descr, Len(Descr), bRW, ByVal 0&)
    If rwResult = 0 Then
        MsgBox "Unable to remove file descriptor", vbCritical
        DeleteFileDescriptor = False
        Exit Function
    End If
    
    DeleteFileDescriptor = True
End Function

Public Function GetFileDataFromDescriptor(vDiskHandle As Long, Descriptor As FILE_DESCRIPTOR, ByRef FileData() As Byte) As Boolean
    Dim bRW As Long, CRCBuffer As Long
    
    Set EngineCRC32 = New clsCRC
    
    'Point to RAW data
    SetFilePointer vDiskHandle, Descriptor.StartOffset, 0, 0
    ReadFile vDiskHandle, FileData(0), UBound(FileData) + 1, bRW, ByVal 0&
    'Decrypt the RAW data
    EngineENCRYPT.DecryptByte FileData, vdskPassword
    
    'Check the CRC32 of decrypted data
    CRCBuffer = EngineCRC32.CalculateBytes(FileData)
    If CRCBuffer <> Descriptor.CRC32 Then
        If MsgBox("CRC32 Checksum Error" & vbCrLf & "File may be corrupted" & vbCrLf & vbCrLf & "Current: " & Hex$(CRCBuffer) & vbCrLf & "Expected: " & Hex$(Descriptor.CRC32) & vbCrLf & vbCrLf & "Extract and execute anyway?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then GetFileDataFromDescriptor = True Else GetFileDataFromDescriptor = False
        Set EngineCRC32 = Nothing
        Exit Function
    End If
    
    Set EngineCRC32 = Nothing
    GetFileDataFromDescriptor = True
End Function

Public Sub UpdateFileList(vDiskHandle As Long, objList As ListView)

'Fill the ListView in main form

    Dim i As Integer, bRW As Long
    Dim FileDescr As FILE_DESCRIPTOR
    
    objList.ListItems.Clear
    
    SetFilePointer vDiskHandle, GetDiskDescriptor(vDiskHandle).FilesDescriptorOffset, 0, 0
    
    For i = 1 To MAX_FILE_NUMBER
        ReadFile vDiskHandle, FileDescr, Len(FileDescr), bRW, ByVal 0&
        
        If FileDescr.fName(0) <> 0 Then
        
            EncryptFileDescriptor FileDescr, DecryptData
        
            objList.ListItems.Add objList.ListItems.Count + 1, "k" & FileDescr.Index, ArrayToString(FileDescr.fName), , GetIconIndex(ArrayToString(FileDescr.fName))
            objList.ListItems(objList.ListItems.Count).Tag = ArrayToString(FileDescr.Comment)
            objList.ListItems(objList.ListItems.Count).ListSubItems.Add , , FileDescr.Size & " bytes"
            objList.ListItems(objList.ListItems.Count).ListSubItems.Add , , CDate(FileDescr.CreationDate)
            objList.ListItems(objList.ListItems.Count).ListSubItems.Add , , Hex$(FileDescr.CRC32)
        End If
    Next i
End Sub

Public Sub GetFileDataFromDisk(Path As String, ByRef FileData() As Byte)
    Dim hFile As Long, bRW As Long, hResult As Long
    
    'Open the specified file
    hFile = CreateFile(Path, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
    If hFile = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to open '" & Path & "'", vbCritical
        Exit Sub
    End If
    
    'Read the RAW data
    hResult = ReadFile(hFile, FileData(0), UBound(FileData) + 1, bRW, ByVal 0&)
    If hResult = 0 Then
        MsgBox "Unable to read from '" & Path & "'", vbCritical
        CloseHandle hFile
        Exit Sub
    End If
    
    CloseHandle hFile
End Sub

Public Function WriteTempFile(Path As String, RAWData() As Byte) As Long
    Dim hFile As Long, bRW As Long, rwResult As Long
    
    'Check if file existing
    If Dir(Path) <> vbNullString Then
        WriteTempFile = 1
        Exit Function
    End If
    
    'Create a new file
    hFile = CreateFile(Path, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, ByVal 0)
    If hFile = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to create temporary file", vbCritical
        WriteTempFile = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    'Write the decrypted RAW data
    rwResult = WriteFile(hFile, RAWData(0), UBound(RAWData) + 1, bRW, ByVal 0&)
    If rwResult = 0 Then
        MsgBox "Unable to write temporary file", vbCritical
        CloseHandle hFile
        WriteTempFile = INVALID_HANDLE_VALUE
        Exit Function
    End If
    
    CloseHandle hFile
    WriteTempFile = 0
End Function

Public Function ArrayToString(strData() As Byte) As String
    Dim i As Integer
    
    For i = 0 To UBound(strData)
        If strData(i) <> 0 Then ArrayToString = ArrayToString & Chr$(strData(i))
    Next i
End Function

Public Sub StringToArray(strData As String, ByRef Buffer() As Byte)
    Dim i As Integer
    
    For i = 0 To Len(strData) - 1
        Buffer(i) = Asc(Mid(strData, i + 1, 1))
    Next i
End Sub

Public Sub ZeroArray(ByRef Buffer() As Byte)
    Dim i As Long
    
    For i = 0 To UBound(Buffer)
        Buffer(i) = 0
    Next i
End Sub

Public Function GetIconIndex(FileName As String) As ICON_INDEX
    
'Returns the icon index of specified file type
    
    Dim Pattern As String, i As Integer
    
    Pattern = vbNullString
    i = 1
    If InStr(1, FileName, ".") Then
        Do
            Pattern = Right(FileName, i)
            i = i + 1
        Loop Until Left(Pattern, 1) = "."
    End If
    
    Select Case LCase(Pattern)
        Case ".txt", ".doc", ".rtf", ".diz":
            GetIconIndex = ICON_TEXT
        
        Case ".bmp", ".jpg", ".gif", ".jpeg", ".psp", ".psd", ".png", ".tif", ".tiff", ".tga", ".emf", ".wmf", ".pcx", ".ico":
            GetIconIndex = ICON_IMAGE
        
        Case ".exe":
            GetIconIndex = ICON_APPLICATION
        
        Case ".htm", ".html", ".asp", ".php", ".aspx":
            GetIconIndex = ICON_WEB
        
        Case ".zip", ".rar", ".cab", ".gzip", ".tar", ".ace", ".arj", ".lzh", ".uue", ".z", ".jar", ".bz2"
            GetIconIndex = ICON_COMPRESSED
        
        Case ".wav", ".mp3", ".wma", ".mid", ".midi", ".rmi", ".aif", ".aiff", ".snd", ".au", ".ra", ".rm":
            GetIconIndex = ICON_AUDIO
            
        Case ".avi", ".mpg", ".mpeg", ".asf", ".wmv":
            GetIconIndex = ICON_VIDEO
            
            
        Case Else:
            GetIconIndex = ICON_UNKNOWN
    End Select
End Function

Public Sub EncryptFileDescriptor(ByRef Descriptor As FILE_DESCRIPTOR, OperationType As EncryptOperation)
    Dim strSize As String, strCreationDate As String, strStartOffset As String, strCRC32 As String
    
    With Descriptor
        strSize = Hex$(.Size)
        strCreationDate = Hex$(.CreationDate)
        strStartOffset = Hex$(.StartOffset)
        strCRC32 = Hex$(.CRC32)
        
        'Format the numeric fields to prevents conversion errors
        If Len(strSize) Mod 2 <> 0 Then strSize = "0" & strSize
        If Len(strCreationDate) Mod 2 <> 0 Then strCreationDate = "0" & strCreationDate
        If Len(strStartOffset) Mod 2 <> 0 Then strStartOffset = "0" & strStartOffset
        If Len(strCRC32) Mod 2 <> 0 Then strCRC32 = "0" & strCRC32
    
        If OperationType = EncryptData Then
            StringToArray EngineENCRYPT.EncryptString(StrConv(.fName, vbUnicode), vdskPassword), .fName
            .Size = CLng("&H" & StrToHex(EngineENCRYPT.EncryptString(HexToStr(strSize), vdskPassword)))
            .CreationDate = CLng("&H" & StrToHex(EngineENCRYPT.EncryptString(HexToStr(strCreationDate), vdskPassword)))
            .StartOffset = CLng("&H" & StrToHex(EngineENCRYPT.EncryptString(HexToStr(strStartOffset), vdskPassword)))
            If .Comment(0) <> 0 Then StringToArray EngineENCRYPT.EncryptString(StrConv(.Comment, vbUnicode), vdskPassword), .Comment
            .CRC32 = CLng("&H" & StrToHex(EngineENCRYPT.EncryptString(HexToStr(strCRC32), vdskPassword)))
            
        ElseIf OperationType = DecryptData Then
            StringToArray EngineENCRYPT.DecryptString(StrConv(.fName, vbUnicode), vdskPassword), .fName
            .Size = CLng("&H" & StrToHex(EngineENCRYPT.DecryptString(HexToStr(strSize), vdskPassword)))
            .CreationDate = CLng("&H" & StrToHex(EngineENCRYPT.DecryptString(HexToStr(strCreationDate), vdskPassword)))
            .StartOffset = CLng("&H" & StrToHex(EngineENCRYPT.DecryptString(HexToStr(strStartOffset), vdskPassword)))
            If .Comment(0) <> 0 Then StringToArray EngineENCRYPT.DecryptString(StrConv(.Comment, vbUnicode), vdskPassword), .Comment
            .CRC32 = CLng("&H" & StrToHex(EngineENCRYPT.DecryptString(HexToStr(strCRC32), vdskPassword)))
            
        End If
    End With
End Sub

Public Function ExtractAllFiles(vDiskHandle As Long, objProgress As XP_ProgressBar, objCurrentFile As Label) As Long
    Dim i As Integer, bRW As Long, tResult As Long
    Dim FileDescr As FILE_DESCRIPTOR
    Dim RAWData() As Byte
    Dim NextDescriptorOffset As Long
    Dim ExtraSize As Long
    
    objProgress.Max = Int(GetFileSize(vDiskHandle, ExtraSize) / 1000)
    objProgress.Value = 0
    
    SetFilePointer vDiskHandle, GetDiskDescriptor(vDiskHandle).FilesDescriptorOffset, 0, 0
    
    For i = 1 To MAX_FILE_NUMBER
        ReadFile vDiskHandle, FileDescr, Len(FileDescr), bRW, ByVal 0&
        NextDescriptorOffset = SetFilePointer(vDiskHandle, 0, 0, FILE_CURRENT)
        
        If FileDescr.fName(0) <> 0 Then
            'Decrypt the file descriptor
            EncryptFileDescriptor FileDescr, DecryptData
            
            'Read the RAW data
            ReDim RAWData(FileDescr.Size - 1)
            If Not GetFileDataFromDescriptor(vDiskHandle, FileDescr, RAWData) Then Exit Function
            
            objCurrentFile.Caption = " Current file: " & ArrayToString(FileDescr.fName)
            If objProgress.Value < objProgress.Max Then objProgress.Value = objProgress.Value + Int(Int(FileDescr.Size / 1000) / 2)
            DoEvents
            
            'Write the RAW data in a temp file
            tResult = WriteTempFile(IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(vDiskHandle).DiskName) & "\" & ArrayToString(FileDescr.fName), RAWData)
            If tResult = INVALID_HANDLE_VALUE Then Exit Function
            
            If objProgress.Value < objProgress.Max Then objProgress.Value = objProgress.Value + Int(Int(FileDescr.Size / 1000) / 2)
            DoEvents
            
            'Point to the next descriptor
            SetFilePointer vDiskHandle, NextDescriptorOffset, 0, 0
        End If
    Next i
    
    objProgress.Value = objProgress.Max
End Function

Public Function FormatBytes(nBytes As Double) As String
    If nBytes < 1024 Then
        FormatBytes = nBytes
        FormatBytes = FormatBytes & IIf(FormatBytes = "1", " Byte", " Bytes")
    ElseIf nBytes / 1024 < 1024 Then
        FormatBytes = Format$((nBytes / 1024), ".0")
        FormatBytes = FormatBytes & IIf(FormatBytes = "1.0", " KByte", " KBytes")
    ElseIf nBytes / 1048576 < 1024 Then
        FormatBytes = Format$((nBytes / 1048576), ".00")
        FormatBytes = FormatBytes & IIf(FormatBytes = "1.00", " MByte", " MBytes")
    Else
        FormatBytes = Format$((nBytes / 1073741824), ".00")
        FormatBytes = FormatBytes & IIf(FormatBytes = "1.00", " GByte", " GBytes")
    End If
End Function
