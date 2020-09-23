VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeviceInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Virtual drive information"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   7575
      Begin MSComctlLib.ListView lstInfo 
         Height          =   2640
         Left            =   15
         TabIndex        =   1
         Top             =   105
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   3
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmDeviceInfo.frx":0000
      End
   End
End
Attribute VB_Name = "frmDeviceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... frmDeviceInfo
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This form show information about
'                the currently mounted virtual drive
'-------------------------------------------------------

Private Enum OPERATION
    GET_FILE_SIZE = True
    COUNT_FILES = False
End Enum

Private Sub Form_Load()
    Dim HalfSize As Double
    Dim ExtendedSize As Long
    
    lstInfo.ColumnHeaders(1).Width = lstInfo.Width / 3
    lstInfo.ColumnHeaders(2).Width = (lstInfo.Width / 3) * 2
    
    With lstInfo
        .ListItems.Add 1, , "Device"
        .ListItems(1).ListSubItems.Add 1, , DeviceID & "\"
        
        .ListItems.Add 2, , "Mapped pathname"
        .ListItems(2).ListSubItems.Add 1, , GetDeviceInfo(DeviceID)
        .ListItems(2).ListSubItems(1).ToolTipText = .ListItems(2).ListSubItems(1).Text
        
        .ListItems.Add 3, , "Total virtual size"
        HalfSize = GetFreeSpace(DeviceID & "\") / 2
        .ListItems(3).ListSubItems.Add 1, , FormatBytes(GetFreeSpace(DeviceID & "\")) & " (" & FormatBytes(HalfSize) & " reserved for VDISK)"
        
        .ListItems.Add 4, , "Available space"
        .ListItems(4).ListSubItems.Add 1, , FormatBytes(HalfSize)
        
        .ListItems.Add 5, , "Space used"
        .ListItems(5).ListSubItems.Add 1, , FormatBytes(GetDeviceFilesInfo(DeviceID & "\", GET_FILE_SIZE))
        
        .ListItems.Add 6, , "Number of files"
        .ListItems(6).ListSubItems.Add 1, , GetDeviceFilesInfo(DeviceID & "\", COUNT_FILES)
    End With
End Sub

Private Function GetDeviceFilesInfo(DeviceName As String, lpOperation As OPERATION) As Double
    Dim FileName As String
    
    FileName = Dir(DeviceName & "*")
    Do While FileName <> vbNullString
        If lpOperation = GET_FILE_SIZE Then
            GetDeviceFilesInfo = GetDeviceFilesInfo + FileLen(DeviceName & FileName)
        ElseIf lpOperation = COUNT_FILES Then
            GetDeviceFilesInfo = GetDeviceFilesInfo + 1
        End If
        FileName = Dir
    Loop
End Function

