VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "VIRTUAL ENCRYPTED DISK UTILITY v2.0"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmAdvanced 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame Frame6 
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   3600
         Width           =   5900
         Begin VB.Label lblCurrentFile 
            BackColor       =   &H80000005&
            Caption         =   " Current File: N/A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   15
            TabIndex        =   15
            Top             =   105
            Width           =   5850
         End
      End
      Begin VB.Frame Frame4 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   3600
         Begin VB.Label lblOperation 
            BackColor       =   &H80000005&
            Caption         =   " Operation: Ready..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   15
            TabIndex        =   13
            Top             =   105
            Width           =   3555
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3240
         TabIndex        =   10
         Top             =   2160
         Width           =   3375
         Begin VB.CommandButton cmdAdvanced 
            Caption         =   "Create virtual drive"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   850
            Left            =   10
            TabIndex        =   11
            Top             =   110
            Width           =   3340
         End
      End
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
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   9615
         Begin VB.Label lblAdvancedDescription 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   8
            Top             =   345
            Width           =   9255
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1440
            Left            =   15
            TabIndex        =   9
            Top             =   105
            Width           =   9570
         End
      End
      Begin vDisk.XP_ProgressBar XP_ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   12937777
         Scrolling       =   1
         ShowText        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   9855
      Begin MSComctlLib.ImageList imgFiles 
         Left            =   6120
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1CFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2294
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":282E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3362
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":38FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4430
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   4320
         Left            =   15
         TabIndex        =   4
         Top             =   105
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   7620
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   3
         _Version        =   393217
         SmallIcons      =   "imgFiles"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   6546
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Creation Date"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Checksum"
            Object.Width           =   2540
         EndProperty
         Picture         =   "Form1.frx":49CA
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8705
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Simple Mode"
            Key             =   "kSimpleMode"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced Mode"
            Key             =   "kAdvancedMode"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgMount 
      Left            =   4920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbMount 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6525
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   12647
            MinWidth        =   2646
            Text            =   "No Virtual Disk Mounted"
            TextSave        =   "No Virtual Disk Mounted"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Label: N/A"
            TextSave        =   "Label: N/A"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2470
            Text            =   "Size: N/A"
            TextSave        =   "Size: N/A"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList lstMenuDisabled 
      Left            =   6000
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AB16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2DB68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3CD02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList lstMenu2 
      Left            =   5400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3FD54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4BE9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4EEEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":54F92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList lstMenu1 
      Left            =   4800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B036
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E088
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":610DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6412C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6717E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A1D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D222
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   1290
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2275
      ButtonWidth     =   2170
      ButtonHeight    =   2223
      Appearance      =   1
      Style           =   1
      ImageList       =   "lstMenu1"
      DisabledImageList=   "lstMenuDisabled"
      HotImageList    =   "lstMenu2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mount VDISK"
            Key             =   "kMountVDISK"
            Object.ToolTipText     =   "Mount an existing Virtual Disk"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Unmount VDISK"
            Key             =   "kUnmountVDISK"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Create VDISK"
            Key             =   "kCreateVDISK"
            Object.ToolTipText     =   "Create a new Virtual Disk"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Volume Info"
            Key             =   "kVolumeInfo"
            Object.ToolTipText     =   "Show Virtual Disk volume information"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Add File"
            Key             =   "kAddFile"
            Object.ToolTipText     =   "Add a new file to Virtual Disk"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete File"
            Key             =   "kDeleteFile"
            Object.ToolTipText     =   "Delete a file from Virtual Disk"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "kAbout"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "kExit"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... frmMain
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This is the main window
'-------------------------------------------------------


Option Explicit

Const BUTTON_MOUNT = 1
Const BUTTON_UNMOUNT = 2
Const BUTTON_CREATE = 3
Const BUTTON_VOLUME_INFO = 4
Const BUTTON_ADD_FILE = 6
Const BUTTON_DELETE_FILE = 7

Const PANEL_VDISK_MOUNTED = 1
Const PANEL_VDISK_LABEL = 2
Const PANEL_VDISK_SIZE = 3

Const SW_SHOWNORMAL = 1
Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private TT As CTooltip
Private m_lCurItemIndex As Long

Dim hDisk As Long

Private Sub cmdAdvanced_Click()
    Dim DeviceResult As String
    
    If ProgramMode = AdvancedMode Then
        Load frmTrayIcon
        Me.Hide
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    DeviceResult = CreateDevice(GetFirstFreeDevice, IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName))
    If DeviceResult = vbNullString Then
        Exit Sub
    End If
    
    DeviceID = DeviceResult
    MsgBox "Virtual drive is mapped to " & DeviceID & vbCrLf & vbCrLf & "Decrypting files data...", vbInformation
    lblOperation.Caption = " Operation: Decrypting files"
    
    ExtractAllFiles hDisk, XP_ProgressBar1, lblCurrentFile
    
    ProgramMode = AdvancedMode
    lblOperation.Caption = " Operation: Ready..."
    lblCurrentFile.Caption = " Current file: N/A"
    Me.MousePointer = vbDefault
    
    Load frmTrayIcon
    Me.Hide
    
    MsgBox "Now you can access to your files direcly from Explorer", vbInformation
End Sub

Private Sub Form_Load()
    Set TT = New CTooltip
    
    Set EngineENCRYPT = New clsCryptAPI
    
    TT.Style = TTBalloon
    TT.Icon = TTIconInfo
    
    Load frmInput
    Load frmPassword
    
    lblAdvancedDescription.Caption = _
    "The 'Advanced Mode' creates a virtual drive accessible via Explorer." & vbCrLf & _
    "You can perform all possible file operations, copy, modify, delete, rename, etc." & vbCrLf & vbCrLf & _
    "When the Virtual Disk is unmounted the virtual drive are wiped."
    
    ProgramMode = SimpleMode
    DeviceID = vbNullString
    
    If CreateIcon Then WriteFileType
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    TabStrip1.Height = Me.Height - 2320
    TabStrip1.Width = Me.Width - 360
    
    Frame2.Width = TabStrip1.Width - 240
    Frame2.Height = TabStrip1.Height - 480
    frmAdvanced.Width = TabStrip1.Width - 240
    frmAdvanced.Height = TabStrip1.Height - 480
    
    lstFiles.Width = Frame2.Width - 45
    lstFiles.Height = Frame2.Height - 135
    
    Frame1.Width = frmAdvanced.Width - 240
    Frame1.Height = frmAdvanced.Height - 2880
    Label1.Width = Frame1.Width - 45
    Label1.Height = Frame1.Height - 135
    lblAdvancedDescription.Width = Label1.Width - 315
    lblAdvancedDescription.Height = Label1.Height - 465
    
    Frame3.Top = Frame1.Height + 600
    Frame3.Left = frmAdvanced.Width / 3
    
    Frame4.Width = frmAdvanced.Width / 2.74
    Frame4.Top = frmAdvanced.Height - 855
    lblOperation.Width = Frame4.Width - 45
    
    Frame6.Left = Frame4.Width + Frame4.Left + 120
    Frame6.Width = (frmAdvanced.Width / 1.67)
    Frame6.Top = frmAdvanced.Height - 855
    lblCurrentFile.Width = Frame6.Width - 45
    
    XP_ProgressBar1.Width = frmAdvanced.Width - 240
    XP_ProgressBar1.Top = frmAdvanced.Height - 375
    
    With lstFiles
        .ColumnHeaders(1).Width = .Width / 2
        .ColumnHeaders(2).Width = (.Width / 2) / 3
        .ColumnHeaders(3).Width = (.Width / 2) / 3
        .ColumnHeaders(4).Width = (.Width / 2) / 3
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hDisk <> 0 And hDisk <> INVALID_HANDLE_VALUE Then
        MsgBox "Please unmount Virtual Disk first", vbExclamation
        Cancel = 1
    Else
        Unload frmInput
        Unload frmPassword
    End If
End Sub

Private Sub lstFiles_DblClick()
    On Error Resume Next
    Dim Descriptor As FILE_DESCRIPTOR
    Dim RAWData() As Byte
    Dim tResult As Long
    
    If lstFiles.SelectedItem.Key <> vbNullString Then
        Me.MousePointer = vbHourglass
        
        Descriptor = GetFileDescriptor(hDisk, CLng(Right(lstFiles.SelectedItem.Key, Len(lstFiles.SelectedItem.Key) - 1)))
        ReDim RAWData(Descriptor.Size - 1)

        If Not GetFileDataFromDescriptor(hDisk, Descriptor, RAWData) Then Exit Sub
        
        tResult = WriteTempFile(IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName) & "\" & ArrayToString(Descriptor.fName), RAWData)
        If tResult <> INVALID_HANDLE_VALUE Then
            If ShellExecute(0, "open", IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName) & "\" & ArrayToString(Descriptor.fName), vbNullString, vbNullString, SW_SHOWNORMAL) <= 32 Then MsgBox "Unable to run '" & IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName) & "\" & ArrayToString(Descriptor.fName) & "'", vbCritical
        End If
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
   
    lvhti.pt.X = X / Screen.TwipsPerPixelX
    lvhti.pt.Y = Y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(lstFiles.hwnd, LVM_HITTEST, 0, lvhti) + 1
    
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        If m_lCurItemIndex = 0 Or lstFiles.ListItems(m_lCurItemIndex).Tag = vbNullString Then    ' no item under the mouse pointer
            TT.Destroy
        ElseIf lstFiles.ListItems(m_lCurItemIndex).Tag <> vbNullString Then
            TT.Title = lstFiles.ListItems(m_lCurItemIndex).Text
            TT.TipText = lstFiles.ListItems(m_lCurItemIndex).Tag
            TT.Create lstFiles.hwnd
        End If
    End If
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Key
        Case "kSimpleMode":
            Frame2.Visible = True
            frmAdvanced.Visible = False
            
        Case "kAdvancedMode":
            Frame2.Visible = False
            frmAdvanced.Visible = True
    End Select
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim pwdArray(MAX_PASSWORD_LENGHT - 1) As Byte
    Dim CreateResult As Long
    Dim FileDataArray() As Byte, FileSize As Long, ExtendedSize As Long
    
    Select Case Button.Key
        Case "kMountVDISK":
            dlgMount.FileName = vbNullString
            dlgMount.DialogTitle = "Select a VDISK image"
            dlgMount.Filter = "Virtual Disk Image [.vdsk]|*.vdsk"
            dlgMount.ShowOpen
            If dlgMount.FileName <> vbNullString Then
                hDisk = MountVirtualDisk(dlgMount.FileName, frmPassword.GetInput(MAX_PASSWORD_LENGHT))
                If hDisk <> INVALID_HANDLE_VALUE Then
                    stbMount.Panels(PANEL_VDISK_MOUNTED).Text = "Mounted: " & dlgMount.FileName
                    stbMount.Panels(PANEL_VDISK_LABEL).Text = "Label: " & ArrayToString(GetDiskDescriptor(hDisk).DiskName)
                    stbMount.Panels(PANEL_VDISK_SIZE).Text = "Size: " & FormatBytes(FileLen(dlgMount.FileName))
                    
                    tlbMenu.Buttons(BUTTON_MOUNT).Enabled = False
                    tlbMenu.Buttons(BUTTON_UNMOUNT).Enabled = True
                    tlbMenu.Buttons(BUTTON_ADD_FILE).Enabled = True
                    tlbMenu.Buttons(BUTTON_VOLUME_INFO).Enabled = True
                    tlbMenu.Buttons(BUTTON_DELETE_FILE).Enabled = True
                    
                    cmdAdvanced.Enabled = True
                    
                    UpdateFileList hDisk, lstFiles
                    
                    If CreateDirectory(IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName), ByVal 0&) = 0 Then
                        MsgBox "Unable to create temporary directory for this VDISK", vbExclamation
                    End If
                End If
            End If
        
        Case "kUnmountVDISK":
            If ProgramMode = AdvancedMode Then
                RemoveDevice DeviceID
                DeviceID = vbNullString
                ProgramMode = SimpleMode
            End If
            
            If Not WipeAllFiles(IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName)) Then
                MsgBox "Unable to delete temporary files" & vbCrLf & vbCrLf & "Please delete manually to preserve security", vbCritical
            End If
            
            If RemoveDirectory(IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & ArrayToString(GetDiskDescriptor(hDisk).DiskName)) = 0 Then
                MsgBox "Unable to remove temporary directory", vbExclamation
            End If
            
            CloseHandle hDisk
            hDisk = 0
            stbMount.Panels(PANEL_VDISK_MOUNTED).Text = "No Virtual Disk Mounted"
            stbMount.Panels(PANEL_VDISK_LABEL).Text = "Label: N/A"
            stbMount.Panels(PANEL_VDISK_SIZE).Text = "Size: N/A"
            MsgBox "Virtual Disk unmounted", vbInformation
            
            tlbMenu.Buttons(BUTTON_ADD_FILE).Enabled = False
            tlbMenu.Buttons(BUTTON_VOLUME_INFO).Enabled = False
            tlbMenu.Buttons(BUTTON_UNMOUNT).Enabled = False
            tlbMenu.Buttons(BUTTON_DELETE_FILE).Enabled = False
            tlbMenu.Buttons(BUTTON_MOUNT).Enabled = True
            
            cmdAdvanced.Enabled = False
            
            lstFiles.ListItems.Clear
        
        Case "kCreateVDISK":
            dlgMount.DialogTitle = "Specify a file name for VDISK"
            dlgMount.FileName = vbNullString
            dlgMount.Filter = "Virtual Disk Image [.vdsk]|*.vdsk"
            dlgMount.ShowSave
            If dlgMount.FileName <> vbNullString Then
                StringToArray frmPassword.GetInput(MAX_PASSWORD_LENGHT), pwdArray
                CreateResult = CreateVirtualDisk(dlgMount.FileName, UCase(frmInput.GetInput("Please enter the label for this Virtual Disk", MAX_DISK_NAME_LENGHT)), DISK_TYPE_NORMAL, pwdArray)
                If CreateResult <> INVALID_HANDLE_VALUE Then MsgBox "Virtual Disk created successfully", vbInformation
            End If
        
        Case "kVolumeInfo":
            Load frmVolumeInfo
            With frmVolumeInfo.lstInfo
                .ListItems.Add 1, , "Volume label"
                .ListItems(1).ListSubItems.Add 1, , ArrayToString(GetDiskDescriptor(hDisk).DiskName)
                
                .ListItems.Add 2, , "Files capacity"
                .ListItems(2).ListSubItems.Add 1, , GetDiskDescriptor(hDisk).MaxFileCapacity & " Files"
                
                .ListItems.Add 3, , "Max name lenght"
                .ListItems(3).ListSubItems.Add 1, , GetDiskDescriptor(hDisk).MaxNameLenght & " Chars"
                
                .ListItems.Add 4, , "VDISK type"
                .ListItems(4).ListSubItems.Add 1, , "0x" & Hex$(GetDiskDescriptor(hDisk).Type)
                
                .ListItems.Add 5, , "File Descriptors offset"
                .ListItems(5).ListSubItems.Add 1, , "0x" & Hex$(GetDiskDescriptor(hDisk).FilesDescriptorOffset)
                
                .ListItems.Add 6, , "RAW Data offset"
                .ListItems(6).ListSubItems.Add 1, , "0x" & Hex$(GetDiskDescriptor(hDisk).StartOfData)
            End With
            frmVolumeInfo.Show vbModal, Me
        
        Case "kAddFile":
            dlgMount.FileName = vbNullString
            dlgMount.DialogTitle = "Select a file to add to VDISK"
            dlgMount.Filter = "All Files [*.*]|*.*"
            dlgMount.ShowOpen
            If dlgMount.FileName <> vbNullString Then
                If Len(dlgMount.FileTitle) > MAX_NAME_LENGHT Then
                    MsgBox "File name is too long" & vbCrLf & "The maximum lenght is " & MAX_NAME_LENGHT, vbCritical
                    Exit Sub
                End If
                
                Me.MousePointer = vbHourglass
                
                FileSize = FileLen(dlgMount.FileName)
                ReDim FileDataArray(FileSize - 1)
                GetFileDataFromDisk dlgMount.FileName, FileDataArray
                AddFile hDisk, dlgMount.FileTitle, FileSize, FileDataArray
                
                UpdateFileList hDisk, lstFiles
                stbMount.Panels(PANEL_VDISK_SIZE).Text = "Size: " & FormatBytes(GetFileSize(hDisk, ExtendedSize))
                
                Me.MousePointer = vbDefault
            End If
        
        Case "kDeleteFile":
            If MsgBox("Are you sure to delete '" & lstFiles.SelectedItem.Text & "'?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                DeleteFile hDisk, CLng(Right(lstFiles.SelectedItem.Key, Len(lstFiles.SelectedItem.Key) - 1))
                
                UpdateFileList hDisk, lstFiles
            End If
        
        Case "kAbout":
            frmAbout.Show vbModal, Me
        
        Case "kExit":
            Unload Me
            
    End Select
End Sub
