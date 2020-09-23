VERSION 5.00
Begin VB.Form frmTrayIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tray"
   ClientHeight    =   1785
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3375
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.Image menuDeviceInfo 
      Height          =   210
      Left            =   3000
      Picture         =   "frmTrayIcon.frx":058A
      Top             =   360
      Width           =   210
   End
   Begin VB.Image menuExplore 
      Height          =   210
      Left            =   3000
      Picture         =   "frmTrayIcon.frx":0AAC
      Top             =   0
      Width           =   210
   End
   Begin VB.Image DriveInfo 
      Height          =   210
      Left            =   3000
      Picture         =   "frmTrayIcon.frx":0FCE
      Top             =   1440
      Width           =   210
   End
   Begin VB.Image menuAbout 
      Height          =   210
      Left            =   3000
      Picture         =   "frmTrayIcon.frx":14F0
      Top             =   1080
      Width           =   210
   End
   Begin VB.Image menuShow 
      Height          =   210
      Left            =   3000
      Picture         =   "frmTrayIcon.frx":1A12
      Top             =   720
      Width           =   210
   End
   Begin VB.Image TrayLogo 
      Height          =   840
      Left            =   120
      Picture         =   "frmTrayIcon.frx":1F34
      Top             =   240
      Width           =   2790
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Begin VB.Menu Logo 
         Caption         =   "Logo"
         Enabled         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu Explore 
         Caption         =   "Explore virtual drive"
      End
      Begin VB.Menu DeviceInfo 
         Caption         =   "Virtual drive info"
      End
      Begin VB.Menu Show 
         Caption         =   "Show main program window"
      End
      Begin VB.Menu About 
         Caption         =   "Some info about this program"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu VDISK_Drive 
         Caption         =   "VDISK Mounted on Drive"
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... frmTrayIcon
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This form show the menu in system tray
'-------------------------------------------------------

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10
Private Const MIIM_ID = &H2
Private Const MFT_STRING = &H0&
Private Const SW_SHOWNORMAL = 1

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const MENU_EXPLORE = 2
Private Const MENU_DEVICE = 3
Private Const MENU_SHOW = 4
Private Const MENU_ABOUT = 5
Private Const MENU_INFO = 7

Private Sub About_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub DeviceInfo_Click()
    frmDeviceInfo.Show vbModal
End Sub

Private Sub Explore_Click()
    If ShellExecute(0, "explore", DeviceID & "\", vbNullString, vbNullString, SW_SHOWNORMAL) <= 32 Then MsgBox "Unable to explore drive " & DeviceID & vbCrLf & vbCrLf & "Please explore manually", vbExclamation
End Sub

Private Sub Form_Load()
    SetMenuLogo Array(0, 0), TrayLogo.Picture
    
    SetMenuBitmaps MENU_EXPLORE, menuExplore.Picture
    SetMenuBitmaps MENU_DEVICE, menuDeviceInfo.Picture
    SetMenuBitmaps MENU_SHOW, menuShow.Picture
    SetMenuBitmaps MENU_ABOUT, menuAbout.Picture
    SetMenuBitmaps MENU_INFO, DriveInfo.Picture
    
    VDISK_Drive.Caption = "VDISK is mounted on drive " & DeviceID
    TrayAdd Me.hwnd, Me.Icon, "VIRTUAL ENCRYPTED DISK UTILITY 2.0" & vbCrLf & "VDISK is mounted on drive " & DeviceID, MouseMove
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cEvent As Single
    
    cEvent = X / Screen.TwipsPerPixelX
    
    If cEvent = RightUp Then PopupMenu TrayMenu
End Sub

Private Sub Show_Click()
    frmMain.Show
    TrayDelete
    Unload Me
End Sub

Public Sub SetMenuLogo(ByVal item_numbers As Variant, ByVal pic As Picture)
    Dim menu_handle As Long
    Dim i As Integer
    Dim MENU_INFO As MENUITEMINFO

    menu_handle = GetMenu(Me.hwnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i
    
    With MENU_INFO
        .cbSize = Len(MENU_INFO)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With
    
    SetMenuItemInfo menu_handle, item_numbers(UBound(item_numbers)), True, MENU_INFO
End Sub

Public Sub SetMenuBitmaps(ItemNumber As Integer, hPicture As Long)
    Dim hMenu As Long, hSubMenu As Long, hID As Long
    
    hMenu& = GetMenu(Me.hwnd)
    hSubMenu& = GetSubMenu(hMenu&, 0)
    hID& = GetMenuItemID(hSubMenu&, ItemNumber)
    
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, hPicture, 0
End Sub
