VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
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
   ScaleHeight     =   7395
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
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
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   6735
      Begin VB.Image Picture4 
         Height          =   1830
         Left            =   6840
         Picture         =   "frmAbout.frx":0000
         Top             =   120
         Width           =   2445
      End
      Begin VB.Image Picture3 
         Height          =   1365
         Left            =   6840
         Picture         =   "frmAbout.frx":25CC
         Top             =   3120
         Width           =   6315
      End
      Begin VB.Image Picture2 
         Height          =   1395
         Left            =   6840
         Picture         =   "frmAbout.frx":5470
         Top             =   1680
         Width           =   5775
      End
      Begin VB.Image Picture1 
         Height          =   1440
         Left            =   6840
         Picture         =   "frmAbout.frx":7F84
         Top             =   240
         Width           =   4410
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   240
         TabIndex        =   1
         Top             =   4560
         Width           =   6255
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   7095
         Left            =   10
         Top             =   105
         Width           =   6705
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... frmAbout
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This is the about form
'-------------------------------------------------------

Private Sub Form_Load()
    Label1.Caption = _
    "Copyright 2005 Sangaletti Federico" & vbCrLf & _
    "email: sangaletti@aliceposta.it" & vbCrLf & vbCrLf & _
    "Version 2.0 features" & vbCrLf & _
    " > Strong encryption with CryptoAPI" & vbCrLf & _
    " > Datas is really unaccessible without password" & vbCrLf & _
    " > Very fast read/write with Win32 API" & vbCrLf & _
    " > Fixed some bugs (CRC32, Encrypt/Decrypt routine)" & vbCrLf & _
    " > Added 2 working modes (Simple and Advanced)" & vbCrLf & _
    " > *REAL* virtual disk access via Windows Explorer"
    
    
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If Picture1.Left > 300 Then
        Picture1.Left = Picture1.Left - 200
    ElseIf Picture2.Left > 300 Then
        Picture2.Left = Picture2.Left - 200
    ElseIf Picture3.Left > 300 Then
        Picture3.Left = Picture3.Left - 200
    ElseIf Picture4.Left > 4200 Then
        Picture4.Left = Picture4.Left - 100
    ElseIf Label1.Height < 2415 Then
        Label1.Height = Label1.Height + 50
    Else
        Timer1.Enabled = False
    End If
End Sub
