VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVDISK 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   500
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5790
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   15
         TabIndex        =   1
         Top             =   105
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Max             =   100
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "frmVDISK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ProgressBar1.Value = ProgressBar1.Value + 1
End Sub
