VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input"
   ClientHeight    =   3135
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6015
   ControlBox      =   0   'False
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
   ScaleHeight     =   3135
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5775
      Begin VB.TextBox txtInput 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   220
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   15
         TabIndex        =   7
         Top             =   105
         Width           =   5730
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
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   50
      Width           =   5775
      Begin VB.TextBox txtDescription 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   5535
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
         Height          =   840
         Left            =   15
         TabIndex        =   4
         Top             =   105
         Width           =   5730
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... frmInput
'   Author...... Sangaletti Federico
'   email....... sangaletti@aliceposta.it
'   License..... FREE (but respect copyright of my work!)
'
'   Decription.. This form handles the user's input
'-------------------------------------------------------

Option Explicit

Private Declare Function SetFocusA Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private TT As CTooltip
Private OK As Boolean
Private Cancel As Boolean
Private MAX_INPUT_LEN As Integer

Private Sub CancelButton_Click()
    Cancel = True
End Sub

Public Function GetInput(Description As String, MAX_LEN As Integer) As String
    OK = False
    Cancel = False
    
    SetFocusA txtInput.hwnd
    
    MAX_INPUT_LEN = MAX_LEN
    txtDescription.Text = Description
    Me.Show
    
    Do
        DoEvents
    Loop Until OK = True Or Cancel = True
    
    If OK = True Then
        GetInput = txtInput.Text
    ElseIf Cancel = True Then
        GetInput = vbNullString
    End If
    
    txtInput.Text = vbNullString
    Me.Hide
End Function

Private Sub Form_Load()
    Set TT = New CTooltip
End Sub

Private Sub OKButton_Click()
    If Len(txtInput) > MAX_INPUT_LEN Then
        With TT
            .Style = TTBalloon
            .Icon = TTIconError
            .Title = "Error"
            .TipText = "Input is too long" & vbCrLf & "The input len must be under " & MAX_INPUT_LEN & " chars"
            .Create OKButton.hwnd
        End With
    Else
        OK = True
    End If
End Sub

Private Sub txtInput_Change()
    TT.Destroy
End Sub
