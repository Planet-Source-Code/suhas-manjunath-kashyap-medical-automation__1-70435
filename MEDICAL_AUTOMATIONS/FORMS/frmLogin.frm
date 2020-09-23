VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN FORM FOR St. TM"
   ClientHeight    =   9690
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5725.171
   ScaleMode       =   0  'User
   ScaleWidth      =   13225.05
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1380
   End
   Begin VB.TextBox txtPassword 
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3600
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   0
      Top             =   3720
      Width           =   1560
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = "srs" Or txtPassword = "SRS" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Me.Hide
        frmmain.Show
        Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword = ""
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

