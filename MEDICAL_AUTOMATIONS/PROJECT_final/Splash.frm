VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11070
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1890
      Top             =   750
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4170
      Top             =   840
   End
   Begin MSComctlLib.ProgressBar Bar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   9225
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblbar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading . . ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   2280
      Width           =   2925
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Suhas Manjunath"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   0
      Left            =   630
      TabIndex        =   3
      Top             =   1770
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   1230
      Width           =   1830
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SRI GOKULA COLLEGE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   5385
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
frmLogin.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
Static A As Integer
Static b As Integer
Static c As Integer
Static d As Integer
Static e As Integer
Static f As Integer

A = A + 1
b = b + 1
c = c + 1
d = d + 1
e = e + 1
f = f + 1
Label1.Caption = Mid("HOSPITAL AUTOMATION", 1, e)
If e = 22 Then e = 22
Label3(0).Caption = Mid("Suhas Manjunath ", 1, f)

If f = 18 Then f = 18


Bar.Value = Bar.Value + 2

Screen.MousePointer = vbHourglass
If Bar.Value = 20 Then
lblbar.Caption = "Loading . . ."
ElseIf Bar.Value = 60 Then
lblbar.Caption = "Please wait . . ."
ElseIf Bar.Value = 100 Then

If Bar.Value = 100 Then

If Timer1.Interval >= 1 Then
Unload Splash
'frmLogin.Show
frmLogin.Show
Screen.MousePointer = vbDefault
End If
End If
End If
End Sub

