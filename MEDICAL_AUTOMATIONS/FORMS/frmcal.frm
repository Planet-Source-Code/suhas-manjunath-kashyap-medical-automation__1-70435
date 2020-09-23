VERSION 5.00
Begin VB.Form FrmCal 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "St. Theresa Medical Pvt Ltd."
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3675
   Icon            =   "frmcal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command18 
      BackColor       =   &H0000FF00&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      MaxLength       =   7
      TabIndex        =   19
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   450
      MaxLength       =   7
      TabIndex        =   18
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   450
      TabIndex        =   17
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "1"
      Height          =   495
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2"
      Height          =   495
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "3"
      Height          =   495
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "4"
      Height          =   495
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "5"
      Height          =   495
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "6"
      Height          =   495
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "7"
      Height          =   495
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "8"
      Height          =   495
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      Height          =   495
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   495
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "AC"
      Height          =   495
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FF00&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000FF00&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0000FF00&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0000FF00&
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2490
      Top             =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "St. Theresa Medical Pvt Ltd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   -30
      TabIndex        =   20
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "FrmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public choice As String

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command1.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command1.Caption
End If
End Sub

Private Sub Command10_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command10.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command10.Caption
End If
End Sub

Private Sub Command11_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command11.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command11.Caption
End If
End Sub

Private Sub Command12_Click()
Text1.Visible = True
Text2.Visible = False
Text3.Visible = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'resets everything
End Sub

Private Sub Command13_Click()
Text1.Visible = False
Text2.Visible = True
'add button
choice = "plus"
End Sub


Private Sub Command14_Click()
Text1.Visible = False
'minus button
Text2.Visible = True
choice = "minus"
End Sub

Private Sub Command15_Click()
Text1.Visible = False
'multiply button
Text2.Visible = True
choice = "multiply"
End Sub

Private Sub Command16_Click()
Text1.Visible = False
'divide button
Text2.Visible = True
choice = "divide"
End Sub

Private Sub Command17_Click()
On Error Resume Next
Text1.Visible = False
Text2.Visible = False
Text3.Visible = True
Text3.Enabled = True
If choice = "plus" Then
Text3 = Val(Text1) + Val(Text2)
End If
If choice = "minus" Then
Text3 = Text1 - Text2
End If
If choice = "multiply" Then
Text3 = Text1 * Text2
End If
If choice = "divide" Then
Text3 = Text1 / Text2
End If
End Sub


Private Sub Command18_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command2.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command2.Caption
End If
End Sub


Private Sub Command3_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command3.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command3.Caption
End If
End Sub


Private Sub Command4_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command4.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command4.Caption
End If
End Sub


Private Sub Command5_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command5.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command5.Caption
End If
End Sub


Private Sub Command6_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command6.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command6.Caption
End If
End Sub


Private Sub Command7_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command7.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command7.Caption
End If
End Sub


Private Sub Command8_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command8.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command8.Caption
End If
End Sub


Private Sub Command9_Click()
If Text1.Visible = True Then
Text1.Text = Text1.Text & Command9.Caption
End If
If Text2.Visible = True Then
Text2.Text = Text2.Text & Command9.Caption
End If
End Sub


Private Sub end_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text2.Visible = False
Text3.Visible = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
End Sub



Private Sub Timer1_Timer()
If Me.Width > 4665 Then
Me.Width = 4665
End If
If Me.Height > 5415 Then
Me.Height = 5415
End If
End Sub
