VERSION 5.00
Begin VB.Form frmentry 
   Caption         =   "ENTRY FORM FOR HOSPITAL"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   2520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   8595
      ScaleWidth      =   6915
      TabIndex        =   2
      Top             =   0
      Width           =   6975
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DESHAPRIYO PARK   KOLKATA-700045"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   7320
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PRIVATE LIMITED"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   5640
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AMKRISHNA HOSPITA"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   9480
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.Image Image1 
         Height          =   465
         Left            =   120
         Picture         =   "Form1.frx":D870
         Top             =   7680
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
frmLogin.Show
End Sub
