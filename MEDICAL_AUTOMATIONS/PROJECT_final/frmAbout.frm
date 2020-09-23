VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10200
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   13860
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7040.221
   ScaleMode       =   0  'User
   ScaleWidth      =   13015.26
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   0
      Left            =   5160
      ScaleHeight     =   4905
      ScaleWidth      =   6105
      TabIndex        =   2
      Top             =   5160
      Width           =   6135
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
         Index           =   1
         Left            =   0
         Picture         =   "frmAbout.frx":0000
         ScaleHeight     =   8595
         ScaleWidth      =   6915
         TabIndex        =   3
         Top             =   0
         Width           =   6975
         Begin VB.PictureBox Picture2 
            Height          =   495
            Left            =   4800
            Picture         =   "frmAbout.frx":D870
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   9
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "St. Theresa Medical"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Left            =   1080
            TabIndex        =   6
            Top             =   120
            Width           =   4695
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
            Index           =   2
            Left            =   2280
            TabIndex        =   5
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "RajajiNagar Bangalore"
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
            TabIndex        =   4
            Top             =   7320
            Width           =   2535
         End
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Suhas Manjunath"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmAbout.frx":E08B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HOSPITAL AUTOMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   6960
      TabIndex        =   1
      Top             =   1200
      Width           =   3480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   2445
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Picture2_Click()
Splash.Show
Me.Hide
End Sub
