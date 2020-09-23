VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAIN FORM R.K.H"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image14 
      Height          =   570
      Left            =   3120
      Picture         =   "Form2.frx":0000
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Image Image13 
      Height          =   1800
      Left            =   9360
      Picture         =   "Form2.frx":056F
      Top             =   240
      Width           =   1800
   End
   Begin VB.Image Image12 
      Height          =   1800
      Left            =   840
      Picture         =   "Form2.frx":0B20
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MAIN FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image Image11 
      Height          =   1425
      Left            =   3840
      Picture         =   "Form2.frx":10D1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Image Image7 
      Height          =   345
      Left            =   7320
      Picture         =   "Form2.frx":52E5
      Top             =   4920
      Width           =   345
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   4680
      Picture         =   "Form2.frx":5500
      Top             =   6120
      Width           =   345
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT DISCHARGED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "O.P.D PATIENT FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " PATIENT REPORT DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   4920
      Width           =   3975
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   1680
      Picture         =   "Form2.frx":571B
      Top             =   4920
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   7320
      Picture         =   "Form2.frx":5936
      Top             =   2760
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   7320
      Picture         =   "Form2.frx":5B51
      Top             =   3840
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   1680
      Picture         =   "Form2.frx":5D6C
      Top             =   3840
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   1680
      Picture         =   "Form2.frx":5F87
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT R.K.H PROJECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   6120
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT ENQUIRY FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING FORM OF R.K.H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADMISSION FORM OF R.K.H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   3375
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label10_Click()
Unload frmmain
FrmOPD.Show
End Sub

Private Sub Label11_Click()
Unload frmmain
FrmDischarge.Show
End Sub

Private Sub Label2_Click()
Unload frmmain
FrmAdmission.Show
End Sub

Private Sub Label3_Click()
Unload frmmain
FrmBill.Show
End Sub

Private Sub Label4_Click()
Unload frmmain
FrmPatientEnq.Show
End Sub
Private Sub Label6_Click()
End
End Sub

Private Sub Label7_Click()
Unload frmmain
frmpatreport.Show
End Sub
