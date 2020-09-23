VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   6510
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   10000
      Left            =   0
      ScaleHeight     =   10005
      ScaleWidth      =   12000
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Login :"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   3480
      TabIndex        =   1
      Top             =   2520
      Width           =   5535
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00808080&
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H00808080&
         Caption         =   "O.K"
         Default         =   -1  'True
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "   PASSWORD:"
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
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnAdmin 
      Caption         =   "&Admission"
      Enabled         =   0   'False
      Begin VB.Menu AdmPatient 
         Caption         =   "Patient &Admission"
      End
      Begin VB.Menu AdmDisc 
         Caption         =   "Patient &Discharge"
      End
      Begin VB.Menu AdmOPD 
         Caption         =   "&O.P.D Patient"
      End
      Begin VB.Menu AdmSep 
         Caption         =   "-"
      End
      Begin VB.Menu AdmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnEnquiry 
      Caption         =   "&Enquiry"
      Enabled         =   0   'False
      Begin VB.Menu EnqPatient 
         Caption         =   "&Patient Enquiry"
      End
   End
   Begin VB.Menu MnReport 
      Caption         =   "&Report"
      Enabled         =   0   'False
      Begin VB.Menu RepAdmin 
         Caption         =   "&Admission Patient Report"
      End
      Begin VB.Menu RepBilling 
         Caption         =   "&Billing Report"
      End
      Begin VB.Menu RepDischarge 
         Caption         =   "&Discharge Patient Report"
      End
      Begin VB.Menu RepOPDPat 
         Caption         =   "&O.P.D Patient Report"
      End
   End
   Begin VB.Menu MnBill 
      Caption         =   "&Billing"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AdmDisc_Click()
StatusBar.SimpleText = "Wait..."
FrmDischarge.Show
StatusBar.SimpleText = "Showing Discharge Form..."
End Sub

Private Sub AdmExit_Click()
StatusBar.SimpleText = "Unloading Forms And Databases..."
End
End Sub

Private Sub AdmMod_Click()
StatusBar.SimpleText = "Wait..."
FrmModification.Show
StatusBar.SimpleText = "Showing Modification Form..."
End Sub

Private Sub AdmOPD_Click()
StatusBar.SimpleText = "Wait..."
FrmOPD.Show
StatusBar.SimpleText = "Showing O.P.D Form..."
End Sub

Private Sub AdmPatient_Click()
StatusBar.SimpleText = "Loading..."
FrmAdmission.Show
StatusBar.SimpleText = "Showing Admission Form..."
End Sub

Private Sub AdmPrinter_Click()
StatusBar.SimpleText = "Wait..."
CommDlg.ShowPrinter
StatusBar.SimpleText = "Showing Printer Properties..."
End Sub


Private Sub EnqBed_Click()
StatusBar.SimpleText = "Wait..."
FrmEnqBed.Show
StatusBar.SimpleText = "Showing Bed Enquiry Form..."
End Sub


Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    Dim LoginSucceeded As Boolean

    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = "college" Then
        MnReport.Enabled = True
        MnEnquiry.Enabled = True
        MnAdmin.Enabled = True
        LoginSucceeded = True
        MnBill.Enabled = True
        Frame1.Visible = False
        Picture1.Visible = True
        StatusBar.SimpleText = "Click Menu Item To Begin..."
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub EnqPatient_Click()
StatusBar.SimpleText = "Wait..."
FrmPatientEnq.Show
StatusBar.SimpleText = "Showing Patient Enquiry Form..."
End Sub


Private Sub MnBill_Click()
FrmBill.Show
End Sub


Private Sub RepAdmin_Click()
StatusBar.SimpleText = "Wait..."
RepAdmn.Show
StatusBar.SimpleText = "Showing Admission Patient Report..."
End Sub

Private Sub RepBilling_Click()
StatusBar.SimpleText = "Wait..."
RepBill.Show
StatusBar.SimpleText = "Showing Admission Patient Report..."
End Sub

Private Sub RepDischarge_Click()
StatusBar.SimpleText = "Wait..."
RepDisc.Show
StatusBar.SimpleText = "Showing Discharge Patient Report..."
End Sub

Private Sub RepOPDPat_Click()
StatusBar.SimpleText = "Wait..."
RepOPD.Show
StatusBar.SimpleText = "Showing O.P.D Report..."
End Sub

