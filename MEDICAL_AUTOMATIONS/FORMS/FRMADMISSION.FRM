VERSION 5.00
Begin VB.Form FrmAdmission 
   BackColor       =   &H0080FFFF&
   Caption         =   "Admission Form"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   -1680
   ClientWidth     =   11880
   Icon            =   "FrmAdmission.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   6975
      Left            =   0
      TabIndex        =   38
      Top             =   1560
      Width           =   2055
      Begin VB.Image Image8 
         Height          =   495
         Left            =   360
         Picture         =   "FrmAdmission.frx":0442
         Top             =   5400
         Width           =   1020
      End
      Begin VB.Image Image7 
         Height          =   675
         Left            =   120
         Picture         =   "FrmAdmission.frx":1ADE
         Top             =   4200
         Width           =   1500
      End
      Begin VB.Image Image6 
         Height          =   675
         Left            =   120
         Picture         =   "FrmAdmission.frx":22FE
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image Image5 
         Height          =   675
         Left            =   120
         Picture         =   "FrmAdmission.frx":2B1E
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Image Image4 
         Height          =   750
         Left            =   360
         Picture         =   "FrmAdmission.frx":333E
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1215
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   12000
      Begin VB.Image Image3 
         Height          =   1155
         Left            =   10440
         Picture         =   "FrmAdmission.frx":3B45
         Top             =   0
         Width           =   1170
      End
      Begin VB.Image Image2 
         Height          =   765
         Left            =   3480
         Picture         =   "FrmAdmission.frx":550F
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ADMISSION FORM OF R.K.H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   40
         Top             =   480
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   600
         Picture         =   "FrmAdmission.frx":58D3
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.Data DataAdmn 
      BackColor       =   &H008080FF&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\HOSPITAL AUTOMATION\hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Patient"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtDate 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton CmdAddNew 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Hospital Details :"
      Height          =   2295
      Left            =   6960
      TabIndex        =   22
      Top             =   4800
      Width           =   4455
      Begin VB.TextBox TxtAdmn 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DOA"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TxtDocDuty 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Doctor"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TxtBedNo 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Bed"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox TxtWard 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Ward"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FFFF&
         Caption         =   "Date Of Admission :"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "Doctor On Duty :"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FFFF&
         Caption         =   "Bed No :"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ward :"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ward :"
      Height          =   2415
      Left            =   6960
      TabIndex        =   17
      Top             =   2280
      Width           =   4455
      Begin VB.OptionButton OptCabin 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cabin"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton OptMaternity 
         BackColor       =   &H0080FFFF&
         Caption         =   "Maternity Ward"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   1360
         Width           =   1575
      End
      Begin VB.OptionButton OptFemale 
         BackColor       =   &H0080FFFF&
         Caption         =   "Female Medical"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   920
         Width           =   1575
      End
      Begin VB.OptionButton OptMale 
         BackColor       =   &H0080FFFF&
         Caption         =   "Male Medical"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Patient Details:"
      Height          =   4815
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox TXTADD 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_Address"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   39
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox TxtConPer 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Con_Per"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox TxtConNo 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Con_Num"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   3342
         Width           =   2415
      End
      Begin VB.TextBox TxtPin 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Zip"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2845
         Width           =   2415
      End
      Begin VB.TextBox TxtCountry 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_Country"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2348
         Width           =   2415
      End
      Begin VB.TextBox TxtState 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_State"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   1851
         Width           =   2415
      End
      Begin VB.TextBox TxtCity 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_City"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1354
         Width           =   2415
      End
      Begin VB.TextBox TxtName 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_Name"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contact Person :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contact No :"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   3462
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Pin Code :"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   2965
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Country :"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   2468
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "State :"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1971
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "City :"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   1474
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Address :"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   977
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Patient Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   30
      X1              =   0
      X2              =   12000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Date:"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "FrmAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAddNew_Click()
CmdSave.Enabled = True
CmdAddNew.Enabled = True
Dim Serial As String
DataAdmn.Recordset.MoveLast
'Serial = (DataAdmn.Recordset.AbsolutePosition + 2)
DataAdmn.Recordset.AddNew
'TxtAdmId.Text = Serial
''TxtAdmId.TabStop = True
'TxtAdmId.SetFocus
TxtDate.Refresh
End Sub

Private Sub cmdCancel_Click()
DataAdmn.Recordset.Edit
DataAdmn.Recordset.CancelUpdate
End Sub

Private Sub CmdExit_Click()
Me.Hide
Unload FrmAdmission
frmmain.Show
DataAdmn.Recordset.Edit
'frmmain.StatusBar.SimpleText = "Click Menu Item to Begin..."
End Sub

Private Sub CmdNext_Click()
CmdSave.Enabled = False
Dim intResponse As Integer
DataAdmn.Recordset.Edit
DataAdmn.Recordset.MoveNext
If DataAdmn.Recordset.EOF Then
intResponse = MsgBox("Already at the end of the record", vbExclamation, "Ramkrishna Hospital Private Limited.")
DataAdmn.Recordset.MoveLast
End If
End Sub

Private Sub CmdPrevious_Click()
CmdSave.Enabled = False
Dim intResponse As Integer
DataAdmn.Recordset.Edit
DataAdmn.Recordset.MovePrevious
If DataAdmn.Recordset.BOF Then
intResponse = MsgBox("Already at the beginning of the record", vbExclamation, "Ramkrishna Hospital Private Limited.")
DataAdmn.Recordset.MoveFirst
End If
End Sub

Private Sub CmdSave_Click()
DataAdmn.Recordset.Edit
DataAdmn.Recordset.Update
DataAdmn.UpdateRecord
CmdSave.Enabled = False
End Sub

Private Sub Form_Load()
TxtDate.Text = Date
FrmAdmission.Height = 10500
End Sub

Private Sub Form_Paint()
CmdSave.Enabled = False
End Sub

Private Sub Image4_Click()
Unload FrmAdmission
frmentry.Show
End Sub

Private Sub Image8_Click()
Unload FrmAdmission
frmmain.Show
End Sub

Private Sub OptCabin_Click(Index As Integer)
TxtWard.Text = "Cabin"
End Sub

Private Sub OptFemale_Click(Index As Integer)
TxtWard.Text = "Female Medical"
End Sub

Private Sub OptMale_Click(Index As Integer)
TxtWard.Text = "Male Medical"
End Sub

Private Sub OptMaternity_Click(Index As Integer)
TxtWard.Text = "Maternity Ward"
End Sub

Private Sub TxtAddress_Change()

End Sub

Private Sub TxtName_Change()
CmdSave.Enabled = True
End Sub
