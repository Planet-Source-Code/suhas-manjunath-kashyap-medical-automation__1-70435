VERSION 5.00
Begin VB.Form FrmDischarge 
   BackColor       =   &H0080FFFF&
   Caption         =   "Discgarge Form"
   ClientHeight    =   8490
   ClientLeft      =   900
   ClientTop       =   1095
   ClientWidth     =   11880
   Icon            =   "FrmDischarge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   7455
      Left            =   0
      TabIndex        =   42
      Top             =   1560
      Width           =   1815
      Begin VB.Image Image7 
         Height          =   675
         Left            =   120
         Picture         =   "FrmDischarge.frx":0442
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Image Image6 
         Height          =   675
         Left            =   120
         Picture         =   "FrmDischarge.frx":0C62
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Image Image5 
         Height          =   675
         Left            =   120
         Picture         =   "FrmDischarge.frx":1482
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Image Image4 
         Height          =   750
         Left            =   360
         Picture         =   "FrmDischarge.frx":1CA2
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   12000
      Begin VB.Image Image3 
         Height          =   1155
         Left            =   10080
         Picture         =   "FrmDischarge.frx":24A9
         Top             =   0
         Width           =   1170
      End
      Begin VB.Image Image2 
         Height          =   765
         Left            =   3600
         Picture         =   "FrmDischarge.frx":3E73
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "DISCHARGE FORM OF R.K.H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   41
         Top             =   480
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   480
         Picture         =   "FrmDischarge.frx":4237
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.Data DataAdmn 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\HOSPITAL AUTOMATION\hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Patient"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00FFC0FF&
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
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton CmdDisc 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Discharge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton CmdSearch 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Patient Details:"
      Height          =   4815
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_Address"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   43
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox TxtName 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_Name"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox TxtCity 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_City"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   1354
         Width           =   2415
      End
      Begin VB.TextBox TxtState 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_State"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   1851
         Width           =   2415
      End
      Begin VB.TextBox TxtCountry 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pat_Country"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2348
         Width           =   2415
      End
      Begin VB.TextBox TxtPin 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Zip"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2845
         Width           =   2415
      End
      Begin VB.TextBox TxtConNo 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Con_Num"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   3342
         Width           =   2415
      End
      Begin VB.TextBox TxtConPerson 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Con_Per"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Patient Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Address :"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   977
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "City :"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   1474
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "State :"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   1971
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Country :"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   2468
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Pin Code :"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2965
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contact No :"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   3462
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contact Person :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Hospital Details :"
      Height          =   3615
      Left            =   6960
      TabIndex        =   1
      Top             =   3480
      Width           =   4695
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Bed"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   39
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox TxtDOD 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DOD"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox TxtStatus 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Status"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TxtWard 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Ward"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtBed 
         DataField       =   "Bed"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtDocDuty 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Doctor"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TxtDOA 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DOA"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FFFF&
         Caption         =   "Date Of Discharge :"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FFFF&
         Caption         =   "Status :"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ward :"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FFFF&
         Caption         =   "Bed No :"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "Doctor On Duty :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FFFF&
         Caption         =   "Date Of Admission :"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtDate 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   30
      X1              =   0
      X2              =   12000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Private Limited"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8040
      TabIndex        =   34
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   11280
      TabIndex        =   33
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "AMKRISHNA HOSPITA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   7320
      TabIndex        =   32
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   6840
      TabIndex        =   31
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Date:"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "FrmDischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdDisc_Click()
TxtStatus.Text = "DISCHARGED"
End Sub

Private Sub CmdExit_Click()
Me.Hide
Unload FrmDischarge
frmmain.Show
'frmmain.StatusBar.SimpleText = "Click Menu Item to Begin..."
End Sub

Private Sub CmdNext_Click()
Dim intResponse As Integer
TxtStatus.Refresh
DataAdmn.Recordset.MoveNext
If DataAdmn.Recordset.EOF Then
intResponse = MsgBox("Already at the end of record", vbExclamation, "Ramkrishna Hospital Pvt Ltd.")
DataAdmn.Recordset.MoveFirst
End If
End Sub

Private Sub CmdSearch_Click()
DataAdmn.Recordset.MoveFirst
Const conBtns = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
Const conMsg = "No matching records found..."
Dim strFindString As String, intUserResponse As Integer
strFindString = InputBox("Enter Patient Name : ", "Find Record")
If strFindString <> "" Then
DataAdmn.Recordset.FindNext "pat_name = '" & strFindString & "'"
If DataAdmn.Recordset.NoMatch = True Then
intUserResponse = MsgBox(conMsg, conBtns, "Ramkrishna Hospital Pvt Ltd")
End If
End If
End Sub

Private Sub Form_Paint()
TxtStatus.Text = "ADMITTED"
End Sub

Private Sub Image4_Click()
Unload FrmDischarge
frmentry.Show
End Sub

Private Sub Label1_Click()

End Sub

Private Sub TxtAddress_Change()

End Sub

Private Sub TxtDate_GotFocus()
TxtDate.Text = Date
End Sub


Private Sub TxtDOD_GotFocus()
TxtDOD.Text = Date
End Sub

Private Sub TxtDOD_LostFocus()
If TxtDOA.Text > TxtDOD.Text Then
MsgBox "Invalid Date"
DataAdmn.Recordset.Edit
DataAdmn.Recordset.CancelUpdate
End If
End Sub

Private Sub TxtPatientID_Change()

End Sub
