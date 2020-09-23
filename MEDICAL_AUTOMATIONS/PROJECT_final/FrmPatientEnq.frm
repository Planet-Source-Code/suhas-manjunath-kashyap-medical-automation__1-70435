VERSION 5.00
Begin VB.Form FrmPatientEnq 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Patient Enquiry"
   ClientHeight    =   9480
   ClientLeft      =   165
   ClientTop       =   -1560
   ClientWidth     =   11820
   Icon            =   "FrmPatientEnq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5895
      Left            =   0
      TabIndex        =   33
      Top             =   1680
      Width           =   1815
      Begin VB.Image Image7 
         Height          =   495
         Left            =   360
         Picture         =   "FrmPatientEnq.frx":0442
         Top             =   5040
         Width           =   1020
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   480
         Picture         =   "FrmPatientEnq.frx":1CE3
         Top             =   3960
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   480
         Picture         =   "FrmPatientEnq.frx":1F99
         Top             =   2760
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   480
         Picture         =   "FrmPatientEnq.frx":227E
         Top             =   1440
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   480
         Picture         =   "FrmPatientEnq.frx":2552
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   12000
      Begin VB.Image Image6 
         Height          =   765
         Left            =   3720
         Picture         =   "FrmPatientEnq.frx":2F44
         Top             =   360
         Width           =   825
      End
      Begin VB.Image Image5 
         Height          =   1170
         Left            =   480
         Picture         =   "FrmPatientEnq.frx":3308
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT ENQUIRY FORM OF St. T.M."
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
         TabIndex        =   35
         Top             =   600
         Width           =   5175
      End
   End
   Begin VB.CommandButton Cmdresetsearch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Reset Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton CmdSerReg 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search By ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton CmdSerName 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search By Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox TxtAdmId 
      DataField       =   "PatientId"
      DataSource      =   "DataAdmn"
      Height          =   375
      Left            =   2880
      TabIndex        =   26
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Patient Details:"
      Height          =   4815
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   4215
      Begin VB.TextBox Text1 
         DataField       =   "Pat_Address"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox TxtName 
         DataField       =   "Pat_Name"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox TxtCity 
         DataField       =   "Pat_City"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   1354
         Width           =   2415
      End
      Begin VB.TextBox TxtState 
         DataField       =   "Pat_State"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   1851
         Width           =   2415
      End
      Begin VB.TextBox TxtCountry 
         DataField       =   "Pat_Country"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2348
         Width           =   2415
      End
      Begin VB.TextBox TxtPin 
         DataField       =   "Zip"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2845
         Width           =   2415
      End
      Begin VB.TextBox TxtConNo 
         DataField       =   "Con_Num"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   3342
         Width           =   2415
      End
      Begin VB.TextBox TxtConPer 
         DataField       =   "Con_Per"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Patient Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Address :"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   977
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "City :"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   1474
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "State :"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   1971
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Country :"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   2468
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pin Code :"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2965
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Contact No :"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   3462
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Contact Person :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Hospital Details :"
      Height          =   2295
      Left            =   6840
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
      Begin VB.TextBox TxtWard 
         Appearance      =   0  'Flat
         DataField       =   "Ward"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtBedNo 
         DataField       =   "Bed"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox TxtDocDuty 
         DataField       =   "Doctor"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TxtAdmn 
         DataField       =   "DOA"
         DataSource      =   "DataAdmn"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ward :"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Bed No :"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Doctor On Duty :"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date Of Admission :"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtDate 
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Data DataAdmn 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\6TH SEM B.C.A PROJECTS\V.B PROJECT\MEDICAL_AUTOMATIONS\DATABASE\HOSPITAL.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Patient"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   16
      X1              =   0
      X2              =   11880
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Patient ID:"
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Date:"
      Height          =   255
      Left            =   8520
      TabIndex        =   27
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "FrmPatientEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdNext_Click()
DataAdmn.Recordset.MoveFirst
End Sub

Private Sub Cmdresetsearch_Click()
DataAdmn.Recordset.MoveFirst
End Sub

Private Sub CmdSerName_Click()
'DataAdmn.Recordset.MoveFirst
Const conBtns = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
Const conMsg = "No matching records found..."
Dim strFindString As String, intUserResponse As Integer
DataAdmn.Recordset.MoveFirst
strFindString = InputBox("Enter Patient Name : ", "Find Record")
While Not DataAdmn.Recordset.EOF
    If DataAdmn.Recordset.Fields(1).Value = strFindString Then
        Exit Sub
    End If
    DataAdmn.Recordset.MoveNext
Wend
intUserResponse = MsgBox(conMsg, conBtns, "Ramkrishna Hospital Pvt Ltd")
DataAdmn.Recordset.MoveFirst
End Sub


Private Sub CmdSerReg_Click()
DataAdmn.Recordset.MoveFirst
Const conBtns = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
Const conMsg = "No matching records found..."
Dim strFindString As String, intUserResponse As Integer
strFindString = InputBox("Enter Patient Registration Number : ", "Find Record")
DataAdmn.Recordset.MoveFirst
While Not DataAdmn.Recordset.EOF
    If DataAdmn.Recordset.Fields(0).Value = strFindString Then
        Exit Sub
    End If
    DataAdmn.Recordset.MoveNext
Wend
intUserResponse = MsgBox(conMsg, conBtns, "Ramkrishna Hospital Pvt Ltd")
DataAdmn.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
TxtDate.Text = Date
End Sub

Private Sub Image1_Click()
Unload FrmPatientEnq
frmentry.Show
End Sub

Private Sub Image7_Click()
Unload FrmPatientEnq
frmmain.Show
End Sub
