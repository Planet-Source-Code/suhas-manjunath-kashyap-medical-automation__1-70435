VERSION 5.00
Begin VB.Form FrmOPD 
   BackColor       =   &H00C0FFC0&
   Caption         =   "O.P.D Patients"
   ClientHeight    =   11010
   ClientLeft      =   2010
   ClientTop       =   -1680
   ClientWidth     =   14010
   Icon            =   "FrmOPD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   14010
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6015
      Left            =   0
      TabIndex        =   25
      Top             =   1440
      Width           =   1935
      Begin VB.Image Image7 
         Height          =   675
         Left            =   120
         Picture         =   "FrmOPD.frx":0442
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Image Image6 
         Height          =   675
         Left            =   120
         Picture         =   "FrmOPD.frx":0C62
         Top             =   4800
         Width           =   1500
      End
      Begin VB.Image Image5 
         Height          =   675
         Left            =   120
         Picture         =   "FrmOPD.frx":1482
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Image Image4 
         Height          =   750
         Left            =   360
         Picture         =   "FrmOPD.frx":1CA2
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11760
      Begin VB.Image Image3 
         Height          =   480
         Left            =   3600
         Picture         =   "FrmOPD.frx":24A9
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   10440
         Picture         =   "FrmOPD.frx":28EB
         Top             =   120
         Width           =   930
      End
      Begin VB.Image Image2 
         Height          =   1125
         Left            =   720
         Picture         =   "FrmOPD.frx":3810
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "O.P.D PATIENT FORM OF St. T. M."
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
         Left            =   4080
         TabIndex        =   26
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.Data DataOPD 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\6TH SEM B.C.A PROJECTS\V.B PROJECT\MEDICAL_AUTOMATIONS\DATABASE\HOSPITAL.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "opd"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0FFC0&
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
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H00C0FFC0&
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
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00C0FFC0&
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
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CmdAddNew 
      BackColor       =   &H00C0FFC0&
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
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0FFC0&
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
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox TxtDate 
      DataField       =   "Date"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox TxtTotal 
      DataField       =   "total"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox TxtConsult 
      DataField       =   "ConsFee"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox TxtDrug 
      DataField       =   "DrugChar"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox TxtLab 
      DataField       =   "LabChar"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox TxtReg 
      DataField       =   "Reg_Char"
      DataSource      =   "DataOPD"
      Height          =   405
      Left            =   3960
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox TxtDoc 
      DataField       =   "DocAttending"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox TxtName 
      DataField       =   "Pt_Name"
      DataSource      =   "DataOPD"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   16
      X1              =   -360
      X2              =   11640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "St. Theresa Medical"
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
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   2400
      Width           =   4095
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
      TabIndex        =   22
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Grand Total :"
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Consultation Charges :"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Drug Charge :"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Lab Charge :"
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Registration Charges :"
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Doctor Attending :"
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Patient Name :"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Date :"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   495
   End
End
Attribute VB_Name = "FrmOPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAddNew_Click()
cmdCancel.Enabled = True
CmdSave.Enabled = True
Dim Serial As String
DataOPD.Recordset.MoveLast
Serial = (DataOPD.Recordset.AbsolutePosition + 2)
DataOPD.Recordset.AddNew
TxtDate.Refresh
End Sub
Private Sub cmdCancel_Click()
DataOPD.Recordset.CancelUpdate
TxtTotal.SetFocus
CmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub
Private Sub CmdExit_Click()
Me.Hide
Unload FrmOPD
frmmain.Show
End Sub
Private Sub CmdNext_Click()
DataOPD.Recordset.MoveNext
If DataOPD.Recordset.EOF <> DataOPD.Recordset.BOF Then
    DataOPD.Recordset.MoveLast
    If DataOPD.Recordset.EOF Then
    MsgBox "Already at the end of record !", vbExclamation, "Ramkrishna Hospital Pvt Ltd"
    DataOPD.Recordset.MoveLast
    End If
MsgBox "Last Record Displayed !", vbExclamation, "Ramkrishna Hospital Pvt Ltd"
DataOPD.Recordset.MoveLast
CmdSave.Enabled = False
End If
TxtTotal.SetFocus
CmdSave.Enabled = False
End Sub
Private Sub CmdPrevious_Click()
DataOPD.Recordset.MovePrevious
TxtTotal.SetFocus
CmdSave.Enabled = False
If DataOPD.Recordset.EOF <> DataOPD.Recordset.BOF Then
    DataOPD.Recordset.MoveFirst
    If DataOPD.Recordset.EOF Then
    'MsgBox "Already at the end of the record !", vbExclamation, "Ramkrishna Hospital Pvt Ltd"
    DataOPD.Recordset.MoveFirst
    End If
MsgBox " Already at the begining of record!", vbExclamation, "Ramkrishna Hospital Pvt Ltd"
CmdSave.Enabled = False
End If
TxtTotal.SetFocus
CmdSave.Enabled = False
End Sub
Private Sub CmdSave_Click()
DataOPD.Recordset.Edit
CmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub
Private Sub Form_Load()
TxtDate.Text = Date
End Sub
Private Sub Form_Paint()
CmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub Image4_Click()
Unload FrmOPD
frmentry.Show
End Sub

Private Sub TxtTotal_GotFocus()
TxtTotal.Text = Val(TxtReg.Text) + Val(TxtDrug.Text) + Val(TxtLab.Text) + Val(TxtConsult.Text)
End Sub
